(() => {
  const state = { files: [], sheets: [], result: null, activeTab: 'orders', templates: {}, busy: false };
  const $ = id => document.getElementById(id);
  const norm = v => String(v ?? '').trim().replace(/\.0$/, '').replace(/\s+/g, ' ').toUpperCase();
  const value = (row, ...keys) => {
    for (const key of keys) { const k = Object.keys(row || {}).find(x => norm(x) === norm(key)); if (k && row[k] !== '' && row[k] != null) return row[k]; }
    return '';
  };
  const cleanDate = v => {
    if (v instanceof Date) return `${v.getFullYear()}-${String(v.getMonth()+1).padStart(2,'0')}-${String(v.getDate()).padStart(2,'0')}`;
    const s = String(v ?? '').trim(); if (!s) return '';
    let m=s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);if(m)return `${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`;
    m=s.match(/^(\d{4})\.(\d{1,2})\.(\d{1,2})$/);if(m)return `${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`;
    m=s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{2,4})$/);if(m){const month=new Date(`${m[2]} 1, 2000`).getMonth()+1;return `${m[3].length===2?'20'+m[3]:m[3]}-${String(month).padStart(2,'0')}-${m[1].padStart(2,'0')}`;}
    m=s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);if(m)return `${m[3]}-${m[1].padStart(2,'0')}-${m[2].padStart(2,'0')}`;
    const d = new Date(s); if (!Number.isNaN(d.valueOf())) return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
    return s;
  };
  window.NBYCleanDate=cleanDate;
  const dateForInput = () => $('asOfDate').value;
  const keyBy = (rows, fn) => { const m = new Map(); rows.forEach(r => { const k = norm(fn(r)); if (!k) return; if (!m.has(k)) m.set(k, []); m.get(k).push(r); }); return m; };
  const kindOf = (rows, fileName='') => {
    const first = rows?.[0] || {}; const keys = Object.keys(first).map(norm); const has = (...k) => k.every(x => keys.includes(norm(x)));
    if (has('Date','Actually shipment','Expect shipment')) return 'Shipping Status Summary';
    if (has('SAP 訂單匹配') || has('核對狀態') || has('Work Order Nbr','SAP SD Document ZRSD')) return '人工成品';
    if (has('Work Order Nbr','Manifest Scan Date','Shipper')) return 'Nike FC';
    if (has('OrderNumber','ConsolidationShipmentId')) return 'AWB Map';
    if (has('Work Order Number','Tracking Number')) return 'SF Baozun';
    if (has('Tracking Number','Ship Date') && (has('Service') || /UPS|Detailed Report/i.test(fileName))) return 'UPS';
    if (has('Customer Reference','SD Document')) return 'SAP ZRSD1019';
    if (has('SD Document','Material') && (has('Unrestricted') || has('Blocked Stock'))) return 'SAP MB52';
    return '';
  };
  function rowsFromSheet(sheet) {
    const grid = XLSX.utils.sheet_to_json(sheet, {header:1, defval:'', raw:false, blankrows:false});
    let headerAt = -1;
    for (let i=0;i<Math.min(grid.length,15);i++) {
      const line = grid[i].map(norm);
      if ((line.includes('WORK ORDER NBR') && line.includes('MANIFEST SCAN DATE')) ||
          (line.includes('CUSTOMER REFERENCE') && line.includes('SD DOCUMENT')) ||
          (line.includes('PLANT') && line.includes('SD DOCUMENT') && line.includes('MATERIAL')) ||
          (line.includes('TRACKING NUMBER') && (line.includes('SHIP DATE') || line.includes('MOTHER WAYBILL'))) ||
          (line.includes('TRACKING NUMBER') && (line.includes('UPS 原始出現次數') || line.includes('紀錄次數') || line.includes('重複筆數'))) ||
          (line.includes('WORK ORDER NUMBER') && line.includes('TRACKING NUMBER')) ||
          (line.includes('DATE') && line.includes('ACTUALLY SHIPMENT') && line.includes('EXPECT SHIPMENT')) ||
          (line.includes('ORDERNUMBER') && line.includes('CONSOLIDATIONSHIPMENTID')) || line.includes('SAP 訂單匹配')) { headerAt=i; break; }
    }
    if (headerAt < 0) return [];
    const seen={}; const headers = grid[headerAt].map((h,i)=>{const k=String(h||`欄位${i+1}`).trim();const n=seen[k]||0;seen[k]=n+1;return n?`${k}.${n}`:k;});
    return grid.slice(headerAt+1).filter(r=>r.some(x=>String(x).trim()!=='')).map(r=>Object.fromEntries(headers.map((h,i)=>[h,r[i]??''])));
  }
  function statusDetailRows(sheet) {
    const grid=XLSX.utils.sheet_to_json(sheet,{header:1,defval:'',raw:false,blankrows:false});
    const headerAt=grid.findIndex(row=>row.some(x=>norm(x)==='PRODUCT NAME'));
    if(headerAt<0)return [];
    const headers=grid[headerAt].map((h,i)=>String(h||`Column ${i+1}`).trim());
    return grid.slice(headerAt+1).filter(row=>row.some(x=>String(x).trim()!=='')).map(row=>Object.fromEntries(headers.map((h,i)=>[h,row[i]??''])));
  }

  async function parseFile(file) {
    const data = new Uint8Array(await file.arrayBuffer());
    const wb = XLSX.read(data, {type:'array', cellDates:true});
    const manualBook = /NBY_1120|NBY_Shipping|Source_Analysis|完整源頭|人工成品/i.test(file.name);
    const statusBook = /SHI+PPING\s+STATUS/i.test(file.name);
    const parsed=[];
    for (const name of wb.SheetNames) {
      if(statusBook && /^summary$/i.test(name)){const rows=rowsFromSheet(wb.Sheets[name]);if(rows.length)parsed.push({file:file.name,sheet:name,kind:'Shipping Status Summary',rows});continue;}
      if(statusBook && /^detail shipment$/i.test(name)){const rows=statusDetailRows(wb.Sheets[name]);if(rows.length)parsed.push({file:file.name,sheet:name,kind:'Shipping Status Detail',rows});continue;}
      const rows = rowsFromSheet(wb.Sheets[name]); if (!rows.length) continue;
      const kind=manualBook?'人工成品':kindOf(rows,file.name); if (kind) parsed.push({file:file.name,sheet:name,kind,rows,processed:rows.some(r=>Object.hasOwn(r,'ZRSD1019')||Object.hasOwn(r,'MB52'))});
    }
    return parsed;
  }
  async function ingest(fileList) {
    const incoming=[...fileList]; if (!incoming.length || state.busy) return; invalidate();state.busy=true;$('runButton').disabled=true;
    $('runStatus').textContent='正在讀取檔案…';
    for (const file of incoming) {
      if(state.files.some(f=>f.name===file.name)){state.files=state.files.filter(f=>f.name!==file.name);state.sheets=state.sheets.filter(s=>s.file!==file.name);}
      try { const parsed=await parseFile(file); state.files.push({name:file.name,size:file.size,importedAt:new Date().toISOString(),types:[...new Set(parsed.map(x=>x.kind))],sheetCount:parsed.length}); state.sheets.push(...parsed); }
      catch(e) { state.files.push({name:file.name,size:file.size,importedAt:new Date().toISOString(),types:['讀取失敗'],sheetCount:0,error:e.message}); }
    }
    state.busy=false;renderSources(); renderSourceChoices(); $('runStatus').textContent='匯入完成。請檢查資料類型後執行核對。'; $('runButton').disabled=state.sheets.length===0; $('clearButton').disabled=false;
  }
  function renderSources() {
    const root=$('sourceList'); if (!state.files.length) {root.className='source-list empty';root.textContent='尚未匯入檔案';return;}
    root.className='source-list'; root.innerHTML=state.files.map((f,i)=>`<div class="source-item"><span class="source-name">${escapeHtml(f.name)}</span><span class="source-types">${f.types.map(x=>`<span class="chip ${x==='讀取失敗'?'':'ok'}">${escapeHtml(x)}</span>`).join('')||'<span class="chip">未辨識</span>'}</span><span class="source-count">${f.sheetCount} 個資料表</span><button type="button" class="source-remove" data-remove="${i}">移除</button>${f.error?`<small class="source-error">${escapeHtml(f.error)}</small>`:''}</div>`).join('');
  }
  $('sourceList').addEventListener('click',e=>{const b=e.target.closest('[data-remove]');if(!b||state.busy)return;const f=state.files[Number(b.dataset.remove)];state.files.splice(Number(b.dataset.remove),1);state.sheets=state.sheets.filter(s=>s.file!==f.name);invalidate();renderSources();renderSourceChoices();$('runButton').disabled=!state.sheets.length;});
  function escapeHtml(s){return String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
  async function calculate() {
    const target=dateForInput();
    const all=k=>state.sheets.filter(s=>s.kind===k).flatMap(s=>s.rows);
    const fcCandidates=state.sheets.filter(s=>s.kind==='Nike FC');
    const source=k=>{const ss=state.sheets.filter(s=>s.kind===k);if(ss.length>1)throw Error(`${k} 有多個來源工作表，請移除重複批次後再執行`);return ss.flatMap(s=>s.rows);};
    const sap=source('SAP ZRSD1019'), mb52Raw=source('SAP MB52'), ups=source('UPS'), sf=source('SF Baozun').filter(r=>!(norm(value(r,'Factory'))==='TOTAL'&&!value(r,'Work Order Number')&&!value(r,'Tracking Number')));
    const mb52=NBYEngine.unique(mb52Raw);
    const selected=fcCandidates.find(x=>`${x.file}::${x.sheet}`===$('fcSource').value);
    if(!selected)throw Error('請先選擇本次 FC 來源工作表');
    const fcRows=selected.rows.filter(r=>cleanDate(value(r,'Manifest Scan Date'))===target);
    if(!fcRows.length)throw Error('選定 FC 來源沒有此 Manifest 日期的資料，請調整日期或來源');
    const selectedFc={s:selected,rows:fcRows};
    const manualSheets=state.sheets.filter(s=>s.kind==='人工成品');
    const {orders,upsRows,sfRows,anomalies}=NBYEngine.reconcile({fc:fcRows,sap,stock:mb52,ups,sf,awb:all('AWB Map'),date:target,cleanDate});
    const add=(severity,category,key,detail)=>anomalies.push({嚴重度:severity,異常類型:category,工單或Tracking:key,說明:detail});
    selectedFc.linked=orders.filter(o=>o.sapMatch).length;selectedFc.rate=selectedFc.linked/orders.length;
    for(const kind of ['SAP ZRSD1019','SAP MB52','UPS','SF Baozun','AWB Map'])if(!all(kind).length)add('中','缺少來源',kind,'未提供，相關核對不視為通過');
    for(const f of state.files)if(f.error||!f.sheetCount)add('高','來源檔讀取失敗或未辨識',f.name,f.error||'沒有可辨識資料表');
    for(const f of state.files)add('低','下載批次待確認',f.name,'匯入時間不代表來源下載時間，請確認同批次');
    if(selected.processed)add('高','使用人工加工資料作為 FC',selected.file,'可測試關聯，但不可視為獨立 FC 原始資料 Golden Sample 驗收通過');
    const otherDates=[...new Set(selected.rows.map(r=>cleanDate(value(r,'Manifest Scan Date'))).filter(x=>x&&x!==target))];
    if(otherDates.length)add('中','FC 含其他日期',selected.file,otherDates.join(', ')+'；僅轉出所選 Manifest 日期');
    if(!/^\d{4}-\d{2}-\d{2}$/.test(target))throw Error('請選擇有效日期');
    const manual=validateManual(manualSheets,orders,upsRows,sfRows,mb52);
    manual.lines.filter(x=>x.結果!=='通過').forEach(x=>add('中','人工成品逐筆比對差異',x.工單,x.差異||x.結果));
    state.result={target,orders,upsRows,sfRows,anomalies,manual,counts:{fcRaw:selectedFc?.s.rows.length||0,fcDate:fcRows.length,orders:orders.length,sap:sap.length,mb52:mb52.length,ups:ups.length,sf:sf.length,manualFiles:new Set(manualSheets.map(x=>x.file)).size},selectedFc:selectedFc?{...selectedFc.s,linked:selectedFc.linked,rate:selectedFc.rate,rows:selectedFc.rows}:null};
    await buildOutputs(); renderResults();
  }
  function validateManual(sheets,orders,upsRows,sfRows,mb52){
    const summaries=[]; const lines=[]; const byWO=new Map(orders.map(o=>[o.wo,o]));
    const byTrack=new Map(upsRows.map(x=>[norm(x.Tracking_Number),x])); const sfByWO=new Map(sfRows.map(x=>[norm(x.Work_Order_Number),x]));
    const mbKey=(doc,item,material)=>[norm(doc),norm(item),norm(material)].join('|');
    const mbSet=new Set(mb52.map(r=>mbKey(value(r,'SD Document'),value(r,'Item (SD)','SD Item'),value(r,'Material'))));
    const summary=(s,type,count,ok,details='')=>summaries.push({檔案:s.file,工作表:s.sheet,比對項目:type,人工筆數:count,成功比對:ok,結果:count===ok?'通過':'需檢查',差異摘要:details});
    for(const s of sheets){
      const rows=s.rows;
      const orderRows=rows.filter(r=>value(r,'Work Order Nbr')&&(value(r,'SAP 訂單匹配')||value(r,'SAP SD Document ZRSD')||value(r,'核對狀態')));
      if(orderRows.length){let matched=0,fieldChecks=0,fieldMatches=0;
        for(const r of orderRows){const wo=norm(value(r,'Work Order Nbr'));const o=byWO.get(wo);let diffs=[];let rowChecks=0;
          if(!o)diffs.push('人工成品有／轉檔結果無');else{
            const checks=[['FC Scan Date',o.fcDate,v=>cleanDate(v)],['Shipper',norm(o.shipper),v=>norm(v)],['Tracking Number',o.track,v=>norm(v)],['SAP SD Document ZRSD',norm(o.sapDoc),v=>norm(v)],['SAP SD Document MB52',norm(o.mbDoc),v=>norm(v)],['SAP 訂單匹配',o.sapMatch?'YES':'NO',v=>norm(v)],['MB52 庫存紀錄匹配',o.stockMatch?'YES':'NO',v=>norm(v)],['UPS Tracking 匹配',NBYEngine.carrier(o.shipper)==='UPS'?(o.upsMatch?'YES':'NO'):'N/A',v=>norm(v)],['AWB / Consolidation ID',norm(o.awb),v=>norm(v)],['AWB 狀態',o.awb?'有':'未填',v=>norm(v)],['核對狀態','工單及來源已匹配（非數量核對）',v=>String(v??'').trim()]];
            for(const [col,expected,coerce] of checks){const got=value(r,col);if(got!==''&&got!=null){fieldChecks++;rowChecks++;if(coerce(got)===coerce(expected))fieldMatches++;else diffs.push(`${col}:人工=${got}／系統=${expected}`);}}
          }
          if(!diffs.length&&rowChecks===0)diffs.push('沒有可比對欄位');
          if(!diffs.length)matched++;
          lines.push({人工檔案:s.file,工作表:s.sheet,比對類型:'工單欄位',工單:wo,結果:diffs.length?'有差異':'通過',差異:diffs.join('；'),欄位已比對:rowChecks});
        }
        summary(s,'工單逐筆／欄位',orderRows.length,matched,`欄位相符 ${fieldMatches}/${fieldChecks}`);
      }
      const trackRows=rows.filter(r=>value(r,'Tracking Number')&&(value(r,'UPS 原始出現次數')!==''||value(r,'FC Data Shipping 出現次數')!==''||value(r,'紀錄次數')!==''||value(r,'重複筆數')!==''));
      if(trackRows.length){let matched=0;
        for(const r of trackRows){const t=norm(value(r,'Tracking Number'));const actual=byTrack.get(t);let diffs=[];if(!actual)diffs.push('人工 Tracking 不在 UPS 原始資料');else{
          for(const [col,expected] of [['UPS 原始出現次數',actual.UPS_紀錄數],['紀錄次數',actual.UPS_紀錄數],['重複筆數',actual.UPS_紀錄數],['FC Data Shipping 出現次數',actual.FC_工單數]]){const got=value(r,col);if(got!==''&&got!=null&&norm(got)!==norm(expected))diffs.push(`${col}:人工=${got}／系統=${expected}`);}
          const service=value(r,'UPS Service','Service');if(service&&norm(service)!==norm(actual.UPS_Service))diffs.push(`Service:人工=${service}／系統=${actual.UPS_Service}`);
          const charge=value(r,'Charge');if(charge!==''&&norm(charge)!==norm(actual.UPS_Charge))diffs.push(`Charge:人工=${charge}／系統=${actual.UPS_Charge}`);
          const ship=value(r,'Ship Date');if(ship&&cleanDate(ship)!==actual.Ship_Date)diffs.push(`Ship Date:人工=${ship}／系統=${actual.Ship_Date}`);
        }
          if(!diffs.length)matched++;lines.push({人工檔案:s.file,工作表:s.sheet,比對類型:'UPS Tracking',Tracking_Number:t,結果:diffs.length?'有差異':'通過',差異:diffs.join('；')});
        }
        summary(s,'UPS Tracking 逐筆',trackRows.length,matched);
      }
      const sfManual=rows.filter(r=>value(r,'Work Order Number')&&value(r,'Mother waybill')&&value(r,'Tracking Number'));
      if(sfManual.length){let matched=0;
        for(const r of sfManual){const wo=norm(value(r,'Work Order Number'));const actual=sfByWO.get(wo);let diffs=[];if(!actual)diffs.push('人工 SF 工單不在 manifest');else{
          for(const [col,got,want] of [['Tracking Number',value(r,'Tracking Number'),actual.Tracking_Number],['Ship Group ID',value(r,'Ship Group ID'),actual.Ship_Group_ID],['Mother waybill',value(r,'Mother waybill'),actual.Mother_Waybill],['Qty',value(r,'Qty'),actual.Qty]])if(got!==''&&norm(got)!==norm(want))diffs.push(`${col}:人工=${got}／系統=${want}`);
          const d=value(r,'Estimated Ship Date','Est Ship Date');if(d&&cleanDate(d)!==actual.Estimated_Ship_Date)diffs.push(`Ship Date:人工=${d}／系統=${actual.Estimated_Ship_Date}`);
        }
          if(!diffs.length)matched++;lines.push({人工檔案:s.file,工作表:s.sheet,比對類型:'SF Baozun manifest',工單:wo,結果:diffs.length?'有差異':'通過',差異:diffs.join('；')});
        }
        summary(s,'SF Baozun 逐筆',sfManual.length,matched);
      }
      const stockRows=rows.filter(r=>value(r,'SD Item','Item (SD)')&&value(r,'Material'));
      if(stockRows.length){let matched=0;
        for(const r of stockRows){const exists=mbSet.has(mbKey(value(r,'SD Document'),value(r,'SD Item','Item (SD)'),value(r,'Material')));if(exists)matched++;lines.push({人工檔案:s.file,工作表:s.sheet,比對類型:'MB52 庫存列',SD_Document:value(r,'SD Document'),SD_Item:value(r,'SD Item','Item (SD)'),Material:value(r,'Material'),結果:exists?'通過':'原始來源無此列',差異:exists?'':'文件／項次／物料組合未匹配'});}
        summary(s,'MB52 庫存列逐筆',stockRows.length,matched);
      }
    }
    if(!summaries.length)summaries.push({檔案:'',工作表:'',比對項目:'人工成品',人工筆數:0,成功比對:0,結果:'尚未匯入人工成品'});
    return {summaries,lines};
  }
  function renderResults(){const r=state.result;$('results').classList.remove('hidden');
    const m=[['FC 唯一工單（非鞋雙數）',r.counts.orders,`${r.counts.fcDate} 筆符合日期`],['SAP ZRSD1019',r.counts.sap,'工單關聯來源'],['SAP MB52',r.counts.mb52,'庫存明細來源'],['UPS Tracking',r.counts.ups,`${r.upsRows.length} 個唯一 Tracking`],['待處理異常',r.anomalies.length,'依來源與關聯規則產生']];
    $('metrics').innerHTML=m.map(x=>`<div class="metric"><span>${x[0]}</span><strong>${x[1]}</strong><small>${x[2]}</small></div>`).join('')+`<div class="source-note"><strong>採用 FC 工作表</strong><span>${escapeHtml(r.selectedFc?.file||'未找到')} · ${escapeHtml(r.selectedFc?.sheet||'')} · ${r.counts.orders} 筆與 SAP 關聯 ${r.counts.orders?Math.round(r.selectedFc.linked/r.selectedFc.rows.length*100):0}%</span><small>由操作者選定 FC 來源，依 Manifest 日期篩選；不依匹配率挑選資料。</small></div>`;renderTable();
  }
  const tabData=()=>{const r=state.result;switch(state.activeTab){case'ups':return r.upsRows;case'sf':return r.sfRows;case'anomalies':return r.anomalies;case'manual':return r.manual.lines.length?r.manual.lines:r.manual.summaries.map(({details,...x})=>x);default:return r.orders.map(o=>({'Work Order Nbr':o.wo,'FC Scan Date':o.fcDate,Shipper:o.shipper,'Tracking Number':o.track,'SAP SD Document ZRSD':o.sapDoc,'SAP SD Document MB52':o.mbDoc,'SAP 訂單匹配':o.sapMatch?'YES':'NO','MB52 庫存紀錄匹配':o.stockMatch?'YES':'NO','UPS Tracking 匹配':o.upsMatch?'YES':'','AWB / Consolidation ID':o.awb,'AWB 狀態':o.awb?'有':'未填'}));}};
  function renderTable(){if(!state.result)return;const data=tabData();const q=norm($('tableSearch').value);const filtered=data.filter(row=>!q||Object.values(row).some(v=>norm(v).includes(q)));const cols=[...new Set(filtered.flatMap(Object.keys))].filter(c=>c!=='details');$('tableHead').innerHTML=`<tr>${cols.map(c=>`<th>${escapeHtml(c)}</th>`).join('')}</tr>`;$('tableBody').innerHTML=filtered.slice(0,500).map(row=>`<tr>${cols.map(c=>`<td>${formatCell(row[c])}</td>`).join('')}</tr>`).join('');$('tableCount').textContent=`${filtered.length.toLocaleString()} 筆${filtered.length>500?'（畫面顯示前 500 筆）':''}`;}
  function formatCell(v){if(typeof v==='boolean')return `<span class="state ${v?'pass':'fail'}">${v?'是':'否'}</span>`;const s=String(v??'');if(['通過','已與 FC Tracking 對上','FC Tracking 一致'].includes(s))return `<span class="state pass">${escapeHtml(s)}</span>`;if(/未匹配|不一致|重複|缺少|待檢查/.test(s))return `<span class="state warn">${escapeHtml(s)}</span>`;return escapeHtml(s);}
  const outputNames=['Orders Update.xlsx','Data Shipping.xlsx','Shipping Status.xlsx','UPS Reconciliation.xlsx','SF Baozun Reconciliation.xlsx','Exception Report.xlsx'];
  async function buildOutputs(){
    const r=state.result;if(!r)return {};
    const cleanRows=rows=>(rows||[]).map(row=>Object.fromEntries(Object.entries(row).map(([k,v])=>[k,v instanceof Date?cleanDate(v):v])));
    const outputHeaderMap={'工單或Tracking':'Work Order or Tracking Number','嚴重度':'Severity','異常類型':'Exception Type','說明':'Details','檔案':'File','人工檔案':'Manual File','工作表':'Worksheet','比對類型':'Check Type','工單':'Work Order','結果':'Result','差異':'Difference','欄位已比對':'Fields Compared','人工筆數':'Manual Rows','成功比對':'Matched Rows','比對項目':'Check Item','差異摘要':'Difference Summary','UPS_紀錄數':'UPS Record Count','FC_工單數':'FC Work Order Count','FC_工單':'FC Work Orders','MB52_明細列數':'MB52 Detail Rows','SAP_SD_Document':'SAP SD Document','SAP_Confirmed_Quantity':'SAP Confirmed Quantity','MB52_Unrestricted':'MB52 Unrestricted Stock','MB52_Blocked_Stock':'MB52 Blocked Stock','UPS_Service':'UPS Service','Ship_Date':'Ship Date','UPS_Charge':'UPS Charge','Work_Order_Number':'Work Order Number','Ship_Group_ID':'Ship Group ID','Mother_Waybill':'Mother Waybill','Estimated_Ship_Date':'Estimated Ship Date','SAP_Confirmed_Qty':'SAP Confirmed Quantity','UPS_原始出現次數':'UPS Source Occurrences','紀錄次數':'Record Count','重複筆數':'Duplicate Count','SAP_訂單匹配':'SAP Order Matched','核對狀態':'Reconciliation Status','MB52_庫存紀錄匹配':'MB52 Stock Row Matched','UPS_Tracking_匹配':'UPS Tracking Matched','AWB_狀態':'AWB Status','Qty與SAP一致':'Quantity Matches SAP','SAP_工單存在':'SAP Work Order Found','MB52_紀錄存在':'MB52 Record Found','FC_工單存在':'FC Work Order Found','FC_Tracking一致':'FC Tracking Matches'};
    const outputTextMap={'高':'High','中':'Medium','低':'Low','通過':'PASS','有差異':'DIFFERENCE','人工成品逐筆比對差異':'Manual output row-by-row difference','FC 工單缺少 SAP ZRSD1019':'FC work order missing from SAP ZRSD1019','SAP 訂單缺少 MB52 庫存紀錄':'SAP order missing from MB52 stock snapshot','FC Tracking 未出現在 UPS':'FC Tracking Number missing from UPS source','FC 缺少 Tracking Number':'FC work order has no Tracking Number','未找到 AWB／Consolidation ID':'AWB / Consolidation ID not found','UPS Tracking 重複':'Duplicate UPS Tracking Number','UPS Tracking 未關聯 FC 工單':'UPS Tracking Number not linked to FC work order','SF manifest 工單未匹配 FC':'SF manifest work order not matched to FC','SF Qty 與 SAP 數量不一致':'SF quantity does not match SAP confirmed quantity','SF／FC Tracking 不一致':'SF and FC Tracking Numbers do not match','FC 工單重複':'Duplicate FC work order','FC 缺少 Work Order Nbr':'FC rows missing Work Order Nbr','依 Customer Reference 找不到對應工單':'No work order found by Customer Reference','無法執行 UPS 逐筆核對':'UPS row-level reconciliation could not run','來源未提供對應 AWB 或 Consolidation ID':'No AWB or Consolidation ID is available in the source','未建立':'Not created','未提供':'Not provided','空白':'Blank','有':'Present','未填':'Blank','工單及來源已匹配（非數量核對）':'Work order and source rows matched (not a quantity check)','原始來源無此列':'Source row not found','人工成品有／轉檔結果無':'Present in manual output but missing from generated output','人工 Tracking 不在 UPS 原始資料':'Tracking Number not found in UPS source','一致（Tracking 存在）':'MATCH (Tracking Number found)','UPS 重複紀錄':'Duplicate UPS record; review required','已與 FC Tracking 對上':'Matched to FC Tracking Number','UPS 有紀錄但 FC 無對應':'UPS record has no matching FC row','SF 工單未出現在 FC':'SF Work Order not found in FC','SF Qty 與 SAP 數量不一致':'SF quantity does not match SAP confirmed quantity','Tracking 不一致':'Tracking Number mismatch','FC Tracking 一致':'FC Tracking matches','工單及來源已匹配（非數量核對）':'Work order and source records matched (not a quantity check)','尚未匯入人工成品':'No manual output file imported','未核對 FC / SAP':'FC / SAP not checked','NBY_1120_完整源頭逐筆比對.xlsx':'NBY_1120_Full_Source_Row_Comparison.xlsx','03_工單逐筆對照':'03_Work_Order_Row_Comparison','04_UPS_異常核對':'04_UPS_Exceptions','UPS 出現 2 筆；需確認是否為分件或重複下載':'UPS contains multiple records; confirm split shipment versus duplicate download','來源未提供對應 AWB 或 Consolidation ID':'No AWB or Consolidation ID is available in the source','物流明細存在，但本次日期 FC 工單中無相同 Tracking Number':'UPS record exists but no matching Tracking Number is present in FC for the selected date','工單欄位':'Work Order Fields','人工成品':'Manual Output','工單逐筆／欄位':'Work Order Row and Field Check','UPS Tracking 逐筆':'UPS Tracking Row Check','欄位相符':'Fields matched','NBY_1120_完整源頭逐筆比對.xlsx':'NBY_1120_Full_Source_Row_Comparison.xlsx','03_工單逐筆對照':'03_Work_Order_Row_Comparison','04_UPS_異常核對':'04_UPS_Exceptions'};
    const englishText=v=>{let s=String(v??'');if(outputTextMap[s])return outputTextMap[s];s=s.replace(/^SAP SD Document (.+?) 在 MB52 無紀錄$/,'SAP SD Document $1 has no matching row in MB52').replace(/^FC 工單 (.+?) 的 Tracking Number 未在 UPS 清單找到$/,'FC work order $1: Tracking Number not found in the UPS source').replace(/^UPS 出現 (\d+) 筆；需確認是否為分件或重複下載$/,'UPS contains $1 records; confirm split shipment versus duplicate download').replace(/^欄位相符 (.+)$/,'Fields matched $1').replace(/^原始資料 (\d+) 筆工單號重複$/,'$1 repeated Work Order Nbr records in source').replace(/([^:／；]+):人工=/g,(m,k)=>`${outputHeaderMap[k]||outputTextMap[k]||k}:Manual=`).replace(/／系統=/g,' / Generated=').replace(/人工=/g,'Manual=');return outputTextMap[s]||s;};
    const englishRows=rows=>cleanRows(rows).map(row=>Object.fromEntries(Object.entries(row).map(([k,v])=>[outputHeaderMap[k]||k,englishText(v)])));
    const headersOf=rows=>[...new Set((rows||[]).flatMap(Object.keys))];
    const matchByWo=new Map(r.orders.map(o=>[norm(o.wo),o]));
    const orderSource=r.selectedFc;
    const orderRows=cleanRows(orderSource?.rows||[]);
    const dataRows=cleanRows((r.selectedFc?.rows||[]).map(src=>{
      const wo=norm(value(src,'Work Order Nbr')),o=matchByWo.get(wo)||{};
      const row={...src};
      row.ZRSD1019=o.sapDoc||'';
      row.MB52=o.mbDoc||'';
      row.AWB=o.awb||'';
      return row;
    }));
    const detail=cleanRows(dataRows);
    const statusGroups=new Map();
    for(const row of detail){
      const date=cleanDate(value(row,'Manifest Scan Date')),product=String(value(row,'Product Name')||'Not provided'),shipper=String(value(row,'Shipper')||'Unassigned'),wo=norm(value(row,'Work Order Nbr'));
      const key=[date,product,shipper].join('|');
      if(!statusGroups.has(key))statusGroups.set(key,{date,product,shipper,workOrders:new Set(),sourceRows:0,unkeyedRows:0});
      const group=statusGroups.get(key);group.sourceRows++;
      if(wo)group.workOrders.add(wo);else group.unkeyedRows++;
    }
    const historySummary=state.sheets.filter(s=>s.kind==='Shipping Status Summary').flatMap(s=>s.rows);
    const historyDetail=state.sheets.filter(s=>s.kind==='Shipping Status Detail').flatMap(s=>s.rows);
    const historicalDates=[];const historicalProducts=new Map();
    for(const row of historyDetail){const product=String(value(row,'Product Name')||'').trim();if(!product)continue;if(!historicalProducts.has(product))historicalProducts.set(product,{});for(const [header,count] of Object.entries(row)){if(norm(header)==='PRODUCT NAME'||count===''||count==null)continue;const date=cleanDate(header);if(!date||date===header&&/[^0-9./-]/.test(header))continue;if(!historicalDates.includes(date))historicalDates.push(date);historicalProducts.get(product)[date]=count;}}
    const productDateOrders=new Map(),totalDateOrders=new Map();
    for(const row of detail){const date=cleanDate(value(row,'Manifest Scan Date')),product=String(value(row,'Product Name')||'Not provided'),wo=norm(value(row,'Work Order Nbr'));if(!date||!wo)continue;const pk=`${date}|${product}`;if(!productDateOrders.has(pk))productDateOrders.set(pk,{date,product,workOrders:new Set()});productDateOrders.get(pk).workOrders.add(wo);if(!totalDateOrders.has(date))totalDateOrders.set(date,new Set());totalDateOrders.get(date).add(wo);}
    for(const x of productDateOrders.values()){if(!historicalProducts.has(x.product))historicalProducts.set(x.product,{});historicalProducts.get(x.product)[x.date]=x.workOrders.size;if(!historicalDates.includes(x.date))historicalDates.push(x.date);}
    const statusSummary=historySummary.filter(row=>cleanDate(value(row,'Date'))!==r.target).map(row=>({'Date':cleanDate(value(row,'Date')),'Actually Shipment':value(row,'Actually Shipment'),'Expect Shipment':value(row,'Expect Shipment')}));
    const actualTarget=totalDateOrders.get(r.target)?.size||0;
    const manualStatusTarget=historySummary.find(row=>cleanDate(value(row,'Date'))===r.target);
    if(manualStatusTarget && Number(value(manualStatusTarget,'Actually Shipment'))!==actualTarget)r.anomalies.push({嚴重度:'高',異常類型:'Manual Shipping Status differs from Manifest Date result',工單或Tracking:r.target,說明:`Manual Actually Shipment ${value(manualStatusTarget,'Actually Shipment')}; same-date unique FC Work Orders ${actualTarget}`});
    statusSummary.push({'Date':r.target,'Actually Shipment':actualTarget,'Expect Shipment':'','Actually Shipment Rule':'Distinct Work Order Nbr by Manifest date; conflicts require review','Expect Shipment Rule':'PENDING — not calculated'});
    const statusDetail=[...historicalProducts].map(([product,counts])=>Object.assign({'Product Name':product},...historicalDates.map(date=>({[date]:counts[date]??''}))));
    const upsRaw=cleanRows(state.sheets.filter(s=>s.kind==='UPS').flatMap(s=>s.rows));
    const upsExport=upsRaw.map((row,i)=>{const track=norm(value(row,'Tracking Number')),ordersForTrack=r.orders.filter(o=>norm(o.shipper)==='UPS'&&norm(o.track)===track);return {...row,'FC Work Order Nbr':ordersForTrack.map(o=>o.wo).join(', '),'FC Match Count':ordersForTrack.length,'Duplicate Tracking in UPS':(r.upsRows.find(u=>norm(u.Tracking_Number)===track)?.UPS_紀錄數||0)>1?'REVIEW':'','Source File':state.sheets.find(s=>s.kind==='UPS')?.file||''};});
    const sfExport=englishRows(r.sfRows);
    const exceptions=englishRows(r.anomalies);
    const testSummary=[
      {Check:'Orders Update source records',Rule:'Retain all imported source rows',Actual:orderRows.length,Status:orderRows.length?'IMPORTED':'MISSING',Source:orderSource?.file||''},
      {Check:'Data Shipping rows for selected Manifest Scan Date',Rule:'Filter by Manifest Scan Date',Actual:dataRows.length,Status:dataRows.length?'IMPORTED':'MISSING',Source:r.selectedFc?.file||''},
      {Check:'UPS source events',Rule:'Retain every UPS source event',Actual:upsRaw.length,Status:upsRaw.length?'IMPORTED':'MISSING',Source:state.sheets.find(s=>s.kind==='UPS')?.file||''},
      {Check:'UPS distinct Tracking Numbers',Rule:'Group by Tracking Number for summary only',Actual:new Set(upsRaw.map(x=>norm(value(x,'Tracking Number'))).filter(Boolean)).size,Status:'CALCULATED',Source:'UPS Tracking Number'},
      {Check:'SF manifest detail rows',Rule:'Keep SF Baozun reconciliation separate',Actual:sfExport.length,Status:sfExport.length?'IMPORTED':'MISSING',Source:state.sheets.find(s=>s.kind==='SF Baozun')?.file||''},
      {Check:'Manual output row comparison',Rule:'Compare uploaded manual outputs field by field',Actual:r.manual.lines.length,Status:r.manual.lines.length?'REVIEW':'NO BASELINE',Source:'See Manual Field Differences worksheet'},
      {Check:'Expect Shipment',Rule:'Business rule must be confirmed before calculation',Actual:'Not calculated',Status:'PENDING RULE',Source:'No estimate is generated'},
      {Check:'Shipping Status history',Rule:'Preserve uploaded Summary and Detail shipment history',Actual:historySummary.length?'History loaded':'No history file imported',Status:historySummary.length?'IMPORTED':'MISSING',Source:state.sheets.find(s=>s.kind==='Shipping Status Summary')?.file||''},
      {Check:'Manual Shipping Status Actual vs same-date FC',Rule:'Compare manual total to distinct FC work orders for the selected Manifest Scan Date',Actual:manualStatusTarget?`${value(manualStatusTarget,'Actually Shipment')} vs ${actualTarget}`:'No manual history row',Status:manualStatusTarget&&Number(value(manualStatusTarget,'Actually Shipment'))!==actualTarget?'DIFFERENCE':manualStatusTarget?'MATCH':'NO BASELINE',Source:state.sheets.find(s=>s.kind==='Shipping Status Summary')?.file||''}
    ];
    const files={};
    const make=(name, sheets)=>{const wb=XLSX.utils.book_new();for(const [sheetName,rows] of sheets){const data=rows?.length?rows:[{'Status':'No rows','Note':'No matching source rows were imported'}];const ws=XLSX.utils.json_to_sheet(data,{header:headersOf(data)});ws['!freeze']={xSplit:0,ySplit:1};if(ws['!ref'])ws['!autofilter']={ref:ws['!ref']};XLSX.utils.book_append_sheet(wb,ws,sheetName.slice(0,31));}files[name]=XLSX.write(wb,{bookType:'xlsx',type:'array',cellStyles:true});};
    make(outputNames[0],[['Orders Update',orderRows]]);
    make(outputNames[1],[['Data Shipping',dataRows]]);
    make(outputNames[2],[['Summary',statusSummary],['Detail shipment',statusDetail]]);
    make(outputNames[3],[['UPS Reconciliation',upsExport],['Tracking Summary',englishRows(r.upsRows)]]);
    make(outputNames[4],[['SF Baozun Reconciliation',sfExport]]);
    make(outputNames[5],[['Exceptions',exceptions],['Processing Log',state.files.map(f=>({File:f.name,Bytes:f.size,ImportedAt:f.importedAt,Types:f.types.join(', '),Error:f.error||'',ManifestDate:r.target}))],['Test Result',testSummary],['Manual Field Differences',englishRows(r.manual.lines)],['Manual Comparison Summary',englishRows(r.manual.summaries.map(({details,...x})=>x))]]);
    r.templateStatus=[];
    for(const [type,name,plans] of [
      ['orders',outputNames[0],[{kind:'fc',rows:orderRows}]],
      ['data',outputNames[1],[{kind:'fc',rows:dataRows},{kind:'sap',rows:state.sheets.filter(s=>s.kind==='SAP ZRSD1019').flatMap(s=>s.rows)},{kind:'stock',rows:state.sheets.filter(s=>s.kind==='SAP MB52').flatMap(s=>s.rows)},{kind:'awb',rows:state.sheets.filter(s=>s.kind==='AWB Map').flatMap(s=>s.rows)}]],
      ['status',outputNames[2],[]]
    ]){
      const template=state.templates[type];
      if(template){try{files[name]=await NBYTemplates.fill(template.bytes,plans,type,{date:r.target,actual:actualTarget,products:[...productDateOrders.values()].map(x=>({product:x.product,count:x.workOrders.size}))});r.templateStatus.push(name+'：原樣板格式');}catch(e){delete files[name];r.templateStatus.push(name+'：未輸出，'+e.message);}}
      else {const draft='DRAFT '+name;files[draft]=files[name];delete files[name];r.templateStatus.push(name+'：未載入樣板，僅提供 DRAFT 資料版');}
    }
    make(outputNames[5],[['Exceptions',exceptions],['Test Result',testSummary],['Template Status',r.templateStatus.map(Status=>({Status}))],['Processing Log',state.files.map(f=>({File:f.name,Bytes:f.size,ImportedAt:f.importedAt,Types:f.types.join(', '),ManifestDate:r.target}))],['Manual Field Differences',englishRows(r.manual.lines)],['Manual Comparison Summary',englishRows(r.manual.summaries)]]);
    const audit={version:'2.0.0',processedAt:new Date().toISOString(),manifestDate:r.target,source:state.files.map(f=>({name:f.name,size:f.size,importedAt:f.importedAt,types:f.types})),selectedFc:{file:r.selectedFc.file,sheet:r.selectedFc.sheet},counts:r.counts,exceptions:r.anomalies.length,templates:r.templateStatus,expectShipment:'PENDING',goldenSample:'NOT CERTIFIED'};
    r.audit=audit;try{const history=JSON.parse(localStorage.getItem('nby-audit')||'[]');history.unshift(audit);localStorage.setItem('nby-audit',JSON.stringify(history.slice(0,30)));}catch(e){};
    r.outputFiles=files;r.outputTestSummary=testSummary;renderOutputFiles();return files;
  }
  function downloadBytes(bytes,name){const blob=new Blob([bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),1500);}
  function renderOutputFiles(){const root=$('outputFiles');if(!root)return;root.innerHTML=Object.keys(state.result.outputFiles).map(name=>`<button class="output-file" data-output="${escapeHtml(name)}">↓ ${escapeHtml(name)}</button>`).join('')+`<span class="output-note">${escapeHtml(state.result.templateStatus.join('；'))}。Expect Shipment 未計算；尚未完成正式驗收。</span>`;root.querySelectorAll('[data-output]').forEach(btn=>btn.addEventListener('click',()=>{const bytes=state.result?.outputFiles?.[btn.dataset.output];if(bytes)downloadBytes(bytes,btn.dataset.output);}));}
  async function exportXlsx(){const files=state.result?.outputFiles||await buildOutputs();if(!window.JSZip){$('runStatus').textContent='Excel 已產生；ZIP 函式庫載入失敗，請個別下載。';return;}const zip=new JSZip();for(const [name,bytes] of Object.entries(files))zip.file(name,bytes);zip.file('Processing Audit.json',JSON.stringify(state.result.audit,null,2));const blob=await zip.generateAsync({type:'blob'});const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=`NBY_Shipping_Outputs_${state.result.target}.zip`;a.click();setTimeout(()=>URL.revokeObjectURL(url),1500);}
  function invalidate(){state.result=null;$('results').classList.add('hidden');$('runStatus').textContent='資料或日期已變更，請重新執行';}
  function renderSourceChoices(){const el=$('fcSource'),old=el.value;const fc=state.sheets.filter(s=>s.kind==='Nike FC');el.innerHTML='<option value="">請選擇 FC 工作表</option>'+fc.map(s=>`<option value="${escapeHtml(s.file+'::'+s.sheet)}">${escapeHtml(s.file+' / '+s.sheet)} (${s.rows.length} 列)</option>`).join('');if(fc.some(s=>s.file+'::'+s.sheet===old))el.value=old;else if(fc.length===1)el.value=fc[0].file+'::'+fc[0].sheet;}
  $('asOfDate').addEventListener('change',invalidate);$('fcSource').addEventListener('change',invalidate);
  document.querySelectorAll('[data-template]').forEach(input=>input.addEventListener('change',async()=>{invalidate();const f=input.files[0];if(!f){delete state.templates[input.dataset.template];return;}state.templates[input.dataset.template]={bytes:new Uint8Array(await f.arrayBuffer()),name:f.name};$('runStatus').textContent='已載入樣板：'+f.name;}));
  $('fileInput').addEventListener('change',e=>ingest(e.target.files));
  $('dropZone').addEventListener('dragover',e=>{e.preventDefault();$('dropZone').classList.add('drag');});$('dropZone').addEventListener('dragleave',()=>$('dropZone').classList.remove('drag'));$('dropZone').addEventListener('drop',e=>{e.preventDefault();$('dropZone').classList.remove('drag');ingest(e.dataTransfer.files);});
  $('runButton').addEventListener('click',async ()=>{if(state.busy)return;state.busy=true;$('runButton').disabled=true;invalidate();$('runStatus').textContent='正在匹配來源並套用 Excel 樣板…';try{await new Promise(resolve=>setTimeout(resolve,0));await calculate();$('runStatus').textContent='轉檔完成；請檢視異常與人工成品驗證。';}catch(e){console.error(e);$('runStatus').textContent=`核對失敗：${e.message}`;}finally{state.busy=false;$('runButton').disabled=false;}});
  $('clearButton').addEventListener('click',()=>{state.templates={};document.querySelectorAll('[data-template]').forEach(e=>e.value='');state.files=[];state.sheets=[];state.result=null;$('fileInput').value='';$('results').classList.add('hidden');$('runStatus').textContent='';$('runButton').disabled=true;$('clearButton').disabled=true;renderSources();renderSourceChoices();});
  $('exportButton').addEventListener('click',exportXlsx);$('tableSearch').addEventListener('input',renderTable);
  document.querySelectorAll('.tab').forEach(t=>t.addEventListener('click',()=>{document.querySelectorAll('.tab').forEach(x=>x.classList.remove('active'));t.classList.add('active');state.activeTab=t.dataset.tab;renderTable();}));
  let language='zh';
  const labels={
    zh:{brand:'NBY Shipping Data Converter',sub:'每日原始資料核對與轉檔 · v2',privacy:'檔案僅在此瀏覽器處理',eyebrow:'出貨資料本機轉檔測試',title:'出貨資料轉檔與逐筆核對',lede:'匯入 Nike FC、SAP、UPS、SF Baozun 來源，再與人工成品逐筆比對。原始檔不會上傳到伺服器。',date:'核對日期',upload:'1. 匯入來源與人工成品',hint:'可一次選取多個 .xls、.xlsx、.csv 檔；系統依欄位辨識資料類型與工作表。',choose:'選擇 Excel / CSV 檔',drop:'拖曳檔案到這裡',dropHint:'支援多檔匯入；不會修改來源檔',run:'執行轉檔與核對',clear:'清除本次資料',result:'2. 核對與轉檔結果',orders:'工單關聯',ups:'UPS 獨立核對',sf:'SF Baozun 核對',anomaly:'異常清單',manual:'人工成品驗證',search:'搜尋工單、Tracking、SAP SD Document…',foot:'轉檔輸出與核對結果皆由本機匯入資料即時計算。缺少必要來源時會標示待確認，不會推測補值。',footer:'內部測試工具 · NBY Digitalization · 不連接正式系統'},
    en:{brand:'NBY Shipping Data Converter',sub:'Daily Source Reconciliation · v2',privacy:'Files stay in this browser',eyebrow:'LOCAL SHIPPING DATA TEST',title:'Shipping Data Conversion & Row Checks',lede:'Import Nike FC, SAP, UPS and SF Baozun files, then compare each row with the manual outputs. Source files are processed in this browser.',date:'Check date',upload:'1. Import source and manual files',hint:'Select multiple .xls, .xlsx or .csv files. The app identifies source types and worksheets from their columns.',choose:'Choose Excel / CSV files',drop:'Drop files here',dropHint:'Batch import supported; source files are never modified',run:'Run conversion and checks',clear:'Clear this session',result:'2. Reconciliation & outputs',orders:'Work Orders',ups:'UPS Reconciliation',sf:'SF Baozun Reconciliation',anomaly:'Exceptions',manual:'Manual Output Check',search:'Search work order, tracking or SAP document…',foot:'Files are processed locally. Missing inputs and unconfirmed rules are flagged; values are never guessed.',footer:'Internal test tool · NBY Digitalization · No connection to production systems'}
  };
  function setLanguage(next){language=next;const t=labels[language];const put=(sel,key)=>{const el=document.querySelector(sel);if(el)el.textContent=t[key];};put('.brand strong','brand');put('.brand small','sub');put('#privacyText','privacy');put('.hero .eyebrow','eyebrow');put('.hero h1','title');put('.lede','lede');put('#dateLabel','date');put('.upload-panel h2','upload');put('.panel-head p','hint');put('.panel-head .primary-button','choose');put('#dropZone strong','drop');put('#dropZone div span','dropHint');put('#runButton','run');put('#clearButton','clear');put('.section-title h2','result');put('[data-tab="orders"]','orders');put('[data-tab="ups"]','ups');put('[data-tab="sf"]','sf');put('[data-tab="anomalies"]','anomaly');put('[data-tab="manual"]','manual');const search=$('tableSearch');search.placeholder=t.search;const foot=document.querySelector('.footnote');if(foot)foot.textContent=t.foot;const footer=document.querySelector('footer');if(footer)footer.textContent=t.footer;$('langToggle').textContent=language==='zh'?'EN':'中文';}
  $('langToggle').addEventListener('click',()=>setLanguage(language==='zh'?'en':'zh'));
})();
