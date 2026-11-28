/* Patch worksheet XML inside the original XLSX package. Private samples are never bundled. */
(function(root){
'use strict';
const parse=s=>{const d=new DOMParser().parseFromString(s,'application/xml');if(d.getElementsByTagName('parsererror').length)throw Error('樣板 XML 無法解析');return d;};
const list=(d,n)=>Array.from(d.getElementsByTagNameNS('*',n));
const text=d=>new XMLSerializer().serializeToString(d);
const col=a=>a.replace(/\d+$/,'');
function setCell(d,row,address,v,formula){
 let c=list(row,'c').find(c=>c.getAttribute('r')===address);if(!c){c=d.createElementNS(d.documentElement.namespaceURI,'c');c.setAttribute('r',address);row.appendChild(c);}
 c.removeAttribute('t');while(c.firstChild)c.removeChild(c.firstChild);
 const node=(tag,t)=>{const el=d.createElementNS(d.documentElement.namespaceURI,tag);el.textContent=String(t);c.appendChild(el);return el;};
 if(formula){node('f',formula);if(typeof v!=='number')c.setAttribute('t','str');node('v',v??'');return;}
 if(v===''||v==null)return;
 if(typeof v==='number'){node('v',v);return;}
 c.setAttribute('t','inlineStr');const is=node('is',''),t=d.createElementNS(d.documentElement.namespaceURI,'t');t.setAttributeNS('http://www.w3.org/XML/1998/namespace','xml:space','preserve');t.textContent=String(v);is.appendChild(t);
}
function shift(formula,delta){
 // Only shift row references outside quoted strings; absolute rows remain fixed.
 return formula.split(/("(?:[^"]|"")*")/).map((s,i)=>i%2?s:s.replace(/(\$?[A-Z]{1,3})(\$?)([1-9]\d*)\b/g,(m,c,a,r)=>a?m:c+(Number(r)+delta))).join('');
}
async function fill(bytes,plans,type,summary){
 const zip=await JSZip.loadAsync(bytes),wb=XLSX.read(bytes,{type:'array',cellFormula:true,cellNF:true});
 const book=parse(await zip.file('xl/workbook.xml').async('string')),rels=parse(await zip.file('xl/_rels/workbook.xml.rels').async('string'));
 const paths=new Map(list(rels,'Relationship').map(r=>[r.getAttribute('Id'),r.getAttribute('Target')]));
 const targets=list(book,'sheet').map(s=>{let path=paths.get(s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id'));path=path.startsWith('/')?path.slice(1):'xl/'+path.replace(/^\.\//,'');return {name:s.getAttribute('name'),path};});
 const matched=new Set();let count=0;
 for(const target of targets){
  const sheet=wb.Sheets[target.name],grid=XLSX.utils.sheet_to_json(sheet,{header:1,raw:false,defval:''});
  const headerAt=grid.findIndex(row=>row.some(x=>NBYEngine.key(x)==='WORK ORDER NBR')||row.some(x=>NBYEngine.key(x)==='CUSTOMER REFERENCE')||row.some(x=>NBYEngine.key(x)==='ORDERNUMBER')||row.some(x=>NBYEngine.key(x)==='SD DOCUMENT'));
  if(type==='status')continue;
  if(headerAt<0)continue;
  const heads=grid[headerAt],keys=heads.map(NBYEngine.key),kind=keys.includes('WORK ORDER NBR')?'fc':keys.includes('CUSTOMER REFERENCE')?'sap':keys.includes('ORDERNUMBER')?'awb':keys.includes('SD DOCUMENT')&&keys.includes('MATERIAL')?'stock':'';
  const plan=plans.find(p=>p.kind===kind);if(!plan)continue;
  const d=parse(await zip.file(target.path).async('string')),sd=list(d,'sheetData')[0],allRows=list(sd,'row'),start=headerAt+2;
  const body=allRows.filter(r=>Number(r.getAttribute('r'))>=start),model=body[0];
  if(!model)throw Error(target.name+' 缺少資料列樣式');
  if(list(d,'tablePart').length)throw Error(target.name+' 含 Excel Table，需先確認表格範圍');
  if(list(d,'mergeCell').some(c=>Number(c.getAttribute('ref').split(':')[1]?.replace(/\D/g,'')||0)>=start))throw Error(target.name+' 資料區有合併儲存格，需確認樣板');
  // FC templates must have a recognizable work-order value on each existing body row.
  const keyCol=keys.indexOf(kind==='fc'?'WORK ORDER NBR':kind==='sap'?'CUSTOMER REFERENCE':kind==='awb'?'ORDERNUMBER':'SD DOCUMENT');
  if(kind==='fc'&&grid.slice(start-1).some(r=>r.some(Boolean)&&!String(r[keyCol]??'').trim()))throw Error(target.name+' 資料區含備註或小計，需確認範圍');
  const originalRows=new Map(body.map(r=>[Number(r.getAttribute('r')),r.cloneNode(true)]));
  body.forEach(r=>sd.removeChild(r));
  const seen={};const headers=heads.map((h,i)=>{const k=String(h||`欄位${i+1}`).trim(),n=seen[k]||0;seen[k]=n+1;return n?`${k}.${n}`:k;});
  for(let i=0;i<plan.rows.length;i++){
   const rowNum=start+i,templateRow=originalRows.get(rowNum)||model,row=templateRow.cloneNode(true),oldNum=Number(templateRow.getAttribute('r'));row.setAttribute('r',rowNum);
   for(const c of list(row,'c'))c.setAttribute('r',col(c.getAttribute('r'))+rowNum);
   const src=plan.rows[i];
   for(let j=0;j<headers.length;j++){
    const address=XLSX.utils.encode_col(j)+rowNum,cell=list(row,'c').find(c=>c.getAttribute('r')===address),f=cell&&list(cell,'f')[0];
    if(f&&f.hasAttribute('t'))throw Error(target.name+' 有共享或陣列公式，需確認樣板');
    const formula=f?shift(f.textContent,rowNum-oldNum):null;
    let v=NBYEngine.val(src,headers[j]);
    const oldCell=sheet[XLSX.utils.encode_col(j)+oldNum];
    if(v!==''&&oldCell?.t==='n'){
      if(oldCell.z&&XLSX.SSF.is_date(oldCell.z)){
        const date=root.NBYCleanDate(v);if(/^\d{4}-\d{2}-\d{2}$/.test(date))v=Date.parse(date+'T00:00:00Z')/86400000+25569;
      }else if(NBYEngine.qty(v)!==null)v=NBYEngine.qty(v);
    }
    // Cache formula results only when source explicitly supplies the output column.
    if(formula&&!Object.keys(src).some(k=>NBYEngine.key(k)===NBYEngine.key(headers[j])))v='';
    setCell(d,row,address,v,formula);
   }
   sd.appendChild(row);
  }
  const end=Math.max(headerAt+1,headerAt+1+plan.rows.length),last=XLSX.utils.encode_col(headers.length-1)+end;
  const dim=list(d,'dimension')[0];if(dim)dim.setAttribute('ref','A1:'+last);
  for(const filter of list(d,'autoFilter'))filter.setAttribute('ref','A'+(headerAt+1)+':'+last);
  zip.file(target.path,text(d));matched.add(kind);count++;
 }
 if(type==='status'){
  const target=targets.find(t=>/^summary$/i.test(t.name)),detail=targets.find(t=>/^detail shipment$/i.test(t.name));
  if(!target||!detail)throw Error('需要 Summary 及 Detail shipment 工作表');
  // Update only an existing date column/row; preserve all history and Expect cells.
  const sd=parse(await zip.file(target.path).async('string')),ss=wb.Sheets[target.name],g=XLSX.utils.sheet_to_json(ss,{header:1,raw:false,defval:''});
  const hi=g.findIndex(r=>r.some(v=>NBYEngine.key(v)==='ACTUALLY SHIPMENT'));if(hi<0)throw Error('Summary 欄位未辨識');
  const dc=g[hi].findIndex(v=>NBYEngine.key(v)==='DATE'),ac=g[hi].findIndex(v=>NBYEngine.key(v)==='ACTUALLY SHIPMENT');
  const ri=g.findIndex((r,i)=>i>hi&&root.NBYCleanDate(r[dc])===summary.date);if(ri<0)throw Error('Summary 尚無所選日期，請在樣板新增日期列');
  const rr=list(sd,'row').find(r=>Number(r.getAttribute('r'))===ri+1),address=XLSX.utils.encode_cell({r:ri,c:ac});
  if(ss[address]?.f)throw Error('Actually Shipment 是公式，需先確認來源規則');
  setCell(sd,rr,address,summary.actual);
  if(summary.expectDate&&summary.expect!=null){
   const ec=g[hi].findIndex(v=>NBYEngine.key(v)==='EXPECT SHIPMENT');
   const er=g.findIndex((r,i)=>i>hi&&root.NBYCleanDate(r[dc])===summary.expectDate);
   if(ec<0||er<0)throw Error('Summary 缺少預計出貨日期或欄位：'+summary.expectDate);
   const ea=XLSX.utils.encode_cell({r:er,c:ec});
   if(ss[ea]?.f)throw Error('Expect Shipment 是公式，需確認樣板');
   const erow=list(sd,'row').find(r=>Number(r.getAttribute('r'))===er+1);
   if(!erow)throw Error('Summary 缺少預計出貨日期列');
   setCell(sd,erow,ea,summary.expect);
  }
  zip.file(target.path,text(sd));
  const dd=parse(await zip.file(detail.path).async('string')),ds=wb.Sheets[detail.name],dg=XLSX.utils.sheet_to_json(ds,{header:1,raw:false,defval:''});
  const dh=dg.findIndex(r=>r.some(v=>NBYEngine.key(v)==='PRODUCT NAME'));if(dh<0)throw Error('Detail shipment 欄位未辨識');
  const pc=dg[dh].findIndex(v=>NBYEngine.key(v)==='PRODUCT NAME'),dateCol=dg[dh].findIndex(v=>root.NBYCleanDate(v)===summary.date);if(dateCol<0)throw Error('Detail shipment 尚無所選日期欄');
  const products=new Map(summary.products.map(x=>[NBYEngine.key(x.product),x.count]));
  for(let i=dh+1;i<dg.length;i++){const product=NBYEngine.key(dg[i][pc]);if(!product||/TOTAL|合計/.test(product))continue;
   const address=XLSX.utils.encode_cell({r:i,c:dateCol});if(ds[address]?.f)continue;
   const row=list(dd,'row').find(r=>Number(r.getAttribute('r'))===i+1);if(row)setCell(dd,row,address,products.get(product)||0);products.delete(product);
  }
  if(products.size)throw Error('Detail shipment 缺少型體：'+[...products.keys()].join(', '));zip.file(detail.path,text(dd));count=2;
 }
 if(!count||type!=='status'&&!matched.has('fc'))throw Error('找不到可用的 Work Order 樣板工作表');
 if(type==='data'&&plans.some(p=>!matched.has(p.kind)))throw Error('Data Shipping 樣板缺少 FC／SAP／MB52／AWB 工作表');
 let calc=list(book,'calcPr')[0];if(!calc){calc=book.createElementNS(book.documentElement.namespaceURI,'calcPr');book.documentElement.appendChild(calc);}calc.setAttribute('fullCalcOnLoad','1');calc.setAttribute('forceFullCalc','1');zip.file('xl/workbook.xml',text(book));
 // Calculation chains refer to old row addresses. Excel rebuilds them on open.
 if(zip.file('xl/calcChain.xml')){zip.remove('xl/calcChain.xml');for(const r of list(rels,'Relationship'))if(/calcChain$/.test(r.getAttribute('Type')))r.parentNode.removeChild(r);zip.file('xl/_rels/workbook.xml.rels',text(rels));const types=parse(await zip.file('[Content_Types].xml').async('string'));for(const n of list(types,'Override'))if(n.getAttribute('PartName')==='/xl/calcChain.xml')n.parentNode.removeChild(n);zip.file('[Content_Types].xml',text(types));}
 return zip.generateAsync({type:'uint8array'});
}
root.NBYTemplates={fill};
})(typeof window!=='undefined'?window:globalThis);
