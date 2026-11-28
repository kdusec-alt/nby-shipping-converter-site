/* Pure reconciliation rules. No network, templates or UI state. */
(function(root){
'use strict';
const key=v=>String(v??'').trim().replace(/\.0$/,'').replace(/\s+/g,' ').toUpperCase();
const val=(r,...names)=>{for(const n of names){const k=Object.keys(r||{}).find(k=>key(k)===key(n));if(k&&r[k]!==''&&r[k]!=null)return r[k];}return '';};
const qty=v=>String(v??'').trim()!==''&&Number.isFinite(Number(String(v).replace(/,/g,'')))?Number(String(v).replace(/,/g,'')):null;
const signature=r=>JSON.stringify(Object.keys(r).sort().map(k=>[k,r[k]]));
const unique=rows=>[...new Map(rows.map(r=>[signature(r),r])).values()];
const index=(rows,fn)=>{const m=new Map();for(const r of rows){const k=key(fn(r));if(k)m.set(k,[...(m.get(k)||[]),r]);}return m;};
const carrier=s=>/UPS/.test(key(s))?'UPS':/BAOZUN|BAOTONG|^SF\b|順豐/.test(key(s))?'SF':'UNKNOWN';
const sum=(rows,column)=>{const ns=rows.map(r=>qty(val(r,column)));return rows.length&&ns.every(n=>n!==null)?ns.reduce((a,b)=>a+b,0):'';};
function reconcile({fc,sap,stock,ups,sf,awb,date,cleanDate}){
 const anomalies=[];const add=(severity,category,id,detail)=>anomalies.push({嚴重度:severity,異常類型:category,工單或Tracking:id,說明:detail});
 const sapIndex=index(unique(sap),r=>val(r,'Customer Reference')),stockIndex=index(unique(stock),r=>val(r,'SD Document')),awbIndex=index(unique(awb),r=>val(r,'OrderNumber'));
 const groups=index(fc,r=>val(r,'Work Order Nbr'));const orders=[];
 for(const row of fc)if(!key(val(row,'Work Order Nbr')))add('高','FC 缺少 Work Order Nbr','', '未納入工單統計；原始列保留於匯出明細');
 for(const [wo,raw] of groups){
  const rows=unique(raw),r=rows[0];if(raw.length>1)add('中','FC 工單重複',wo,`原始 ${raw.length} 列，不同內容 ${rows.length} 列；工單僅計一次，明細全部保留`);
  const ambiguous=rows.length>1;if(ambiguous)add('高','FC 工單明細粒度待確認',wo,'同工單有不同內容，SAP／物流關聯不採用任一列作為唯一答案');
  const sr=sapIndex.get(wo)||[],docs=[...new Set(sr.map(x=>key(val(x,'SD Document'))).filter(Boolean))];
  const itemGroups=index(sr,x=>[val(x,'SD Document'),val(x,'Item (SD)','SD Item','Item'),val(x,'Material')].map(key).join('|'));
  const sapConflict=[...itemGroups.values()].some(v=>v.length>1);
  const sapDoc=!ambiguous&&docs.length===1&&!sapConflict?docs[0]:'';
  if(docs.length>1||sapConflict)add('高','SAP 關聯不唯一',wo,'不同 SAP 文件或同項次存在衝突，未選取第一筆');
  const stockRows=sapDoc?(stockIndex.get(sapDoc)||[]).filter(st=>sr.some(s=>
    (!val(s,'Material')||key(val(st,'Material'))===key(val(s,'Material')))&&(!val(s,'Item (SD)','SD Item','Item')||key(val(st,'Item (SD)','SD Item','Item'))===key(val(s,'Item (SD)','SD Item','Item'))))):[];
  const tracks=[...new Set(rows.map(x=>key(val(x,'Tracking Number'))))];const track=tracks.length===1?tracks[0]:'';
  const shipper=val(r,'Shipper'),c=carrier(shipper);const ar=awbIndex.get(wo)||[],av=[...new Set(ar.map(x=>key(val(x,'ConsolidationShipmentId'))).filter(Boolean))];
  const awbValue=av.length===1?val(ar[0],'ConsolidationShipmentId'):'';if(av.length>1)add('高','AWB 對照衝突',wo,'同 OrderNumber 有多個 ConsolidationShipmentId');
  const ur=c==='UPS'&&!ambiguous?ups.filter(x=>key(val(x,'Tracking Number'))===track&&track):[];
  const fr=c==='SF'&&!ambiguous?sf.filter(x=>key(val(x,'Work Order Number','Work Order Nbr'))===wo):[];
  const dates=[...new Set(rows.map(x=>cleanDate(val(x,'Manifest Scan Date'))).filter(Boolean))];const o={wo,fcDate:dates.length===1?dates[0]:dates.join(', '),shipper,carrier:c,track,sapDoc,confirmedQty:sapDoc?sum(sr,'Confirmed Quantity'):'',pairQty:sapDoc&&sr.length&&sr.every(x=>{const own=key(val(x,'Unit of Measure','Sales Unit','Base Unit of Measure'));const matching=stockRows.filter(st=>key(val(st,'Material'))===key(val(x,'Material'))&&(!val(x,'Item (SD)','SD Item','Item')||key(val(st,'Item (SD)','SD Item','Item'))===key(val(x,'Item (SD)','SD Item','Item'))));const unit=own||((matching.length&&matching.every(st=>key(val(st,'Base Unit of Measure'))==='PAA'))?'PAA':'');return unit==='PAA';})&&sr.every(x=>true&&qty(val(x,'Confirmed Quantity'))!==null&&qty(val(x,'Confirmed Quantity'))>=0)?sum(sr,'Confirmed Quantity'):null,sapRows:sr,mbDoc:stockRows.length?sapDoc:'',sapMatch:!!sapDoc,stockMatch:!!stockRows.length,upsCount:ur.length,upsMatch:!!ur.length,awb:awbValue,sfRows:fr,source:r,sourceRows:rows,stockRows,ambiguous};orders.push(o);
  if(!o.sapMatch)add('高','FC 工單缺少 SAP ZRSD1019',wo,'找不到唯一可用的 SAP 對照');
  if(o.sapMatch&&!o.stockMatch)add('中','SAP 訂單缺少 MB52 庫存紀錄',wo,'文件／物料／項次沒有庫存明細；不減少實際出貨');
  if(!track)add('中','FC 缺少 Tracking Number',wo,'追蹤號缺漏或同工單有多個追蹤號');
  // Blank AWB is not an exception until applicability is configured.
  if(c==='UNKNOWN')add('高','未辨識物流公司',wo,String(shipper));
  if(c==='UPS'&&!ur.length)add('中','FC Tracking 未出現在 UPS',wo,track);
  if(c==='SF'&&!fr.length)add('中','FC 工單未出現在 SF',wo,'SF 樣本可能不完整');
 }
 const ft=index(orders,o=>o.track);for(const [t,os] of ft)if(os.length>1)add('中','FC Tracking 重複',t,`${os.length} 個工單共用追蹤號，需確認包裹粒度`);
 const upsIndex=index(ups,r=>val(r,'Tracking Number'));
 for(const r of ups)if(!key(val(r,'Tracking Number')))add('高','UPS 缺少 Tracking Number','', '原始列已保留，無法匹配');
 const upsRows=[...upsIndex].map(([track,items])=>{
  const os=orders.filter(o=>o.carrier==='UPS'&&o.track===track&&!o.ambiguous),stocks=unique(os.flatMap(o=>o.stockRows));
  if(items.length>1)add('中','UPS Tracking 重複',track,`${items.length} 筆物流紀錄，不能當成鞋雙數`);
  if(!os.length)add('中','UPS Tracking 未關聯 FC 工單',track,'選定日期 FC 無對應');
  for(const r of items){const d=cleanDate(val(r,'Ship Date'));if(d!==date)add('中','UPS 資料日期不一致',track,`Ship Date ${d||'缺漏'}；選定 ${date}`);}
  return {Tracking_Number:track,UPS_紀錄數:items.length,FC_工單數:os.length,FC_工單:os.map(o=>o.wo).join(', '),SAP_SD_Document:os.map(o=>o.sapDoc).join(', '),SAP_Confirmed_Quantity:os.length&&os.every(o=>qty(o.confirmedQty)!==null)?os.reduce((a,o)=>a+qty(o.confirmedQty),0):'',MB52_明細列數:stocks.length,MB52_Unrestricted:sum(stocks,'Unrestricted'),MB52_Blocked_Stock:sum(stocks,'Blocked Stock'),UPS_Service:val(items[0],'Service'),Ship_Date:cleanDate(val(items[0],'Ship Date')),UPS_Charge:val(items[0],'Charge'),數量核對:'待確認 UPS 鞋雙數欄位及單位；未以紀錄數代替',結果:items.length>1?'UPS 重複紀錄':os.length?'已與 FC Tracking 對上':'UPS 有紀錄但 FC 無對應'};
 });
 const sfGroups=index(sf,r=>val(r,'Work Order Number','Work Order Nbr'));
 const sfRows=sf.map(r=>{const wo=key(val(r,'Work Order Number','Work Order Nbr')),o=orders.find(o=>o.wo===wo&&o.carrier==='SF'&&!o.ambiguous),track=key(val(r,'Tracking Number')),group=val(r,'Ship Group ID','Ship Group Nbr'),many=(sfGroups.get(wo)||[]).length>1;
  const n=qty(val(r,'Qty','Quantity')),sn=o?qty(o.confirmedQty):null,refMatch=!!o&&!many&&n!==null&&sn!==null?n===sn:null;
  const unrestricted=o?qty(sum(o.stockRows,'Unrestricted')):null,blocked=o?qty(sum(o.stockRows,'Blocked Stock')):null;
  const mb52Delta=n!==null&&unrestricted!==null?n-unrestricted:'';
  const groupMatch=!!o&&(!group||key(group)===key(val(o.source,'Ship Group Nbr')));
  const result=!o?'SF 工單未出現在 FC':many?'SF 多筆明細，數量粒度待確認':!groupMatch?'Ship Group 不一致':track!==o.track?'Tracking 不一致':refMatch===false?'FC Tracking 一致；SAP Confirmed Qty 僅供參考，MB52 數量待判讀':'FC Tracking 一致；數量未認證';
  if(!o||many||!groupMatch||track!==o?.track)add('中',result,wo,'SF 與 FC 比對；SAP Confirmed Qty 僅供參考。MB52 是庫存快照，差額不代表出貨差異');
  if(refMatch===false)add('低','SF Qty 與 SAP Confirmed Qty 參考差異（待判讀）',wo,`SF Qty ${n}；SAP Confirmed Quantity ${sn}。此欄僅供參考，尚非核准數量規則`);
  if(mb52Delta!==''&&mb52Delta!==0)add('低','MB52 庫存快照差額（待業務判讀）',wo,`SF Qty 減 MB52 Unrestricted = ${mb52Delta}；不代表日出貨量差異`);
  return {Work_Order_Number:wo,Ship_Group_ID:group,Tracking_Number:track,Mother_Waybill:val(r,'Mother waybill'),Estimated_Ship_Date:cleanDate(val(r,'Estimated Ship Date','Est Ship Date')),Qty:val(r,'Qty','Quantity'),SAP_SD_Document:o?.sapDoc||'',SAP_Confirmed_Qty_Reference:o?.confirmedQty??'',Qty_vs_SAP_Reference:refMatch===null?'待確認':refMatch?'相同':'不同',Qty與SAP一致:refMatch===null?'待確認':refMatch?'參考相同':'參考不同',MB52_Unrestricted_Reference:unrestricted??'',MB52_Blocked_Stock_Reference:blocked??'',SF_Qty_minus_MB52_Unrestricted_Reference:mb52Delta,MB52_Interpretation:'庫存快照，不等同當日出貨量；差額未認證',SAP_工單存在:!!o?.sapMatch,MB52_紀錄存在:!!o?.stockMatch,FC_工單存在:!!o,FC_Tracking一致:!!o&&track===o.track,數量核對:'待確認 SF Qty 單位及 MB52 業務規則',結果:result};
 });
 return {orders,upsRows,sfRows,anomalies};
}
const api={key,val,qty,unique,carrier,reconcile};root.NBYEngine=api;if(typeof module!=='undefined')module.exports=api;
})(typeof window!=='undefined'?window:globalThis);
