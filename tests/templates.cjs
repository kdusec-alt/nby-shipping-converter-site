const fs=require('fs'),assert=require('node:assert/strict');const {JSDOM}=require('jsdom');
const w=new JSDOM('',{runScripts:'outside-only'}).window;
w.eval(fs.readFileSync('vendor/xlsx.full.min.js','utf8'));w.eval(fs.readFileSync('engine.js','utf8'));w.eval(fs.readFileSync('templates.js','utf8'));
w.JSZip=require('../vendor/jszip.min.js');const load=w.JSZip.loadAsync;w.JSZip.loadAsync=b=>load(Buffer.from(b));w.NBYCleanDate=String;
(async()=>{
 const wb=w.XLSX.utils.book_new();
 w.XLSX.utils.book_append_sheet(wb,w.XLSX.utils.aoa_to_sheet([['Date','Actually Shipment','Expect Shipment'],['2026-11-19',456,500],['2026-11-20',9,400]]),'Summary');
 w.XLSX.utils.book_append_sheet(wb,w.XLSX.utils.aoa_to_sheet([['Product Name','2026-11-19','2026-11-20'],['Sabrina',456,9],['TOTAL',456,9]]),'Detail shipment');
 for(let i=0;i<5;i++)w.XLSX.utils.book_append_sheet(wb,w.XLSX.utils.aoa_to_sheet([['history',i]]),'History'+i);
 const bytes=w.XLSX.write(wb,{type:'array',bookType:'xlsx'}),pending=[];
 const out=await w.NBYTemplates.fill(bytes,[],'status',{date:'2026-11-20',products:[{product:'Sabrina'},{product:'Unknown'}],productAliases:{},resolveProduct:async()=>null,pendingMappings:pending});
 const result=w.XLSX.read(out,{type:'array'});assert.equal(result.SheetNames.length,7);assert.equal(result.Sheets.Summary.B3.v,9);assert.equal(result.Sheets.Summary.B2.v,456);assert.equal(result.Sheets.Summary.C3.v,400);assert.equal(result.Sheets['Detail shipment'].C2.v,9);assert.equal(pending[0].reason,'UNCONFIRMED');
 const again=await w.NBYTemplates.fill(out,[],'status',{date:'2026-11-20',products:[{product:'Unknown'}],productAliases:{},resolveProduct:async()=>null,pendingMappings:[]});const rerun=w.XLSX.read(again,{type:'array'});assert.equal(rerun.Sheets.Summary.B3.v,9);assert.equal(rerun.Sheets['Detail shipment'].C2.v,9);
 console.log('PASS: unknown product remains pending without blocking; actual and history remain unchanged on rerun.');w.close();
})().catch(e=>{console.error(e);process.exit(1)});
