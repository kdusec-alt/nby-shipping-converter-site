const fs=require('fs'),assert=require('node:assert/strict');const {JSDOM}=require('jsdom');
const w=new JSDOM('',{runScripts:'outside-only'}).window;
w.eval(fs.readFileSync('vendor/xlsx.full.min.js','utf8'));w.eval(fs.readFileSync('engine.js','utf8'));w.eval(fs.readFileSync('templates.js','utf8'));
w.JSZip=require('../vendor/jszip.min.js');const load=w.JSZip.loadAsync;w.JSZip.loadAsync=b=>load(Buffer.from(b));w.NBYCleanDate=String;
(async()=>{
 const wb=w.XLSX.utils.book_new();
 w.XLSX.utils.book_append_sheet(wb,w.XLSX.utils.aoa_to_sheet([['Date','Actually Shipment','Expect Shipment'],['2026-11-19',456,500],['2026-11-20',9,400]]),'Summary');
 w.XLSX.utils.book_append_sheet(wb,w.XLSX.utils.aoa_to_sheet([['Product Name','2026-11-19','2026-11-20'],['Sabrina',456,9]]),'Detail shipment');
 const bytes=w.XLSX.write(wb,{type:'array',bookType:'xlsx'});
 const out=await w.NBYTemplates.fill(bytes,[],'status',{date:'2026-11-20',actual:364,products:[{product:'Sabrina',count:364}]});
 const result=w.XLSX.read(out,{type:'array'});assert.equal(result.Sheets.Summary.B3.v,364);assert.equal(result.Sheets.Summary.B2.v,456);assert.equal(result.Sheets.Summary.C3.v,400);assert.equal(result.Sheets['Detail shipment'].C2.v,364);
 await assert.rejects(w.NBYTemplates.fill(bytes,[],'status',{date:'2026-11-21',actual:1,products:[]}),/尚無所選日期/);
 console.log('PASS: status template preserves history and Expect values, updates selected date, blocks absent dates.');w.close();
})().catch(e=>{console.error(e);process.exit(1)});
