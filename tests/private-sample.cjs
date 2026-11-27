const fs=require('fs'),path=require('path'),assert=require('assert');
const base=process.cwd();const {JSDOM}=require(path.join(base,'node_modules/jsdom'));
const dom=new JSDOM(fs.readFileSync('index.html','utf8'),{runScripts:'outside-only',url:'http://localhost'}),w=dom.window;
for(const f of ['vendor/xlsx.full.min.js','vendor/jszip.min.js','engine.js','templates.js'])w.eval(fs.readFileSync(f,'utf8'));
w.JSZip=require(path.join(base,'vendor/jszip.min.js'));const nativeLoad=w.JSZip.loadAsync;w.JSZip.loadAsync=(b)=>nativeLoad(Buffer.from(b));
w.eval(fs.readFileSync('app.js','utf8').replace(/\}\)\(\);\s*$/,'window.TEST={state,parseFile,ingest,calculate,renderSourceChoices};})();'));
const dir=process.argv[2];if(!dir)throw Error('Pass a private sample directory; never commit samples');const names=['DATA 11.20.xlsx','Detailed Report-4.csv','order list download - shipments-20261120.xlsx','NBY_1120_完整源頭逐筆比對.xlsx'];
const keep=setInterval(()=>{},1000);
(async()=>{
 for(const name of names){const b=fs.readFileSync(path.join(dir,name));const file={name,size:b.length,arrayBuffer:async()=>Uint8Array.from(b).buffer};await w.TEST.ingest([file]);console.log('Imported',name,w.TEST.state.sheets.length);}
 w.document.getElementById('fcSource').value='DATA 11.20.xlsx::Export_WorkOrders';
 w.TEST.state.templates.data={name:'DATA 11.20.xlsx',bytes:new w.Uint8Array(fs.readFileSync(path.join(dir,'DATA 11.20.xlsx')))};
 console.log('Calculating');await w.TEST.calculate();const r=w.TEST.state.result;
 console.log(JSON.stringify({counts:r.counts,match:r.orders.filter(o=>o.sapMatch).length,stock:r.orders.filter(o=>o.stockMatch).length,ups:r.upsRows.length,sf:r.sfRows.length,templates:r.templateStatus,manual:r.manual.summaries},null,2));
 assert.equal(r.orders.length,364);assert.equal(r.orders.filter(o=>o.sapMatch).length,364);assert.equal(r.upsRows.length,364);assert.equal(r.sfRows.length,1);
 assert.ok(r.outputFiles['Data Shipping.xlsx'],'template output must succeed');
 fs.mkdirSync('qa-output',{recursive:true});for(const [n,b]of Object.entries(r.outputFiles))fs.writeFileSync(path.join('qa-output',n),Buffer.from(b));
 const out=w.XLSX.read(r.outputFiles['Data Shipping.xlsx'],{type:'array'});assert.equal(out.SheetNames.length,5);assert.ok(out.Sheets.Export_WorkOrders.D2.f);
 const before=await w.JSZip.loadAsync(w.TEST.state.templates.data.bytes),after=await w.JSZip.loadAsync(r.outputFiles['Data Shipping.xlsx']);assert.equal(await before.file('xl/styles.xml').async('string'),await after.file('xl/styles.xml').async('string'));
 // Generic template works with a different date and row count, with shifted formulas.
 const orderBytes=new w.Uint8Array(fs.readFileSync(path.join(dir,'2026 11 20.xlsx')));
 const rows=r.orders.slice(0,2).map(o=>o.source);const b=await w.NBYTemplates.fill(orderBytes,[{kind:'fc',rows}],'orders',{});
 const ow=w.XLSX.read(b,{type:'array'});assert.equal(ow.Sheets.Export_WorkOrders['!ref'],'A1:AX3');assert.equal(ow.Sheets.Export_WorkOrders.W3.f,'L3');
 console.log('PASS: 364 sample work orders; 5 worksheets, styles and formulas preserved; resized template verified.');
 clearInterval(keep);dom.window.close();
})().catch(e=>{console.error(e);process.exit(1)});
