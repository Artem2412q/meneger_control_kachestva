const fs=require('node:fs'),vm=require('node:vm'),path=require('node:path'),assert=require('node:assert/strict');
const basedir=path.resolve(__dirname,'..');
const context=vm.createContext({Blob,DecompressionStream,TextDecoder,TextEncoder,Uint8Array,ArrayBuffer,DataView,Date,Map,Set,console,Error});
context.self=context;
context.importScripts=(...files)=>{for(const f of files){vm.runInContext(fs.readFileSync(path.resolve(basedir,'assets',f),'utf8'),context,{filename:f});}};
vm.runInContext(fs.readFileSync(path.resolve(basedir,'assets/worker.js'),'utf8'),context,{filename:'worker.js'});
function buf(file){const b=fs.readFileSync('/mnt/data/'+file);return b.buffer.slice(b.byteOffset,b.byteOffset+b.byteLength);}
function callWorker(action,data){return new Promise((resolve,reject)=>{
 context.postMessage=(message)=>{if(message.type==='error')reject(Error(message.message));if(message.payload)resolve(message.payload);};
 context.onmessage({data:{action,...data}}).catch(reject);
});}
(async()=>{
 const start=Date.now();
 const stock=await callWorker('stocks',{saldo:buf('rep_turnover_saldo_for_logist_excel.xlsx'),service:buf('rpt_levelservice_api.level_service_hd_gk_eday_hm.xlsx')});
 assert.ok(stock.meta.saldoCount>500);assert.ok(stock.meta.serviceCount>500);
 assert.ok(Object.keys(stock.meta.dates).length>0);
 console.log('SOURCE',stock.meta.saldoCount,stock.meta.serviceCount,stock.meta.dates,'ms',Date.now()-start);
 const template=await callWorker('template',{buffer:buf('site_output_test.xlsx')});
 console.log('NORM',Object.keys(template.norms).length,'snapshots',template.snapshots.map(x=>[x.date,Object.keys(x.norms).length]),'flowers',template.flowerCodes.length,'catalog',template.catalog.length);
 assert.ok(Object.keys(template.norms).length>1000);assert.equal(template.snapshots.length,3);
 assert.ok(template.flowerCodes.length>500);
 const C=require('../assets/core.js'), R=require('../assets/report-builder.js'), X=require('../assets/export.js');
 const e={codes:['1000036169','1000036170','1000198550'],sourceLines:3,duplicates:0,anomalies:[]};
 const date=Object.keys(stock.meta.dates)[0];
 const report=C.buildReport({...stock,norms:template.norms,etalonCodes:e.codes},{date,mode:'etalon'});
 const sheets=R.prepareSheets({stock:report,stockSources:stock,template,etalon:e});
 const blob=X.makeWorkbook(sheets);
 assert.ok(blob.size>10000);
 console.log('WORKBOOK',sheets.map(x=>x.name),blob.size,'ms',Date.now()-start);
 const context2=vm.createContext({window:{}});
 vm.runInContext(fs.readFileSync(path.resolve(basedir,'assets/etalon.js'),'utf8'),context2);
 const fullEtalon=context2.window.MagnitEtalon;
 const fullReport=C.buildReport({...stock,norms:template.norms,etalonCodes:fullEtalon.codes},{date,mode:'etalon'});
 assert.ok(fullReport.records.length>0 && fullReport.records.length<4287);
 assert.equal(fullReport.stats.missingBoth,0);
 assert.equal(fullReport.stats.etalonNotInSources,4287-fullReport.records.length);
 const evidence=new Set([...Object.keys(stock.saldo),...Object.values(stock.service).filter(r=>r.date===date).map(r=>r.code)]);
 assert.ok(fullReport.records.every(r=>evidence.has(r.code)));
 const phantom='1000184772';
 if(!evidence.has(phantom))assert.ok(!fullReport.records.some(r=>r.code===phantom));
 const fullSheets=R.prepareSheets({stock:fullReport,stockSources:stock,template,etalon:fullEtalon});
 const etalonSheet=fullSheets.find(s=>s.name==='Эталон ТП');
 assert.ok(etalonSheet.rows.every(r=>evidence.has(r[0])));
 assert.equal(etalonSheet.rows.filter(r=>r[6]==='Нет в обеих').length,0);
 const output=X.makeWorkbook(fullSheets);
 assert.ok(output.size>100000);
 // Verification uses the site's own Excel generator; keep data local and avoid republishing a large commercial file.
 console.log('FULL EXPORT:',fullSheets.map(x=>x.name).join(', '),'codes',fullReport.records.length,'zipBytes',output.size,'ms',Date.now()-start);

})().catch(e=>{console.error('FAILED',e);process.exitCode=1});
