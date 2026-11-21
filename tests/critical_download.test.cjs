/* Regression for the separate "today/today" Critical Excel workflow. */
const fs=require('node:fs'),vm=require('node:vm'),path=require('node:path'),assert=require('node:assert/strict');
const dirname=path.resolve(__dirname,'..');
const Core=require('../assets/core.js');
const R=require('../assets/report-builder.js');
const X=require('../assets/export.js');
const date='21.11.2026';
const item='1000036169';
function workbook(def){return X.makeWorkbook([def]).arrayBuffer();}
const saldoDef={headers:['Код товара','Название','Начальное количество','Конечное количество','ГР_ПНАЗВ'],rows:[[item,'Помидор тест',18,12,'Овощи']]};
const serviceDef={headers:['Код товара','Наим_ТП','Заказ','Импорт','Отгрузка','Vicherk','Дата операции','Тип_Отбор'],rows:[[item,'Помидор тест',5,5,3,1,date,'Овощи-Фрукты']]};
const templateDef={name:'НТЗ',headers:['Код товара','Наименование товара','НК ТЗ для РЦ','КК ТЗ для РЦ','21.11.2026'],rows:[[item,'Помидор тест',2,6]]};
const context=vm.createContext({Blob,DecompressionStream,TextDecoder,TextEncoder,Uint8Array,ArrayBuffer,DataView,Date,Map,Set,console,Error});
context.self=context;
context.importScripts=(...files)=>{for(const f of files)vm.runInContext(fs.readFileSync(path.resolve(dirname,'assets',f),'utf8'),context,{filename:f});};
vm.runInContext(fs.readFileSync(path.resolve(dirname,'assets/worker.js'),'utf8'),context,{filename:'worker.js'});
function call(action, payload){return new Promise((resolve,reject)=>{
 context.postMessage=msg=>{if(msg.type==='error')reject(new Error(msg.message));else if(msg.payload)resolve(msg.payload);};
 context.onmessage({data:{action,...payload}}).catch(reject);
});}
(async()=>{
 const [saldo,service,templateBuf]=await Promise.all([workbook(saldoDef),workbook(serviceDef),workbook(templateDef)]);
 const src=await call('stocks',{saldo,service});
 assert.equal(src.meta.saldoCount,1);
 assert.equal(src.meta.dates[date],1);
 const tmpl=await call('template',{buffer:templateBuf});
 assert.equal(tmpl.norms[item].nk,2);
 assert.equal(tmpl.norms[item].kk,6);
 const payload={...src,norms:tmpl.norms,etalonCodes:[item,'1000036170']};
 const crit=Core.buildCriticalReport(payload,{today:'2026-11-21',mode:'fresh'});
 assert.equal(crit.date,date);assert.equal(crit.records.length,1);assert.equal(crit.records[0].actualDays,4);
 assert.throws(()=>Core.buildCriticalReport(payload,{today:'2026-11-22',mode:'fresh'}),/строго за сегодня/);
 const master=Core.buildCriticalReport(payload,{today:'2026-11-21',mode:'etalon'});
 assert.equal(master.records.length,2);
 const sheets=R.prepareSheets({critical:crit,criticalSources:src,template:tmpl,etalon:{codes:[item,'1000036170'],sourceLines:2,duplicates:0}});
 assert.equal(sheets[0].name,'Критика '+date);
 assert.deepEqual(sheets.map(s=>s.name),['Критика '+date,'Эталон ТП','НТЗ','Архив НТЗ','код цветы','Контроль']);
 assert.equal(sheets[0].rows[0][9].v,2);assert.equal(sheets[0].rows[0][10].v,6);
 assert.equal(sheets[1].rows[0][2],'Да');assert.equal(sheets[1].rows[0][3],'Да');
 const xml=X.sheetXml(sheets[0]);
 assert.match(xml,/r="G2" s="15"/);assert.match(xml,/r="H2" s="15"/);
 assert.match(xml,/VLOOKUP/);
 const bytes=Buffer.from(await X.makeWorkbook(sheets).arrayBuffer());
 assert.ok(bytes.byteLength>3000);
 assert.equal(bytes.toString('ascii',0,2),'PK');
 console.log('Critical date, real XLSX parse, NТЗ, export, Etalon status: PASS',bytes.byteLength,'bytes');
})().catch(e=>{console.error(e);process.exitCode=1});
