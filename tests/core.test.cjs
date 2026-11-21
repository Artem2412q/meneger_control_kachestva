const assert=require('node:assert/strict');
const C=require('../assets/core.js');
const p={
 saldo:{'123':{name:'Тестовый товар',start:12,end:8,group:'Плодоовощные товары'}},
 service:{'20.11.2026|123':{code:'123',name:'Тестовый товар',date:'20.11.2026',order:4,imp:5,ship:2,vich:-2,fresh:true}},
 catalog:['123'],norms:{'123':{nk:2,kk:3}}
};
let report=C.buildReport(p,{date:'20.11.2026',mode:'reference'});
assert.equal(report.records.length,1);
let r=report.records[0];
assert.equal(r.usOrder,0.5);
assert.equal(r.usImport,0.4);
assert.equal(r.actualDays,4);
assert.equal(r.orderDays,2);
assert.equal(r.targetComment,2);
assert.equal(r.maxComment,1);
assert.equal(r.vich,-2);
report=C.buildReport({...p,service:{}},{date:'20.11.2026',mode:'reference'});
assert.equal(report.records[0].targetComment,'Нет данных');
assert.equal(report.records[0].order,null);
assert.equal(report.records[0].ship,null);
report=C.buildReport({...p,norms:{}},{date:'20.11.2026',mode:'all'});
assert.equal(report.records[0].nk,null);
assert.equal(report.records[0].targetComment,'Нет НТЗ');
assert.equal(report.stats.normMissing,1);
assert.equal(C.parseRuDate('20.11.2026'),'2026-11-20');
assert.equal(C.toRuDate('2026-11-20'),'20.11.2026');
assert.equal(C.normalize(1234567890),'1234567890');
console.log('Все 13 проверок бизнес-логики пройдены.');
const reportBuilder=require('../assets/report-builder.js');
const exporter=require('../assets/export.js');
const assertThrows=(fn,part)=>assert.throws(fn,e=>String(e.message).includes(part));
assert.equal(C.todayIso(new Date(2026,10,21)),'2026-11-21');
assertThrows(()=>C.buildCriticalReport(p,{today:'2026-11-21'}),'строго за сегодня');
const crit=C.buildCriticalReport(p,{today:'2026-11-20'});
assert.equal(crit.date,'20.11.2026');
assert.equal(crit.records.length,1);
const first={mainDate:'24.10.2026',norms:{'123':{nk:1,kk:5}},snapshots:[{date:'15.05.2026',norms:{'123':{nk:2,kk:4}}}],flowerCodes:['F1'],catalog:['123']};
const second={mainDate:'21.11.2026',norms:{'123':{nk:3,kk:6}},snapshots:[],flowerCodes:[],catalog:[]};
const merged=C.mergeNormHistory(first,second);
assert.equal(merged.snapshots.length,2);
assert.equal(merged.snapshots[1].date,'24.10.2026');
assert.equal(merged.flowerCodes[0],'F1');
assert.equal(C.selectedNorms(merged,'1')['123'].nk,1);
assert.equal(C.selectedNorms(merged,'main')['123'].nk,3);
const defs=reportBuilder.prepareSheets({stock:C.buildReport({...p,norms:second.norms},{date:'20.11.2026'}),template:{...merged,catalog:['123'],flowerCodes:['123']},stockSources:p});
assert.deepEqual(defs.map(d=>d.name),['20.11.2026','НТЗ','Архив НТЗ','код цветы','Цветы','УС','Контроль']);
assert.equal(defs[0].rows[0][10].v,3);
assert.match(defs[0].rows[0][10].f,/VLOOKUP/);
assert.equal(defs[0].rows[0][11].v,6);
assert.equal(defs[4].rows.length,1);
assert.match(exporter.sheetXml(defs[0]),/s="3"/);
assert.match(exporter.sheetXml(defs[0]),/НЕТ В НТЗ/);
// Export style regression: retain precision in XML, display % with two decimal places,
// and show other numeric report columns with exactly two decimal places.
const stylesXml=exporter.stylesXml();
assert.match(stylesXml,/<cellXfs count="16">/);
const xfs=[...stylesXml.matchAll(/<xf numFmtId="(\d+)" fontId=/g)].map(x=>Number(x[1]));
assert.equal(xfs[xfs.length-1],10); // built-in 0.00%
assert.equal(xfs[xfs.length-2],2);  // built-in 0.00
const dailyXml=exporter.sheetXml(defs[0]);
assert.match(dailyXml,/<c r="G2" s="15">/);
assert.match(dailyXml,/<c r="H2" s="15">/);
assert.match(dailyXml,/<c r="M2" s="14">/);
assert.match(dailyXml,/<c r="N2" s="14">/);
assert.match(dailyXml,/<c r="O2" s="10">/);
assert.match(dailyXml,/<c r="P2" s="10">/);
assert.match(dailyXml,/0\.5/); // ratio is stored without percent multiplication
const alertReport=C.buildReport({...p,norms:{'123':{nk:1,kk:2}}},{date:'20.11.2026'});
assert.match(exporter.sheetXml(reportBuilder.makeDaySheet(alertReport,{normLength:1})),/<c r="N2" s="11">/);
assertThrows(()=>reportBuilder.prepareSheets({stock:crit,template:null}),'НТЗ');
console.log('Дополнительные проверки: строгая дата критики, объединение архива, цветы, НК/КК, оформление Excel.');

// ETALON 2.6: the master is an allowlist, NEVER a source of fictitious daily rows.
const {readFileSync}=require('node:fs');
const vm=require('node:vm');
const {resolve}=require('node:path');
const context={window:{}};
vm.runInNewContext(readFileSync(resolve(__dirname,'../assets/etalon.js'),'utf8'),context);
const E=context.window.MagnitEtalon;
assert.equal(E.sourceLines,46555);
assert.equal(E.codes.length,4287);
assert.equal(E.duplicates,42268);
assert.equal(E.anomalies.length,2);
assert.ok(E.codes.includes('144428298'));
assert.ok(E.codes.includes('80000129271000465885'));
const parsed=C.parseEtalon('00234\n00234\n3434\n\n');
assert.deepEqual(parsed.codes,['00234','3434']);
assert.equal(parsed.duplicates,1);
assert.equal(parsed.anomalies.length,2);
const master=C.buildReport({...p,etalonCodes:['123','456','789','123'],norms:{'123':{nk:2,kk:3},'456':{nk:1,kk:5}}},{date:'20.11.2026',mode:'etalon'});
assert.deepEqual(master.records.map(r=>r.code),['123']);
assert.equal(master.stats.missingBoth,0);
assert.equal(master.stats.etalonNotInSources,2);
assert.equal(master.stats.normMissing,0);
assert.equal(master.stats.saldoMissing,0);
assert.equal(master.stats.serviceMissing,0);
assert.equal(master.stats.outsideEtalon.length,0);
// If present in only one file, the row exists, but missing columns stay empty.
const oneSource=C.buildReport({saldo:{'ABC':{name:'Реальный товар',start:3,end:5}},etalonCodes:['ABC','MISSING'],norms:{MISSING:{nk:7,kk:9}}},{date:'20.11.2026',mode:'etalon'});
assert.deepEqual(oneSource.records.map(x=>x.code),['ABC']);
assert.equal(oneSource.records[0].ship,null);
assert.equal(oneSource.records[0].name,'Реальный товар');
assert.equal(oneSource.stats.etalonNotInSources,1);
const noName=C.buildReport({saldo:{'ABC':{start:1,end:2}},norms:{}},{date:'20.11.2026',mode:'all'});
assert.equal(noName.records[0].name,''); // No invented 'Название не найдено' text.
const reduced={...p,saldo:{'123':p.saldo['123'],'999':{name:'вне Эталона',start:1,end:2}}};
const outside=C.buildReport({...reduced,etalonCodes:['123','456']},{date:'20.11.2026',mode:'etalon'});
assert.deepEqual(outside.stats.outsideEtalon,['999']);
const realCode=E.codes[0];
const realSources={saldo:{[realCode]:{name:'Подтверждённый товар',start:2,end:4}},service:{}};
const fullMaster=C.buildReport({...realSources,norms:{},etalonCodes:E.codes},{date:'20.11.2026',mode:'etalon'});
assert.equal(fullMaster.records.length,1);
assert.equal(fullMaster.records[0].code,realCode);
assert.equal(fullMaster.stats.missingBoth,0);
assert.equal(fullMaster.stats.etalonNotInSources,4286);
const fullSheets=reportBuilder.prepareSheets({stock:fullMaster,stockSources:realSources,template:{...merged,catalog:['123'],flowerCodes:['123']},etalon:E});
const masterSheet=fullSheets.find(x=>x.name==='Эталон ТП');
assert.ok(masterSheet);
assert.equal(masterSheet.rows.length,1);
assert.equal(masterSheet.rows[0][0],realCode);
assert.ok(!masterSheet.rows.some(r=>r[0]==='80000129271000465885'));
const main=fullSheets[0];
assert.equal(main.rows.length,1);
assert.equal(main.rows[0][0],realCode);
assert.equal(main.rows[0][3],null);
assert.equal(main.rows[0][12].v,'');
const masterXml=exporter.sheetXml(main);
assert.match(masterXml,/<c r="A2" s="1" t="inlineStr">/);
assert.doesNotMatch(masterXml,/Название не найдено/);
assert.doesNotMatch(masterXml,/80000129271000465885/);
console.log('Эталон 2.6: только фактические коды; 4 286 неподтверждённых строк исключены; фальшивые названия запрещены.');
