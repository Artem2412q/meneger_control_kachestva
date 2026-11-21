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
assert.equal(report.records[0].targetComment,'Без расхода');
report=C.buildReport({...p,norms:{}},{date:'20.11.2026',mode:'all'});
assert.equal(report.records[0].nk,null);
assert.equal(report.records[0].targetComment,'');
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
