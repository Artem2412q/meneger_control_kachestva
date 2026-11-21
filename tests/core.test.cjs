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
