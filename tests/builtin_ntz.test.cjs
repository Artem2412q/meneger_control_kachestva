/* Regression v2.5: a clean browser has built-in НТЗ; only two uploads needed. */
const assert=require('node:assert/strict');
const fs=require('node:fs'),vm=require('node:vm'),path=require('node:path');
const root=path.resolve(__dirname,'..');
const Core=require('../assets/core.js');
const Report=require('../assets/report-builder.js');
const Excel=require('../assets/export.js');
function domEl(tag='div') {
 const ev={};
 const el={tag,children:[],options:[],value:'',hidden:false,textContent:'',disabled:false,
  classList:{toggle(){},add(){},remove(){}},lastChild:{textContent:''},
  addEventListener(name,func){ev[name]=func;},click(){if(ev.click)ev.click();},
  append(...x){this.children.push(...x);},
  replaceChildren(...x){this.children=x;if(this.tag==='select'){this.options=x;this.value=x[0]?.value||'';this.selectedIndex=0;}},
  querySelector(q){return this._q||(this._q={...domEl(q),replaceChildren(){this.children=[];},append(x){this.children.push(x);}});},
  setAttribute(){},get files(){return []}
 };
 return el;
}
const els={};
const document={getElementById(id){if(!els[id])els[id]=domEl(id.includes('Snapshot')?'select':'div');return els[id];},
 querySelectorAll(){return []},querySelector(){return domEl()},createElement(tag){return domEl(tag);}};
const ctx=vm.createContext({window:{},document,localStorage:{getItem(){return null},removeItem(){},setItem(){}},
 Option:class Option{constructor(text,value){this.text=text;this.value=value}},
 console,Date,Number,String,Error,Map,Set,setTimeout,clearTimeout,URL,TextEncoder,TextDecoder,Blob,Uint8Array,DataView,ArrayBuffer,Event:class Event{constructor(type){this.type=type;}}});
ctx.self=ctx.window;
for(const file of ['core.js','etalon.js','norms.js','export.js','report-builder.js','app.js']){
 vm.runInContext(fs.readFileSync(path.resolve(root,'assets',file),'utf8'),ctx,{filename:file});
}
const t=ctx.window.MagnitNorms;
assert.equal(Object.keys(t.norms).length,1148);
assert.equal(t.snapshots.length,3);
assert.deepEqual(Array.from(t.snapshots,x=>x.date),['15.05.2026','16.05.2026','09.10.2026']);
assert.equal(t.flowerCodes.length,627);
assert.equal(els.normSnapshot.options.length,4);
assert.match(els.templateStatus.textContent,/Встроенный НТЗ: 1148 кодов/);
assert.match(els.criticalTemplateStatus.textContent,/НТЗ подключён автоматически/);
assert.doesNotMatch(els.criticalExportHint.textContent,/подключить НТЗ|загрузите НТЗ/i);
assert.equal(els.clearTemplate.hidden,true);
const code='1000001491';assert.equal(t.norms[code].nk,4);assert.equal(t.norms[code].kk,8);
const day=Core.toRuDate(Core.todayIso());
const input={saldo:{[code]:{name:'Тест',start:10,end:8,group:'Овощи',group2:''}},
 service:{[day+'|'+code]:{code,date:day,name:'Тест',order:4,imp:4,ship:2,vich:0,fresh:true}},
 norms:Core.selectedNorms(t,'main'),etalonCodes:[code],flowerCodes:t.flowerCodes};
const crit=Core.buildCriticalReport(input,{today:Core.todayIso(),mode:'etalon'});
assert.equal(crit.records.length,1);assert.equal(crit.records[0].nk,4);assert.equal(crit.records[0].kk,8);
const ss=Report.prepareSheets({critical:crit,criticalSources:input,template:t,etalon:{codes:[code],sourceLines:1,duplicates:0}});
assert.deepEqual(ss.map(x=>x.name),['Критика '+day,'Эталон ТП','НТЗ','Архив НТЗ','код цветы','Контроль']);
assert.equal(ss[2].rows.length,1148);
assert.equal(ss[3].rows.length,1148);
assert.equal(ss[4].rows.length,627);
Excel.makeWorkbook(ss).arrayBuffer().then(bytes=>{assert.ok(bytes.byteLength>5000);console.log('PASS: clean browser, builtin norms, 2-input Critical, main+archive+flowers, XLSX export',bytes.byteLength);}).catch(e=>{console.error(e);process.exitCode=1;});
