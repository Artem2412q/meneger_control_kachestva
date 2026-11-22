/* MAGNIT RC — factual store complaint aggregation. Pure logic, no network or storage. */
(function(root,factory){const core=factory();if(typeof module==='object'&&module.exports)module.exports=core;root.MagnitStoreClaimsCore=core;})(typeof self==='undefined'?globalThis:self,function(){
'use strict';
const REQUIRED=['Код товара','Наименование товара','Дата обращения','Характеристика товара/прихода','Производитель'];
const text=v=>String(v==null?'':v).replace(/\s+/g,' ').trim();
const dateISO=v=>{
  if(v===null||v===undefined||v==='')return null;
  if(typeof v==='number'&&v>1&&v<100000){const d=new Date(Date.UTC(1899,11,30)+Math.floor(v)*86400000);return d.toISOString().slice(0,10);}
  const s=text(v);let m=s.match(/^(\d{2})\.(\d{2})\.(\d{4})(?:\s.*)?$/);if(m){const d=`${m[3]}-${m[2]}-${m[1]}`;return validISO(d)?d:null;}
  m=s.match(/^(\d{4}-\d{2}-\d{2})(?:[T\s].*)?$/);return m&&validISO(m[1])?m[1]:null;
};
function validISO(s){if(!/^\d{4}-\d{2}-\d{2}$/.test(s))return false;const [y,m,d]=s.split('-').map(Number);const v=new Date(Date.UTC(y,m-1,d));return v.getUTCFullYear()===y&&v.getUTCMonth()+1===m&&v.getUTCDate()===d;}
const ruDate=iso=>iso&&/^\d{4}-\d{2}-\d{2}$/.test(iso)?iso.slice(8)+'.'+iso.slice(5,7)+'.'+iso.slice(0,4):'';
const previousDate=iso=>{if(!validISO(iso))throw Error('Выберите корректную дату отчёта.');const [y,m,d]=iso.split('-').map(Number);const day=new Date(Date.UTC(y,m-1,d)-86400000);return day.toISOString().slice(0,10);};
const numeric=v=>{if(v===null||v===undefined||text(v)==='')return null;const n=Number(typeof v==='string'?v.replace(/\s/g,'').replace(',','.'):v);return Number.isFinite(n)?n:null;};
const keyText=s=>text(s).toLocaleLowerCase('ru-RU');
function parseTable(table,source='выгрузка'){
  const [headers,...rows]=table;
  if(!Array.isArray(headers))throw Error('В Excel отсутствует строка с названиями столбцов.');
  const col=new Map(headers.map((v,i)=>[text(v),i]));
  const missing=REQUIRED.filter(h=>!col.has(h));
  if(missing.length)throw Error(`В файле «${source}» отсутствуют столбцы: ${missing.join(', ')}.`);
  const val=(row,name)=>col.has(name)?row[col.get(name)]:null;
  const records=[];let empty=0,invalidDate=0;
  for(let i=0;i<rows.length;i++){
    const row=rows[i]||[];
    const code=text(val(row,'Код товара'));const name=text(val(row,'Наименование товара'));
    if(!code&&!name){empty++;continue;}
    // Do not synthesize or repair codes: only actual nonblank numeric codes + names.
    if(!/^\d+$/.test(code)||!name){empty++;continue;}
    const date=dateISO(val(row,'Дата обращения'));
    if(!date){invalidDate++;continue;}
    const guid=text(val(row,'guid')),id=text(val(row,'id'));
    const obj={guid,id,code,name,date,
      characteristic:text(val(row,'Характеристика товара/прихода')),
      manufacturer:text(val(row,'Производитель')),
      storeCode:text(val(row,'Код ТО')),
      store:text(val(row,'Наименование ТО')),
      branch:text(val(row,'Филиал')),
      status:text(val(row,'Статус')),
      badQuantity:numeric(val(row,'Кол-во брак')),
      comment:text(val(row,'Комментарий ТО')),
      reason:text(val(row,'Причина списания')),
      source};
    const fallback=[code,name,date,obj.storeCode,obj.store,obj.characteristic,obj.comment,obj.badQuantity].join('|');
    obj.key=guid?'g:'+guid:id?'id:'+id:'f:'+fallback;
    records.push(obj);
  }
  return {records,stats:{source,totalRows:rows.length,valid:records.length,empty,invalidDate}};
}
function merge(parsed){const byId=new Map();let duplicates=0;for(const item of parsed){for(const rec of item.records){if(byId.has(rec.key))duplicates++;byId.set(rec.key,rec);}}return {records:[...byId.values()],duplicates,files:parsed.map(a=>a.stats)};}
function aggregate(records,{asOf,threshold=2,query=''}={}){
  if(!validISO(asOf))throw Error('Выберите корректную дату отчёта.');
  const yesterday=previousDate(asOf),q=keyText(query);const selected=records.filter(r=>r.date===asOf||r.date===yesterday);
  const excludedDate=records.length-selected.length;
  const groups=new Map(),characteristics=new Map();
  const create=(key,r)=>({key,code:r.code,name:r.name,characteristic:r.characteristic,complaints:[],count:0,stores:new Set(),manufacturers:new Set(),dates:new Map(),badQuantity:0,hasQuantity:false});
  const add=(g,r)=>{g.complaints.push(r);g.count++;if(r.storeCode)g.stores.add(r.storeCode);else if(r.store)g.stores.add('name:'+r.store);if(r.manufacturer)g.manufacturers.add(r.manufacturer);g.dates.set(r.date,(g.dates.get(r.date)||0)+1);if(r.badQuantity!==null){g.hasQuantity=true;g.badQuantity+=r.badQuantity;}};
  for(const r of selected){
    if(!groups.has(r.code))groups.set(r.code,create(r.code,r));add(groups.get(r.code),r);
    const k=r.code+'\u0000'+keyText(r.characteristic);if(!characteristics.has(k))characteristics.set(k,create(k,r));add(characteristics.get(k),r);
  }
  const sorted=a=>[...a].sort((x,y)=>y.count-x.count||x.code.localeCompare(y.code,'ru'));
  const matches=g=>!q||keyText(g.code+' '+g.name+' '+g.characteristic+' '+[...g.manufacturers].join(' ')+' '+g.complaints.map(r=>r.store+' '+r.storeCode+' '+r.comment).join(' ')).includes(q);
  const totalGroup=sorted(groups.values()),totalCharacteristics=sorted(characteristics.values());
  const thresh=Math.max(1,Math.floor(Number(threshold)||2));
  return {asOf,yesterday,selected,excludedDate,
    groups:totalGroup.filter(matches),characteristics:totalCharacteristics.filter(matches),
    mass:totalGroup.filter(g=>g.count>=thresh&&matches(g)),critical:totalCharacteristics.filter(g=>g.count>=thresh&&matches(g)),
    stats:{total:selected.length,today:selected.filter(r=>r.date===asOf).length,yesterday:selected.filter(r=>r.date===yesterday).length,
      products:groups.size,massProducts:totalGroup.filter(g=>g.count>=thresh).length,
      massIssues:totalCharacteristics.filter(g=>g.count>=thresh).length,
      stores:new Set(selected.map(r=>r.storeCode||r.store).filter(Boolean)).size,
      excludedDate,threshold:thresh}};
}
return {REQUIRED,text,dateISO,ruDate,previousDate,parseTable,merge,aggregate};
});
