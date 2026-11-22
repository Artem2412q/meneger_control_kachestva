/* MAGNIT RC — pallet inspection sample. All decisions are based on factual XLSX rows.
   Does not execute VBA. Deterministic random selection by saved seed for audit. */
(function(root,factory){const lib=factory();if(typeof module==='object'&&module.exports)module.exports=lib;root.MagnitPalletMacro=lib;})(typeof self==='undefined'?globalThis:self,function(){
'use strict';
const txt=x=>String(x==null?'':x).trim();
const norm=x=>txt(x).toLocaleUpperCase('ru-RU').replace(/Ё/g,'Е').replace(/\s+/g,' ');
const code=x=>txt(x).replace(/\.0+$/,'');
const num=x=>{const n=Number(txt(x).replace(/\s/g,'').replace(',','.'));return Number.isFinite(n)?n:NaN;};
const freshStart=/^(?:ТОМАТЫ?|ПОМИДОРЫ?|ОГУРЦЫ?|КАРТОФЕЛЬ|МОРКОВЬ|СВЕКЛА|КАПУСТА|БРОККОЛИ|ЛУК|ЧЕСНОК|ПЕРЕЦ|БАКЛАЖАНЫ?|КАБАЧКИ?|КАБАЧОК|ТЫКВА|РЕДИС|РЕДЬКА|ДАЙКОН|СЕЛЬДЕРЕЙ|САЛАТ|РУККОЛА|ШПИНАТ|УКРОП|ПЕТРУШКА|КИНЗА|БАЗИЛИК|ЩАВЕЛЬ|ЗЕЛЕНЬ|ЯБЛОКИ?|ГРУШИ?|БАНАНЫ?|МАНДАРИНЫ?|АПЕЛЬСИНЫ?|ЛИМОНЫ?|ЛАЙМЫ?|ГРЕЙПФРУТЫ?|ПОМЕЛО|КИВИ|ВИНОГРАД|СЛИВА|ПЕРСИКИ?|НЕКТАРИНЫ?|АБРИКОСЫ?|ХУРМА|ГРАНАТЫ?|АНАНАСЫ?|МАНГО|АВОКАДО|АРБУЗЫ?|ДЫНИ?|КЛУБНИКА|ЗЕМЛЯНИКА|МАЛИНА|ЕЖЕВИКА|ГОЛУБИКА|ЧЕРНИКА|СМОРОДИНА|КРЫЖОВНИК|ЧЕРЕШНЯ|ВИШНЯ|КЛЮКВА|БРУСНИКА|ИНЖИР|ФЕЙХОА|КОКОСЫ?|КУКУРУЗА)(?:\s|[:(]|$)/;
// Conservative exclusion: a product mentioning fruit as flavor/ingredient is not fresh produce.
const processed=/(?:СОК\b|НЕКТАР\s+\d|ПЮРЕ|ДЖЕМ|ВАРЕНЬ|ЗАМОРОЖ|МОРОЖЕН|КОНСЕРВ|МАРИН|СУШЕН|СУШЁН|КВАШЕН|ТУШЕН|ЧИПС|ПОЛУФАБРИКАТ|СОУС|КЕТЧУП|СУП\b|ПО-КОРЕЙСКИ|САЛАТ\s+(?:ЦЕЗАРЬ|ОЛИВЬЕ|ВИТАМИННЫЙ|ГРЕЧЕСКИЙ|КРАБОВЫЙ)|ФРИ\b|ЗАПРАВК|КОКТЕЙЛ|ЙОГУРТ|НАПИТОК|ВЯЛЕН|ОТВАРН|ВАРЕН|ВАРЁН|НАРЕЗАННЫЙ|ЖАРЕН|ПРОТЕР|ПРОТЁР|ПЕРЕТЕР|САХАРОМ|СИРОП|В СОБСТВЕННОМ СОКУ|ПАСТА|ЗАПЕЧЕН|ЗАПЕЧЁН|ПЕЧЕНЫЙ)/;
const REQ={number:['номер','номер палеты','№ палеты','палета'],code:['код','код товара'],name:['назв','наименование товара','наименование','товар'],mo:['мо'],mp:['мп']};
const OPT={active:['актив'],quantity:['колич','количество'],manufacturer:['производитель'],arrival:['№ прихода','номер прихода'],country:['страна']};
function normalizeHead(x){return norm(x).toLowerCase().replace(/\s+/g,' ').replace(/ё/g,'е').trim();}
function detectHeaders(values){
  const entries=values.map((v,i)=>[normalizeHead(v),i]); const find=aliases=>entries.find(([v])=>aliases.some(s=>v===s));
  const map={};for(const [k,aliases] of Object.entries({...REQ,...OPT})){const match=find(aliases);map[k]=match?match[1]:-1;}
  const missing=Object.keys(REQ).filter(k=>map[k]<0);
  if(missing.length)throw Error('В Excel не найдены обязательные колонки: '+missing.map(k=>REQ[k][0]).join(', ')+'.');
  return map;
}
function makePallet(row,fields,rowNo){
  const get=k=>fields[k]<0?'':row[fields[k]];
  return {number:code(get('number')),code:code(get('code')),name:txt(get('name')),
    mo:txt(get('mo')),mp:txt(get('mp')),active:txt(get('active')),
    quantity:fields.quantity<0?null:num(get('quantity')),
    manufacturer:txt(get('manufacturer')),arrival:txt(get('arrival')),country:txt(get('country')),rowNo};
}
function isFreshProduce(name){const t=norm(name);return freshStart.test(t)&&!processed.test(t);}
function isBuffer(mp,mode='12-16'){
 const t=norm(mp);if(!/(БАФ|БУФ)/.test(t))return false;
 if(mode==='all')return true;
 const match=t.match(/(?:БАФЕР|БУФЕР)\s*(\d{1,2})[-\s]/);
 return Boolean(match&&Number(match[1])>=12&&Number(match[1])<=16);
}
function validPallet(p,mode='12-16'){
 if(!p||!p.number||!p.code||!p.name||!p.mo||!p.mp)return false;
 if(p.active!==''&&!['1','ДА','TRUE'].includes(norm(p.active)))return false;
 if(p.quantity!==null&&(!Number.isFinite(p.quantity)||p.quantity<=0))return false;
 return isFreshProduce(p.name)&&isBuffer(p.mp,mode);
}
function supplierMatches(p,term){
 if(!txt(term))return true;
 const search=norm(term).replace(/[\s.\-–—«»"']/g,'');
 const fields=[p.manufacturer,p.arrival,p.country];
 if(fields.some(v=>norm(v).replace(/[\s.\-–—«»"']/g,'').includes(search)))return true;
 // Known user-provided alias: GT/ГТ = Гранд-Трейд; incoming numbers such as ГТАдг… are identifiers.
 const isGT=['ГТ','GT','ГРАНДТРЕЙД','GRANDTRADE'].includes(search);
 if(isGT&&(fields.some(v=>/^(?:ГТ|GT)АДГ/i.test(norm(v)))||norm(p.country).includes('ТУРЦИЯ')))return true;
 return false;
}
function compareStorage(a,b){
 const parts=s=>{const m=txt(s).match(/(\d{1,2})-(\d{1,3})-(\d{1,2})/);return m?[+m[1],+m[2],+m[3]]:[999,999,999];};
 const x=parts(a.mp),y=parts(b.mp);for(let i=0;i<3;i++)if(x[i]!==y[i])return x[i]-y[i];
 return String(a.number).localeCompare(String(b.number),'ru',{numeric:true});
}
function shuffled(input,random){const arr=[...input];for(let i=arr.length-1;i>0;i--){const j=Math.floor(random()*(i+1));[arr[i],arr[j]]=[arr[j],arr[i]];}return arr;}
function seeded(seed){let x=(Number(seed)>>>0)||0x9E3779B9;return ()=>{x^=x<<13;x^=x>>>17;x^=x<<5;return (x>>>0)/4294967296;};}
function build(pallets,{percent=3,supplier='',productCode='',zone='12-16',seed=123456789}={}){
 const pct=Number(percent);if(!Number.isFinite(pct)||pct<=0||pct>100)throw Error('Укажите процент выборки от 0,1 до 100.');
 const qualified=pallets.filter(p=>validPallet(p,zone));
 const qSupplier=txt(supplier);const filtered=qSupplier?qualified.filter(p=>supplierMatches(p,qSupplier)):qualified;
 if(!filtered.length)throw Error(qSupplier?'В исходном Excel нет подтверждённых овощей или фруктов по поставщику «'+qSupplier+'» в выбранных буферных зонах. Уточните поставщика или очистите поле.':'В файле нет подходящих палет: нужны фактические овощи/фрукты с остатком на буферных адресах.');
 const target=code(productCode);
 const hasOtherCodes=target?filtered.some(p=>p.code!==target):false;
 if(target&&!/^[0-9]+$/.test(target))throw Error('Код товара должен содержать только цифры.');
 const prioritized=target?filtered.filter(p=>p.code===target):[];
 if(target&&!prioritized.length)throw Error('Код '+target+' не найден среди подтверждённых палет овощей/фруктов для выбранных условий. Непроверенные товары не добавляются.');
 const size=Math.min(filtered.length,Math.max(1,Math.ceil(filtered.length*pct/100)));
 const rand=seeded(seed);let selected=[],selectedKey=new Set();
 const key=p=>[p.number,p.code,p.mp].join('\u0000');
 const add=p=>{const k=key(p);if(!selectedKey.has(k)){selectedKey.add(k);selected.push(p);}};
 const priorityQuota=target?Math.min(prioritized.length,size,Math.max(1,Math.min(5,Math.ceil(size*.25)))):0;
 shuffled(prioritized,rand).slice(0,priorityQuota).forEach(add);
 // Random draw across distinct item codes first: prevents large-volume apple SKUs dominating a small inspection round.
 const groups=new Map();for(const p of shuffled(filtered,rand)){
   if(selectedKey.has(key(p)))continue;
   if(target && p.code===target && hasOtherCodes)continue;
   if(!groups.has(p.code))groups.set(p.code,[]);
   groups.get(p.code).push(p);
 }
 for(const code of shuffled([...groups.keys()],rand)){
   if(selected.length>=size)break;
   const pool=groups.get(code);if(pool?.length)add(pool[0]);
 }
 if(selected.length<size){for(const p of shuffled(filtered,rand)){if(selected.length>=size)break;if(target && p.code===target && hasOtherCodes)continue;add(p);}}
 selected=selected.sort(compareStorage);
 return {selected,qualified:qualified.length,eligible:filtered.length,count:size,percent:pct,seed:Number(seed)>>>0,
   uniqueCodes:new Set(selected.map(p=>p.code)).size,priorityCount:selected.filter(p=>target&&p.code===target).length,
   supplier:qSupplier,productCode:target,zone,sourcePallets:pallets.length,remainingSupplierMissing:pallets.filter(p=>!p.manufacturer).length};
}
return {txt,norm,detectHeaders,makePallet,isFreshProduce,isBuffer,validPallet,supplierMatches,compareStorage,build};
});
