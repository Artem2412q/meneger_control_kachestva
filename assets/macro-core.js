/* MAGNIT RC — pallet inspection selection 2.17: fresh FROV including all identified fresh herbs and leafy greens.
   The source is the uploaded report; no synthetic pallets or post-inspection comments.
   Two independently adjustable sampling rates, reproducible random seed.
*/
(function(root,factory){const api=factory();if(typeof module==='object'&&module.exports)module.exports=api;root.MagnitPalletMacro=api;})(typeof self==='undefined'?globalThis:self,function(){
'use strict';
const txt=x=>String(x==null?'':x).trim();
const norm=x=>txt(x).toLocaleUpperCase('ru-RU').replace(/Ё/g,'Е').replace(/\s+/g,' ');
const code=x=>txt(x).replace(/\.0+$/,'');
const num=x=>{if(x===null||x===undefined||txt(x)==='')return NaN;const n=Number(txt(x).replace(/\s/g,'').replace(',','.'));return Number.isFinite(n)?n:NaN;};
const freshStart=/^(?:ТОМАТЫ?|ПОМИДОРЫ?|ОГУРЦЫ?|КАРТОФЕЛЬ|МОРКОВЬ|СВЕКЛА|КАПУСТА|БРОККОЛИ|ЛУК|ЧЕСНОК|ПЕРЕЦ|БАКЛАЖАНЫ?|КАБАЧКИ?|КАБАЧОК|ТЫКВА|РЕДИС|РЕДЬКА|ДАЙКОН|СЕЛЬДЕРЕЙ|САЛАТ|РУККОЛА|ШПИНАТ|УКРОП|ПЕТРУШКА|КИНЗА|БАЗИЛИК|ЩАВЕЛЬ|ЗЕЛЕНЬ|ЯБЛОКИ?|ГРУШИ?|БАНАНЫ?|МАНДАРИНЫ?|АПЕЛЬСИНЫ?|ЛИМОНЫ?|ЛАЙМЫ?|ГРЕЙПФРУТЫ?|ПОМЕЛО|КИВИ|ВИНОГРАД|СЛИВА|ПЕРСИКИ?|НЕКТАРИНЫ?|АБРИКОСЫ?|ХУРМА|ГРАНАТЫ?|АНАНАСЫ?|МАНГО|АВОКАДО|АРБУЗЫ?|ДЫНИ?|КЛУБНИКА|ЗЕМЛЯНИКА|МАЛИНА|ЕЖЕВИКА|ГОЛУБИКА|ЧЕРНИКА|СМОРОДИНА|КРЫЖОВНИК|ЧЕРЕШНЯ|ВИШНЯ|КЛЮКВА|БРУСНИКА|ИНЖИР|ФЕЙХОА|КОКОСЫ?|КУКУРУЗА)(?:\s|[:(/-]|$)/;
// ФРЕШ/ФРОВ: свежая зелень и листовые салаты — отдельные группы от овощей и фруктов.
// Названия из реальных выгрузок включают «САЛАТНЫЙ МИКС», «УКРОП/ПЕТРУШКА», «САЛАТ Фриллис».
const freshHerbsStart=/^(?:МЯТА|МЕЛИССА|ЭСТРАГОН|ТАРХУН|РОЗМАРИН|ТИМЬЯН|ЧАБРЕЦ|ШАЛФЕЙ|МАЙОРАН|ОРЕГАНО|ДУШИЦА|МИКРОЗЕЛЕНЬ|МИКРОЗЕЛЕНЬЮ|ЧЕРЕМША|РУКОЛА|РУККОЛА|КРЕСС(?:[-\s]+САЛАТ)?|МАНГОЛЬД|МИЗУНА|РАДИЧИО|РАДИККИО|ПАК[-\s]?ЧОЙ|БОК[-\s]?ЧОЙ|КЕЙЛ|ШНИТТ[-\s]?ЛУК|ЛИСТЬЯ\s+САЛАТА|ЛИСТОВОЙ\s+САЛАТ|ЛУК\s+(?:ЗЕЛЕНЫЙ|ЗЕЛЕНЫЙ|ПЕРО)|ЗЕЛЕНЫЙ\s+ЛУК|ЗЕЛЕНЬ\s+СВЕЖАЯ|СВЕЖАЯ\s+ЗЕЛЕНЬ)(?:\s|[:(/-]|$)/;
const freshMixStart=/^(?:САЛАТНЫЙ\s+МИКС|САЛАТ[-\s]+МИКС|МИКС\s+(?:САЛАТОВ|САЛАТНЫЙ|ЗЕЛЕНИ)|СМЕСЬ\s+(?:САЛАТОВ|ЗЕЛЕНИ)|НАБОР\s+(?:ЗЕЛЕНИ|САЛАТНЫЙ)|АССОРТИ\s+ЗЕЛЕНИ|ЗЕЛЕНЬ\s+МИКС|БЭБИ[-\s]?САЛАТ)(?:\s|[:(/-]|$)/;
// Другие свежие ФРОВ: шампиньоны/грибы и позиции, которые отсутствовали в первом перечне.
const freshOtherFrovStart=/^(?:ГРИБЫ|ШАМПИНЬОНЫ|ВЕШЕНКИ|ЛИСИЧКИ|ОПЯТА|РЕДИСКА|АЙВА|ПАПАЙЯ|ПИТАХАЙЯ|МАРАКУЙЯ|ПАСТЕРНАК|РЕВЕНЬ|КОЛЬРАБИ|СПАРЖА|АРТИШОК|ФИЗАЛИС)(?:\s|[:(/-]|$)/;
const processed=/(?:СОК\b|НЕКТАР\s+\d|ПЮРЕ|ДЖЕМ|ВАРЕНЬ|ЗАМОРОЖ|МОРОЖЕН|КОНСЕРВ|(?:^|[^А-ЯЁ])МАРИН|СУШЕН|СУШЁН|КВАШЕН|ТУШЕН|ЧИПС|ПОЛУФАБРИКАТ|СОУС|КЕТЧУП|СУП\b|ПО-КОРЕЙСКИ|САЛАТ\s+(?:ЦЕЗАРЬ|ОЛИВЬЕ|ВИТАМИННЫЙ|ГРЕЧЕСКИЙ|КРАБОВЫЙ)|ФРИ\b|ЗАПРАВК|КОКТЕЙЛ|ЙОГУРТ|НАПИТОК|ЛИМОНАД|ГАЗИРОВ|ВЯЛЕН|ОТВАРН|ВАРЕН|ВАРЁН|НАРЕЗАННЫЙ|ЖАРЕН|ПРОТЕР|ПРОТЁР|ПЕРЕТЕР|САХАРОМ|СИРОП|СОЛЕН|СОЛЁН|ГОРОШКОМ|ЛЕПЕСТК|МОРСКОЙ КАПУСТЫ|В\/МАЙОН|СПЕЦИ|В СОБСТВЕННОМ СОКУ|ПАСТА|МОЛОТ|ПРИПРАВ|ПОРОШОК|МАЙОН|ПАШТЕТ|ЗАПЕЧЕН|ЗАПЕЧЁН|ПЕЧЕНЫЙ)/;
const REQ={number:['номер','номер палеты','№ палеты','палета'],code:['код','код товара'],name:['назв','наименование товара','наименование','товар'],mp:['мп']};
const OPT={mo:['мо'],active:['актив'],quantity:['колич','количество'],manufacturer:['производитель'],arrival:['№ прихода','номер прихода'],country:['страна'],formed:['дата формирования'],cep:['цеп','sscc/центр.пэ','sscc','центр.пэ'],avgWeight:['средний вес коробки'],defect:['процент брака'],defectDate:['дата установки (% брака)'],note:['примечание'],days:['дней на хранении']};
const INSPECTION_HEADERS=['актив','Дата формирования','номер','Дней на хранении','Сегодня','ЦЕП','Снять на проверку','код','назв','МО','МП','колич','Средний вес коробки','Процент брака','Дата установки (% брака)','Примечание','Результат проверки','Процент бракапо итогам разбраковки','Комментарий (ОКК)'];
function normalizeHead(x){return norm(x).toLowerCase().replace(/\s+/g,' ').trim();}
function detectHeaders(values){
 const arr=values.map((v,i)=>[normalizeHead(v),i]);const found=aliases=>arr.find(([s])=>aliases.includes(s));const map={};
 for(const [key,aliases] of Object.entries({...REQ,...OPT})){const q=found(aliases);map[key]=q?q[1]:-1;}
 const missing=Object.keys(REQ).filter(k=>map[k]<0);if(missing.length)throw Error('В выгрузке отсутствуют обязательные столбцы: '+missing.map(k=>REQ[k][0]).join(', '));
 if(map.formed<0&&map.days<0)throw Error('Для двух возрастных выборок нужен столбец «Дата формирования» или «Дней на хранении».');
 return map;
}
function makePallet(row,cols,rowNo){const get=k=>cols[k]===undefined||cols[k]<0?'':row[cols[k]];
 return {number:code(get('number')),code:code(get('code')),name:txt(get('name')),mo:txt(get('mo')),mp:txt(get('mp')),
 active:txt(get('active')),quantity:cols.quantity<0?null:num(get('quantity')),manufacturer:txt(get('manufacturer')),
 arrival:txt(get('arrival')),country:txt(get('country')),formed:txt(get('formed')),cep:txt(get('cep')),
 avgWeight:txt(get('avgWeight')),defect:txt(get('defect')),defectDate:txt(get('defectDate')),
 note:txt(get('note')),sourceDays:cols.days<0?NaN:num(get('days')),rowNo};
}
function parseDate(x){
 const raw=txt(x);if(!raw)return null;
 if(/^\d+(?:\.\d+)?$/.test(raw)){const n=Number(raw);if(n>30000&&n<100000)return Math.round((n-25569)*86400000);}
 let m=raw.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2,4})(?:\D|$)/);
 if(m){let yy=Number(m[3]);if(yy<100)yy+=yy<50?2000:1900;const d=Date.UTC(yy,Number(m[2])-1,Number(m[1]));if(new Date(d).getUTCDate()===Number(m[1])&&new Date(d).getUTCMonth()===Number(m[2])-1)return d;return null;}
 m=raw.match(/^(\d{4})-(\d{2})-(\d{2})/);if(m){const d=Date.UTC(+m[1],+m[2]-1,+m[3]);if(new Date(d).getUTCDate()===+m[3])return d;}
 return null;
}
function isoToDate(x){return parseDate(x);}
function daysOnSite(p,asOfDate){const end=isoToDate(asOfDate);if(end===null)return null;
 const formed=parseDate(p.formed);if(formed!==null){const diff=Math.floor((end-formed)/86400000);return diff>=0?diff:null;}
 if(Number.isFinite(p.sourceDays)&&p.sourceDays>=0)return Math.floor(p.sourceDays);
 return null;
}
function isFreshProduce(name){
 const t=norm(name);
 if(!t||processed.test(t))return false;
 return freshStart.test(t)||freshHerbsStart.test(t)||freshMixStart.test(t)||freshOtherFrovStart.test(t);
}
function isBuffer(mp,mode='all'){const t=norm(mp);if(!/(БАФ|БУФ)/.test(t))return false;
 if(mode==='all')return true;const m=t.match(/(?:БАФЕР|БУФЕР)\s*(\d{1,2})[-\s]/);return !!(m&&Number(m[1])>=12&&Number(m[1])<=16);
}
function validPallet(p,zone='all'){
 if(!p?.number||!p.code||!p.name||!p.mp)return false;
 if(p.active!==''&&!['1','ДА','TRUE'].includes(norm(p.active)))return false;
 if(p.quantity!==null&&(!Number.isFinite(p.quantity)||p.quantity<=0))return false;
 return isFreshProduce(p.name)&&isBuffer(p.mp,zone);
}
function supplierMatches(p,term){if(!txt(term))return true;const search=norm(term).replace(/[\s.\-–—«»"']/g,'');const fields=[p.manufacturer,p.arrival,p.country];
 if(fields.some(v=>norm(v).replace(/[\s.\-–—«»"']/g,'').includes(search)))return true;
 const gt=['ГТ','GT','ГРАНДТРЕЙД','GRANDTRADE'].includes(search);
 return gt&&(fields.some(v=>/^(?:ГТ|GT)АДГ/i.test(norm(v)))||norm(p.country).includes('ТУРЦИЯ'));
}
function compareStorage(a,b){const parts=s=>{const m=txt(s).match(/(\d{1,2})-(\d{1,3})-(\d{1,2})/);return m?[+m[1],+m[2],+m[3]]:[999,999,999]};const x=parts(a.mp),y=parts(b.mp);for(let i=0;i<3;i++)if(x[i]!==y[i])return x[i]-y[i];return a.number.localeCompare(b.number,'ru',{numeric:true});}
function shuffled(arr,random){const a=[...arr];for(let i=a.length-1;i>0;i--){const j=Math.floor(random()*(i+1));[a[i],a[j]]=[a[j],a[i]];}return a;}
function seeded(seed){let x=(Number(seed)>>>0)||0x9E3779B9;return ()=>{x^=x<<13;x^=x>>>17;x^=x<<5;return (x>>>0)/4294967296;};}
function sampleGroup(pool,count,random,priorityCode,priorityRemaining){const picked=[],seen=new Set();const key=p=>[p.number,p.code,p.mp].join('\x00');const add=p=>{const k=key(p);if(!seen.has(k)){seen.add(k);picked.push(p);}};
 if(priorityCode&&priorityRemaining>0)for(const p of shuffled(pool.filter(p=>p.code===priorityCode),random).slice(0,Math.min(count,priorityRemaining)))add(p);
 const hasOthers=priorityCode&&pool.some(q=>q.code!==priorityCode);
 const groups=new Map();for(const p of shuffled(pool,random)){if(seen.has(key(p)))continue;if(priorityCode&&p.code===priorityCode&&hasOthers)continue;if(!groups.has(p.code))groups.set(p.code,[]);groups.get(p.code).push(p);}
 for(const name of shuffled([...groups.keys()],random)){if(picked.length>=count)break;add(groups.get(name)[0]);}
 if(picked.length<count)for(const p of shuffled(pool,random)){if(picked.length>=count)break;add(p);}
 return picked;
}
function build(pallets,{recentPercent=1.5,olderPercent=1.5,recentDays=3,asOfDate='',supplier='',productCode='',zone='all',seed=123456789,totalPalettes=null,priorityQty=5}={}){
 const a=Number(recentPercent),b=Number(olderPercent),cutoff=Number(recentDays),date=isoToDate(asOfDate);
 if(date===null)throw Error('Укажите корректную дату проверки.');
 if(!Number.isFinite(a)||a<=0||a>100||!Number.isFinite(b)||b<=0||b>100)throw Error('Каждый процент должен быть больше 0 и не превышать 100.');
 if(!Number.isInteger(cutoff)||cutoff<0||cutoff>365)throw Error('Граница новых палет должна быть в пределах от 0 до 365 дней.');
 const product=code(productCode);if(product&&!/^\d+$/.test(product))throw Error('Код товара должен содержать только цифры.');
 const allBuffers=pallets.filter(p=>p?.number&&p.mp&&isBuffer(p.mp,'all'));
 const fresh=allBuffers.filter(p=>validPallet(p,zone));
 const eligible=supplier?fresh.filter(p=>supplierMatches(p,supplier)):fresh;
 if(supplier&&!eligible.length)throw Error('По поставщику «'+supplier+'» нет подтверждённых палет ФРЕШ/ФРОВ. Уточните фильтр.');
 if(product&&!eligible.some(p=>p.code===product))throw Error('Код '+product+' не найден среди подходящих палет из загруженного Excel.');
 const recent=[],older=[],unknown=[];
 for(const p of eligible){const days=daysOnSite(p,asOfDate);if(days===null)unknown.push(p);else if(days<=cutoff)recent.push(p);else older.push(p);}
 const baseR=Math.min(recent.length,Math.ceil(recent.length*a/100));
 const baseO=Math.min(older.length,Math.ceil(older.length*b/100));
 let quotaRecent=baseR,quotaOlder=baseO;
 if(totalPalettes!==null && totalPalettes!=='' && totalPalettes!==undefined){
  const requested=Number(totalPalettes);
  if(!Number.isInteger(requested)||requested<1||requested>100000)throw Error('Количество палет должно быть целым числом от 1 до 100 000.');
  const target=Math.min(requested,recent.length+older.length);
  const weight=(baseR+baseO)?baseR/(baseR+baseO):recent.length/(recent.length+older.length||1);
  quotaRecent=Math.min(recent.length,Math.round(target*weight));
  quotaOlder=Math.min(older.length,target-quotaRecent);
  // Дозаполнение по фактической доступности, без превышения общего количества.
  while(quotaRecent+quotaOlder<target){
   if(quotaRecent<recent.length && (quotaOlder===older.length || quotaRecent/(recent.length||1)<=quotaOlder/(older.length||1)))quotaRecent++;
   else if(quotaOlder<older.length)quotaOlder++;
   else if(quotaRecent<recent.length)quotaRecent++;
   else break;
  }
 }
 const priorityRequested=Number(priorityQty);
 if(!Number.isInteger(priorityRequested)||priorityRequested<0||priorityRequested>100000)throw Error('Количество приоритетных палет должно быть целым числом от 0 до 100 000.');
 // Сохраняем общий объём и по возможности перемещаем квоту между возрастными группами,
 // чтобы приоритетный код попал в выборку в указанном количестве.
 if(product&&priorityRequested>0){
  const pr=recent.filter(p=>p.code===product).length,po=older.filter(p=>p.code===product).length;
  const required=Math.min(priorityRequested,quotaRecent+quotaOlder,pr+po);
  while(Math.min(pr,quotaRecent)+Math.min(po,quotaOlder)<required){
   if(po>quotaOlder && quotaRecent>Math.min(pr,quotaRecent) && quotaOlder<older.length){quotaRecent--;quotaOlder++;}
   else if(pr>quotaRecent && quotaOlder>Math.min(po,quotaOlder) && quotaRecent<recent.length){quotaOlder--;quotaRecent++;}
   else break;
  }
 }
 const rng=seeded(seed);let limit=product?priorityRequested:0;
 const selectedRecent=sampleGroup(recent,quotaRecent,rng,product,limit);limit-=selectedRecent.filter(p=>p.code===product).length;
 const selectedOlder=sampleGroup(older,quotaOlder,rng,product,limit);
 selectedRecent.sort(compareStorage);selectedOlder.sort(compareStorage);
 const selected=[...selectedRecent,...selectedOlder];const toKeys=new Set(selected.map(p=>p.number+'\x00'+p.code+'\x00'+p.mp));
 const allBufferKeys=new Set(allBuffers.map(p=>p.number+'\x00'+p.code+'\x00'+p.mp));
 return {selected,selectedRecent,selectedOlder,recent,older,unknown,allBuffers,eligibleCount:eligible.length,
 sourcePallets:pallets.length,bufferCount:allBuffers.length,bufferUnique:allBufferKeys.size,
 recentPercent:a,olderPercent:b,recentDays:cutoff,asOfDate,zone,supplier:txt(supplier),productCode:product,seed:Number(seed)>>>0, totalPalettes:totalPalettes==null||totalPalettes===''?null:Number(totalPalettes),priorityRequested:product?priorityRequested:0,
 priorityCount:selected.filter(p=>p.code===product&&!!product).length,uniqueCodes:new Set(selected.map(p=>p.code)).size,
 selectedUnique:toKeys.size};
}
function formatRussianDate(iso){const m=txt(iso).match(/^(\d{4})-(\d{2})-(\d{2})$/);return m?`${m[3]}.${m[2]}.${m[1]}`:iso;}
function sourceNumber(value){const n=num(value);return Number.isFinite(n)?n:txt(value);}
function inspectionRow(p,asOf){const age=daysOnSite(p,asOf);return [sourceNumber(p.active),p.formed,sourceNumber(p.number),age===null?'':age,formatRussianDate(asOf),p.cep,'Снять на проверку',String(p.code),p.name,p.mo,p.mp,sourceNumber(p.quantity),sourceNumber(p.avgWeight),sourceNumber(p.defect),p.defectDate,p.note,'','',''];}
return {txt,norm,detectHeaders,makePallet,parseDate,daysOnSite,isFreshProduce,isBuffer,validPallet,supplierMatches,compareStorage,INSPECTION_HEADERS,build,inspectionRow,formatRussianDate};
});
