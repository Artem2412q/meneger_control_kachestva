/* Pallet inspector UI — v2.13. Keeps files in-memory, no backend. */
(function(){
'use strict';
const $=id=>document.getElementById(id), macro=window.MagnitPalletMacro;
if(!macro||!window.MagnitXlsxLite||!window.MagnitExcel){console.error('Не удалось загрузить модуль макроса.');return;}
const state={file:null,rows:[],report:null,download:null,loading:false,sourceName:''};
const status=(message,error=false)=>{const el=$('macroMessage');el.textContent=message;el.style.color=error?'#ac2937':'#526e60';};
const format=n=>Number(n||0).toLocaleString('ru-RU');
const cell=(tag,text,cls)=>{const el=document.createElement(tag);el.textContent=text==null?'':String(text);if(cls)el.className=cls;return el;};
function clearDownload(){if(state.download){URL.revokeObjectURL(state.download.url);state.download=null;}$('macroDownloadLink').hidden=true;}
function hideResult(){state.report=null;$('macroOutput').hidden=true;$('macroExport').disabled=true;$('macroReroll').disabled=true;clearDownload();}
function currentOptions(seed){return {percent:Number($('macroRate').value),productCode:$('macroCode').value.trim(),supplier:$('macroSupplier').value.trim(),zone:$('macroZoneMode').value,seed};}
function freshSeed(){if(globalThis.crypto?.getRandomValues){let bytes=new Uint32Array(1);crypto.getRandomValues(bytes);return bytes[0]||1;}return (Math.random()*0xffffffff)>>>0||1;}
function render(result){
 $('macroOutput').hidden=false;$('macroExport').disabled=false;$('macroReroll').disabled=false;
 $('macroSummary').textContent=`Отобрано ${format(result.selected.length)} из ${format(result.eligible)} подходящих палет · ${result.percent}% · сортировка по МП`;
 const kpis=[['К проверке',result.selected.length],['Подходящих палет',result.eligible],['Разных товаров',result.uniqueCodes],['По приоритетному коду',result.priorityCount]];
 const metrics=$('macroKpis');metrics.replaceChildren();for(const [label,value] of kpis){const div=cell('div','','macro-kpi');div.append(cell('span',label),cell('strong',format(value)));metrics.append(div);}
 const tbody=$('macroRows');tbody.replaceChildren();
 for(const p of result.selected.slice(0,250)){
  const tr=document.createElement('tr');if(result.productCode&&p.code===result.productCode)tr.className='macro-highlight';
  for(const text of [p.number,p.code,p.name,p.mo,p.mp,''])tr.append(cell('td',text));tbody.append(tr);
 }
 if(result.selected.length>250){const tr=document.createElement('tr');const td=cell('td',`Показаны первые 250 строк. В Excel попадут все ${result.selected.length} палет.`);td.colSpan=6;tr.append(td);tbody.append(tr);}
 const notes=[];if(result.productCode)notes.push(`Приоритетный код ${result.productCode}: ${result.priorityCount} палет.`);
 if(result.qualified>result.eligible)notes.push(`Исключено по фильтру поставщика: ${result.qualified-result.eligible}.`);
 notes.push('Поставщик в исходном Excel может быть не указан: отсутствующие сведения не восстанавливаются автоматически.');
 $('macroWarnings').textContent=notes.join(' ');
 $('macroAuditInfo').textContent=`Случайное зерно: ${result.seed} · Зона: ${result.zone==='all'?'все буферные':'12–16'} · Полный список — в XLSX`;
}
function doBuild(seed=freshSeed()){
 if(!state.rows.length){status('Сначала загрузите Excel с палетами.',true);return;}
 try{
  const result=macro.build(state.rows,currentOptions(seed));state.report=result;clearDownload();render(result);
  status(`Готово: ${format(result.count)} палет из ${format(result.eligible)}. Можно скачать Excel по образцу менеджера.`);
 }catch(e){hideResult();status(e?.message||String(e),true);}
}
async function parseFile(file){
 if(!file)return;
 if(!/\.xlsx$/i.test(file.name)){hideResult();status('Нужен исходный Excel в формате .xlsx. Файлы .xls/.xlsm здесь не обрабатываются.',true);return;}
 hideResult();state.file=file;state.rows=[];state.loading=true;$('macroBuild').disabled=true;
 $('macroFileStatus').textContent='Чтение файла '+file.name+'…';$('macroZone').classList.remove('valid');
 status('Читаю Excel, проверяю код, номер палеты, МО, МП и название товара…');
 try{
  const book=await MagnitXlsxLite.open(await file.arrayBuffer());if(!book.names?.length)throw Error('В Excel отсутствуют листы.');
  let parsed=null,lastError=null;
  for(const sheetName of book.names){
   let fields=null,head=false,number=0;const found=[],seen=new Set();
   await book.eachRow(sheetName,raw=>{
    number++;
    const row=raw.map(c=>c?.v??'');
    if(!head){
      if(number>35)return;
      try{fields=macro.detectHeaders(row);head=true;}catch(e){lastError=e;}
      return;
    }
    const p=macro.makePallet(row,fields,number);
    if(!p.number||!p.code||!p.name) return;
    const key=p.number+'\u0001'+p.code+'\u0001'+p.mp;
    if(seen.has(key))return;
    seen.add(key);found.push(p);
   });
   if(head&&found.length){parsed=found;break;}
  }
  if(!parsed?.length)throw lastError||Error('Не удалось найти таблицу с палетами.');
  state.rows=parsed;state.sourceName=file.name;
  $('macroFileStatus').textContent=`${file.name} · ${format(parsed.length)} реальных палет`;
  $('macroZone').classList.add('valid');
  $('macroSourceInfo').textContent=`Источник: ${file.name}. Палет прочитано: ${format(parsed.length)}. Автоматически исключаются готовые и прочие продукты.`;
  $('macroBuild').disabled=false;
  doBuild();
 }catch(e){state.rows=[];state.file=null;hideResult();$('macroFileStatus').textContent='Ошибка чтения XLSX';$('macroZone').classList.add('invalid');status(e?.message||String(e),true);}
 finally{state.loading=false;}
}
function save(){
 const r=state.report;if(!r?.selected?.length){status('Нечего выгружать: сначала сформируйте фактическую выборку.',true);return;}
 const data=r.selected.map(p=>[p.number,String(p.code),p.name,p.mo,p.mp,'','']);
 const description=[
  ['Параметр','Значение'],['Источник',state.sourceName],['Дата формирования',new Date().toLocaleString('ru-RU')],
  ['Доля выборки',r.percent+'%'],['Буферные зоны',r.zone==='12-16'?'12–16':'Все буферные зоны'],
  ['Поставщик (фильтр)',r.supplier||'Не задан'],['Приоритетный код',r.productCode||'Не задан'],
  ['Кодовые позиции в выборке',r.uniqueCodes],['Палет из файла',r.sourcePallets],
  ['Подходящих свежих овощей/фруктов',r.qualified],['После фильтра поставщика',r.eligible],
  ['Отобрано палет',r.selected.length],['Приоритетный код (палет)',r.priorityCount],
  ['Зерно случайного отбора',String(r.seed)],
  ['Принцип отбора','Случайные коды товаров, затем палеты; приоритетный код — в первую очередь'],
  ['Принцип сортировки','По адресу МП, по возрастанию'],
  ['Данные для проверки','Восстановлены исключительно из исходного файла; неподтверждённые позиции исключены'],
  ['Результат проверки','Пустой — заполнить после физического осмотра']
 ];
 const sheets=[
  {name:'Проверка палет',type:'pallet',headers:['номер','код','назв','МО','МП','','Результат проверки'],rows:data,widths:[13,16,59,18,27,9,26]},
  {name:'Протокол отбора',headers:description[0],rows:description.slice(1),widths:[42,93]}
 ];
 try{
  clearDownload();const name='MAGNIT_MACRO_PALLETS_'+new Date().toISOString().slice(0,10)+'.xlsx';
  const out=MagnitExcel.save(sheets,name);state.download=out;const a=$('macroDownloadLink');a.href=out.url;a.download=out.filename;a.hidden=false;
  status(`Excel сформирован: ${r.selected.length} палет, лист проверки и протокол отбора. Если загрузка заблокирована, используйте резервную ссылку.`);
 }catch(e){status('Не удалось создать XLSX: '+(e?.message||e),true);}
}
$('macroFile').addEventListener('change',e=>{parseFile(e.target.files?.[0]);e.target.value='';});
for(const event of ['dragenter','dragover'])$('macroZone').addEventListener(event,e=>{e.preventDefault();$('macroZone').classList.add('dragging');});
for(const event of ['dragleave','drop'])$('macroZone').addEventListener(event,e=>{e.preventDefault();$('macroZone').classList.remove('dragging');});
$('macroZone').addEventListener('drop',e=>{e.preventDefault();parseFile(e.dataTransfer?.files?.[0]);});
$('macroZone').addEventListener('keydown',e=>{if(e.target===$('macroZone')&&(e.key==='Enter'||e.key===' ')){e.preventDefault();$('macroFile').click();}});
$('macroBuild').addEventListener('click',()=>doBuild());
$('macroReroll').addEventListener('click',()=>doBuild());
$('macroExport').addEventListener('click',save);
for(const id of ['macroRate','macroZoneMode','macroSupplier','macroCode'])$(id).addEventListener('change',()=>{if(state.rows.length)doBuild();});
})();
