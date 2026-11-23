/* MAGNIT RC — 2.17: fresh FROV with greens; original two XLSX templates unchanged. */
(function(){
'use strict';
const $=id=>document.getElementById(id), macro=window.MagnitPalletMacro;
if(!macro||!window.MagnitXlsxLite||!window.MagnitExcel){console.error('Не удалось загрузить модуль макроса.');return;}
const state={files:[],rows:new Map(),report:null,downloads:{},loading:false};
const status=(message,error=false)=>{const el=$('macroMessage');el.textContent=message;el.style.color=error?'#aa2937':'#526e60';};
const format=n=>Number(n||0).toLocaleString('ru-RU');
const cell=(tag,text,cls)=>{const el=document.createElement(tag);el.textContent=text==null?'':String(text);if(cls)el.className=cls;return el;};
const getRows=()=>[...state.rows.values()];
function clearDownload(){for(const out of Object.values(state.downloads)){if(out?.url)URL.revokeObjectURL(out.url);}state.downloads={};$('macroDownloadLink').hidden=true;$('macroRequestDownloadLink').hidden=true;}
function hideResult(){state.report=null;$('macroOutput').hidden=true;$('macroExport').disabled=true;$('macroExportRequest').disabled=true;$('macroReroll').disabled=true;clearDownload();}
function freshSeed(){if(globalThis.crypto?.getRandomValues){const a=new Uint32Array(1);crypto.getRandomValues(a);return a[0]||1;}return (Math.random()*0xffffffff)>>>0||1;}
function options(seed){return {recentPercent:Number($('macroRecentRate').value),olderPercent:Number($('macroOlderRate').value),recentDays:Number($('macroRecentDays').value),asOfDate:$('macroAsOf').value,productCode:$('macroCode').value.trim(),totalPalettes:$('macroTotalPallets').value.trim()===''?null:Number($('macroTotalPallets').value),priorityQty:Number($('macroPriorityQty').value),supplier:$('macroSupplier').value.trim(),zone:$('macroZoneMode').value,seed};}
function render(r){
 $('macroOutput').hidden=false;$('macroExport').disabled=false;$('macroExportRequest').disabled=false;$('macroReroll').disabled=false;
 $('macroSummary').textContent=`Недавно: ${format(r.selectedRecent.length)} (${r.recentPercent}%) · Давно: ${format(r.selectedOlder.length)} (${r.olderPercent}%) · Всего ${format(r.selected.length)} палет · «Лист1»: ${format(r.bufferCount)} буферов`;
 const kpis=[['Недавно / отбор',r.selectedRecent.length],['Давно / отбор',r.selectedOlder.length],['Буферов в Лист1',r.bufferCount],['Всего к спуску',r.selected.length]];
 const box=$('macroKpis');box.replaceChildren();for(const [label,value] of kpis){const d=cell('div','','macro-kpi');d.append(cell('span',label),cell('strong',format(value)));box.append(d);}
 const tbody=$('macroRows');tbody.replaceChildren();let count=0;
 for(const [label,group] of [['Недавно',r.selectedRecent],['Давно',r.selectedOlder]]){
   for(const p of group){if(count++>=250)break;const tr=document.createElement('tr');if(r.productCode&&p.code===r.productCode)tr.className='macro-highlight';
    for(const t of [label,p.number,p.code,p.name,p.mo,p.mp,macro.daysOnSite(p,r.asOfDate),''])tr.append(cell('td',t));tbody.append(tr);
   }
 }
 if(r.selected.length>250){const tr=document.createElement('tr');const td=cell('td',`Первые 250 из ${format(r.selected.length)} палет. В Excel будут все.`);td.colSpan=8;tr.append(td);tbody.append(tr);}
 const warnings=[];
 if(!r.recent.length)warnings.push('В группе «Недавно на РЦ» нет подтверждённых палет.');
 if(!r.older.length)warnings.push('В группе «Давно на РЦ» нет подтверждённых палет.');
 if(r.unknown.length)warnings.push(`У ${format(r.unknown.length)} подходящих палет не удалось определить возраст — они остались только в накопительном списке.`);
 if(r.productCode){warnings.push(`Приоритетный код ${r.productCode}: ${format(r.priorityCount)} из запрошенных ${format(r.priorityRequested)} палет.`);if(r.priorityCount<r.priorityRequested)warnings.push('Приоритет не достиг запрошенного количества: недостаточно подходящих палет либо общая выборка меньше приоритета.');} if(r.totalPalettes!==null&&r.selected.length<r.totalPalettes)warnings.push(`Запрошено ${format(r.totalPalettes)} палет, найдено только ${format(r.selected.length)} подходящих.`);
 warnings.push('Дата формирования не всегда совпадает с поступлением. Поля осмотра в Excel пустые.');
 $('macroWarnings').textContent=warnings.join(' ');
 $('macroAuditInfo').textContent=`Рандомное зерно: ${r.seed} · Граница: ${r.recentDays} дн. · Источников: ${state.files.length} · Дата: ${r.asOfDate}`;
}
function doBuild(seed=freshSeed()){
 if(!state.rows.size){status('Сначала загрузите Excel с палетами.',true);return;}
 try{const r=macro.build(getRows(),options(seed));state.report=r;clearDownload();render(r);
  status(`Готово: ${r.selectedRecent.length} недавних + ${r.selectedOlder.length} старых палет. В «Лист1» — ${format(r.bufferCount)} фактических буферов.`);
 }catch(e){hideResult();status(e?.message||String(e),true);}
}
function updateFileInfo(){const files=state.files;const n=state.rows.size;
 $('macroFileStatus').textContent=files.length?`${files.length} файл(ов) · ${format(n)} палет в накоплении`:'Выберите или перетащите Excel';
 $('macroSourceInfo').textContent=files.length?`Накоплено ${format(n)} уникальных палет по ${files.length} файлу(ам). Повторная запись того же номера и кода обновляется последней загрузкой.`:'Данные находятся в памяти этого окна браузера, на сервер не отправляются.';
 $('macroZone').classList.toggle('valid',files.length>0);
 $('macroBuild').disabled=!n;
}
async function parseOne(file){
 if(!/\.xlsx$/i.test(file.name))throw Error('Файл '+file.name+' не .xlsx. Загрузите стандартную складскую выгрузку.');
 const book=await MagnitXlsxLite.open(await file.arrayBuffer());if(!book.names?.length)throw Error('В файле '+file.name+' нет листов.');
 let parsed=null,lastError=null;
 for(const sheetName of book.names){let cols=null,head=false,line=0;const found=[];
  await book.eachRow(sheetName,raw=>{line++;const row=raw.map(c=>c?.v??'');
   if(!head){if(line>35)return;try{cols=macro.detectHeaders(row);head=true;}catch(e){lastError=e;}return;}
   const p=macro.makePallet(row,cols,line);if(!p.number||!p.mp||!macro.isBuffer(p.mp,'all'))return;
   found.push(p);
  });
  if(head&&found.length){parsed=found;break;}
 }
 if(!parsed?.length)throw lastError||Error('В '+file.name+' не найдены буферные палеты.');
 return parsed;
}
async function parseFiles(files){if(!files?.length)return;
 state.loading=true;hideResult();$('macroBuild').disabled=true;status('Читаю Excel и накапливаю все буферные адреса…');
 let success=0;const errors=[];
 for(const file of files){try{const parsed=await parseOne(file);
    // Newest uploaded state supersedes previous records with the same pallet number/code.
    for(const p of parsed){const k=p.number+'\x00'+p.code;state.rows.set(k,p);}
    state.files.push(file.name);success++;
  }catch(e){errors.push(e?.message||String(e));}
 }
 state.loading=false;updateFileInfo();
 if(success){doBuild();if(errors.length)status(`Загружено ${success} файл(ов). Ошибки: ${errors.join(' ')}`,true);}
 else status(errors.join(' ')||'Не получилось прочитать файлы.',true);
}
// Строго два файла, по структуре переданных пользователем шаблонов.
// №1: MAGNIT_MACRO_PALLETS_2026-11-22.xlsx — 2 листа.
// №2: Запрос на Спуск паллет 22.11.2026.xlsx — 4 листа.
const REQUEST_HEADERS=macro.INSPECTION_HEADERS;
const COMMENTARY_REFERENCE=[
 ' -\u00a0визуальная проверка, критики нет\u00a0 - указывается для паллет, по которым в ходе осмотра визуальная критика не выявлена',
 ' -\u00a0наличие критики, разбраковка\u00a0– указывается для паллет, по которым проведена разбраковка, но % брака/%нестандарта не превышает допустимый для отгрузки в сеть. Полученный % брака по результату разбракови необходимо фиксировать в журнале.',
 ' -\u00a0выставлен акт поставщику\u00a0- указывается для паллет, по которым проведена разбраковка, % \u00a0брака превышает допустимый для отгрузки в сеть. Выставлен Акт с хранения',
 ' -\u00a0не спущен\u00a0– указывается для паллет, не спущенных для контроля. '
];
function prepareDefs(which,r){
 const selected=[...r.selected].sort(macro.compareStorage);
 if(which==='short'){
  // В исходном шаблоне F — пустой разделитель, G — результат проверки.
  const protocol=[
   ['Источник',state.files.join(' + ')],
   ['Дата формирования',new Date().toLocaleString('ru-RU')],
   ['Доля выборки',`${r.recentPercent}% (недавние) + ${r.olderPercent}% (давние)${r.totalPalettes===null?'':'; ручное количество '+r.totalPalettes}`],
   ['Буферные зоны',r.zone==='all'?'Все':'12–16'],
   ['Поставщик (фильтр)',r.supplier||'Не задан'],
   ['Приоритетный код',r.productCode||'Не задан'],
   ['Кодовые позиции в выборке',r.uniqueCodes],
   ['Палет из файла',r.sourcePallets],
   ['Подходящих палет ФРЕШ/ФРОВ',r.recent.length+r.older.length+r.unknown.length],
   ['После фильтра поставщика',r.eligibleCount],
   ['Отобрано палет',r.selected.length],
   ['Приоритетный код (палет)',`${r.priorityCount} из ${r.priorityRequested}`],
   ['Зерно случайного отбора',String(r.seed)],
   ['Принцип отбора',`Недавно: ${r.selectedRecent.length} / ${r.recentPercent}%; давно: ${r.selectedOlder.length} / ${r.olderPercent}%; приоритет ${r.priorityRequested}`],
   ['Принцип сортировки','По адресу МП, по возрастанию'],
   ['Данные для проверки','Только подтверждённые палеты из загруженной выгрузки'],
   ['Результат проверки','Пустой — заполнить после физического осмотра']
  ];
  return [
    {name:'Проверка палет',type:'pallet',headers:['номер','код','назв','МО','МП','','Результат проверки'],rows:selected.map(p=>[Number.isSafeInteger(Number(p.number))?Number(p.number):p.number,p.code,p.name,p.mo,p.mp,'','']),widths:[13,16,59,18,27,9,26],noFilter:true},
    {name:'Протокол отбора',headers:['Параметр','Значение'],rows:protocol,widths:[42,93]}
  ];
 }
 if(which==='request'){
  const chosenShift=$('macroShift').value;
  const dayRows=chosenShift==='День'?selected.map(p=>macro.inspectionRow(p,r.asOfDate)):[];
  const nightRows=chosenShift==='Ночь'?selected.map(p=>macro.inspectionRow(p,r.asOfDate)):[];
  const buffers=[...r.allBuffers].sort(macro.compareStorage).map(p=>[Number.isSafeInteger(Number(p.number))?Number(p.number):p.number,p.mp]);
  // Колонки Q, R и S у всех строк намеренно пусты до решения менеджера.
  for(const row of [...dayRows,...nightRows]){row[16]='';row[17]='';row[18]='';}
  return [
   {name:'День',type:'pallet',headers:REQUEST_HEADERS,rows:dayRows,widths:[11,21,12,11,17,17,21,13,57,13,18,12,19,17,22,22,34,35,33]},
   {name:'Ночь',type:'pallet',headers:REQUEST_HEADERS,rows:nightRows,widths:[11,21,12,11,17,17,21,13,49,13,17,12,19,17,22,22,34,35,33]},
   {name:'Лист1',headers:['номер','МП'],rows:buffers,widths:[16,22],noFilter:true},
   {name:'Унифицированный коментарий',headers:[],rows:COMMENTARY_REFERENCE.map(v=>[v]),widths:[160],noFilter:true}
  ];
 }
 throw Error('Неизвестный шаблон');
}
function save(which){
 const r=state.report;if(!r){status('Сначала сформируйте выборку.',true);return;}
 try{
  const filename=which==='short'?`MAGNIT_MACRO_PALLETS_${r.asOfDate}.xlsx`:`Запрос на Спуск паллет ${macro.formatRussianDate(r.asOfDate)}.xlsx`;
  if(state.downloads[which]?.url)URL.revokeObjectURL(state.downloads[which].url);
  const out=MagnitExcel.save(prepareDefs(which,r),filename);state.downloads[which]=out;
  const link=$(which==='short'?'macroDownloadLink':'macroRequestDownloadLink');link.href=out.url;link.download=out.filename;link.hidden=false;
  status(which==='short'?`Таблица №1 готова: ${format(r.selected.length)} палет, результат проверки пустой.`:`Таблица №2 готова: ${format(r.selected.length)} палет в смене «${$('macroShift').value}», ${format(r.bufferCount)} записей в «Лист1». Комментарии пусты.`);
 }catch(e){status('Не удалось сформировать файл по шаблону: '+(e?.message||String(e)),true);}
}
function localDate(){const d=new Date();return [d.getFullYear(),String(d.getMonth()+1).padStart(2,'0'),String(d.getDate()).padStart(2,'0')].join('-');}
$('macroAsOf').value=localDate();
$('macroFile').addEventListener('change',e=>{parseFiles([...e.target.files||[]]);e.target.value='';});
for(const name of ['dragenter','dragover'])$('macroZone').addEventListener(name,e=>{e.preventDefault();$('macroZone').classList.add('dragging');});
for(const name of ['dragleave','drop'])$('macroZone').addEventListener(name,e=>{e.preventDefault();$('macroZone').classList.remove('dragging');});
$('macroZone').addEventListener('drop',e=>{e.preventDefault();parseFiles([...e.dataTransfer?.files||[]]);});
$('macroZone').addEventListener('keydown',e=>{if(e.target===$('macroZone')&&(e.key==='Enter'||e.key===' ')){e.preventDefault();$('macroFile').click();}});
$('macroClear').addEventListener('click',()=>{state.rows.clear();state.files=[];hideResult();updateFileInfo();$('macroZone').classList.remove('invalid');status('Накопление очищено. Загрузите новую сводную выгрузку.');});
$('macroBuild').addEventListener('click',()=>doBuild());$('macroReroll').addEventListener('click',()=>doBuild());$('macroExport').addEventListener('click',()=>save('short'));$('macroExportRequest').addEventListener('click',()=>save('request'));
for(const id of ['macroRecentRate','macroOlderRate','macroRecentDays','macroAsOf','macroZoneMode','macroSupplier','macroCode','macroTotalPallets','macroPriorityQty','macroShift'])$(id).addEventListener('change',()=>{if(state.rows.size)doBuild();});
})();
