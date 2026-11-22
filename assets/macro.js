/* MAGNIT RC — 2.14: two age groups and cumulative buffer register, in-browser only. */
(function(){
'use strict';
const $=id=>document.getElementById(id), macro=window.MagnitPalletMacro;
if(!macro||!window.MagnitXlsxLite||!window.MagnitExcel){console.error('Не удалось загрузить модуль макроса.');return;}
const state={files:[],rows:new Map(),report:null,download:null,loading:false};
const status=(message,error=false)=>{const el=$('macroMessage');el.textContent=message;el.style.color=error?'#aa2937':'#526e60';};
const format=n=>Number(n||0).toLocaleString('ru-RU');
const cell=(tag,text,cls)=>{const el=document.createElement(tag);el.textContent=text==null?'':String(text);if(cls)el.className=cls;return el;};
const getRows=()=>[...state.rows.values()];
function clearDownload(){if(state.download){URL.revokeObjectURL(state.download.url);state.download=null;}$('macroDownloadLink').hidden=true;}
function hideResult(){state.report=null;$('macroOutput').hidden=true;$('macroExport').disabled=true;$('macroReroll').disabled=true;clearDownload();}
function freshSeed(){if(globalThis.crypto?.getRandomValues){const a=new Uint32Array(1);crypto.getRandomValues(a);return a[0]||1;}return (Math.random()*0xffffffff)>>>0||1;}
function options(seed){return {recentPercent:Number($('macroRecentRate').value),olderPercent:Number($('macroOlderRate').value),recentDays:Number($('macroRecentDays').value),asOfDate:$('macroAsOf').value,productCode:$('macroCode').value.trim(),supplier:$('macroSupplier').value.trim(),zone:$('macroZoneMode').value,seed};}
function render(r){
 $('macroOutput').hidden=false;$('macroExport').disabled=false;$('macroReroll').disabled=false;
 $('macroSummary').textContent=`Недавно: ${format(r.selectedRecent.length)} (${r.recentPercent}%) · Давно: ${format(r.selectedOlder.length)} (${r.olderPercent}%) · «Лист1»: ${format(r.bufferCount)} буферов`;
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
 if(r.productCode)warnings.push(`Приоритетный код ${r.productCode}: ${format(r.priorityCount)} палет в рамках квот.`);
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
function save(){const r=state.report;if(!r){status('Сначала сформируйте выборку.',true);return;}
 const widths=[10,24,13,18,15,29,21,16,58,19,24,14,21,20,25,35,40,39,38];
 const rows=group=>group.map(p=>macro.inspectionRow(p,r.asOfDate));
 const both=[...r.selectedRecent,...r.selectedOlder].sort(macro.compareStorage);
 const bufferList=[...r.allBuffers].sort(macro.compareStorage).map(p=>[Number.isSafeInteger(Number(p.number))?Number(p.number):String(p.number),p.mp]);
 const protocol=[
  ['Параметр','Значение'],['Источник / файлы',state.files.join(' + ')],['Дата проверки',r.asOfDate],
  ['Недавно на РЦ — дней','0–'+r.recentDays],['Давно на РЦ — дней','больше '+r.recentDays],
  ['Недавно, %',r.recentPercent],['Давно, %',r.olderPercent],
  ['Палет недавно — кандидаты',r.recent.length],['Палет давно — кандидаты',r.older.length],
  ['Недавно — к спуску',r.selectedRecent.length],['Давно — к спуску',r.selectedOlder.length],
  ['Неизвестен возраст',r.unknown.length],['Всего буферов в Лист1',r.bufferCount],
  ['Входных записей после накопления',r.sourcePallets],['Зона выборки',r.zone==='all'?'Все буферные зоны':'Бафер 12–16'],
  ['Фильтр поставщика',r.supplier||'Не задан'],['Приоритетный код',r.productCode||'Не задан'],
  ['Приоритетный код — палет',r.priorityCount],['Случайное зерно',String(r.seed)],
  ['Способ отбора','Два независимых пула. В каждом — отдельная случайная выборка из свежих овощей/фруктов.'],
  ['Объяснение возраста','Дата формирования из источника; если отсутствует — явно указанное поле «Дней на хранении»'],
  ['Сортировка','По адресу МП внутри каждой возрастной группы'],
  ['Результат проверки / % итоговой разбраковки / Комментарий ОКК','ПУСТО — заполняет менеджер после осмотра'],
  ['Лист1','Все реально существующие буферы из накопленных выгрузок, включая другие категории продуктов']
 ];
 const defs=[
  {name:'Проверка палет',type:'pallet',headers:macro.INSPECTION_HEADERS,rows:rows(both),widths},
  {name:'Недавно на РЦ',type:'pallet',headers:macro.INSPECTION_HEADERS,rows:rows(r.selectedRecent),widths},
  {name:'Давно на РЦ',type:'pallet',headers:macro.INSPECTION_HEADERS,rows:rows(r.selectedOlder),widths},
  {name:'Лист1',headers:['номер','МП'],rows:bufferList,widths:[17,29]},
  {name:'Протокол отбора',headers:protocol[0],rows:protocol.slice(1),widths:[53,100]}
 ];
 try{clearDownload();const filename='MAGNIT_SPUSK_PALET_'+r.asOfDate+'.xlsx';const out=MagnitExcel.save(defs,filename);state.download=out;
  const link=$('macroDownloadLink');link.href=out.url;link.download=out.filename;link.hidden=false;
  status(`XLSX подготовлен: ${format(r.selected.length)} палет к спуску и ${format(r.bufferCount)} буферов в «Лист1». Все комментарии осмотра пустые.`);
 }catch(e){status('Не удалось сформировать Excel: '+(e?.message||String(e)),true);}
}
function localDate(){const d=new Date();return [d.getFullYear(),String(d.getMonth()+1).padStart(2,'0'),String(d.getDate()).padStart(2,'0')].join('-');}
$('macroAsOf').value=localDate();
$('macroFile').addEventListener('change',e=>{parseFiles([...e.target.files||[]]);e.target.value='';});
for(const name of ['dragenter','dragover'])$('macroZone').addEventListener(name,e=>{e.preventDefault();$('macroZone').classList.add('dragging');});
for(const name of ['dragleave','drop'])$('macroZone').addEventListener(name,e=>{e.preventDefault();$('macroZone').classList.remove('dragging');});
$('macroZone').addEventListener('drop',e=>{e.preventDefault();parseFiles([...e.dataTransfer?.files||[]]);});
$('macroZone').addEventListener('keydown',e=>{if(e.target===$('macroZone')&&(e.key==='Enter'||e.key===' ')){e.preventDefault();$('macroFile').click();}});
$('macroClear').addEventListener('click',()=>{state.rows.clear();state.files=[];hideResult();updateFileInfo();$('macroZone').classList.remove('invalid');status('Накопление очищено. Загрузите новую сводную выгрузку.');});
$('macroBuild').addEventListener('click',()=>doBuild());$('macroReroll').addEventListener('click',()=>doBuild());$('macroExport').addEventListener('click',save);
for(const id of ['macroRecentRate','macroOlderRate','macroRecentDays','macroAsOf','macroZoneMode','macroSupplier','macroCode'])$(id).addEventListener('change',()=>{if(state.rows.size)doBuild();});
})();
