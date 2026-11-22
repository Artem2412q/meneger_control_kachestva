/* Browser UI for claims received from MAGNIT stores. No data uploaded to server. */
(function(){'use strict';
const core=window.MagnitStoreClaimsCore,reader=window.MagnitXlsxLite,$=id=>document.getElementById(id);
const STORAGE_RULES='magnit_supplier_rules_v1',STORAGE_OVERRIDES='magnit_supplier_overrides_v1';
function loadRules(){try{return core.validSupplierRules(JSON.parse(localStorage.getItem(STORAGE_RULES)||'[]'));}catch(e){return [];}}
function loadOverrides(){try{const data=JSON.parse(localStorage.getItem(STORAGE_OVERRIDES)||'{}'),result=Object.create(null);if(!data||typeof data!=='object'||Array.isArray(data))return result;for(const [k,v] of Object.entries(data).slice(0,5000))if(typeof v==='string'&&k.length<=600&&v.length<=140)result[k]=v;return result;}catch(e){return Object.create(null);}}
const state={parsed:[],merged:null,report:null,lastLink:null,selectedGroup:null,aliases:loadRules(),overrides:loadOverrides()};
const day=new Date(),today=[day.getFullYear(),String(day.getMonth()+1).padStart(2,'0'),String(day.getDate()).padStart(2,'0')].join('-');
$('storeClaimsDate').value=today;
const count=n=>Number(n).toLocaleString('ru-RU');
const showDate=core.ruDate;
const fmtNum=n=>n==null?'—':Number(n).toLocaleString('ru-RU',{maximumFractionDigits:2});
function notice(message,error=false){const e=$('storeClaimsMessage');e.textContent=message;e.classList.toggle('issue',error);}
function cell(tag,value,klass=''){const el=document.createElement(tag);el.textContent=value==null?'':String(value);if(klass)el.className=klass;return el;}
function emptyTable(body,label){body.replaceChildren();const tr=document.createElement('tr'),td=cell('td',label);td.colSpan=10;td.className='store-empty';tr.append(td);body.append(tr);}
function persistSuppliers(){
  try{localStorage.setItem(STORAGE_RULES,JSON.stringify(state.aliases));localStorage.setItem(STORAGE_OVERRIDES,JSON.stringify(state.overrides));return true;}
  catch(e){feedbackSupplier('Изменения применены к текущему отчёту, но браузер не сохранил их. Выгрузите JSON для резервной копии.',true);return false;}
}
function feedbackSupplier(message,error=false){const el=$('storeSupplierFeedback');el.textContent=message;el.classList.toggle('issue',error);}
function updateSuppliers(){
  if(!state.merged)return;
  // The previous export no longer reflects the corrected supplier data.
  if(state.lastLink){URL.revokeObjectURL(state.lastLink);state.lastLink=null;}
  $('storeClaimsDownloadLink').hidden=true;
  const previous=state.selectedGroup?.key;
  core.applySupplierEdits(state.merged.records,state.aliases,state.overrides);
  build();
  if(previous&&state.report){const group=[...state.report.groups,...state.report.characteristics].find(g=>g.key===previous);if(group)details(group,false);}
}
function putAlias(pattern,supplier){
  const p=core.text(pattern),v=core.text(supplier),key=core.supplierKey(p);
  if(!p||!v)throw Error('Заполните название из Excel и поставщика.');
  if(p.length>250||v.length>140)throw Error('Слишком длинное наименование для справочника.');
  if(state.aliases.length>=500&&!state.aliases.some(r=>core.supplierKey(r.pattern)===key))throw Error('Достигнут предел в 500 пользовательских правил.');
  state.aliases=state.aliases.filter(r=>core.supplierKey(r.pattern)!==key);
  state.aliases.push({pattern:p,supplier:v});persistSuppliers();renderSupplierRules();updateSuppliers();
}
function renderSupplierRules(){
  const root=$('storeSupplierRulesList');root.replaceChildren();
  const rules=core.effectiveSupplierRules(state.aliases),customKeys=new Set(state.aliases.map(r=>core.supplierKey(r.pattern)));
  for(const rule of rules){
    const row=document.createElement('div');row.className='store-supplier-rule';
    const patlabel=cell('label','Название в Excel');const pat=document.createElement('input');pat.type='text';pat.value=rule.pattern;pat.maxLength=250;patlabel.append(pat);
    const targetlabel=cell('label','Поставщик в отчёте');const target=document.createElement('input');target.type='text';target.value=rule.supplier;target.maxLength=140;targetlabel.append(target);
    const buttons=document.createElement('div');buttons.className='store-supplier-rule-actions';
    const save=cell('button','Сохранить');save.type='button';save.addEventListener('click',()=>{
      try{if(core.supplierKey(rule.pattern)!==core.supplierKey(pat.value))state.aliases=state.aliases.filter(r=>core.supplierKey(r.pattern)!==core.supplierKey(rule.pattern));putAlias(pat.value,target.value);feedbackSupplier('Правило обновлено. Отчёт и массовость пересчитаны.');}
      catch(e){feedbackSupplier(e.message,true);}
    });buttons.append(save);
    const hasOverride=customKeys.has(core.supplierKey(rule.pattern));
    if(hasOverride){const remove=cell('button',core.DEFAULT_SUPPLIER_RULES.some(d=>core.supplierKey(d.pattern)===core.supplierKey(rule.pattern))?'Восстановить':'Удалить');remove.type='button';remove.addEventListener('click',()=>{
      state.aliases=state.aliases.filter(r=>core.supplierKey(r.pattern)!==core.supplierKey(rule.pattern));persistSuppliers();renderSupplierRules();updateSuppliers();feedbackSupplier('Пользовательское правило удалено или восстановлено встроенное.');
    });buttons.append(remove);}
    else {const label=cell('span','Встроенное','store-supplier-origin');buttons.append(label);}
    row.append(patlabel,targetlabel,buttons);root.append(row);
  }
}
$('storeSupplierForm').addEventListener('submit',e=>{
  e.preventDefault();try{putAlias($('storeSupplierPattern').value,$('storeSupplierTarget').value);$('storeSupplierPattern').value='';feedbackSupplier('Соответствие сохранено и применено ко всем подходящим обращениям.');}
  catch(err){feedbackSupplier(err.message,true);}
});
$('storeSupplierDownload').addEventListener('click',()=>{
  const content=JSON.stringify({format:'magnit-supplier-rules-v1',rules:state.aliases,overrides:state.overrides},null,2);
  const blob=new Blob([content],{type:'application/json;charset=utf-8'}),url=URL.createObjectURL(blob),a=document.createElement('a');
  a.href=url;a.download='MAGNIT_SUPPLIER_RULES.json';document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1000);
  feedbackSupplier('Справочник и ручные исправления скачаны в JSON. Храните файл конфиденциально.');
});
$('storeSupplierImport').addEventListener('change',async e=>{
  const file=e.target.files?.[0];if(!file)return;
  try{
    if(file.size>1000000)throw Error('Слишком большой JSON (более 1 МБ).');
    const data=JSON.parse(await file.text());if(data?.format!=='magnit-supplier-rules-v1'||!Array.isArray(data.rules)||!data.overrides||typeof data.overrides!=='object'||Array.isArray(data.overrides))throw Error('Неверный формат справочника JSON.');
    const rules=core.validSupplierRules(data.rules),overrides=Object.create(null);
    for(const [k,v] of Object.entries(data.overrides).slice(0,5000))if(typeof v==='string'&&k.length<=600&&v.length<=140)overrides[k]=v;
    state.aliases=rules;state.overrides=overrides;persistSuppliers();renderSupplierRules();updateSuppliers();feedbackSupplier('Справочник импортирован. Данные в отчёте обновлены.');
  }catch(err){feedbackSupplier('Импорт не выполнен: '+err.message,true);}
  e.target.value='';
});
renderSupplierRules();
async function parseFile(file){
  if(!/\.xlsx$/i.test(file.name))throw Error(`Файл «${file.name}»: поддерживается Excel .xlsx.`);
  const book=await reader.open(await file.arrayBuffer());
  let lastError;for(const sheet of book.names){const table=[];await book.eachRow(sheet,row=>{table.push(row.map(c=>c?.v??null));});
    if(!table.length)continue;try{return core.parseTable(table,file.name);}catch(e){lastError=e;}
  }
  throw lastError||Error(`В файле «${file.name}» не найдена таблица претензий.`);
}
function hideOutput(){state.report=null;state.selectedGroup=null;$('storeClaimsDashboard').hidden=true;$('storeClaimsExport').disabled=true;$('storeClaimsDownloadLink').hidden=true;$('storeClaimsDetailPanel').hidden=true;}
$('storeClaimsFiles').addEventListener('change',async function(e){
  const files=[...e.target.files];if(!files.length)return;
  hideOutput();
  $('storeClaimsFileNames').textContent='Добавление: '+files.map(f=>f.name).join(' · ');
  notice('Читаю Excel и проверяю обязательные столбцы…');$('storeClaimsBuild').disabled=true;
  try{
    const added=[];for(const f of files)added.push(await parseFile(f));
    state.parsed.push(...added);
    state.merged=core.merge(state.parsed);
    core.applySupplierEdits(state.merged.records,state.aliases,state.overrides);
    $('storeClaimsFileNames').textContent=state.parsed.map(f=>f.stats.source).join(' · ');
    const uniqueDates=[...new Set(state.merged.records.map(r=>r.date))].sort();
    $('storeClaimsSourceInfo').textContent=`Загружено файлов: ${state.parsed.length}. Уникальных обращений: ${count(state.merged.records.length)}. Удалено повторов: ${count(state.merged.duplicates)}. Даты: ${uniqueDates.map(showDate).join(', ')||'нет'}.`;
    $('storeClaimsBuild').disabled=!state.merged.records.length;
    notice(state.merged.records.length?'Excel прочитан. Нажмите «Сформировать массовость».':'Нет подтверждённых обращений с кодом, названием и датой.',!state.merged.records.length);
  }catch(err){notice(err.message,true);$('storeClaimsSourceInfo').textContent=state.merged?`Ранее загружено обращений: ${count(state.merged.records.length)}. Ошибка нового файла.`:'Файл не загружен.';$('storeClaimsBuild').disabled=!state.merged?.records.length;}
  e.target.value='';
});
function formatProducers(g){const a=[...g.manufacturers];return a.length?a.slice(0,2).join('; ')+(a.length>2?` · ещё ${a.length-2}`:''):'—';}
function formatSuppliers(g){
  const hits=new Map();for(const r of g.complaints)if(r.supplier)hits.set(r.supplier,(hits.get(r.supplier)||0)+1);
  return hits.size?[...hits].map(([name,n])=>name+(n<g.count||hits.size>1?` (${n} из ${g.count})`:'' )).join('; '):'—';
}
function formatDates(g){return [state.report.yesterday,state.report.asOf].map(d=>g.dates.has(d)?`${showDate(d)}: ${g.dates.get(d)}`:null).filter(Boolean).join(' · ');}
function rowForGroup(g,type){
  const tr=document.createElement('tr');
  tr.append(cell('td',g.code,'store-code'),cell('td',g.name,'store-name'));
  if(type==='crit')tr.append(cell('td',g.characteristic||'—'));
  tr.append(cell('td',count(g.count),'store-number'),cell('td',count(g.stores.size),'store-number'),cell('td',formatDates(g)),cell('td',formatSuppliers(g)),cell('td',formatProducers(g)),cell('td',g.hasQuantity?fmtNum(g.badQuantity):'—','store-number'));
  const action=cell('td','');const btn=cell('button','Обращения ↗','store-detail-button');btn.type='button';btn.addEventListener('click',()=>details(g));action.append(btn);tr.append(action);return tr;
}
function renderTable(id,groups,type){const target=$(id);target.replaceChildren();if(!groups.length){emptyTable(target,'Нет записей по выбранным условиям.');return;}for(const g of groups)target.append(rowForGroup(g,type));}
function details(g,scroll=true){
  state.selectedGroup=g;
  $('storeClaimsDetailTitle').textContent=`${g.code} · ${g.name}`;
  $('storeClaimsDetailMeta').textContent=`${g.characteristic&&g.key.includes('\u0000')?g.characteristic+' · ':''}${count(g.count)} обращений · поставщик редактируется прямо в строке`;
  const body=$('storeClaimsDetailRows');body.replaceChildren();
  for(const r of [...g.complaints].sort((a,b)=>b.date.localeCompare(a.date))){
    const tr=document.createElement('tr');
    for(const value of [showDate(r.date),r.characteristic||'—',r.manufacturer||'—'])tr.append(cell('td',value));
    const td=document.createElement('td'),edit=document.createElement('div');edit.className='store-record-supplier';
    const input=document.createElement('input');input.type='text';input.value=r.supplier||'';input.placeholder='Не определён';input.maxLength=140;input.setAttribute('aria-label',`Поставщик для обращения ${r.guid||r.id||r.code}`);
    const actions=document.createElement('div');actions.className='store-record-supplier-actions';
    const save=cell('button','Сохранить');save.type='button';save.title='Изменить поставщика только для этого обращения';
    save.addEventListener('click',()=>{
      state.overrides[r.key]=core.text(input.value);persistSuppliers();updateSuppliers();notice('Поставщик для обращения обновлён. Массовость и экспорт Excel пересчитаны.');
    });actions.append(save);
    const all=cell('button','Для производителя');all.type='button';all.title='Применить к всем обращениям с этим наименованием производителя';all.disabled=!r.manufacturer;
    all.addEventListener('click',()=>{
      try{
        const supplier=core.text(input.value);
        if(!supplier)throw Error('Для правила производителя укажите наименование поставщика.');
        // Manufacturer-wide mapping supersedes manual edits on these same records.
        if(state.merged)for(const rec of state.merged.records){if(core.supplierKey(rec.manufacturer)===core.supplierKey(r.manufacturer))delete state.overrides[rec.key];}
        putAlias(r.manufacturer,supplier);notice('Правило применено ко всем обращениям этого производителя и сохранено.');
      }catch(err){notice(err.message,true);}
    });actions.append(all);
    edit.append(input,actions);td.append(edit);tr.append(td);
    for(const value of [r.storeCode||'—',r.store||'—',r.badQuantity===null?'—':fmtNum(r.badQuantity),r.comment||'—'])tr.append(cell('td',value));
    body.append(tr);
  }
  $('storeClaimsDetailPanel').hidden=false;
  if(scroll)$('storeClaimsDetailPanel').scrollIntoView({behavior:'smooth',block:'nearest'});
}
function render(){const r=state.report,s=r.stats;
  $('storeClaimsPeriod').textContent=`${showDate(r.yesterday)} — ${showDate(r.asOf)}`;
  $('storeClaimsKpis').replaceChildren();
  for(const [label,value] of [['Обращений',s.total],['За вчера',s.yesterday],['За сегодня',s.today],['Разных товаров',s.products],['Массовых товаров',s.massProducts],['Торговых объектов',s.stores],['Гранд-Трейд',s.grandTrade]]){
    const node=document.createElement('div');node.className='store-kpi';node.append(cell('span',label),cell('strong',count(value)));$('storeClaimsKpis').append(node);
  }
  const filtered=r.groups,threshold=s.threshold;
  $('storeClaimsTableInfo').textContent=`Товаров с обращениями: ${filtered.length}. Массовых: ${r.mass.length}. Порог: от ${threshold} обращений`;
  renderTable('storeClaimsMassRows',r.mass,'mass');
  renderTable('storeClaimsCriticalRows',r.critical,'crit');
  $('storeClaimsMassSubtitle').textContent=`Массовых товаров: ${r.mass.length}. Групп по характеристике: ${r.critical.length}.`;
  const warnings=[];
  if(!s.today)warnings.push(`За ${showDate(r.asOf)} нет обращений в загруженных файлах.`);
  if(!s.yesterday)warnings.push(`За ${showDate(r.yesterday)} нет обращений в загруженных файлах.`);
  if(r.excludedDate)warnings.push(`Вне двухдневного периода: ${count(r.excludedDate)} обращений (не включены).`);
  if(state.merged.duplicates)warnings.push(`Повторы по идентификатору: ${count(state.merged.duplicates)} (не учитываются повторно).`);
  $('storeClaimsWarnings').textContent=warnings.length?warnings.join(' '):'Период подтверждён данными выгрузки.';
  $('storeClaimsDashboard').hidden=false;$('storeClaimsExport').disabled=!s.total;
  $('storeClaimsDetailPanel').hidden=true;
  notice(s.total?`Сформировано: ${count(s.total)} обращений за ${showDate(r.yesterday)} — ${showDate(r.asOf)}.`:'В выбранные два дня обращений нет — проверьте дату отчёта.',!s.total);
}
function build(){
  if(!state.merged){notice('Сначала загрузите Excel-выгрузку.',true);return;}
  try{state.report=core.aggregate(state.merged.records,{asOf:$('storeClaimsDate').value,threshold:Number($('storeClaimsThreshold').value),query:$('storeClaimsSearch').value});render();}
  catch(err){notice(err.message,true);}
}
$('storeClaimsBuild').addEventListener('click',build);
$('storeClaimsClear').addEventListener('click',()=>{
  state.parsed=[];state.merged=null;hideOutput();$('storeClaimsBuild').disabled=true;
  $('storeClaimsFileNames').textContent='Выберите Excel-файл(ы) для объединения';
  $('storeClaimsSourceInfo').textContent='Файлы не загружены. Названия столбцов сверяются автоматически с образцом.';
  $('storeClaimsSearch').value='';notice('Файлы очищены. Загрузите новый отчёт.');
});
['storeClaimsDate','storeClaimsThreshold'].forEach(id=>$(id).addEventListener('change',()=>{if(state.merged)build();}));
$('storeClaimsSearch').addEventListener('input',()=>{if(state.merged)build();});
$('storeClaimsDetailClose').addEventListener('click',()=>{$('storeClaimsDetailPanel').hidden=true;});
const csvSafe=v=>v==null?'':v;
function worksheetGroups(groups,critical=false){return groups.map(g=>[
 g.code,g.name,...(critical?[g.characteristic]:[]),g.count,g.stores.size,
 g.dates.get(state.report.yesterday)||0,g.dates.get(state.report.asOf)||0,
 formatSuppliers(g)==='—'?'':formatSuppliers(g),[...g.manufacturers].join('; '),g.hasQuantity?g.badQuantity:null
]);}
$('storeClaimsExport').addEventListener('click',()=>{
  if(!state.report||!state.report.stats.total)return notice('Нет обращений в выбранном периоде.',true);
  const r=state.report,head=['Код товара','Наименование товара','Обращений','Магазинов','Вчера','Сегодня','Поставщик','Производители','Кол-во брак'];
  const sheets=[
    {name:'Массовость',headers:head,rows:worksheetGroups(r.mass),widths:[17,45,15,14,14,14,23,70,16]},
    {name:'Критика',headers:[...head.slice(0,2),'Характеристика товара',...head.slice(2)],rows:worksheetGroups(r.critical,true),widths:[17,42,50,14,14,14,14,23,60,16]},
    {name:'Все товары',headers:head,rows:worksheetGroups(r.groups),widths:[17,42,14,14,14,14,23,60,16]},
    {name:'Обращения',headers:['guid','id','Дата обращения','Код товара','Наименование товара','Характеристика товара','Производитель (из Excel)','Поставщик','Основание определения поставщика','Код ТО','Наименование ТО','Филиал','Кол-во брак','Комментарий ТО','Причина списания','Статус'],rows:r.selected.map(x=>[x.guid,x.id,showDate(x.date),x.code,x.name,x.characteristic,x.manufacturer,x.supplier,x.supplierEvidence,x.storeCode,x.store,x.branch,x.badQuantity,x.comment,x.reason,x.status]),widths:[39,25,18,17,42,43,50,23,60,15,26,22,16,65,36,24]},
    {name:'Справочник поставщиков',headers:['Название в Excel','Поставщик','Тип правила'],rows:core.effectiveSupplierRules(state.aliases).map(rule=>[rule.pattern,rule.supplier,rule.origin]),widths:[70,32,24]},
    {name:'Правки обращений',headers:['Ключ обращения','Поставщик','Тип'],rows:Object.entries(state.overrides).map(([key,value])=>[key,value||'(не определён)','Ручная правка']),widths:[60,34,23]},
    {name:'Контроль',headers:['Параметр','Значение'],rows:[['Начало периода',showDate(r.yesterday)],['Конец периода',showDate(r.asOf)],['Обращений',r.stats.total],['За вчера',r.stats.yesterday],['За сегодня',r.stats.today],['Дубли исключены',state.merged.duplicates],['Вне периода',r.excludedDate],['Порог массовости',r.stats.threshold],['Товары в периоде',r.stats.products],['Массовых товаров',r.stats.massProducts],['Массовых дефектов',r.stats.massIssues],['Обращений поставщика Гранд-Трейд',r.stats.grandTrade],['Пользовательских правил поставщика',state.aliases.length],['Ручных правок обращений',Object.keys(state.overrides).length],['Файлов загружено',state.parsed.length]],widths:[31,30]}
  ];
  try{
    const download=window.MagnitExcel.save(sheets,`MAGNIT_STORE_CLAIMS_${r.yesterday}_${r.asOf}.xlsx`);
    if(state.lastLink)URL.revokeObjectURL(state.lastLink);
    state.lastLink=download.url;
    const a=$('storeClaimsDownloadLink');a.href=download.url;a.download=download.filename;a.hidden=false;
    notice('Отчёт Excel готов. Если автоматическое скачивание не началось, нажмите резервную ссылку.');
  }catch(err){notice(`Не удалось создать Excel: ${err.message}`,true);}
});
})();
