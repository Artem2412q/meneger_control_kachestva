(() => {
  'use strict';
  const Core = window.MagnetCore;
  const $ = id => document.getElementById(id);
  const state = {
    files: {saldo:null, service:null, gt:null,criticalSaldo:null,criticalService:null},
    sources:null, template:null, report:null, gt:null,criticalSources:null,criticalReport:null
  };
  const STORAGE_KEY = 'magnit-rc-v2-reference';
  const moneyNumber = n => Number.isFinite(Number(n)) ? Number(n).toLocaleString('ru-RU',{maximumFractionDigits:2}) : '—';
  const decimalNumber = n => Number(n).toLocaleString('ru-RU',{minimumFractionDigits:2,maximumFractionDigits:2});
  const percentNumber = n => Number(n).toLocaleString('ru-RU',{style:'percent',minimumFractionDigits:2,maximumFractionDigits:2});
  let toastTimeout;

  function toast(message, error=false) {
    const node=$('toast');node.textContent=message;node.hidden=false;node.classList.toggle('error',error);
    clearTimeout(toastTimeout);toastTimeout=setTimeout(()=>node.hidden=true,6500);
  }
  function changeTab(tab) {
    document.querySelectorAll('.view').forEach(el=>el.classList.toggle('active',el.id===`view-${tab}`));
    document.querySelectorAll('.nav-button').forEach(el=>{
      const active=el.dataset.tab===tab;
      el.classList.toggle('active',active);
      if(active)el.setAttribute('aria-current','page');else el.removeAttribute('aria-current');
    });
    document.querySelector('.sidebar').classList.remove('open');
    window.scrollTo({top:0,behavior:'smooth'});
  }
  document.querySelectorAll('[data-tab],[data-go]').forEach(el=>el.addEventListener('click',()=>changeTab(el.dataset.tab||el.dataset.go)));
  $('mobileMenu').addEventListener('click',()=>document.querySelector('.sidebar').classList.toggle('open'));

  function baseName(filename){ return filename.replace(/\.(xlsx|xls)$/i,'').toLowerCase(); }
  function validateName(file, expected) {
    return !!file && /\.xlsx$/i.test(file.name) && baseName(file.name)===expected.toLowerCase();
  }
  function updateChecks() {
    const names=validateName(state.files.saldo,'rep_turnover_saldo_for_logist_excel') &&
      validateName(state.files.service,'rpt_levelservice_api.level_service_hd_gk_eday_hm');
    setCheck('checkNames',names,'Точные имена файлов');
    setCheck('checkHeaders',Boolean(state.sources),'Проверка столбцов');
    setCheck('checkDate',Boolean(state.report),'Дата уровня сервиса');
    setCheck('checkNorms',Boolean(state.template && Object.keys(Core.selectedNorms(state.template,$('normSnapshot').value)).length),'Нормативы НК / КК');
    $('processStock').disabled=!names;
  }
  function setCheck(id,ok,text) {
    const el=$(id);el.classList.toggle('ok',ok);
    el.querySelector('.check-indicator').textContent=ok?'✓':'○';
    el.lastChild.textContent=` ${text}`;
  }
  function invalidateStock() {
    state.sources=null;state.report=null;$('exportStock').disabled=true;$('previewStock').disabled=true;
    $('reportPreviewPanel').hidden=true;updateChecks();
  }
  function setFile(kind,file) {
    if(!file)return;
    const inputId={saldo:'saldoFile',service:'serviceFile',gt:'gtFile'}[kind];
    const statusId={saldo:'saldoStatus',service:'serviceStatus',gt:'gtStatus'}[kind];
    const zoneId={saldo:'saldoZone',service:'serviceZone',gt:'gtZone'}[kind];
    state.files[kind]=file;
    const valid=kind==='gt'?/\.xlsx?$/i.test(file.name):validateName(file,kind==='saldo'?'rep_turnover_saldo_for_logist_excel':'rpt_levelservice_api.level_service_hd_gk_eday_hm');
    $(zoneId).classList.toggle('valid',valid);$(zoneId).classList.toggle('invalid',!valid);
    $(statusId).textContent= valid ? `${file.name} · ${(file.size/1024/1024).toFixed(2)} МБ` : `Ошибка: неверное название или формат — ${file.name}`;
    if(kind!=='gt') {invalidateStock();}
    else if(valid) {importGT(file);} else {state.gt=null;$('gtPreviewPanel').hidden=true;toast('Для возвратов нужен файл Excel .xls или .xlsx.',true);}
    if(!valid && kind!=='gt')toast('Имя файла должно в точности совпадать с указанным в карточке.',true);
  }
  function bindFileZone(kind) {
    const ids={saldo:['saldoFile','saldoZone'],service:['serviceFile','serviceZone'],gt:['gtFile','gtZone']}[kind];
    const input=$(ids[0]),zone=$(ids[1]);
    input.addEventListener('change',()=>setFile(kind,input.files?.[0]));
    zone.addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();input.click();}});
    ['dragenter','dragover'].forEach(event=>zone.addEventListener(event,e=>{e.preventDefault();zone.classList.add('dragging');}));
    ['dragleave','drop'].forEach(event=>zone.addEventListener(event,e=>{e.preventDefault();zone.classList.remove('dragging');}));
    zone.addEventListener('drop',e=>setFile(kind,e.dataTransfer?.files?.[0]));
  }
  ['saldo','service','gt'].forEach(bindFileZone);

  function runWorker(action, data) {
    return new Promise((resolve,reject)=>{
      if(typeof Worker === 'undefined')return reject(new Error('Этот браузер не поддерживает обработку файлов в отдельном потоке.'));
      let worker;
      try{worker=new Worker('assets/worker.js');}catch(e){return reject(e);}
      let done=false;
      worker.onmessage=e=>{
        const msg=e.data;
        if(msg.type==='progress'){toast(msg.message);return;}
        done=true;worker.terminate();
        if(msg.type==='error')reject(new Error(msg.message));else resolve(msg.payload);
      };
      worker.onerror=e=>{if(done)return;done=true;worker.terminate();reject(new Error('Не удалось загрузить библиотеку обработки Excel. Проверьте доступ к интернету.'));};
      const transfers=[];
      for(const key of ['saldo','service','buffer'])if(data[key] instanceof ArrayBuffer)transfers.push(data[key]);
      worker.postMessage({action,...data},transfers);
    });
  }

  async function parseStock() {
    if(!validateName(state.files.saldo,'rep_turnover_saldo_for_logist_excel')||
      !validateName(state.files.service,'rpt_levelservice_api.level_service_hd_gk_eday_hm')){
      return toast('Проверьте оба имени Excel-файлов.',true);
    }
    const button=$('processStock');button.disabled=true;button.textContent='Обрабатываются файлы…';
    try {
      const [saldo,service]=await Promise.all([state.files.saldo.arrayBuffer(),state.files.service.arrayBuffer()]);
      state.sources=await runWorker('stocks',{saldo,service});
      const dates=Object.keys(state.sources.meta.dates).sort((a,b)=>Core.parseRuDate(b).localeCompare(Core.parseRuDate(a)));
      if(!dates.length) throw new Error('В уровне сервиса не найдена дата операции.');
      $('reportDate').value=Core.parseRuDate(dates[0]);
      $('dateInfo').textContent=`В файле найдены даты: ${dates.join(', ')}. Отчёт по выбранной дате; РЦ: ${state.sources.meta.rcs.join(', ')||'не указан'}.`;
      $('previewStock').disabled=false;
      recalculateReport();
      toast(`Проверка пройдена. Остатков: ${state.sources.meta.saldoCount}, позиций уровня сервиса: ${state.sources.meta.serviceCount}.`);
    }catch(e){invalidateStock();toast(e.message,true);}
    finally{button.textContent='Проверить и объединить файлы →';updateChecks();}
  }
  $('processStock').addEventListener('click',parseStock);

  function recalculateReport() {
    if(!state.sources)return;
    const dateIso=$('reportDate').value;
    const date=Core.toRuDate(dateIso);
    if(!state.sources.meta.dates[date]) {
      state.report=null;$('exportStock').disabled=true;$('reportPreviewPanel').hidden=true;
      return toast(`Дата ${date} не найдена в «Уровне сервиса». Выберите дату из списка обнаруженных.`,true);
    }
    const payload={...state.sources, norms:Core.selectedNorms(state.template,$('normSnapshot').value),catalog:state.template?.catalog||[],flowerCodes:state.template?.flowerCodes||[]};
    state.report=Core.buildReport(payload,{date,mode:$('reportMode').value});
    $('exportStock').disabled=!state.report.records.length||!state.template;
    $('reportPreviewPanel').hidden=false;
    $('previewSubtitle').textContent=`${state.report.records.length} позиций · ${date} · режим «${modeLabel(state.report.stats.appliedMode)}»`;
    renderStockTable();updateChecks();
    if(state.report.stats.appliedMode!==$('reportMode').value) toast('Образец не загружен: временно применён отбор овощей и фруктов.');
    if(state.report.stats.normMissing)toast(`Внимание: ${state.report.stats.normMissing} товаров без НК/КК в выбранном НТЗ. При выгрузке они будут явно отмечены.`,true);
  }
  function modeLabel(mode){return {reference:'По списку из образца',fresh:'Овощи и фрукты',all:'Все товары',flowers:'Код цветы'}[mode]||mode;}
  ['reportMode','reportDate'].forEach(id=>$(id).addEventListener('change',recalculateReport));
  $('previewStock').addEventListener('click',recalculateReport);
  $('stockSearch').addEventListener('input',renderStockTable);
  function renderStockTable() {
    if(!state.report)return;
    const {records,stats}=state.report;
    $('stockMetrics').innerHTML='';
    [
      [moneyNumber(stats.total),'Товарных позиций'],
      [moneyNumber(stats.totalOrder),'Заказано единиц'],
      [moneyNumber(stats.overMax),'Выше максимума'],
      [moneyNumber(stats.normMissing),'Без норм НК / КК',true]
    ].forEach(([v,label,warn])=>{
      const box=document.createElement('div');box.className='metric'+(warn&&stats.normMissing?' warn':'');
      const val=document.createElement('strong');val.textContent=v;
      const small=document.createElement('small');small.textContent=label;
      box.append(val,small);$('stockMetrics').append(box);
    });
    const headers=['Код товара','Название','Нач. кол-во','Заказ','Импорт','Отгрузка','УС / заказ, %','УС / импорт, %','Vicherk','Кон. кол-во','НК (из НТЗ)','КК (из НТЗ)','ТЗ, дни','ТЗ / заказ','Откл. НК','Откл. КК'];
    renderHead($('stockTable'),headers);
    const query=$('stockSearch').value.trim().toLowerCase();
    const filtered=query?records.filter(r=>r.code.includes(query)||r.name.toLowerCase().includes(query)):records;
    const body=$('stockTable').querySelector('tbody');body.replaceChildren();
    filtered.slice(0,80).forEach(r=>{
      const vals=[r.code,r.name,r.start,r.order,r.imp,r.ship,r.usOrder,r.usImport,r.vich,r.ending,r.nk,r.kk,r.actualDays,r.orderDays,r.targetComment,r.maxComment];
      const tr=document.createElement('tr');
      vals.forEach((v,i)=>{const td=document.createElement('td');td.textContent= i<=1 ? String(v):typeof v==='number'?(i===6||i===7?percentNumber(v):decimalNumber(v)):v==null?'НЕТ В НТЗ':String(v);if((i===10||i===11)&&v===null)td.classList.add('caution');if((r.maxComment==='Без расхода'||(typeof r.maxComment==='number'&&r.maxComment>0))&&i>=13)td.classList.add('overmax');tr.append(td);});
      body.append(tr);
    });
    $('stockFooter').textContent=`Показано ${Math.min(filtered.length,80)} из ${filtered.length}. Экспорт включает все ${records.length} позиций, не только первые строки предпросмотра.`;
  }
  function renderHead(table,labels){const tr=document.createElement('tr');labels.forEach(s=>{const th=document.createElement('th');th.textContent=s;tr.append(th);});table.querySelector('thead').replaceChildren(tr);}

  function downloadWorkbook(book,filename) {
    if(typeof XLSX==='undefined')throw new Error('Библиотека Excel не загружена. Подключитесь к интернету и обновите страницу.');
    XLSX.writeFile(book,filename,{bookType:'xlsx',compression:true});
  }
  function exportStock() {
    if(!state.report)return toast('Сначала выполните расчёт.',true);
    if(!state.template)return toast('Сначала загрузите справочник НТЗ.',true);
    try {
      const critical=state.criticalReport?.date===Core.toRuDate(Core.todayIso())?state.criticalReport:null;
      const sheets=window.MagnitReport.prepareSheets({stock:state.report,critical,
        stockSources:state.sources,template:state.template,snapshot:$('normSnapshot').value});
      window.MagnitExcel.save(sheets,`Отчёт по товарному запасу ${state.report.date}.xlsx`);
      toast(`Готовый Excel: НТЗ, полный архив, коды цветов, «Цветы», «Контроль»${critical?', Критика':''}.`);
    }catch(e){toast(e.message,true);}
  }
  $('exportStock').addEventListener('click',exportStock);

  function updateTemplateStatus() {
    const template=state.template;
    const present=!!template&&Object.keys(template.norms||{}).length>0;
    const select=$('normSnapshot'),oldValue=select.value;
    select.replaceChildren(new Option(`Основной НТЗ (${template?.mainDate||'без даты'})`,'main'));
    for(let i=0;i<(template?.snapshots||[]).length;i++){
      const snapshot=template.snapshots[i];
      select.append(new Option(`Архив ${snapshot.date} (${Object.keys(snapshot.norms||{}).length} кодов)`,String(i)));
    }
    if([...select.options].some(o=>o.value===oldValue))select.value=oldValue;
    select.disabled=!present;
    $('templateStatus').textContent=present?`НТЗ: ${Object.keys(template.norms).length} норм · ${template.catalog.length} кодов образца · цветов: ${template.flowerCodes?.length||0}`:'Справочник ещё не загружен';
    $('archiveInfo').textContent=present?`Архив: ${(template.snapshots||[]).map(s=>`${s.date} (${Object.keys(s.norms||{}).length} товаров)`).join('; ')||'не найден'}. Выбранный срез полностью включается в итоговый Excel.`:'Архив НТЗ будет доступен после импорта образца.';
    $('clearTemplate').hidden=!present;updateChecks();
  }
  $('normSnapshot').addEventListener('change',()=>{if(state.sources)recalculateReport();if(state.criticalSources)recalculateCritical();updateChecks();});
  try {
    const stored=localStorage.getItem(STORAGE_KEY);if(stored)state.template=JSON.parse(stored);
  }catch(e){console.warn('localStorage not available',e);}
  updateTemplateStatus();
  $('templateFile').addEventListener('change',async e=>{
    const file=e.target.files?.[0];if(!file)return;
    if(!/\.xlsx?$/i.test(file.name))return toast('Образец должен иметь формат .xls или .xlsx.',true);
    $('templateStatus').textContent='Чтение справочника...';
    try {
      const buffer=await file.arrayBuffer();
      state.template=Core.mergeNormHistory(state.template,await runWorker('template',{buffer}));
      try{localStorage.setItem(STORAGE_KEY,JSON.stringify(state.template));}catch(err){toast('Справочник загружен, но браузер не разрешил сохранить его на будущее.');}
      updateTemplateStatus();if(state.sources)recalculateReport();if(state.criticalSources)recalculateCritical();
      toast('Импортированы НТЗ, полный архив НТЗ и коды цветов.');
    }catch(err){updateTemplateStatus();toast(err.message,true);}
    e.target.value='';
  });
  $('clearTemplate').addEventListener('click',()=>{
    state.template=null;
    try{localStorage.removeItem(STORAGE_KEY);}catch(e){}
    updateTemplateStatus();if(state.sources)recalculateReport();if(state.criticalSources)recalculateCritical();
    $('exportStock').disabled=true;$('exportCritical').disabled=true;
    toast('Локальный справочник очищен.');
  });

  // Независимая «Критика»: две новые выгрузки, дата только локальная сегодня.
  function criticalToday(){return Core.toRuDate(Core.todayIso());}
  $('todayDateLabel').textContent=criticalToday();
  function criticalName(kind){return kind==='criticalSaldo'?'rep_turnover_saldo_for_logist_excel':'rpt_levelservice_api.level_service_hd_gk_eday_hm';}
  function updateCriticalButton(){
    $('processCritical').disabled=!validateName(state.files.criticalSaldo,criticalName('criticalSaldo'))||
      !validateName(state.files.criticalService,criticalName('criticalService'));
  }
  function invalidateCritical(){
    state.criticalSources=null;state.criticalReport=null;
    $('exportCritical').disabled=true;$('criticalPanel').hidden=true;updateCriticalButton();
  }
  function setCriticalFile(kind,file){
    if(!file)return;state.files[kind]=file;invalidateCritical();
    const zone=$(kind==='criticalSaldo'?'criticalSaldoZone':'criticalServiceZone');
    const status=$(kind==='criticalSaldo'?'criticalSaldoStatus':'criticalServiceStatus');
    const valid=validateName(file,criticalName(kind));
    zone.classList.toggle('valid',valid);zone.classList.toggle('invalid',!valid);
    status.textContent=valid?`${file.name} · ${(file.size/1024/1024).toFixed(2)} МБ`:`Неверное имя/формат: ${file.name}`;
    if(!valid)toast('Для критики требуется файл .xlsx с точным названием.',true);
    updateCriticalButton();
  }
  for(const [kind,inputId,zoneId] of [['criticalSaldo','criticalSaldoFile','criticalSaldoZone'],['criticalService','criticalServiceFile','criticalServiceZone']]){
    const input=$(inputId),zone=$(zoneId);
    input.addEventListener('change',()=>setCriticalFile(kind,input.files?.[0]));
    zone.addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();input.click();}});
    ['dragenter','dragover'].forEach(event=>zone.addEventListener(event,e=>{e.preventDefault();zone.classList.add('dragging');}));
    ['dragleave','drop'].forEach(event=>zone.addEventListener(event,e=>{e.preventDefault();zone.classList.remove('dragging');}));
    zone.addEventListener('drop',e=>setCriticalFile(kind,e.dataTransfer?.files?.[0]));
  }
  function recalculateCritical(){
    if(!state.criticalSources)return;
    const date=criticalToday();
    if(!state.criticalSources.meta.dates[date]){
      state.criticalReport=null;$('criticalPanel').hidden=true;$('exportCritical').disabled=true;
      return toast(`Критика строго за сегодня (${date}). В выгрузке уровня сервиса нет такой даты.`,true);
    }
    const payload={...state.criticalSources,norms:Core.selectedNorms(state.template,$('normSnapshot').value),
      catalog:state.template?.catalog||[]};
    state.criticalReport=Core.buildCriticalReport(payload,{today:Core.todayIso(),mode:'fresh'});
    $('exportCritical').disabled=!state.template||!state.criticalReport.records.length;
    $('criticalPanel').hidden=false;
    const {stats,records}=state.criticalReport;
    $('criticalSummary').textContent=`Сегодня ${date} · ${records.length} товаров (отбор овощи/фрукты — предварительный)`;
    $('criticalMetrics').replaceChildren();
    for(const [value,name] of [[stats.total,'Товарных позиций'],[stats.totalOrder,'Заказ'],[stats.overMax,'Выше КК'],[stats.normMissing,'Нормы не найдены']]){
      const box=document.createElement('div');box.className='metric';
      const strong=document.createElement('strong');strong.textContent=moneyNumber(value);
      const small=document.createElement('small');small.textContent=name;box.append(strong,small);$('criticalMetrics').append(box);
    }
    renderHead($('criticalTable'),Core.HEADERS.filter((_,i)=>i!==8));
    const body=$('criticalTable').querySelector('tbody');body.replaceChildren();
    records.slice(0,70).forEach(r=>{
      const vals=[r.code,r.name,r.start,r.order,r.imp,r.ship,r.usOrder,r.usImport,r.ending,r.nk,r.kk,r.actualDays,r.orderDays,r.targetComment,r.maxComment];
      const tr=document.createElement('tr');vals.forEach((v,i)=>{
        const td=document.createElement('td');td.textContent=v==null?'НЕТ В НТЗ':typeof v==='number'?(i===6||i===7?percentNumber(v):decimalNumber(v)):String(v);
        if((i===9||i===10)&&v===null)td.classList.add('caution');
        if((r.maxComment==='Без расхода'||(typeof r.maxComment==='number'&&r.maxComment>0))&&i>=12)td.classList.add('overmax');tr.append(td);
      });body.append(tr);
    });
    if(stats.normMissing)toast(`Критика: отсутствуют НК/КК для ${stats.normMissing} товаров. Проверьте «Контроль».`,true);
  }
  $('processCritical').addEventListener('click',async()=>{
    if(!validateName(state.files.criticalSaldo,criticalName('criticalSaldo'))||!validateName(state.files.criticalService,criticalName('criticalService')))
      return toast('Выберите две отдельные корректные выгрузки для критики.',true);
    const button=$('processCritical');button.disabled=true;button.textContent='Проверяем критику…';
    try{
      const [saldo,service]=await Promise.all([state.files.criticalSaldo.arrayBuffer(),state.files.criticalService.arrayBuffer()]);
      state.criticalSources=await runWorker('stocks',{saldo,service});
      recalculateCritical();
    }catch(e){invalidateCritical();toast(e.message,true);}finally{button.textContent='Проверить критику за сегодня →';updateCriticalButton();}
  });
  $('exportCritical').addEventListener('click',()=>{
    if(!state.criticalReport||state.criticalReport.date!==criticalToday())return toast('Критика устарела. Загрузите текущие выгрузки.',true);
    try{
      const sheets=window.MagnitReport.prepareSheets({critical:state.criticalReport,template:state.template,snapshot:$('normSnapshot').value});
      window.MagnitExcel.save(sheets,`Критика ${state.criticalReport.date}.xlsx`);
      toast('Критика выгружена с нормативами НТЗ и архивом.');
    }catch(e){toast(e.message,true);}
  });

  async function importGT(file) {
    $('gtStatus').textContent='Чтение реестра возвратов...';
    try{
      const buffer=await file.arrayBuffer();
      state.gt=await runWorker('gt',{buffer});
      $('gtStatus').textContent=`${file.name} · ${state.gt.records.length} записей`;
      $('gtPreviewPanel').hidden=false;renderGT();
      toast('Реестр возвратов открыт. Доступны поиск и экспорт.');
    }catch(e){state.gt=null;$('gtPreviewPanel').hidden=true;toast(e.message,true);}
  }
  function getFilteredGT() {
    if(!state.gt)return[];
    const q=$('gtSearch').value.toLocaleLowerCase('ru').trim();
    return q?state.gt.records.filter(r=>r.some(v=>String(v==null?'':v).toLocaleLowerCase('ru').includes(q))):state.gt.records;
  }
  function renderGT() {
    if(!state.gt)return;
    const filtered=getFilteredGT();
    renderHead($('gtTable'),state.gt.headers);
    const body=$('gtTable').querySelector('tbody');body.replaceChildren();
    filtered.slice(0,70).forEach(row=>{
      const tr=document.createElement('tr');row.forEach(value=>{const td=document.createElement('td');td.textContent=value==null?'':String(value);tr.append(td);});body.append(tr);
    });
    $('gtRowCount').textContent=`Найдено: ${filtered.length} из ${state.gt.records.length}. Показаны первые ${Math.min(70,filtered.length)} записей.`;
  }
  $('gtSearch').addEventListener('input',renderGT);
  $('exportGT').addEventListener('click',()=>{
    if(!state.gt)return;
    try {
      const book=XLSX.utils.book_new();const rows=[state.gt.headers,...getFilteredGT()];
      const ws=XLSX.utils.aoa_to_sheet(rows);
      ws['!cols']=state.gt.headers.map(h=>({wch:Math.min(54,Math.max(17,h.length+4))}));
      for(let c=0;c<state.gt.headers.length;c++)if(/дата/i.test(state.gt.headers[c])){
        for(let r=1;r<rows.length;r++){
          const addr=XLSX.utils.encode_cell({r,c});
          if(ws[addr]?.t==='n'&&ws[addr].v>=20000&&ws[addr].v<=80000)ws[addr].z='dd.mm.yyyy';
        }
      }
      ws['!autofilter']={ref:`A1:${XLSX.utils.encode_cell({r:Math.max(0,rows.length-1),c:Math.max(0,state.gt.headers.length-1)})}`};
      XLSX.utils.book_append_sheet(book,ws,'Возвраты ГТ');downloadWorkbook(book,'Возвраты ГТ - выборка.xlsx');
      toast('Сформирован Excel с отфильтрованными строками.');
    }catch(e){toast(e.message,true);}
  });
})();
