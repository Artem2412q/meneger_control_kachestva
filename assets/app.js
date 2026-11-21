(() => {
  'use strict';
  const Core = window.MagnetCore;
  const $ = id => document.getElementById(id);
  const state = {
    files: {saldo:null, service:null, gt:null},
    sources:null, template:null, report:null, gt:null
  };
  const STORAGE_KEY = 'magnit-rc-v1-reference';
  const moneyNumber = n => Number.isFinite(Number(n)) ? Number(n).toLocaleString('ru-RU',{maximumFractionDigits:2}) : '—';
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
    setCheck('checkNorms',Boolean(state.template && Object.keys(state.template.norms).length),'Нормативы НК / КК');
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
    const payload={...state.sources, norms:state.template?.norms||{},catalog:state.template?.catalog||[]};
    state.report=Core.buildReport(payload,{date,mode:$('reportMode').value});
    $('exportStock').disabled=!state.report.records.length;
    $('reportPreviewPanel').hidden=false;
    $('previewSubtitle').textContent=`${state.report.records.length} позиций · ${date} · режим «${modeLabel(state.report.stats.appliedMode)}»`;
    renderStockTable();updateChecks();
    if(state.report.stats.appliedMode!==$('reportMode').value) toast('Образец не загружен: временно применён отбор овощей и фруктов.');
  }
  function modeLabel(mode){return {reference:'По списку из образца',fresh:'Овощи и фрукты',all:'Все товары'}[mode]||mode;}
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
    const headers=['Код товара','Название','Нач. кол-во','Заказ','Импорт','Отгрузка','УС / заказ','УС / импорт','Vicherk','Кон. кол-во','НК','КК','ТЗ, дни','ТЗ / заказ','Откл. НК','Откл. КК'];
    renderHead($('stockTable'),headers);
    const query=$('stockSearch').value.trim().toLowerCase();
    const filtered=query?records.filter(r=>r.code.includes(query)||r.name.toLowerCase().includes(query)):records;
    const body=$('stockTable').querySelector('tbody');body.replaceChildren();
    filtered.slice(0,80).forEach(r=>{
      const vals=[r.code,r.name,r.start,r.order,r.imp,r.ship,r.usOrder,r.usImport,r.vich,r.ending,r.nk,r.kk,r.actualDays,r.orderDays,r.targetComment,r.maxComment];
      const tr=document.createElement('tr');
      vals.forEach((v,i)=>{const td=document.createElement('td');td.textContent= i<=1 ? String(v):typeof v==='number'?moneyNumber(v):v==null?'—':String(v);if(i>=14&&typeof v==='number'&&v>0)td.className='caution';tr.append(td);});
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
    try {
      const records=state.report.records, day=state.report.date;
      const book=XLSX.utils.book_new();
      const matrix=[Core.HEADERS];
      for(const r of records)matrix.push([r.code,r.name,r.start,r.order,r.imp,r.ship,r.usOrder,r.usImport,r.vich,r.ending,r.nk,r.kk,r.actualDays,r.orderDays,r.targetComment,r.maxComment]);
      const sheet=XLSX.utils.aoa_to_sheet(matrix);
      sheet['!cols']=[{wch:16},{wch:64},{wch:18},{wch:17},{wch:17},{wch:17},{wch:22},{wch:23},{wch:16},{wch:20},{wch:22},{wch:22},{wch:20},{wch:16},{wch:25},{wch:28}];
      sheet['!autofilter']={ref:`A1:P${matrix.length}`};
      const putFormula=(row,col,formula,value)=>{
        const addr=XLSX.utils.encode_cell({r:row-1,c:col});
        sheet[addr]={t:typeof value==='string'?'s':'n',f:formula,v:value};
        if(typeof value==='number')sheet[addr].z=col===6||col===7?'0.00%':'0.00';
      };
      records.forEach((r,i)=>{
        const n=i+2;
        putFormula(n,6,`IFERROR(F${n}/D${n},0)`,r.usOrder);
        putFormula(n,7,`IFERROR(F${n}/E${n},0)`,r.usImport);
        putFormula(n,12,`IFERROR(J${n}/F${n},0)`,r.actualDays);
        putFormula(n,13,`IFERROR(J${n}/D${n},0)`,r.orderDays);
        putFormula(n,14,`IF(D${n}=0,"Без расхода",IF(K${n}="","",M${n}-K${n}))`,r.targetComment);
        putFormula(n,15,`IF(D${n}=0,"Без расхода",IF(L${n}="","",M${n}-L${n}))`,r.maxComment);
      });
      XLSX.utils.book_append_sheet(book,sheet,day);
      const normRows=[['Код товара','Наименование товара','НК\nТЗ для РЦ','КК\nТЗ для РЦ']];
      for(const [code,n] of Object.entries(state.template?.norms||{}))normRows.push([code,n.name||'',n.nk,n.kk]);
      const normSheet=XLSX.utils.aoa_to_sheet(normRows);normSheet['!cols']=[{wch:16},{wch:61},{wch:18},{wch:18}];
      XLSX.utils.book_append_sheet(book,normSheet,'НТЗ');
      const usRows=[['Код товара','Сумма по полю Заказ','Сумма по полю Импорт','Сумма по полю Отгрузка','Сумма по полю Vicherk']];
      for(const item of Object.values(state.sources.service).filter(x=>x.date===day).sort((a,b)=>a.code.localeCompare(b.code)))usRows.push([item.code,item.order,item.imp,item.ship,item.vich]);
      const usSheet=XLSX.utils.aoa_to_sheet(usRows);usSheet['!cols']=[{wch:18},{wch:22},{wch:22},{wch:24},{wch:23}];
      XLSX.utils.book_append_sheet(book,usSheet,'УС');
      const controls=[['Проверка','Значение / товар','Пояснение'],
        ['Дата отчёта',day,'Определена по уровню сервиса'],
        ['Режим отбора',modeLabel(state.report.stats.appliedMode),'Состав товаров будет уточняться по вашим правилам'],
        ['Товарных позиций',records.length,'Объединение по коду товара'],
        ['Позиции без норматива',state.report.stats.normMissing,'НК/КК для этих кодов оставлены пустыми'],
        ['Позиции без сальдовой ведомости',state.report.stats.saldoMissing,'Для таких кодов остатки условно равны нулю'],
        ['Повторяющиеся коды в сальдовой ведомости',(state.sources.meta.duplicateSaldo||[]).length,'При повторе оставлена последняя строка — требует проверки']];
      for(const r of records){if(r.missingNorm) controls.push(['Не найден норматив',r.code,r.name]);if(r.missingSaldo)controls.push(['Не найден в ведомости',r.code,r.name]);}
      const ctl=XLSX.utils.aoa_to_sheet(controls);ctl['!cols']=[{wch:36},{wch:30},{wch:73}];XLSX.utils.book_append_sheet(book,ctl,'Контроль');
      const [d,m,y]=day.split('.');downloadWorkbook(book,`Отчёт по товарному запасу ${d}.${m}.${y}.xlsx`);
      toast('Excel сформирован. Проверьте лист «Контроль» перед использованием.');
    } catch(e){toast(e.message,true);}
  }
  $('exportStock').addEventListener('click',exportStock);

  function updateTemplateStatus() {
    const present=!!state.template&&Object.keys(state.template.norms).length>0;
    $('templateStatus').textContent=present?`Сохранено норм: ${Object.keys(state.template.norms).length} · товары образца: ${state.template.catalog.length} · лист: ${state.template.catalogSheet||'не найден'}`:'Справочник ещё не загружен';
    $('clearTemplate').hidden=!present;updateChecks();
  }
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
      state.template=await runWorker('template',{buffer});
      try{localStorage.setItem(STORAGE_KEY,JSON.stringify(state.template));}catch(err){toast('Справочник загружен, но браузер не разрешил сохранить его на будущее.');}
      updateTemplateStatus();if(state.sources)recalculateReport();
      toast('Нормативы НТЗ и коды товаров из образца успешно загружены.');
    }catch(err){updateTemplateStatus();toast(err.message,true);}
    e.target.value='';
  });
  $('clearTemplate').addEventListener('click',()=>{
    state.template=null;
    try{localStorage.removeItem(STORAGE_KEY);}catch(e){}
    updateTemplateStatus();if(state.sources)recalculateReport();
    toast('Локальный справочник очищен.');
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
