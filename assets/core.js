/* MAGNIT RC — isolated, testable report business logic. No network / storage. */
(function (root, factory) {
  const lib = factory();
  if (typeof module === 'object' && module.exports) module.exports = lib;
  root.MagnetCore = lib;
})(typeof self !== 'undefined' ? self : globalThis, function () {
  'use strict';
  const HEADERS = [
    'Код товара', 'Название', 'Начальное количество', 'Заказ', 'Импорт',
    'Отгрузка', 'УС отгрузки к заказу', 'УС отгрузки к импорту', ' Vicherk',
    'Конечное количество', 'Целевой ТЗ конечного РЦ (НК)',
    'Максимальный ТЗ конечного РЦ (КК)', 'Фактический тов.запас в днях ',
    'ТЗ/Заказ', 'Целевой ТЗ Комментарии', 'Максимальный ТЗ Комментарии'
  ];
  const normalize = value => String(value == null ? '' : value).trim().replace(/\.0$/, '');
  const num = value => {
    if (value === null || value === undefined || value === '') return 0;
    const n = Number(typeof value === 'string' ? value.replace(/\s/g, '').replace(',', '.') : value);
    return Number.isFinite(n) ? n : 0;
  };
  const roundForPreview = n => Number.isFinite(n) ? n : 0;
  function isFresh(saldo, service, code, catalog) {
    const text = `${saldo?.group || ''} ${saldo?.group2 || ''}`.toLowerCase();
    return Boolean(service?.fresh) || /овощ|фрукт|плодоовощ|плод[ыо]|ягод|зелень|гриб/.test(text) || Boolean(catalog?.includes(code));
  }
  // Normalize a plain-text master of product codes without inventing corrections.
  function parseEtalon(text) {
    const lines = String(text || '').replace(/^\uFEFF/, '').split(/\r?\n/);
    const codes=[], seen=new Set(), anomalies=[], invalid=[];
    let sourceLines=0,duplicates=0;
    lines.forEach((line,i)=>{
      const code=line.trim();if(!code)return;
      sourceLines++;
      if(!/^\d+$/.test(code)){invalid.push({line:i+1,code});return;}
      if(seen.has(code)){duplicates++;return;}
      seen.add(code);codes.push(code);
      if(code.length!==10)anomalies.push({line:i+1,code,length:code.length});
    });
    return {codes,sourceLines,duplicates,anomalies,invalid};
  }
  function makeRecord(code, saldo, service, rule) {
    const hasSaldo=Boolean(saldo),hasService=Boolean(service);
    // Missing source data is NULL, not fictitious 0.
    const start = hasSaldo ? num(saldo.start) : null;
    const ending = hasSaldo ? num(saldo.end) : null;
    const order = hasService ? num(service.order) : null;
    const imp = hasService ? num(service.imp) : null;
    const ship = hasService ? num(service.ship) : null;
    const vich = hasService ? num(service.vich) : null;
    const nk = rule && rule.nk != null ? num(rule.nk) : null;
    const kk = rule && rule.kk != null ? num(rule.kk) : null;
    const usOrder = hasService ? (order ? ship / order : 0) : null;
    const usImport = hasService ? (imp ? ship / imp : 0) : null;
    const actualDays = hasSaldo && hasService ? (ship ? ending / ship : 0) : null;
    const orderDays = hasSaldo && hasService ? (order ? ending / order : 0) : null;
    const noData=!hasSaldo||!hasService;
    const noOrder=order===0;
    return {
      code, name: saldo?.name || service?.name || rule?.name || 'Название не найдено',
      start, order, imp, ship, usOrder, usImport, vich, ending,
      nk, kk, actualDays, orderDays,
      targetComment: noData ? 'Нет данных' : noOrder ? 'Без расхода' : nk===null ? 'Нет НТЗ' : actualDays - nk,
      maxComment: noData ? 'Нет данных' : noOrder ? 'Без расхода' : kk===null ? 'Нет НТЗ' : actualDays - kk,
      missingNorm: nk===null || kk===null, missingSaldo: !hasSaldo, missingService: !hasService,
      missingBoth: noData && !hasSaldo && !hasService,
      hasService
    };
  }
  function buildReport(payload, { date, mode = 'reference' } = {}) {
    const saldo = payload.saldo || {};
    const perDate = payload.service || {};
    const catalog = payload.catalog || [];
    const etalonCodes = payload.etalonCodes || [];
    const catalogSet = new Set(catalog);
    const etalonSet = new Set(etalonCodes);
    const flowerSet=new Set(payload.flowerCodes||[]);
    const norms = payload.norms || {};
    const selected = new Set(Object.keys(saldo));
    const serviceAtDate = Object.create(null);
    for (const entry of Object.values(perDate)) {
      if(entry.date!==date)continue;
      serviceAtDate[entry.code]=entry;
      selected.add(entry.code);
    }
    let appliedMode = mode;
    if (mode === 'reference' && !catalog.length) appliedMode = 'fresh';
    if (mode === 'etalon' && !etalonCodes.length)throw new Error('Эталон товарных позиций не загружен.');
    const records = [];
    // In master mode the ETALON is the left-hand side of the join.
    // Preserve its order; all rows appear once, including unmatched codes.
    const candidates=appliedMode==='etalon'?etalonSet:selected;
    for (const code of candidates) {
      const s = saldo[code], u = serviceAtDate[code];
      if (appliedMode === 'reference' && !catalogSet.has(code)) continue;
      if (appliedMode === 'fresh' && !isFresh(s, u, code, catalog)) continue;
      if (appliedMode === 'flowers' && !flowerSet.has(code)) continue;
      records.push(makeRecord(code, s, u, norms[code]));
    }
    if(appliedMode!=='etalon')records.sort((a,b) => a.name.localeCompare(b.name, 'ru') || a.code.localeCompare(b.code));
    const outsideEtalon=etalonCodes.length?[...selected].filter(code=>!etalonSet.has(code)):[];
    const stats = {
      total: records.length, normMissing: records.filter(r => r.missingNorm).length,
      saldoMissing: records.filter(r => r.missingSaldo).length,
      serviceMissing: records.filter(r => r.missingService).length,
      missingBoth: records.filter(r => r.missingBoth).length,
      withoutOrder: records.filter(r => r.order === 0).length,
      overMax: records.filter(r => r.kk !== null && r.order !== null && r.order !== 0 && r.actualDays !== null && r.actualDays > r.kk).length,
      withService: records.filter(r => r.hasService).length,
      withSaldo: records.filter(r => !r.missingSaldo).length,
      totalOrder: records.reduce((s,r) => s+(r.order||0), 0),
      totalShipment: records.reduce((s,r) => s+(r.ship||0), 0),
      etalonCount: etalonCodes.length, outsideEtalon, appliedMode
    };
    return {records, stats, date};
  }
  function parseRuDate(value) {
    const t=String(value||'').trim();
    const m=t.match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
    if (m) return `${m[3]}-${m[2]}-${m[1]}`;
    if (/^\d{4}-\d{2}-\d{2}$/.test(t)) return t;
    return null;
  }
  function toRuDate(iso) {
    if (!iso || !/^\d{4}-\d{2}-\d{2}$/.test(iso)) return iso||'';
    const [y,m,d]=iso.split('-');return `${d}.${m}.${y}`;
  }
  function selectedNorms(template, index){
    if(!template) return {};
    if(index==='main'||!index)return template.norms||{};
    const numIndex=Number(index);
    return template.snapshots?.[numIndex]?.norms || {};
  }
  function todayIso(now=new Date()){
    const d=String(now.getDate()).padStart(2,'0'),m=String(now.getMonth()+1).padStart(2,'0');
    return `${now.getFullYear()}-${m}-${d}`;
  }
  function buildCriticalReport(payload,{today=todayIso(),mode='fresh'}={}){
    const date=toRuDate(today);
    if(!Object.values(payload.service||{}).some(item=>item.date===date))
      throw new Error(`Критика строго за сегодня (${date}). В уровне сервиса нет этой даты.`);
    return buildReport(payload,{date,mode});
  }
  function mergeNormHistory(previous,next){
    if(!previous)return next;
    if(!next)return previous;
    const map=new Map();
    for(const s of previous.snapshots||[])if(s?.date)map.set(s.date,s);
    if(previous.mainDate&&previous.mainDate!==next.mainDate&&previous.norms)
      map.set(previous.mainDate,{date:previous.mainDate,norms:previous.norms});
    for(const s of next.snapshots||[])if(s?.date)map.set(s.date,s);
    map.delete(next.mainDate);
    return {...next,snapshots:[...map.values()].sort((a,b)=>String(parseRuDate(a.date)||a.date).localeCompare(String(parseRuDate(b.date)||b.date))),
      flowerCodes:next.flowerCodes?.length?next.flowerCodes:(previous.flowerCodes||[]),
      catalog:next.catalog?.length?next.catalog:(previous.catalog||[])};
  }
  return { HEADERS, normalize, num, parseEtalon, buildReport, makeRecord, parseRuDate, toRuDate, isFresh,selectedNorms,todayIso,mergeNormHistory,buildCriticalReport };
});
