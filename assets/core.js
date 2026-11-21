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
  function makeRecord(code, saldo, service, rule) {
    const start = num(saldo?.start);
    const ending = num(saldo?.end);
    const order = num(service?.order);
    const imp = num(service?.imp);
    const ship = num(service?.ship);
    const vich = num(service?.vich);
    const nk = rule && rule.nk != null ? num(rule.nk) : null;
    const kk = rule && rule.kk != null ? num(rule.kk) : null;
    const usOrder = order ? ship / order : 0;
    const usImport = imp ? ship / imp : 0;
    const actualDays = ship ? ending / ship : 0;
    const orderDays = order ? ending / order : 0;
    const noOrder = order === 0;
    return {
      code, name: saldo?.name || service?.name || rule?.name || 'Название не найдено',
      start, order, imp, ship, usOrder: roundForPreview(usOrder),
      usImport: roundForPreview(usImport), vich, ending,
      nk, kk, actualDays: roundForPreview(actualDays), orderDays: roundForPreview(orderDays),
      targetComment: noOrder ? 'Без расхода' : nk === null ? '' : actualDays - nk,
      maxComment: noOrder ? 'Без расхода' : kk === null ? '' : actualDays - kk,
      missingNorm: nk === null || kk === null, missingSaldo: !saldo, hasService: Boolean(service)
    };
  }
  function buildReport(payload, { date, mode = 'reference' } = {}) {
    const saldo = payload.saldo || {};
    const perDate = payload.service || {};
    const catalog = payload.catalog || [];
    const catalogSet = new Set(catalog);
    const flowerSet=new Set(payload.flowerCodes||[]);
    const norms = payload.norms || {};
    const selected = Object.create(null);
    for (const code of Object.keys(saldo)) selected[code] = 1;
    for (const item of Object.values(perDate)) {
      if (item.date === date) selected[item.code] = 1;
    }
    let appliedMode = mode;
    if (mode === 'reference' && !catalog.length) appliedMode = 'fresh';
    const serviceAtDate = Object.create(null);
    for (const entry of Object.values(perDate)) {
      if (entry.date === date) serviceAtDate[entry.code] = entry;
    }
    const records = [];
    for (const code of Object.keys(selected)) {
      const s = saldo[code]; const u = serviceAtDate[code];
      if (appliedMode === 'reference' && !catalogSet.has(code)) continue;
      if (appliedMode === 'fresh' && !isFresh(s, u, code, catalog)) continue;
      if (appliedMode === 'flowers' && !flowerSet.has(code)) continue;
      records.push(makeRecord(code, s, u, norms[code]));
    }
    records.sort((a,b) => a.name.localeCompare(b.name, 'ru') || a.code.localeCompare(b.code));
    const stats = {
      total: records.length, normMissing: records.filter(r => r.missingNorm).length,
      saldoMissing: records.filter(r => r.missingSaldo).length,
      withoutOrder: records.filter(r => r.order === 0).length,
      overMax: records.filter(r => r.kk !== null && r.order !== 0 && r.actualDays > r.kk).length,
      withService: records.filter(r => r.hasService).length,
      totalOrder: records.reduce((s,r) => s+r.order, 0),
      totalShipment: records.reduce((s,r) => s+r.ship, 0), appliedMode
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
  return { HEADERS, normalize, num, buildReport, makeRecord, parseRuDate, toRuDate, isFresh,selectedNorms,todayIso,mergeNormHistory,buildCriticalReport };
});
