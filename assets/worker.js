/* All input files are processed inside the user's browser. Nothing is uploaded. */
importScripts('https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js');
importScripts('./core.js');
const Core = self.MagnetCore;
function reportProgress(message) {self.postMessage({type:'progress',message});}
function readBook(buffer) {
  return XLSX.read(buffer, { type:'array', dense:true, cellDates:false, cellText:false, cellNF:false, bookVBA:false });
}
function getRows(book, sheetName) {
  const sheet = book.Sheets[sheetName || book.SheetNames[0]];
  if (!sheet) throw new Error(`Лист «${sheetName}» не найден`);
  if (!Array.isArray(sheet['!data'])) throw new Error('Не удалось прочитать строки Excel.');
  return sheet['!data'];
}
const raw = cell => cell == null ? null : cell.v;
function cellText(cell) { const v=raw(cell); return v == null ? '' : String(v).trim(); }
function mapHeaders(headerRow) {
  const out=Object.create(null);
  for (let i=0;i<(headerRow||[]).length;i++) {
    const val=cellText(headerRow[i]).replace(/\s+/g, ' ').toLowerCase();
    if (val) out[val]=i;
  }
  return out;
}
function requireColumns(header, expected, filename) {
  const indices={};const missing=[];
  for (const name of expected) {
    const idx=header[name.toLowerCase()];
    if (idx === undefined) missing.push(name);else indices[name]=idx;
  }
  if (missing.length) throw new Error(`Файл «${filename}»: отсутствуют столбцы: ${missing.join(', ')}`);
  return indices;
}
function value(row, cols, name) {return raw(row[cols[name]]);}
function productCode(v){ return Core.normalize(v); }
function serviceDate(value){
  if (typeof value === 'number') {
    const info=XLSX.SSF.parse_date_code(value);
    if (info) return `${String(info.d).padStart(2,'0')}.${String(info.m).padStart(2,'0')}.${info.y}`;
  }
  const v=String(value||'').trim();
  if (Core.parseRuDate(v)) return Core.toRuDate(Core.parseRuDate(v));
  return v;
}
function parseSaldo(book) {
  const rows=getRows(book);
  const h=mapHeaders(rows[0]);
  const cols=requireColumns(h, ['Код товара','Название','Начальное количество','Конечное количество','ГР_ПНАЗВ'], 'Сальдово-оборотная ведомость');
  const group2=h['полное название группы 2'];
  const out=Object.create(null), duplicates=[];
  for (let i=1;i<rows.length;i++) {
    const row=rows[i];if(!row)continue;
    const code=productCode(value(row,cols,'Код товара'));if(!code)continue;
    const item={name:value(row,cols,'Название')||'',start:Core.num(value(row,cols,'Начальное количество')),
      end:Core.num(value(row,cols,'Конечное количество')),
      group:value(row,cols,'ГР_ПНАЗВ')||'',group2:group2==null?'':(raw(row[group2])||'')};
    if (out[code]) duplicates.push(code);
    out[code]=item;
  }
  return {items:out, duplicates};
}
function parseService(book) {
  const rows=getRows(book);
  const h=mapHeaders(rows[0]);
  const cols=requireColumns(h, ['Код товара','Наим_ТП','Заказ','Импорт','Отгрузка','Vicherk','Дата операции','Тип_Отбор'], 'Уровень сервиса');
  const out=Object.create(null),dates=Object.create(null),rcs=new Set();
  const rcIndex=h['рц'];
  for (let i=1;i<rows.length;i++) {
    const row=rows[i];if(!row)continue;
    const code=productCode(value(row,cols,'Код товара'));
    if(!code)continue;
    const date=serviceDate(value(row,cols,'Дата операции'));
    if(!Core.parseRuDate(date)) throw new Error(`Неизвестный формат даты операции в строке ${i+1}: ${date}`);
    dates[date]=(dates[date]||0)+1;
    if(rcIndex!==undefined && raw(row[rcIndex])!=null)rcs.add(String(raw(row[rcIndex])));
    const key=`${date}|${code}`;
    if (!out[key]) out[key]={code,date,name:value(row,cols,'Наим_ТП')||'',order:0,imp:0,ship:0,vich:0,fresh:false};
    const target=out[key];
    target.order += Core.num(value(row,cols,'Заказ'));
    target.imp += Core.num(value(row,cols,'Импорт'));
    target.ship += Core.num(value(row,cols,'Отгрузка'));
    target.vich += Core.num(value(row,cols,'Vicherk'));
    if(String(value(row,cols,'Тип_Отбор')||'').toLowerCase().includes('овощи-фрукты')) target.fresh=true;
  }
  return {items:out, dates, rcs:[...rcs]};
}
function parseTemplate(book) {
  const candidates=book.SheetNames;
  if (!candidates.includes('НТЗ')) throw new Error('В образце не найден лист «НТЗ» с нормативами.');
  const rows=getRows(book,'НТЗ');
  const header=mapHeaders(rows[0]);
  const codeIdx=header['код товара'];
  const nkIdx=header['нк тз для рц'];
  const kkIdx=header['кк тз для рц'];
  if (codeIdx==null||nkIdx==null||kkIdx==null)throw new Error('Лист «НТЗ» должен содержать столбцы «Код товара», «НК ТЗ для РЦ», «КК ТЗ для РЦ».');
  const norms=Object.create(null);
  for (let i=1;i<rows.length;i++) {
    const row=rows[i];if(!row)continue;
    const code=productCode(raw(row[codeIdx]));if(!code)continue;
    norms[code]={name:cellText(row[1]),nk:Core.num(raw(row[nkIdx])),kk:Core.num(raw(row[kkIdx]))};
  }
  const dated=candidates.filter(name=>/^\d{2}\.\d{2}\.\d{4}$/.test(name)).sort((a,b)=>Core.parseRuDate(b).localeCompare(Core.parseRuDate(a)));
  const catalog=[];
  if(dated.length) {
    const daily=getRows(book,dated[0]);
    for(let i=1;i<daily.length;i++) {
      const row=daily[i];if(!row)continue;
      const code=productCode(raw(row[0]));if(code)catalog.push(code);
    }
  }
  return {norms,catalog:[...new Set(catalog)],catalogSheet:dated[0]||null};
}
function parseGT(book) {
  const rows=getRows(book);
  const h=mapHeaders(rows[0]);
  requireColumns(h,['дата приемки','код ТП','Наименование ТП'], 'Отчёт возвраты ГТ');
  const headers=(rows[0]||[]).map(cellText);
  const records=[];
  for(let i=1;i<rows.length;i++){
    const row=rows[i];if(!row)continue;
    const values=headers.map((_,j)=>raw(row[j]));
    if(values.some(v=>v!==null && v!==''))records.push(values);
  }
  return {headers,records};
}
self.onmessage=async ({data}) => {
  try {
    if(data.action==='stocks') {
      reportProgress('Чтение сальдово-оборотной ведомости...');
      const saldo=parseSaldo(readBook(data.saldo));
      reportProgress('Обработка операций уровня сервиса и объединение кодов...');
      const service=parseService(readBook(data.service));
      self.postMessage({type:'stocks',payload:{saldo:saldo.items,service:service.items,
        meta:{saldoCount:Object.keys(saldo.items).length,serviceCount:Object.keys(service.items).length,
          duplicateSaldo:saldo.duplicates.slice(0,20), dates:service.dates, rcs:service.rcs}}});
    } else if(data.action==='template') {
      reportProgress('Чтение листа «НТЗ» и списка товаров из образца...');
      const payload=parseTemplate(readBook(data.buffer));
      self.postMessage({type:'template',payload});
    } else if(data.action==='gt') {
      reportProgress('Чтение реестра возвратов ГТ...');
      self.postMessage({type:'gt',payload:parseGT(readBook(data.buffer))});
    } else throw new Error('Неизвестная команда.');
  } catch(error) {
    self.postMessage({type:'error',message:error?.message||String(error)});
  }
};
