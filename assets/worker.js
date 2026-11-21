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
function archiveDate(value){
  if(typeof value==='number'){
    const d=XLSX.SSF.parse_date_code(value);if(d)return `${String(d.d).padStart(2,'0')}.${String(d.m).padStart(2,'0')}.${d.y}`;
  }
  return serviceDate(value);
}
function normRow(row,offset){
  const code=productCode(raw(row[offset]));if(!code)return null;
  const nk=raw(row[offset+2]),kk=raw(row[offset+3]);
  const numeric=v=>{if(v===null||v===undefined||v==='')return null;const n=Number(String(v).replace(',','.'));return Number.isFinite(n)?n:null;};
  return {code,name:cellText(row[offset+1]),nk:numeric(nk),kk:numeric(kk)};
}
function parseTemplate(book) {
  if (!book.SheetNames.includes('НТЗ')) throw new Error('В образце не найден лист «НТЗ» с нормативами.');
  const rows=getRows(book,'НТЗ');
  const h=mapHeaders(rows[0]);
  if(h['код товара']==null||h['нк тз для рц']==null||h['кк тз для рц']==null)
    throw new Error('Лист «НТЗ» должен содержать «Код товара», «НК ТЗ для РЦ», «КК ТЗ для РЦ».');
  const norms=Object.create(null);
  for(let i=1;i<rows.length;i++){
    const r=rows[i];if(!r)continue;
    const code=productCode(raw(r[h['код товара']]));if(!code)continue;
    const n=raw(r[h['нк тз для рц']]),k=raw(r[h['кк тз для рц']]);
    norms[code]={name:cellText(r[1]),nk:n==null||n===''?null:Core.num(n),kk:k==null||k===''?null:Core.num(k)};
  }
  const mainDate=archiveDate(raw(rows[0][4]))||'НТЗ (основной)';
  const snapshots=[];
  if(book.SheetNames.includes('Архив НТЗ')){
    const ar=getRows(book,'Архив НТЗ'),hdr=ar[0]||[];
    for(let i=0;i<hdr.length-3;i+=6){
      if(cellText(hdr[i]).toLowerCase()!=='код товара')continue;
      const archive=Object.create(null);
      for(let j=1;j<ar.length;j++){
        const row=ar[j];if(!row)continue;
        const rec=normRow(row,i);if(rec)archive[rec.code]={name:rec.name,nk:rec.nk,kk:rec.kk};
      }
      snapshots.push({date:archiveDate(raw(hdr[i+4]))||`Архив ${snapshots.length+1}`,norms:archive});
    }
  }
  const dated=book.SheetNames.filter(name=>/^\d{2}\.\d{2}\.\d{4}$/.test(name))
    .sort((a,b)=>Core.parseRuDate(b).localeCompare(Core.parseRuDate(a)));
  const catalog=[];
  if(dated.length){const daily=getRows(book,dated[0]);for(let i=1;i<daily.length;i++){
    const row=daily[i];if(!row)continue;const code=productCode(raw(row[0]));if(code)catalog.push(code);
  }}
  const flowerCodes=[];
  if(book.SheetNames.includes('код цветы')){
    const codes=getRows(book,'код цветы');
    for(let i=1;i<codes.length;i++){const c=productCode(raw(codes[i]?.[0]));if(c)flowerCodes.push(c);}
  }
  return {norms,mainDate,snapshots,flowerCodes:[...new Set(flowerCodes)],
    catalog:[...new Set(catalog)],catalogSheet:dated[0]||null};
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
