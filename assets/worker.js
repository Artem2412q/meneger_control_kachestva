/* All input files are processed inside the user's browser. Nothing is uploaded. */
importScripts('./xlsx-lite.js');
importScripts('./core.js');
const Core = self.MagnetCore;
const Lite = self.MagnitXlsxLite;
function reportProgress(message){self.postMessage({type:'progress',message});}
async function openBook(buffer){
  if(Lite.isXlsx(buffer))return Lite.open(buffer);
  // Legacy .xls is a binary BIFF format, not an XLSX ZIP. A compatible
  // parser is loaded only when needed; modern .xlsx works entirely offline.
  if(typeof XLSX==='undefined'){
    try{importScripts('https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js');}
    catch(e){throw Error('Старый формат .xls требует доступ к библиотеке чтения. Сохраните образец как .xlsx в Excel и загрузите снова. Новые .xlsx работают без интернета.');}
  }
  const book=XLSX.read(buffer,{type:'array',dense:true,cellDates:false,cellText:false,cellNF:false});
  return {names:book.SheetNames,async eachRow(name,callback){
    const sheet=book.Sheets[name||book.SheetNames[0]],rows=sheet?.['!data'];
    if(!Array.isArray(rows))throw Error(`Лист «${name}» не найден`);
    for(let i=0;i<rows.length;i++)callback(rows[i]||[],i);
  }};
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
// A blank source cell is missing data, not a confirmed numeric zero.
function nullableQuantity(value){
  return value===null || value===undefined || String(value).trim()==='' ? null : Core.num(value);
}
function productCode(v){ return Core.normalize(v); }
function serviceDate(value){
  if (typeof value === 'number') {
    const info=Lite.excelDate(value);
    if (info) return `${String(info.d).padStart(2,'0')}.${String(info.m).padStart(2,'0')}.${info.y}`;
  }
  const v=String(value||'').trim();
  if (Core.parseRuDate(v)) return Core.toRuDate(Core.parseRuDate(v));
  return v;
}
async function parseSaldo(book){
  let cols,group2;const out=Object.create(null),duplicates=[];
  await book.eachRow(null,(row,i)=>{
    if(i===0){const h=mapHeaders(row);cols=requireColumns(h,['Код товара','Название','Начальное количество','Конечное количество','ГР_ПНАЗВ'],'Сальдово-оборотная ведомость');group2=h['полное название группы 2'];return;}
    const code=productCode(value(row,cols,'Код товара'));if(!/^\d+$/.test(code))return;
    const item={name:value(row,cols,'Название')||'',start:nullableQuantity(value(row,cols,'Начальное количество')),
      end:nullableQuantity(value(row,cols,'Конечное количество')),group:value(row,cols,'ГР_ПНАЗВ')||'',group2:group2==null?'':(raw(row[group2])||'')};
    if(out[code])duplicates.push(code);
    out[code]=item;
  });
  return {items:out,duplicates};
}
async function parseService(book){
  let cols,rcIndex;const out=Object.create(null),dates=Object.create(null),rcs=new Set();
  await book.eachRow(null,(row,i)=>{
    if(i===0){const h=mapHeaders(row);cols=requireColumns(h,['Код товара','Наим_ТП','Заказ','Импорт','Отгрузка','Vicherk','Дата операции','Тип_Отбор'],'Уровень сервиса');rcIndex=h['рц'];return;}
    const code=productCode(value(row,cols,'Код товара'));if(!/^\d+$/.test(code))return;
    const date=serviceDate(value(row,cols,'Дата операции'));
    if(!Core.parseRuDate(date))throw Error(`Неизвестный формат даты операции в строке ${i+1}: ${date}`);
    dates[date]=(dates[date]||0)+1;
    if(rcIndex!==undefined&&raw(row[rcIndex])!=null)rcs.add(String(raw(row[rcIndex])));
    const key=`${date}|${code}`;
    if(!out[key])out[key]={code,date,name:value(row,cols,'Наим_ТП')||'',order:null,imp:null,ship:null,vich:null,fresh:false};
    const target=out[key];
    for(const [field,header] of [['order','Заказ'],['imp','Импорт'],['ship','Отгрузка'],['vich','Vicherk']]){
      const quantity=nullableQuantity(value(row,cols,header));
      if(quantity!==null)target[field]=(target[field]??0)+quantity;
    }
    if(String(value(row,cols,'Тип_Отбор')||'').toLowerCase().includes('овощи-фрукты'))target.fresh=true;
  });
  return {items:out,dates,rcs:[...rcs]};
}
function archiveDate(value){
  if(typeof value==='number'){
    const d=Lite.excelDate(value);if(d)return `${String(d.d).padStart(2,'0')}.${String(d.m).padStart(2,'0')}.${d.y}`;
  }
  return serviceDate(value);
}
function normRow(row,offset){
  const code=productCode(raw(row[offset]));if(!/^\d+$/.test(code))return null;
  const nk=raw(row[offset+2]),kk=raw(row[offset+3]);
  const numeric=v=>{if(v===null||v===undefined||v==='')return null;const n=Number(String(v).replace(',','.'));return Number.isFinite(n)?n:null;};
  return {code,name:cellText(row[offset+1]),nk:numeric(nk),kk:numeric(kk)};
}
async function parseTemplate(book){
  if(!book.names.includes('НТЗ'))throw Error('В образце не найден лист «НТЗ» с нормативами.');
  let header, cols, mainDate;const norms=Object.create(null);
  await book.eachRow('НТЗ',(row,i)=>{
    if(i===0){header=row;const h=mapHeaders(row);
      if(h['код товара']==null||h['нк тз для рц']==null||h['кк тз для рц']==null)throw Error('Лист «НТЗ» должен содержать «Код товара», «НК ТЗ для РЦ», «КК ТЗ для РЦ».');
      cols=h;mainDate=archiveDate(raw(row[4]))||'НТЗ (основной)';return;}
    const code=productCode(raw(row[cols['код товара']]));if(!code)return;
    const n=raw(row[cols['нк тз для рц']]),k=raw(row[cols['кк тз для рц']]);
    norms[code]={name:cellText(row[1]),nk:n==null||n===''?null:Core.num(n),kk:k==null||k===''?null:Core.num(k)};
  });
  const snapshots=[];
  if(book.names.includes('Архив НТЗ')){
    let archiveStarts=[];
    await book.eachRow('Архив НТЗ',(row,i)=>{
      if(i===0){
        for(let j=0;j<row.length-3;j+=6){if(cellText(row[j]).toLowerCase()!=='код товара')continue;
          const snap={date:archiveDate(raw(row[j+4]))||`Архив ${snapshots.length+1}`,norms:Object.create(null)};
          archiveStarts.push({offset:j,snap});snapshots.push(snap);
        }return;
      }
      for(const {offset,snap} of archiveStarts){const rec=normRow(row,offset);if(rec)snap.norms[rec.code]={name:rec.name,nk:rec.nk,kk:rec.kk};}
    });
  }
  const dated=book.names.filter(name=>/^\d{2}\.\d{2}\.\d{4}$/.test(name)).sort((a,b)=>Core.parseRuDate(b).localeCompare(Core.parseRuDate(a)));
  const catalog=[];
  if(dated.length)await book.eachRow(dated[0],(row,i)=>{if(!i)return;const code=productCode(raw(row[0]));if(code)catalog.push(code);});
  const flowerCodes=[];
  if(book.names.includes('код цветы'))await book.eachRow('код цветы',(row,i)=>{if(!i)return;const code=productCode(raw(row[0]));if(code)flowerCodes.push(code);});
  return {norms,mainDate,snapshots,flowerCodes:[...new Set(flowerCodes)],catalog:[...new Set(catalog)],catalogSheet:dated[0]||null};
}
async function parseGT(book){
  let headers=[];const records=[];
  await book.eachRow(null,(row,i)=>{
    if(i===0){const h=mapHeaders(row);requireColumns(h,['дата приемки','код ТП','Наименование ТП'],'Отчёт возвраты ГТ');headers=row.map(cellText);return;}
    const values=headers.map((_,j)=>raw(row[j]));if(values.some(v=>v!==null&&v!==undefined&&v!==''))records.push(values);
  });
  return {headers,records};
}
self.onmessage=async ({data}) => {
  try {
    if(data.action==='stocks') {
      reportProgress('Чтение сальдово-оборотной ведомости...');
      const saldo=await parseSaldo(await openBook(data.saldo));
      reportProgress('Обработка операций уровня сервиса и объединение кодов...');
      const service=await parseService(await openBook(data.service));
      self.postMessage({type:'stocks',payload:{saldo:saldo.items,service:service.items,
        meta:{saldoCount:Object.keys(saldo.items).length,serviceCount:Object.keys(service.items).length,
          duplicateSaldo:saldo.duplicates.slice(0,20), dates:service.dates, rcs:service.rcs}}});
    } else if(data.action==='template') {
      reportProgress('Чтение листа «НТЗ» и списка товаров из образца...');
      const payload=await parseTemplate(await openBook(data.buffer));
      self.postMessage({type:'template',payload});
    } else if(data.action==='gt') {
      reportProgress('Чтение реестра возвратов ГТ...');
      self.postMessage({type:'gt',payload:await parseGT(await openBook(data.buffer))});
    } else throw new Error('Неизвестная команда.');
  } catch(error) {
    self.postMessage({type:'error',message:error?.message||String(error)});
  }
};
