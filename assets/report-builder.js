/* Business reporting workbook layout, kept independent of browser File APIs. */
(function(root,factory){const lib=factory(root.MagnetCore||(typeof require==='function'?require('./core.js'):null));if(typeof module==='object'&&module.exports)module.exports=lib;root.MagnitReport=lib;})(typeof window==='undefined'?globalThis:window,function(Core){
'use strict';
const CRITICAL_HEADERS=Core.HEADERS.filter((_,i)=>i!==8);
const NO_NORM='НЕТ В НТЗ';
function formatCode(code){return String(code||'').replace(/\.0$/,'');}
function makeNormSheet(template,chosenNorms){
 const label=template?(template.chosenDate||template.mainDate||''):'Не загружено';
 return {name:'НТЗ',headers:['Код товара','Наименование товара','НК\nТЗ для РЦ','КК\nТЗ для РЦ',`Срез: ${label}`],rows:Object.entries(chosenNorms).map(([code,n])=>[formatCode(code),n.name||'',n.nk,n.kk,'']),widths:[17,64,18,18,24]};
}
function makeArchive(template){
 const snapshots=template?.snapshots||[],headers=[],rows=[];
 for(let i=0;i<snapshots.length;i++){
  const off=i*6;headers[off]='Код товара';headers[off+1]='Наименование товара';headers[off+2]='НК\nТЗ для РЦ';headers[off+3]='КК\nТЗ для РЦ';headers[off+4]=snapshots[i].date;
  let index=0;for(const [code,n] of Object.entries(snapshots[i].norms||{})){
   if(!rows[index])rows[index]=[];
   rows[index][off]=formatCode(code);rows[index][off+1]=n.name||'';rows[index][off+2]=n.nk;rows[index][off+3]=n.kk;index++;
  }
 }
 if(!snapshots.length) return {name:'Архив НТЗ',headers:['Архив НТЗ отсутствует в образце'],rows:[],widths:[35]};
 const len=(snapshots.length-1)*6+5;
 return {name:'Архив НТЗ',headers:Array.from({length:len},(_,i)=>headers[i]||''),rows:rows.map(r=>Array.from({length:len},(_,i)=>r?.[i]??'')),
  widths:Array.from({length:len},(_,i)=>i%6===1?53:i%6===5?4:19)};
}
function makeDaySheet(report,{critical=false, flowers=false,normLength=1}={}){
 const rows=report.records.map((r,i)=>{
  const n=i+2,nn=Math.max(2,normLength+1);
  const ref=idx=>`IFERROR(IF(VLOOKUP(A${n},'НТЗ'!$A$2:$D$${nn},${idx},FALSE)=\"\",\"${NO_NORM}\",VLOOKUP(A${n},'НТЗ'!$A$2:$D$${nn},${idx},FALSE)),\"${NO_NORM}\")`;
  const nk={f:ref(3),v:r.nk===null?NO_NORM:r.nk,missing:r.nk===null},kk={f:ref(4),v:r.kk===null?NO_NORM:r.kk,missing:r.kk===null};
  if(!critical&&!flowers){
   return [r.code,r.name,r.start,r.order,r.imp,r.ship,{f:`IFERROR(F${n}/D${n},0)`,v:r.usOrder},{f:`IFERROR(F${n}/E${n},0)`,v:r.usImport},r.vich,r.ending,nk,kk,
   {f:`IFERROR(J${n}/F${n},0)`,v:r.actualDays},{f:`IFERROR(J${n}/D${n},0)`,v:r.orderDays},
   {f:`IF(D${n}=0,\"Без расхода\",IF(ISNUMBER(K${n}),M${n}-K${n},\"Нет НТЗ\"))`,v:r.order===0?'Без расхода':r.nk===null?'Нет НТЗ':r.targetComment},
   {f:`IF(D${n}=0,\"Без расхода\",IF(ISNUMBER(L${n}),M${n}-L${n},\"Нет НТЗ\"))`,v:r.order===0?'Без расхода':r.kk===null?'Нет НТЗ':r.maxComment}];
  }
  return [r.code,r.name,r.start,r.order,r.imp,r.ship,{f:`IFERROR(F${n}/D${n},0)`,v:r.usOrder},{f:`IFERROR(F${n}/E${n},0)`,v:r.usImport},r.ending,nk,kk,
   {f:`IFERROR(I${n}/F${n},0)`,v:r.actualDays},{f:`IFERROR(I${n}/D${n},0)`,v:r.orderDays},
   {f:`IF(D${n}=0,\"Без расхода\",IF(ISNUMBER(J${n}),L${n}-J${n},\"Нет НТЗ\"))`,v:r.order===0?'Без расхода':r.nk===null?'Нет НТЗ':r.targetComment},
   {f:`IF(D${n}=0,\"Без расхода\",IF(ISNUMBER(K${n}),L${n}-K${n},\"Нет НТЗ\"))`,v:r.order===0?'Без расхода':r.kk===null?'Нет НТЗ':r.maxComment}];
 });
 const isMain=!critical&&!flowers;
 return {name:flowers?'Цветы':critical?`Критика ${report.date}`:report.date,type:flowers?'flowers':critical?'critical':'report',headers:isMain?Core.HEADERS:CRITICAL_HEADERS,rows,
  widths:isMain?[17,62,20,15,15,15,22,22,17,19,23,23,21,18,27,29]:[17,62,20,15,15,15,22,22,19,23,23,21,18,27,29]};
}
function makeUsSheet(sources,date){
 const us=Object.values(sources?.service||{}).filter(x=>x.date===date).sort((a,b)=>String(a.code).localeCompare(String(b.code)));
 return {name:'УС',headers:['Код товара','Заказ','Импорт','Отгрузка','Vicherk','Дата операции'],rows:us.map(r=>[r.code,r.order,r.imp,r.ship,r.vich,r.date]),widths:[19,20,20,20,20,20]};
}
function makeControl(stock,critical,template,selectedLabel){
 const rows=[['Выбранный срез нормативов',selectedLabel,'НК/КК в отчёте связаны формулами VLOOKUP с листом «НТЗ»'],
 ['Исторических срезов НТЗ',template?.snapshots?.length||0,'Сохранены в листе «Архив НТЗ»'],
 ['Кодов цветов',template?.flowerCodes?.length||0,'Лист «код цветы»; лист «Цветы» формируется по совпадениям'],
 ['Дата товарного запаса',stock?.date||'—','Историческая дата возможна'],
 ['Дата критики',critical?.date||'не сформирована','Только текущий день, файлы загружаются отдельно'],
 ['Отбор для критики','Овощи и фрукты (предварительно)','Точные правила отбора требуют сверки с рабочим примером']];
 for(const [section,rep] of [['Товарный запас',stock],['Критика',critical]]){
  if(!rep)continue;
  rows.push([`${section}: товаров`,rep.records.length,'']);
  rows.push([`${section}: нет НК/КК`,rep.stats.normMissing,'Не подставляем 0; отсутствующие нормативы отмечены в отчёте']);
  rows.push([`${section}: нет сальдо`,rep.stats.saldoMissing,'Недостающие позиции отдельно проверяются']);
  for(const r of rep.records){
   if(r.missingNorm)rows.push([`${section}: НЕТ В НТЗ`,r.code,r.name]);
   if(r.missingSaldo)rows.push([`${section}: НЕТ В САЛЬДО`,r.code,r.name]);
  }
 }
 return {name:'Контроль',headers:['Проверка','Значение или код','Комментарий'],rows,widths:[42,28,81]};
}
function prepareSheets({stock=null,critical=null,stockSources=null,template=null,snapshot='main'}={}){
 if(!template||!Object.keys(template.norms||{}).length)throw new Error('Сначала загрузите образец с листом «НТЗ». Без НТЗ отчёт нельзя выгружать.');
 if(!stock&&!critical)throw new Error('Нет подготовленного отчёта.');
 const norm=Core.selectedNorms(template,snapshot);
 const selectedLabel=snapshot==='main'?`Основной НТЗ (${template.mainDate||'без даты'})`:`Архив ${template.snapshots?.[Number(snapshot)]?.date||'без даты'}`;
 const normCount=Object.keys(norm).length;
 const sheets=[];
 if(stock)sheets.push(makeDaySheet(stock,{normLength:normCount}));
 if(critical)sheets.push(makeDaySheet(critical,{critical:true,normLength:normCount}));
 sheets.push(makeNormSheet({...template,chosenDate:selectedLabel},norm));
 sheets.push(makeArchive(template));
 sheets.push({name:'код цветы',headers:['Код'],rows:(template.flowerCodes||[]).map(c=>[c]),widths:[22]});
 if(stock&&stockSources){
  const flowerReport=Core.buildReport({saldo:stockSources.saldo,service:stockSources.service,flowerCodes:template.flowerCodes||[],norms:norm},{date:stock.date,mode:'flowers'});
  sheets.push(makeDaySheet(flowerReport,{flowers:true,normLength:normCount}));
  sheets.push(makeUsSheet(stockSources,stock.date));
 }
 sheets.push(makeControl(stock,critical,template,selectedLabel));
 return sheets;
}
return {prepareSheets,makeDaySheet,makeArchive,makeNormSheet,NO_NORM};
});
