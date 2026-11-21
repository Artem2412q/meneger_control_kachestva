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
  const ref=idx=>`IFERROR(IF(VLOOKUP(A${n},'НТЗ'!$A$2:$D$${nn},${idx},FALSE)="","${NO_NORM}",VLOOKUP(A${n},'НТЗ'!$A$2:$D$${nn},${idx},FALSE)),"${NO_NORM}")`;
  const nk={f:ref(3),v:r.nk===null?NO_NORM:r.nk,missing:r.nk===null},kk={f:ref(4),v:r.kk===null?NO_NORM:r.kk,missing:r.kk===null};
  // Missing data must stay blank in the final Excel and in the cached formula values.
  const usOrder={f:`IF(OR(COUNTBLANK(D${n})>0,COUNTBLANK(F${n})>0),"",IFERROR(F${n}/D${n},0))`,v:r.usOrder??''};
  const usImport={f:`IF(OR(COUNTBLANK(E${n})>0,COUNTBLANK(F${n})>0),"",IFERROR(F${n}/E${n},0))`,v:r.usImport??''};
  const stockCol=critical||flowers?'I':'J';
  const actualCol=critical||flowers?'L':'M';
  const orderCol=critical||flowers?'M':'N';
  const nkCol=critical||flowers?'J':'K';
  const kkCol=critical||flowers?'K':'L';
  const days={f:`IF(OR(COUNTBLANK(${stockCol}${n})>0,COUNTBLANK(F${n})>0),"",IFERROR(${stockCol}${n}/F${n},0))`,v:r.actualDays??''};
  const orderDays={f:`IF(OR(COUNTBLANK(${stockCol}${n})>0,COUNTBLANK(D${n})>0),"",IFERROR(${stockCol}${n}/D${n},0))`,v:r.orderDays??''};
  const comment=which=>{
   const normCol=which==='nk'?nkCol:kkCol;
   const val=which==='nk'?r.targetComment:r.maxComment;
   const fx=`IF(OR(COUNTBLANK(${stockCol}${n})>0,COUNTBLANK(D${n})>0,COUNTBLANK(F${n})>0),"Нет данных",IF(D${n}=0,"Без расхода",IF(ISNUMBER(${normCol}${n}),${actualCol}${n}-${normCol}${n},"Нет НТЗ")))`;
   return {f:fx,v:val};
  };
  if(!critical&&!flowers){
   return [r.code,r.name,r.start,r.order,r.imp,r.ship,usOrder,usImport,r.vich,r.ending,nk,kk,days,orderDays,comment('nk'),comment('kk')];
  }
  return [r.code,r.name,r.start,r.order,r.imp,r.ship,usOrder,usImport,r.ending,nk,kk,days,orderDays,comment('nk'),comment('kk')];
 });
 const isMain=!critical&&!flowers;
 return {name:flowers?'Цветы':critical?`Критика ${report.date}`:report.date,type:flowers?'flowers':critical?'critical':'report',headers:isMain?Core.HEADERS:CRITICAL_HEADERS,rows,
  widths:isMain?[17,62,20,15,15,15,22,22,17,19,23,23,21,18,27,29]:[17,62,20,15,15,15,22,22,19,23,23,21,18,27,29]};
}
function makeEtalonSheet(etalon,stock,stockSources,norm){
 const allCodes=etalon?.codes||[];
 const byCode=new Map((stock?.records||[]).map(r=>[r.code,r]));
 const date=stock?.date;
 const serviceRecords=Object.values(stockSources?.service||{}).filter(r=>r.date===date);
 const serviceDate=new Set(serviceRecords.map(r=>r.code));
 const serviceByCode=new Map(serviceRecords.map(r=>[r.code,r]));
 const hasSources=Boolean(stockSources);
 // Only traceable daily items belong in the generated reconciliation sheet.
 // The complete master stays in assets/etalon.js, not in this daily Excel.
 const codes=allCodes.filter(code=>hasSources?Core.hasFacts(stockSources.saldo?.[code],serviceByCode.get(code)):byCode.has(code));
 const rows=codes.map(code=>{
  const saldo=stockSources?.saldo?.[code], service=serviceDate.has(code), rule=norm?.[code];
  const record=byCode.get(code);
  const status=!hasSources?'Не проверено':!saldo?'Нет в сальдо':!service?'Нет в УС':'ОК';
  const names=record?.name||saldo?.name||rule?.name||'';
  return [code,names,hasSources?(saldo?'Да':'Нет'):'—',hasSources?(service?'Да':'Нет'):'—',rule?.nk??NO_NORM,rule?.kk??NO_NORM,status,code.length===10?'':`Проверьте: ${code.length} цифр`];
 });
 return {name:'Эталон ТП',headers:['Код товара','Название','Есть в сальдо','Есть в УС за дату','НК (из НТЗ)','КК (из НТЗ)','Сверка','Контроль кода'],rows,widths:[19,60,20,20,20,20,23,27]};
}
function makeUsSheet(sources,date){
 const us=Object.values(sources?.service||{}).filter(x=>x.date===date && Core.hasFacts(null,x)).sort((a,b)=>String(a.code).localeCompare(String(b.code)));
 return {name:'УС',headers:['Код товара','Заказ','Импорт','Отгрузка','Vicherk','Дата операции'],rows:us.map(r=>[r.code,r.order,r.imp,r.ship,r.vich,r.date]),widths:[19,20,20,20,20,20]};
}
function makeControl(stock,critical,template,selectedLabel,etalon){
 const rows=[['Выбранный срез нормативов',selectedLabel,'НК/КК в отчёте связаны формулами VLOOKUP с листом «НТЗ»'],
 ['Строк исходного Эталона',etalon?.sourceLines||0,'В том числе повторы; в отчёте каждый код один раз'],
 ['Уникальных ТП в Эталоне',etalon?.codes?.length||0,'Справочник допустимых кодов, не источник ежедневных строк'],
 ['Повторов в Эталоне',etalon?.duplicates||0,'Исключены из итоговых строк'],
 ['Кодов нестандартной длины',etalon?.anomalies?.length||0,'Не исправлены автоматически; проверьте в листе «Эталон ТП»'],
 ['Исторических срезов НТЗ',template?.snapshots?.length||0,'Сохранены в листе «Архив НТЗ»'],
 ['Кодов цветов',template?.flowerCodes?.length||0,'Лист «код цветы»; лист «Цветы» формируется по совпадениям'],
 ['Дата товарного запаса',stock?.date||'—','Историческая дата возможна'],
 ['Дата критики',critical?.date||'не сформирована','Только текущий день, файлы загружаются отдельно'],
 ['Отбор для критики',critical?.stats?.appliedMode||'не сформирована','Выбран пользователем; бизнес-правила критических ТП уточняются']];
 for(const [section,rep] of [['Товарный запас',stock],['Критика',critical]]){
  if(!rep)continue;
  rows.push([`${section}: товаров`,rep.records.length,'Только товары с фактическими остатками или движением']);
  rows.push([`${section}: пустых позиций исключено`,rep.stats.excludedEmpty||0,'Нет ненулевых остатков, заказа, импорта, отгрузки или Vicherk']);
  rows.push([`${section}: кодов Эталона без ежедневных данных`,rep.stats.etalonNotInSources||0,'Исключены из ежедневных листов и сверки; не создаём фиктивные строки']);
  rows.push([`${section}: нет НК/КК`,rep.stats.normMissing,'Не подставляем 0; отсутствующие нормативы отмечены в отчёте']);
  rows.push([`${section}: нет сальдо`,rep.stats.saldoMissing,'Недостающие позиции отдельно проверяются']);
  rows.push([`${section}: нет уровня сервиса`,rep.stats.serviceMissing,'Не подставляем ноль вместо отсутствующих операций']);
  rows.push([`${section}: нет в обеих выгрузках`,rep.stats.missingBoth,'Такие строки запрещены в итоговом отчёте']);
  if(rep.stats.outsideEtalon?.length)rows.push([`${section}: коды вне Эталона`,rep.stats.outsideEtalon.length,'Из исходных файлов, отсутствуют в Эталоне']);
  for(const code of rep.stats.outsideEtalon||[])rows.push([`${section}: ВНЕ ЭТАЛОНА`,code,'Есть в исходной выгрузке, но не включён в режим «Эталон»']);
  for(const r of rep.records){
   if(r.missingNorm)rows.push([`${section}: НЕТ В НТЗ`,r.code,r.name]);
   if(r.missingSaldo)rows.push([`${section}: НЕТ В САЛЬДО`,r.code,r.name]);
   if(r.missingService)rows.push([`${section}: НЕТ В УС`,r.code,r.name]);
  }
 }
 return {name:'Контроль',headers:['Проверка','Значение или код','Комментарий'],rows,widths:[42,28,81]};
}
function prepareSheets({stock=null,critical=null,stockSources=null,criticalSources=null,template=null,etalon=null,snapshot='main'}={}){
 if(!template||!Object.keys(template.norms||{}).length)throw new Error('Встроенный НТЗ не загрузился. Проверьте файл assets/norms.js в версии сайта 2.7.');
 if(!stock&&!critical)throw new Error('Нет подготовленного отчёта.');
 const norm=Core.selectedNorms(template,snapshot);
 const selectedLabel=snapshot==='main'?`Основной НТЗ (${template.mainDate||'без даты'})`:`Архив ${template.snapshots?.[Number(snapshot)]?.date||'без даты'}`;
 const normCount=Object.keys(norm).length;
 const sheets=[];
 if(stock)sheets.push(makeDaySheet(stock,{normLength:normCount}));
 if(critical)sheets.push(makeDaySheet(critical,{critical:true,normLength:normCount}));
 if(etalon?.codes?.length)sheets.push(makeEtalonSheet(etalon,stock||critical,stockSources||criticalSources,norm));
 sheets.push(makeNormSheet({...template,chosenDate:selectedLabel},norm));
 sheets.push(makeArchive(template));
 sheets.push({name:'код цветы',headers:['Код'],rows:(template.flowerCodes||[]).map(c=>[c]),widths:[22]});
 if(stock&&stockSources){
  const flowerReport=Core.buildReport({saldo:stockSources.saldo,service:stockSources.service,flowerCodes:template.flowerCodes||[],norms:norm},{date:stock.date,mode:'flowers'});
  sheets.push(makeDaySheet(flowerReport,{flowers:true,normLength:normCount}));
  sheets.push(makeUsSheet(stockSources,stock.date));
 }
 sheets.push(makeControl(stock,critical,template,selectedLabel,etalon));
 return sheets;
}
return {prepareSheets,makeDaySheet,makeArchive,makeNormSheet,makeEtalonSheet,NO_NORM};
});
