/* Local XLSX writer with real Excel colors, formulas and original NTZ archive.
   ZIP files are stored (no compression) to avoid additional external dependencies. */
(function(root,factory){const lib=factory();if(typeof module==='object'&&module.exports)module.exports=lib;root.MagnitExcel=lib;})(typeof window==='undefined'?globalThis:window,function(){
'use strict';
const esc=x=>String(x==null?'':x).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;');
function colName(index){let s='';for(let n=index+1;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;}
const enc=new TextEncoder();
const color={blue:'238FFA',blueBody:'E9F3FF',green:'B4E67D',greenBody:'EDF9E1',ink:'112A44',white:'FFFFFF',red:'FFD9D9',grey:'EEF2F7'};
function stylesXml(){
 const fills=['<fill><patternFill patternType="none"/></fill>','<fill><patternFill patternType="gray125"/></fill>',...[
 color.blue,color.blueBody,color.green,color.greenBody,color.red,color.grey].map(c=>`<fill><patternFill patternType="solid"><fgColor rgb="FF${c}"/><bgColor indexed="64"/></patternFill></fill>`)].join('');
 const fonts=['<font><sz val="10"/><color rgb="FF1F2937"/><name val="Aptos"/></font>',
 '<font><b/><sz val="10"/><color rgb="FF0B1525"/><name val="Aptos"/></font>',
 '<font><b/><sz val="10"/><color rgb="FF16391C"/><name val="Aptos"/></font>'].join('');
 const borders='<border><left/><right/><top/><bottom/><diagonal/></border><border><left style="hair"><color rgb="FFD8E2EC"/></left><right style="hair"><color rgb="FFD8E2EC"/></right><top style="hair"><color rgb="FFD8E2EC"/></top><bottom style="hair"><color rgb="FFD8E2EC"/></bottom><diagonal/></border>';
 const xf=(fill,font=0,number=0,wrap=false)=>`<xf numFmtId="${number}" fontId="${font}" fillId="${fill}" borderId="1" xfId="0" applyFill="1" applyAlignment="1"${number?' applyNumberFormat="1"':''}><alignment horizontal="center" vertical="center"${wrap?' wrapText="1"':''}/></xf>`;
 const xfs=[xf(0),xf(3,0),xf(2,1,0,true),xf(4,2,0,true),xf(5,0),xf(6,0),xf(7,0),xf(2,1,0,true),xf(2,1,10,true),xf(3,0,2),xf(5,0,2),xf(6,0,2),xf(3,0,10),xf(5,0,10),xf(0,0,2),xf(0,0,10)].join('');
 return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="3">${fonts}</fonts><fills count="8">${fills}</fills><borders count="2">${borders}</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="16">${xfs}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
}
function cellXml(value,row,col,style){
 const ref=colName(col)+row;
 let formula=null, v=value;
 if(value&&typeof value==='object'&&!Array.isArray(value)){formula=value.f;v=value.v;if(Number.isInteger(value.s))style=value.s;}
 if(v===undefined||v===null)return '';
 const attr=` r="${ref}" s="${style||0}"`;
 if(formula){const fx=String(formula).replace(/^=/,'');
  if(typeof v==='number'&&Number.isFinite(v))return `<c${attr}><f>${esc(fx)}</f><v>${v}</v></c>`;
  return `<c${attr} t="str"><f>${esc(fx)}</f><v>${esc(v)}</v></c>`;
 }
 if(typeof v==='number'&&Number.isFinite(v))return `<c${attr}><v>${v}</v></c>`;
 return `<c${attr} t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`;
}
function sheetXml(def){
 const head=def.headers||[],rows=def.rows||[];
 const all=[head,...rows];const maxCol=Math.max(head.length,...rows.slice(0,10).map(r=>r.length),1);
 let table='';for(let ri=0;ri<all.length;ri++){
   const isHead=ri===0, row=all[ri];if(!row)continue;
   const isReport=['report','critical','flowers'].includes(def.type);
   const isMain=def.type==='report';
   const maxComment=(row[isMain?15:14]&&typeof row[isMain?15:14]==='object')?row[isMain?15:14].v:row[isMain?15:14];
   const overMax=(typeof maxComment==='number'&&maxComment>0)||maxComment==='Без расхода';
   let contents='';for(let c=0;c<row.length;c++){
     const item=row[c];
     let style=isHead?(isReport?((isMain?[10,11,14,15]:[9,10,13,14]).includes(c)?3:2):7):(c===0?1:0);
     if(!isHead&&isReport){
       // Excel built-in format 2 = 0.00; format 10 = 0.00%.
       // The original formula/cache retains full precision, only the visible format changes.
       if(c>1)style=14;
       if(c===6||c===7)style=15;
       const normCols=isMain?[10,11]:[9,10];
       const commentCols=isMain?[14,15]:[13,14];
       const redCols=isMain?[13,14,15]:[12,13,14];
       if(normCols.includes(c))style=(item&&typeof item==='object'&&item.missing)?11:10;
       if(commentCols.includes(c))style=10;
       if(overMax&&redCols.includes(c))style=11;
     }
     contents+=cellXml(item,ri+1,c,style);
   }
   if(contents)table+=`<row r="${ri+1}" ht="${isHead?58:19}" customHeight="1">${contents}</row>`;
 }
 const widths=def.widths||[];
 const cols=Array.from({length:maxCol},(_,c)=>`<col min="${c+1}" max="${c+1}" width="${widths[c]|| (c===1?57:c===0?17:18)}" customWidth="1"/>`).join('');
 const ref=`A1:${colName(maxCol-1)}${Math.max(all.length,1)}`;
 const filter=def.noFilter?'':`<autoFilter ref="${ref}"/>`;
 return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="${ref}"/><sheetViews><sheetView workbookViewId="0"><pane xSplit="2" ySplit="1" topLeftCell="C2" activePane="bottomRight" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="18"/><cols>${cols}</cols><sheetData>${table}</sheetData>${filter}</worksheet>`;
}
const crcTable=(()=>{const a=new Uint32Array(256);for(let i=0;i<256;i++){let c=i;for(let j=0;j<8;j++)c=c&1?0xEDB88320^(c>>>1):c>>>1;a[i]=c>>>0;}return a;})();
function crc32(bytes){let n=0xFFFFFFFF;for(const b of bytes)n=crcTable[(n^b)&0xFF]^(n>>>8);return (n^0xFFFFFFFF)>>>0;}
function zipStored(files){
 let offset=0;const chunks=[],central=[];
 for(const [filename,source] of files){
  const name=enc.encode(filename),data=typeof source==='string'?enc.encode(source):source;
  const crc=crc32(data),local=new Uint8Array(30+name.length),dv=new DataView(local.buffer);
  dv.setUint32(0,0x04034b50,true);dv.setUint16(4,20,true);dv.setUint16(6,0,true);dv.setUint16(8,0,true);dv.setUint32(14,crc,true);dv.setUint32(18,data.length,true);dv.setUint32(22,data.length,true);dv.setUint16(26,name.length,true);local.set(name,30);
  chunks.push(local,data);
  const ce=new Uint8Array(46+name.length),cdv=new DataView(ce.buffer);cdv.setUint32(0,0x02014b50,true);cdv.setUint16(4,20,true);cdv.setUint16(6,20,true);cdv.setUint32(16,crc,true);cdv.setUint32(20,data.length,true);cdv.setUint32(24,data.length,true);cdv.setUint16(28,name.length,true);cdv.setUint32(42,offset,true);ce.set(name,46);central.push(ce);
  offset+=local.length+data.length;
 }
 let centralSize=0;for(const item of central){chunks.push(item);centralSize+=item.length;}
 const eocd=new Uint8Array(22),end=new DataView(eocd.buffer);end.setUint32(0,0x06054b50,true);end.setUint16(8,files.length,true);end.setUint16(10,files.length,true);end.setUint32(12,centralSize,true);end.setUint32(16,offset,true);chunks.push(eocd);
 return new Blob(chunks,{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}
function makeWorkbook(sheetDefs){
 const sheets=sheetDefs.slice(0,80).map((s,i)=>({...s,name:String(s.name||`Лист ${i+1}`).slice(0,31)}));
 const base='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
 const workbook=`${base}<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets>${sheets.map((s,i)=>`<sheet name="${esc(s.name)}" sheetId="${i+1}" r:id="rId${i+1}"/>`).join('')}</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`;
 const relationships=`${base}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_,i)=>`<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i+1}.xml"/>`).join('')}<Relationship Id="rId${sheets.length+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;
 const types=`${base}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets.map((_,i)=>`<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>`;
 const rels=`${base}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
 const entries=[['[Content_Types].xml',types],['_rels/.rels',rels],['xl/workbook.xml',workbook],['xl/_rels/workbook.xml.rels',relationships],['xl/styles.xml',stylesXml()]];
 sheets.forEach((s,i)=>entries.push([`xl/worksheets/sheet${i+1}.xml`,sheetXml(s)]));return zipStored(entries);
}
function save(sheets,filename){const blob=makeWorkbook(sheets);const url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=filename;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),30000);}
return {makeWorkbook,save,sheetXml,stylesXml};
});
