/* Offline XLSX reader for the MAGNIT report inputs.
   Streams worksheet XML from ZIP; only requested sheets are parsed and rows are
   consumed immediately. No spreadsheet data is sent over the network. */
(function(root,factory){
  const lib=factory();
  if(typeof module==='object'&&module.exports)module.exports=lib;
  root.MagnitXlsxLite=lib;
})(typeof self==='undefined'?globalThis:self,function(){
  'use strict';
  const utf8=new TextDecoder('utf-8');
  const ZIP_LOCAL=0x04034b50, ZIP_CENTRAL=0x02014b50, ZIP_END=0x06054b50;
  function attr(tag,name){const r=new RegExp('(?:\\s|^)'+name.replace(':','\\:')+'="([^"]*)"');return tag.match(r)?.[1]||'';}
  function decodeXml(s){
    return String(s||'').replace(/&#(?:x([\da-fA-F]+)|(\d+));/g,(_,hex,dec)=>{
      const n=parseInt(hex||dec,hex?16:10);
      try{return String.fromCodePoint(n);}catch(e){return '\uFFFD';}
    }).replace(/&(amp|lt|gt|quot|apos);/g,(_,e)=>({amp:'&',lt:'<',gt:'>',quot:'"',apos:"'"}[e]));
  }
  function colIndex(name){let n=0;for(const ch of name.toUpperCase())n=n*26+ch.charCodeAt(0)-64;return n-1;}
  function zipEntries(buffer){
    const u8=new Uint8Array(buffer),v=new DataView(buffer),min=Math.max(0,u8.length-65557);
    let end=-1;
    for(let p=u8.length-22;p>=min;p--)if(v.getUint32(p,true)===ZIP_END){end=p;break;}
    if(end<0)throw Error('Некорректный XLSX: не найден каталог ZIP.');
    const size=v.getUint16(end+10,true),off=v.getUint32(end+16,true);
    const entries=new Map();let pos=off;
    for(let i=0;i<size;i++){
      if(v.getUint32(pos,true)!==ZIP_CENTRAL)throw Error('Некорректная структура ZIP (запись каталога).');
      const method=v.getUint16(pos+10,true),compressed=v.getUint32(pos+20,true),
        filenameSize=v.getUint16(pos+28,true),extra=v.getUint16(pos+30,true),comment=v.getUint16(pos+32,true),
        local=v.getUint32(pos+42,true),name=utf8.decode(u8.subarray(pos+46,pos+46+filenameSize));
      entries.set(name.replace(/^\//,''),{method,compressed,local});
      pos+=46+filenameSize+extra+comment;
    }
    return {buffer,u8,v,entries};
  }
  function entryStream(zip,name){
    const it=zip.entries.get(name);
    if(!it)throw Error(`В XLSX отсутствует ${name}`);
    const p=it.local,v=zip.v;
    if(v.getUint32(p,true)!==ZIP_LOCAL)throw Error('Некорректная ZIP-запись.');
    const offset=p+30+v.getUint16(p+26,true)+v.getUint16(p+28,true);
    const bytes=zip.u8.subarray(offset,offset+it.compressed);
    if(it.method===0)return new Blob([bytes]).stream();
    if(it.method!==8)throw Error(`Неподдерживаемый алгоритм ZIP: ${it.method}`);
    if(typeof DecompressionStream==='undefined')throw Error('Браузер слишком старый для чтения XLSX. Обновите Chrome или Edge.');
    return new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  }
  async function entryText(zip,name){
    const reader=entryStream(zip,name).getReader(),d=new TextDecoder();let result='';
    while(true){const {done,value}=await reader.read();if(done)break;result+=d.decode(value,{stream:true});}
    return result+d.decode();
  }
  function workbookSheets(xml,rels){
    const map=new Map(),lookup=new Map();
    for(const m of rels.matchAll(/<(?:\w+:)?Relationship\b[^>]*\/?\s*>/g)){
      const rid=attr(m[0],'Id'),target=attr(m[0],'Target');
      if(!rid||!target)continue;
      let path=target.replace(/^\//,'');
      if(!path.startsWith('xl/'))path='xl/'+path;
      path=path.replace(/\/\.\//g,'/');lookup.set(rid,path);
    }
    for(const m of xml.matchAll(/<(?:\w+:)?sheet\b[^>]*\/?\s*>/g)){
      const name=decodeXml(attr(m[0],'name'));
      const id=attr(m[0],'r:id');
      if(name)map.set(name,lookup.get(id)||`xl/worksheets/sheet${map.size+1}.xml`);
    }
    return map;
  }
  async function loadSharedStrings(zip){
    if(!zip.entries.has('xl/sharedStrings.xml'))return [];
    const xml=await entryText(zip,'xl/sharedStrings.xml');
    const items=[];
    for(const item of xml.matchAll(/<(?:\w+:)?si(?:\s[^>]*)?>([\s\S]*?)<\/(?:\w+:)?si>/g)){
      const parts=[...item[1].matchAll(/<(?:\w+:)?t(?:\s[^>]*)?>([\s\S]*?)<\/(?:\w+:)?t>/g)].map(m=>decodeXml(m[1]));
      items.push(parts.join(''));
    }
    return items;
  }
  function parseRow(xml,shared){
    const row=[];
    // Supports numeric cells, shared strings, inline strings, cached formula results.
    for(const c of xml.matchAll(/<(?:\w+:)?c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?c>)/g)){
      const r=attr(c[1],'r'),letters=r.match(/^[A-Za-z]+/)?.[0];
      if(!letters)continue;
      const idx=colIndex(letters),type=attr(c[1],'t'),content=c[2]||'';
      let value;
      if(type==='inlineStr'){
        value=[...content.matchAll(/<(?:\w+:)?t(?:\s[^>]*)?>([\s\S]*?)<\/(?:\w+:)?t>/g)].map(m=>decodeXml(m[1])).join('');
      }else{
        const v=content.match(/<(?:\w+:)?v(?:\s[^>]*)?>([\s\S]*?)<\/(?:\w+:)?v>/)?.[1];
        if(v===undefined)continue;
        if(type==='s')value=shared[Number(v)]||'';
        else if(type==='str'||type==='e'||type==='d')value=decodeXml(v);
        else if(type==='b')value=v==='1';
        else {const n=Number(v);value=Number.isFinite(n)?n:decodeXml(v);}
      }
      row[idx]={v:value};
    }
    return row;
  }
  async function eachRow(zip,path,shared,onRow){
    const reader=entryStream(zip,path).getReader(),decoder=new TextDecoder('utf-8');
    let buffer='',count=0;
    while(true){
      const {done,value}=await reader.read();
      buffer+=done?decoder.decode():decoder.decode(value,{stream:true});
      while(true){
        const start=buffer.search(/<(?:\w+:)?row(?:\s|>)/);
        if(start<0){buffer=buffer.slice(-20);break;}
        if(start>0)buffer=buffer.slice(start);
        const endTag=buffer.match(/<\/(?:\w+:)?row>/);
        if(!endTag)break;
        const end=endTag.index, endLen=endTag[0].length;
        const rowXml=buffer.slice(0,end+endLen);
        buffer=buffer.slice(end+endLen);
        onRow(parseRow(rowXml,shared),count++);
      }
      if(done)break;
    }
    return count;
  }
  async function open(buffer){
    const zip=zipEntries(buffer);
    const sheets=workbookSheets(await entryText(zip,'xl/workbook.xml'),await entryText(zip,'xl/_rels/workbook.xml.rels'));
    const shared=await loadSharedStrings(zip);
    return {
      names:[...sheets.keys()],
      async eachRow(name,cb){
        const path=sheets.get(name||[...sheets.keys()][0]);
        if(!path)throw Error(`Лист «${name}» не найден`);
        return eachRow(zip,path,shared,cb);
      }
    };
  }
  function isXlsx(buffer){return new DataView(buffer).getUint32(0,true)===ZIP_LOCAL;}
  function excelDate(n){const d=new Date(Date.UTC(1899,11,30)+Math.round(n*86400000));return {y:d.getUTCFullYear(),m:d.getUTCMonth()+1,d:d.getUTCDate()};}
  return {open,isXlsx,parseRow,decodeXml,zipEntries,excelDate};
});
