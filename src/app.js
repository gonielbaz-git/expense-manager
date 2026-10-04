
/* =============================================================
   ניהול הוצאות חכם — single-file personal expense manager
   Data model: state.months['YYYY-MM'] = {rows, incomes, recon}
   ============================================================= */
'use strict';

/* ==PARSERS== (this block is also exercised by the Node test harness) */

/* ---------- tiny utils ---------- */
function normName(s){ return String(s||'').toLowerCase().replace(/[\d"'().,*#?!|/\\-]/g,' ').replace(/\s+/g,' ').trim(); }
function round2(x){ return Math.round((x+Number.EPSILON)*100)/100; }
function pad2(n){ return String(n).padStart(2,'0'); }
const HEB_MONTHS = ['ינואר','פברואר','מרץ','אפריל','מאי','יוני','יולי','אוגוסט','ספטמבר','אוקטובר','נובמבר','דצמבר'];
function monthKeyFromDate(d){ return d.y + '-' + pad2(d.m); }
function monthLabel(key){ const [y,m]=key.split('-').map(Number); return HEB_MONTHS[m-1]+' '+y; }
function keyFromHebLabel(label){
  const m = String(label||'').trim().match(/^(\S+)\s+(\d{4})$/);
  if(!m) return null;
  const idx = HEB_MONTHS.indexOf(m[1]);
  if(idx<0) return null;
  return m[2]+'-'+pad2(idx+1);
}
function prevMonthKey(key){ let [y,m]=key.split('-').map(Number); m--; if(m===0){m=12;y--;} return y+'-'+pad2(m); }
/* parse dd.mm.yy / dd.mm.yyyy / dd-mm-yyyy → {y,m,d} or null */
function parseDMY(s){
  const m = String(s||'').trim().match(/^(\d{1,2})[.\-\/](\d{1,2})[.\-\/](\d{2,4})$/);
  if(!m) return null;
  let y = Number(m[3]); if(y<100) y += 2000;
  const mo=Number(m[2]), d=Number(m[1]);
  if(mo<1||mo>12||d<1||d>31) return null;
  return {y:y,m:mo,d:d};
}
function fmtDMY(dt){ return pad2(dt.d)+'.'+pad2(dt.m)+'.'+dt.y; }
function excelSerialToDMY(n){
  const ms = Math.round((n - 25569) * 86400 * 1000); // days since 1970-01-01 (excel epoch 1899-12-30)
  const dt = new Date(ms);
  return {y:dt.getUTCFullYear(), m:dt.getUTCMonth()+1, d:dt.getUTCDate()};
}
function parseAmount(v){
  if(typeof v === 'number') return v;
  const s = String(v||'').replace(/[₪$€,\s]/g,'');
  const f = parseFloat(s);
  return isNaN(f) ? null : f;
}
function uid(){ return Math.random().toString(36).slice(2,10); }

/* ---------- CSV parser (handles quotes + "" escapes) ---------- */
function parseCSV(text){
  text = text.replace(/^﻿/,'');
  const rows=[]; let row=[], cur='', inQ=false;
  for(let i=0;i<text.length;i++){
    const c=text[i];
    if(inQ){
      if(c==='"'){ if(text[i+1]==='"'){cur+='"';i++;} else inQ=false; }
      else cur+=c;
    } else {
      if(c==='"') inQ=true;
      else if(c===','){ row.push(cur); cur=''; }
      else if(c==='\n'){ row.push(cur); rows.push(row); row=[]; cur=''; }
      else if(c!=='\r') cur+=c;
    }
  }
  if(cur!=='' || row.length){ row.push(cur); rows.push(row); }
  return rows.filter(r=>r.some(c=>c.trim()!==''));
}

/* ---------- HTML table export (e.g. Fibi "שמירה" as html) → grid ---------- */
function readHtmlGrid(text){
  if(typeof DOMParser==='undefined') return [];
  const doc=new DOMParser().parseFromString(text,'text/html');
  const out=[];
  doc.querySelectorAll('tr').forEach(tr=>{
    const cells=[...tr.children].filter(c=>/^t[dh]$/i.test(c.tagName)).map(c=>c.textContent.replace(/\s+/g,' ').trim());
    if(cells.some(c=>c)) out.push(cells.map(c=>{ const n=c.replace(/,/g,''); return c!=='' && /^-?\d+(\.\d+)?$/.test(n)? Number(n) : c; }));
  });
  return out;
}

/* ---------- minimal XLSX (zip + sheet xml) reader ---------- */
async function inflateRawBytes(bytes){
  if(globalThis.__inflateRaw) return globalThis.__inflateRaw(bytes); // test harness hook
  const ds = new DecompressionStream('deflate-raw');
  const stream = new Blob([bytes]).stream().pipeThrough(ds);
  const buf = await new Response(stream).arrayBuffer();
  return new Uint8Array(buf);
}
function u16(b,o){ return b[o] | (b[o+1]<<8); }
function u32(b,o){ return (b[o] | (b[o+1]<<8) | (b[o+2]<<16) | (b[o+3]<<24)) >>> 0; }
async function unzip(buf){
  const b = new Uint8Array(buf);
  // find EOCD
  let e = -1;
  for(let i=b.length-22; i>=Math.max(0,b.length-22-65536); i--){
    if(u32(b,i)===0x06054b50){ e=i; break; }
  }
  if(e<0) throw new Error('not a zip');
  const count = u16(b,e+10); let off = u32(b,e+16);
  const entries = {};
  const dec = new TextDecoder();
  for(let i=0;i<count;i++){
    if(u32(b,off)!==0x02014b50) break;
    const method=u16(b,off+10), csize=u32(b,off+20), nlen=u16(b,off+28), elen=u16(b,off+30), clen=u16(b,off+32), lho=u32(b,off+42);
    const name = dec.decode(b.subarray(off+46, off+46+nlen));
    entries[name] = {method,csize,lho};
    off += 46+nlen+elen+clen;
  }
  async function read(name){
    const ent = entries[name]; if(!ent) return null;
    const lo = ent.lho;
    if(u32(b,lo)!==0x04034b50) throw new Error('bad local header');
    const nlen=u16(b,lo+26), elen=u16(b,lo+28);
    const start = lo+30+nlen+elen;
    const data = b.subarray(start, start+ent.csize);
    if(ent.method===0) return data.slice();
    if(ent.method===8) return inflateRawBytes(data);
    throw new Error('unsupported compression '+ent.method);
  }
  return {names:Object.keys(entries), read};
}
function xmlUnescape(s){ return s.replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&#x([0-9a-fA-F]+);/g,(_,h)=>String.fromCodePoint(parseInt(h,16))).replace(/&#(\d+);/g,(_,d)=>String.fromCodePoint(Number(d))).replace(/&amp;/g,'&'); }
function colToIdx(ref){ // 'BC12' → column index 0-based
  let c=0,i=0;
  for(;i<ref.length;i++){ const ch=ref.charCodeAt(i); if(ch>=65&&ch<=90) c=c*26+(ch-64); else break; }
  return c-1;
}
/* returns array of rows; each row is array of cell values (string|number|null) */
async function readXlsxGrid(arrayBuf){
  const z = await unzip(arrayBuf);
  const dec = new TextDecoder();
  // shared strings (tags may carry a namespace prefix, e.g. <x:t>)
  const T_RE=/<(?:\w+:)?t(?:\s[^>]*)?>([\s\S]*?)<\/(?:\w+:)?t>/g;
  const V_RE=/<(?:\w+:)?v(?:\s[^>]*)?>([\s\S]*?)<\/(?:\w+:)?v>/;
  let shared=[];
  const ssRaw = await z.read('xl/sharedStrings.xml');
  if(ssRaw){
    const xml = dec.decode(ssRaw);
    const sis = xml.split(/<(?:\w+:)?si[ >]/).slice(1);
    shared = sis.map(si=>{
      let out='';
      const re=new RegExp(T_RE.source,'g'); let m;
      while((m=re.exec(si))) out+=xmlUnescape(m[1]);
      return out;
    });
  }
  // pick first sheet file
  const sheetName = z.names.filter(n=>/^xl\/worksheets\/sheet\d+\.xml$/.test(n)).sort((a,b)=>parseInt(a.match(/\d+/)[0])-parseInt(b.match(/\d+/)[0]))[0];
  if(!sheetName) throw new Error('no worksheet');
  const xml = dec.decode(await z.read(sheetName));
  const rows=[];
  const rowRe=/<(?:\w+:)?row[^>]*\br="(\d+)"[^>]*>([\s\S]*?)<\/(?:\w+:)?row>/g;
  let rm;
  while((rm=rowRe.exec(xml))){
    const rIdx = Number(rm[1])-1;
    const cells=[];
    const cellRe=/<(?:\w+:)?c\s+([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?c>)/g;
    let cm;
    while((cm=cellRe.exec(rm[2]))){
      const attrs=cm[1], inner=cm[2]||'';
      const refM = attrs.match(/\br="([A-Z]+)\d+"/);
      const col = refM? colToIdx(refM[1]) : cells.length;
      const tM = attrs.match(/\bt="(\w+)"/); const t = tM? tM[1] : 'n';
      let val=null;
      if(t==='s'){ const vM=inner.match(V_RE); val = vM? shared[Number(vM[1])] : null; }
      else if(t==='inlineStr'){ let out=''; const re=new RegExp(T_RE.source,'g'); let m2; while((m2=re.exec(inner))) out+=xmlUnescape(m2[1]); val=out; }
      else if(t==='str'){ const vM=inner.match(V_RE); val = vM? xmlUnescape(vM[1]) : null; }
      else { const vM=inner.match(V_RE); val = vM? Number(vM[1]) : null; }
      if(val==='') val=null;
      cells[col]=val;
    }
    rows[rIdx]=cells;
  }
  for(let i=0;i<rows.length;i++) if(!rows[i]) rows[i]=[];
  return rows;
}

/* ---------- file type detection + specific parsers ---------- */
/* ---------- minimal legacy .xls (BIFF8 in OLE2/CFB) reader → grid of rows ---------- */
function readXlsGrid(arrayBuf){
  const b=new Uint8Array(arrayBuf);
  const dv=new DataView(b.buffer, b.byteOffset, b.byteLength);
  if(!(b[0]===0xD0&&b[1]===0xCF&&b[2]===0x11&&b[3]===0xE0)) throw new Error('not an OLE2 file');
  const secSize=1<<dv.getUint16(30,true), miniSize=1<<dv.getUint16(32,true);
  const numFat=dv.getUint32(44,true), dirStart=dv.getUint32(48,true);
  const miniCut=dv.getUint32(56,true), miniFatStart=dv.getUint32(60,true), numMiniFat=dv.getUint32(64,true);
  const difatStart=dv.getUint32(68,true), numDifat=dv.getUint32(72,true);
  const secOff=s=>(s+1)*secSize;
  // FAT sector list
  const fatSecs=[];
  for(let i=0;i<109&&fatSecs.length<numFat;i++){ const s=dv.getUint32(76+i*4,true); if(s<0xFFFFFFFE) fatSecs.push(s); }
  let ds=difatStart;
  for(let n=0;n<numDifat&&ds<0xFFFFFFFE;n++){
    const o=secOff(ds);
    for(let i=0;i<secSize/4-1;i++){ const s=dv.getUint32(o+i*4,true); if(s<0xFFFFFFFE) fatSecs.push(s); }
    ds=dv.getUint32(o+secSize-4,true);
  }
  const fat=[];
  for(const fs of fatSecs){ const o=secOff(fs); for(let i=0;i<secSize/4;i++) fat.push(dv.getUint32(o+i*4,true)); }
  const chain=(start,table)=>{ const out=[]; let s=start, guard=0; while(s<0xFFFFFFFE && guard++<1e6){ out.push(s); s=table[s]; } return out; };
  const readChain=(start,size)=>{
    const secs=chain(start,fat); const out=new Uint8Array(secs.length*secSize);
    secs.forEach((s,i)=>out.set(b.subarray(secOff(s),secOff(s)+secSize), i*secSize));
    return size!=null? out.subarray(0,size) : out;
  };
  // directory
  const dir=readChain(dirStart);
  const entries=[];
  for(let o=0;o+128<=dir.length;o+=128){
    const nameLen=dir[o+64]|(dir[o+65]<<8); if(!nameLen) continue;
    let name=''; for(let i=0;i<nameLen-2;i+=2) name+=String.fromCharCode(dir[o+i]|(dir[o+i+1]<<8));
    const type=dir[o+66];
    const start=(dir[o+116]|(dir[o+117]<<8)|(dir[o+118]<<16)|(dir[o+119]<<24))>>>0;
    const size=(dir[o+120]|(dir[o+121]<<8)|(dir[o+122]<<16)|(dir[o+123]<<24))>>>0;
    entries.push({name,type,start,size});
  }
  const root=entries.find(e=>e.type===5);
  const wb=entries.find(e=>e.type===2 && /^(Workbook|Book)$/i.test(e.name));
  if(!wb) throw new Error('no Workbook stream');
  let stream;
  if(wb.size<miniCut){
    const miniFat=[]; for(const s of chain(miniFatStart,fat)){ const o=secOff(s); for(let i=0;i<secSize/4;i++) miniFat.push(dv.getUint32(o+i*4,true)); }
    const mini=readChain(root.start, root.size);
    const secs=chain(wb.start,miniFat); stream=new Uint8Array(secs.length*miniSize);
    secs.forEach((s,i)=>stream.set(mini.subarray(s*miniSize,(s+1)*miniSize), i*miniSize));
    stream=stream.subarray(0,wb.size);
  } else stream=readChain(wb.start,wb.size);
  // BIFF records
  const sd=new DataView(stream.buffer, stream.byteOffset, stream.byteLength);
  const recs=[]; let p=0;
  while(p+4<=stream.length){ const id=sd.getUint16(p,true), len=sd.getUint16(p+2,true); recs.push({id,off:p+4,len}); p+=4+len; }
  // shared string table (SST 0xFC + CONTINUE 0x3C)
  const sst=[];
  for(let ri=0;ri<recs.length;ri++){
    const r=recs[ri]; if(r.id!==0xFC) continue;
    // build list of chunks: SST data then CONTINUE datas
    const chunks=[stream.subarray(r.off,r.off+r.len)];
    let rj=ri+1; while(rj<recs.length && recs[rj].id===0x3C){ chunks.push(stream.subarray(recs[rj].off,recs[rj].off+recs[rj].len)); rj++; }
    let ci=0, cp=8; // skip cstTotal, cstUnique
    const total=sd.getUint32(r.off+4,true);
    const avail=()=>chunks[ci].length-cp;
    const nextChunk=()=>{ ci++; cp=0; };
    const u8=()=>{ if(avail()<=0) nextChunk(); return chunks[ci][cp++]; };
    const u16=()=>{ const lo=u8(); const hi=u8(); return lo|(hi<<8); };
    const u32=()=>{ const a=u16(); const b2=u16(); return (a|(b2<<16))>>>0; };
    for(let n=0;n<total && ci<chunks.length;n++){
      if(avail()<=0) nextChunk();
      if(ci>=chunks.length) break;
      const cch=u16(); let flags=u8();
      const rich=flags&8, ext=flags&4;
      const cRun=rich? u16():0; const cbExt=ext? u32():0;
      let s=''; let wide=!!(flags&1); let left=cch;
      while(left>0){
        if(avail()<=0){ nextChunk(); if(ci>=chunks.length) break; flags=u8(); wide=!!(flags&1); }
        if(wide){ const can=Math.min(left, Math.floor(avail()/2)); for(let i=0;i<can;i++) s+=String.fromCharCode(u16()); left-=can; if(can===0 && avail()===1){ nextChunk(); if(ci<chunks.length){flags=u8(); wide=!!(flags&1);} } }
        else { const can=Math.min(left, avail()); for(let i=0;i<can;i++) s+=String.fromCharCode(u8()); left-=can; }
      }
      // skip rich runs / ext data (may span chunks)
      let skip=cRun*4+cbExt;
      while(skip>0){ if(avail()<=0){ nextChunk(); if(ci>=chunks.length) break; } const k=Math.min(skip,avail()); cp+=k; skip-=k; }
      sst.push(s);
    }
    break;
  }
  // sheets: take the first worksheet BOF → EOF cells (all sheets concatenated as separate grids; return first non-empty)
  const grids=[]; let cur=null, depth=0;
  const rk=v=>{ const frac=v&2, i100=v&1; let x; if(frac){ x=v>>2; } else { const hi=(v&0xFFFFFFFC)>>>0; const buf=new DataView(new ArrayBuffer(8)); buf.setUint32(0,hi,false); buf.setUint32(4,0,false); x=buf.getFloat64(0,false); } return i100? x/100 : x; };
  const setCell=(row,col,val)=>{ if(!cur) return; (cur[row]=cur[row]||[])[col]=val; };
  for(let ri=0;ri<recs.length;ri++){
    const r=recs[ri], o=r.off;
    if(r.id===0x809){ depth++; if(depth>=1 && r.len>=4 && sd.getUint16(o+2,true)===0x10){ cur=[]; grids.push(cur); } else if(depth>1) cur=null; continue; }
    if(r.id===0x0A){ depth--; if(depth<=1) cur=null; continue; }
    if(!cur) continue;
    if(r.id===0xFD){ setCell(sd.getUint16(o,true), sd.getUint16(o+2,true), sst[sd.getUint32(o+6,true)]); }
    else if(r.id===0x203){ setCell(sd.getUint16(o,true), sd.getUint16(o+2,true), sd.getFloat64(o+6,true)); }
    else if(r.id===0x27E){ setCell(sd.getUint16(o,true), sd.getUint16(o+2,true), rk(sd.getUint32(o+6,true))); }
    else if(r.id===0xBD){ const row=sd.getUint16(o,true), c0=sd.getUint16(o+2,true); const n=(r.len-6)/6; for(let i=0;i<n;i++) setCell(row,c0+i,rk(sd.getUint32(o+4+i*6+2,true))); }
    else if(r.id===0x204){ const row=sd.getUint16(o,true), col=sd.getUint16(o+2,true); const cch=sd.getUint16(o+6,true), fl=stream[o+8]; let s=''; if(fl&1){ for(let i=0;i<cch;i++) s+=String.fromCharCode(sd.getUint16(o+9+i*2,true)); } else { for(let i=0;i<cch;i++) s+=String.fromCharCode(stream[o+9+i]); } setCell(row,col,s); }
    else if(r.id===0x06){ const row=sd.getUint16(o,true), col=sd.getUint16(o+2,true); const b6=stream[o+12], b7=stream[o+13]; if(b6===0xFF && b7===0xFF){ const t=stream[o+6]; if(t===0){ const nx=recs[ri+1]; if(nx && nx.id===0x207){ const cch=sd.getUint16(nx.off,true), fl=stream[nx.off+2]; let s=''; if(fl&1){ for(let i=0;i<cch;i++) s+=String.fromCharCode(sd.getUint16(nx.off+3+i*2,true)); } else { for(let i=0;i<cch;i++) s+=String.fromCharCode(stream[nx.off+3+i]); } setCell(row,col,s); } } else if(t===1) setCell(row,col,!!stream[o+8]); } else setCell(row,col,sd.getFloat64(o+6,true)); }
  }
  const grid=grids.find(g=>g.some(r=>r&&r.length)) || [];
  const out=[]; for(let i=0;i<grid.length;i++){ const row=grid[i]||[]; out.push(Array.from({length:row.length},(_,j)=>row[j]==null?null:row[j])); }
  return out;
}

function gridCell(row,i){ return row && row[i]!=null ? row[i] : null; }
function gridHasText(rows, txt){ return rows.some(r=>r.some(c=>typeof c==='string' && c.includes(txt))); }

/* Isracard/Fibi credit card "פירוט עסקאות" export */
function parseCardGrid(rows){
  // card number
  let card='';
  outer:
  for(const r of rows){ for(const c of r){ const m=String(c||'').match(/(?:מסטרקארד|ויזה|כרטיס)[^\d]*-\s*(\d{4})/); if(m){ card=m[1]; break outer; } } }
  if(!card){ // fallback: any cell like "‫... - 1234"
    outer2:
    for(const r of rows){ for(const c of r){ const m=String(c||'').match(/-\s*(\d{4})\s*$/); if(m && String(c).length<40){ card=m[1]; break outer2; } } }
  }
  // billing month from title line "פירוט עסקאות | אוגוסט 2026"
  let billingKey=null;
  for(const r of rows){ for(const c of r){ const k=keyFromHebLabel(c); if(k){ billingKey=k; break; } } if(billingKey) break; }
  const out=[]; let section=null;
  for(let i=0;i<rows.length;i++){
    const r=rows[i];
    const joined = r.map(c=>String(c==null?'':c)).join(' ');
    if(joined.includes('תאריך רכישה') && joined.includes('שם בית עסק')){
      // identify section from preceding non-empty line
      let sec='למועד';
      for(let j=i-1;j>=Math.max(0,i-4);j--){
        const t=rows[j].map(c=>String(c==null?'':c)).join(' ');
        if(t.includes('שטרם נקלטו')){sec='טרם נקלט';break;}
        if(t.includes('מחוץ למועד')){sec='מחוץ למועד';break;}
        if(t.includes('למועד חיוב')){sec='למועד';break;}
      }
      section={header:r, sec};
      continue;
    }
    if(!section) continue;
    const dt = parseDMY(gridCell(r,0));
    if(!dt){
      const t=joined.trim();
      if(t && !t.startsWith('סה"כ')) section=null; // section ended
      else if(t.startsWith('סה"כ')) section=null;
      continue;
    }
    const merchant = String(gridCell(r,1)||'').trim();
    if(!merchant) continue;
    const txAmount = parseAmount(gridCell(r,2));
    const txCurrency = String(gridCell(r,3)||'₪').trim();
    let amount = parseAmount(gridCell(r,4));
    let currency = String(gridCell(r,5)||'').trim();
    let voucher = gridCell(r,6)!=null ? String(gridCell(r,6)).replace(/\.0$/,'') : '';
    let details = String(gridCell(r,7)||'').replace(/\s*\n\s*/g,' ').trim();
    const pending = section.sec==='טרם נקלט';
    if(pending){ amount = null; currency=''; voucher=''; details=''; }
    if(amount==null && !pending) { amount = txAmount; currency = txCurrency; }
    const isInstallment = /תשלום\s+\d+\s+מתוך/.test(details);
    let mk;
    if(isInstallment && billingKey) mk = prevMonthKey(billingKey);
    else mk = monthKeyFromDate(dt);
    // which bank charge cycle covers this row
    let billedKey = null;
    if(pending) billedKey = null;
    else if(section.sec==='מחוץ למועד'){
      const bd = parseDMY(gridCell(r,8));
      billedKey = bd? monthKeyFromDate(bd) : (billingKey? prevMonthKey(billingKey) : null);
    } else billedKey = billingKey? prevMonthKey(billingKey) : null;
    out.push({
      date: fmtDMY(dt), monthKey: mk, billedKey, merchant, card,
      amount: amount!=null? round2(amount): null, currency: currency||'₪',
      txAmount: txAmount!=null? round2(txAmount): null, txCurrency,
      voucher, details, pending,
      isRefund: amount!=null && amount<0
    });
  }
  return {card, billingKey, rows: out};
}

/* Max / transaction-details export */
function parseMaxGrid(rows){
  let hi=-1, cols={};
  for(let i=0;i<rows.length;i++){
    const idxDate = rows[i].findIndex(c=>String(c||'').trim()==='תאריך עסקה');
    if(idxDate>=0){
      hi=i;
      rows[i].forEach((c,j)=>{ cols[String(c||'').trim()]=j; });
      break;
    }
  }
  if(hi<0) return {rows:[]};
  const out=[];
  for(let i=hi+1;i<rows.length;i++){
    const r=rows[i];
    const dt = parseDMY(gridCell(r,cols['תאריך עסקה']));
    if(!dt) continue;
    const merchant = String(gridCell(r,cols['שם בית העסק'])||'').trim();
    if(!merchant || merchant==='סך הכל') continue;
    const amount = parseAmount(gridCell(r,cols['סכום חיוב']));
    const txAmount = parseAmount(gridCell(r,cols['סכום עסקה מקורי']));
    const last4 = String(gridCell(r,cols['4 ספרות אחרונות של כרטיס האשראי'])||'').replace(/\.0$/,'').trim();
    const notes = String(gridCell(r,cols['הערות'])||'').trim();
    const kind = String(gridCell(r,cols['סוג עסקה'])||'').trim();
    const chargeDt = parseDMY(gridCell(r,cols['תאריך חיוב']));
    const isInstallment = /תשלום\s+\d+\s+מתוך/.test(notes) || kind==='תשלומים';
    let mk;
    if(isInstallment && chargeDt) mk = prevMonthKey(monthKeyFromDate(chargeDt));
    else mk = monthKeyFromDate(dt);
    out.push({
      date: fmtDMY(dt), monthKey: mk, billedKey: chargeDt? prevMonthKey(monthKeyFromDate(chargeDt)) : mk,
      merchant, card: last4||'מקס',
      amount: amount!=null? round2(amount): null, currency:'₪',
      txAmount: txAmount!=null? round2(txAmount): txAmount, txCurrency:'₪',
      voucher:'', details: [kind,notes].filter(Boolean).join(' · '),
      pending:false, isRefund: amount!=null && amount<0
    });
  }
  return {rows:out};
}

/* Fibi bank account report — incomes + card reconciliation */
const CARD_COMPANY_RE = /(ישראכרט|ישרכארט|מקס איט|max איט|כרטיסי אשראי|כאל|לאומי קארד|דיינרס|אמריקן אקספרס)/;
function parseBankGrid(rows){
  let hi=-1, cols={};
  for(let i=0;i<rows.length;i++){
    const names = rows[i].map(c=>String(c||'').trim());
    if(names.includes('זכות') && names.includes('חובה') && (names.includes('תאריך')||names.includes('תאריך ערך'))){
      hi=i; names.forEach((c,j)=>{ if(c) cols[c]=j; }); break;
    }
  }
  if(hi<0) return {incomes:[], cardDebits:[], bankRows:[]};
  const dateCol = cols['תאריך']!=null? cols['תאריך'] : cols['תאריך ערך'];
  const incomes=[], cardDebits=[], bankRows=[];
  for(let i=hi+1;i<rows.length;i++){
    const r=rows[i];
    let dv = gridCell(r,dateCol);
    let dt=null;
    if(typeof dv==='number') dt=excelSerialToDMY(dv);
    else dt=parseDMY(String(dv||'').slice(0,10).split(' ')[0]) || (function(){ const m=String(dv||'').match(/^(\d{4})-(\d{2})-(\d{2})/); return m? {y:+m[1],m:+m[2],d:+m[3]} : null; })();
    if(!dt) continue;
    const credit = parseAmount(gridCell(r,cols['זכות']));
    const debit = parseAmount(gridCell(r,cols['חובה']));
    const desc = String(gridCell(r,cols['תיאור'])||'').trim();
    const mk = monthKeyFromDate(dt);
    const cardName = ()=>{ const cm = desc.match(/(\d{4})/); return cm? cm[1] : (/מקס/.test(desc)?'מקס':desc); };
    if(credit!=null && credit>0){
      const fromCardCo = CARD_COMPANY_RE.test(desc);
      let cat='החזרים';
      if(/משכ/.test(desc)) cat='משכורת';
      else if(/מלגה/.test(desc)) cat='מלגה';
      else if(/ריבית|דיבידנד/.test(desc)) cat='אחר';
      incomes.push({date:fmtDMY(dt), monthKey:mk, source:desc, amount:round2(credit), category:cat, fromCardCo});
      // a refund from the card company is part of that card's bank-side balance for the month
      if(fromCardCo) cardDebits.push({date:fmtDMY(dt), monthKey: dt.d<=7? prevMonthKey(mk) : mk, card:cardName(), bankAmount:round2(-credit)});
    }
    if(debit!=null && debit>0 && CARD_COMPANY_RE.test(desc)){
      // a card charge in the first days of month M covers month M-1
      cardDebits.push({date:fmtDMY(dt), monthKey: dt.d<=7? prevMonthKey(mk) : mk, card:cardName(), bankAmount:round2(debit)});
    } else if(debit!=null && debit>0){
      // any other debit from the account (rent, standing orders, transfers) is an expense paid directly from the bank
      const ref = String(gridCell(r,cols['אסמכתא'])!=null? gridCell(r,cols['אסמכתא']) : '').replace(/\.0$/,'').trim();
      let cat='';
      if(/שכר דירה|שכ"ד|ועד בית|ארנונה|חשמל|מים|גז|ביטוח|עמל|מס ר\.זכות/.test(desc)) cat='שכר דירה, חשבונות ובעלי מקצוע';
      else if(/חסכון|חיסכון|פק"מ|פקדון|פיקדון|קרן השתלמות|קופת גמל/.test(desc)) cat=BANK_SAVINGS_CAT;
      bankRows.push({
        date:fmtDMY(dt), monthKey:mk, billedKey:mk, merchant:desc, card:BANK_CARD,
        amount:round2(debit), currency:'₪', txAmount:round2(debit), txCurrency:'₪',
        voucher:'bank:'+fmtDMY(dt)+':'+(ref||round2(debit)), details:'חיוב ישיר בחשבון', pending:false, isRefund:false, category:cat
      });
    }
  }
  return {incomes, cardDebits, bankRows};
}
const BANK_CARD='בנק';
const BANK_SAVINGS_CAT='חיסכון';
const MANUAL_CARD='ידני';

/* App's own CSV export (הוצאות_*.csv) */
function parseAppCsv(text){
  const rows = parseCSV(text);
  if(!rows.length) return {rows:[]};
  const H = rows[0].map(s=>s.trim());
  const idx = n=>H.indexOf(n);
  const iCard=idx('כרטיס'), iDate=idx('תאריך'), iMerch=idx('הוצאה'), iTx=idx('עסקה'), iTxCur=idx('מטבע'),
        iAmt=idx('חיוב'), iCur=idx('מטבע חיוב'), iVou=idx('מספר שובר'), iCat=idx('פירוט'), iStat=idx('סטטוס'), iNotes=idx('הערות');
  if(iDate<0||iMerch<0||iAmt<0) return {rows:[]};
  const out=[];
  for(let i=1;i<rows.length;i++){
    const r=rows[i];
    const dt=parseDMY(r[iDate]); if(!dt) continue;
    const amount=parseAmount(r[iAmt]); if(amount==null) continue;
    out.push({
      date:fmtDMY(dt), monthKey:monthKeyFromDate(dt),
      merchant:String(r[iMerch]||'').trim(), card:iCard>=0? String(r[iCard]||'').trim():'',
      amount:round2(amount), currency:iCur>=0? (r[iCur]||'₪'):'₪',
      txAmount:iTx>=0? parseAmount(r[iTx]) : null, txCurrency:iTxCur>=0? (r[iTxCur]||'₪'):'₪',
      voucher:iVou>=0? String(r[iVou]||'').trim():'',
      details:iStat>=0? String(r[iStat]||'').trim():'',
      category:iCat>=0? String(r[iCat]||'').trim():'',
      notes:iNotes>=0? String(r[iNotes]||'').trim():'',
      pending:false, isRefund:amount<0
    });
  }
  return {rows:out};
}

/* detect + parse one uploaded File-like {name, arrayBuffer(), text()} */
async function parseUpload(name, getBuf, getText){
  const lower = name.toLowerCase();
  if(lower.endsWith('.json')){
    const data = JSON.parse(await getText());
    return {kind:'backup', data};
  }
  if(lower.endsWith('.csv')){
    const text = await getText();
    if(text.includes('כרטיס') && text.includes('שובר')) return {kind:'appcsv', data:parseAppCsv(text)};
    return {kind:'unknown'};
  }
  if(lower.endsWith('.html') || lower.endsWith('.htm')){
    const grid = readHtmlGrid(await getText());
    if(gridHasText(grid,'תאריך רכישה')) return {kind:'card', data:parseCardGrid(grid)};
    if(gridHasText(grid,'זכות') && gridHasText(grid,'חובה')) return {kind:'bank', data:parseBankGrid(grid)};
    return {kind:'unknown'};
  }
  if(lower.endsWith('.xlsx') || lower.endsWith('.xls')){
    const buf = await getBuf();
    const head = new Uint8Array(buf, 0, 4);
    const grid = (head[0]===0xD0 && head[1]===0xCF) ? readXlsGrid(buf) : await readXlsxGrid(buf);
    if(gridHasText(grid,'תאריך רכישה')) return {kind:'card', data:parseCardGrid(grid)};
    if(gridHasText(grid,'תאריך עסקה')) return {kind:'max', data:parseMaxGrid(grid)};
    if(gridHasText(grid,'זכות') && gridHasText(grid,'חובה')) return {kind:'bank', data:parseBankGrid(grid)};
    return {kind:'unknown'};
  }
  return {kind:'unknown'};
}
/* ==END PARSERS== */

/* ---------- state ---------- */
const LS_KEY='em2_state';
const DEFAULT_CATS=['מסעדות, קפה וברים','סופר וירקות','בגדים','תרבות','תחבצ','חנייה ודלק','שכר דירה, חשבונות ובעלי מקצוע','תחזוקה, רהיטים, אלקטרוניקה','טיפוח','חופשות','ספורט','רפואה','מתנות','לא ברור','תרומות','דברים שאני מקבלת עליהם החזר'];
let state=null, ui={month:null, filter:'all', search:'', dirty:false, saveStatus:'saved', readOnly:false, artifactNS:undefined};
let normMerchants={}; // normalized merchant name → category

const DEFAULT_INC_CATS=['משכורת','מלגה','הורים','הכנסה מהשקעות','החזרים','אחר'];
/* exchange rates used to fold foreign-currency charges (billed separately from a €/$ account) into ₪ totals */
const DEFAULT_FX={'€':3.4854,'$':3.35,'£':4.4};
function fxRate(cur){ const fx=(state&&state.settings&&state.settings.fx)||DEFAULT_FX; const v=Number(fx[cur]); return v>0? v : (DEFAULT_FX[cur]||1); }
/* ₪ value of an expense row (foreign rows converted at the configured rate) */
function ils(r){ if(r.amount==null) return 0; return (!r.currency || r.currency==='₪')? r.amount : round2(r.amount*fxRate(r.currency)); }
function blankState(){
  return {version:2, ts:0, categories:DEFAULT_CATS.slice(), incomeCategories:DEFAULT_INC_CATS.slice(),
          merchants:{}, months:{}, investments:[], settings:{autoPublish:true, fx:Object.assign({},DEFAULT_FX)}};
}
/* make any loaded state safe to render: fill whatever an older / partial copy lacks */
function normalizeState(st){
  if(!st || typeof st!=='object') st=blankState();
  st.version=2;
  if(!Array.isArray(st.categories) || !st.categories.length) st.categories=DEFAULT_CATS.slice();
  if(!Array.isArray(st.incomeCategories)) st.incomeCategories=[];
  for(const c of DEFAULT_INC_CATS) if(!st.incomeCategories.includes(c)) st.incomeCategories.push(c);
  if(!st.months || typeof st.months!=='object') st.months={};
  for(const k of Object.keys(st.months)){
    const m=st.months[k];
    if(!/^\d{4}-\d{2}$/.test(k) || !m || typeof m!=='object'){ delete st.months[k]; continue; }
    if(!Array.isArray(m.rows)) m.rows=[];
    if(!Array.isArray(m.incomes)) m.incomes=[];
    if(!Array.isArray(m.recon)) m.recon=[];
    for(const r of m.rows){ if(!r.id) r.id=uid(); if(r.merchant==null) r.merchant=''; if(r.details==null) r.details=''; }
  }
  if(!st.merchants || typeof st.merchants!=='object') st.merchants={};
  if(!Array.isArray(st.investments)) st.investments=[];
  if(!st.settings || typeof st.settings!=='object') st.settings={autoPublish:true};
  st.settings.fx=Object.assign({},DEFAULT_FX,st.settings.fx||{});
  if(!Array.isArray(st.settings.splitKeywords)) st.settings.splitKeywords=[];
  if(!st.settings.cardAliases || typeof st.settings.cardAliases!=='object') st.settings.cardAliases={};
  if(!st.splitMerchants || typeof st.splitMerchants!=='object') st.splitMerchants={};
  return st;
}
function ensureMonth(k){
  if(!state.months[k]) state.months[k]={rows:[],incomes:[],recon:[]};
  return state.months[k];
}
function rebuildNormMerchants(){
  normMerchants={};
  for(const [n,c] of Object.entries(state.merchants)){ const nn=normName(n); if(nn) normMerchants[nn]=c; }
}
/* generic transfer names mean something different every time — never learn or auto-match a category from them */
const GENERIC_MERCHANT=/^(העברה|תשלום מיידי)/;
/* shared household bills that should also be entered in Splitwise: learned per merchant, on top of the
   keywords listed in settings.splitKeywords (substring match on the merchant name) */
function splitKeywords(){ return ((state.settings&&state.settings.splitKeywords)||[]).filter(Boolean); }
function isSplitMerchant(merchant){
  const nn=normName(merchant), sm=state.splitMerchants||{};
  if(nn && Object.prototype.hasOwnProperty.call(sm,nn)) return !!sm[nn];
  const mm=String(merchant||''); return splitKeywords().some(k=>mm.includes(k));
}
function setSplit(r,val){
  r.split=val||'';
  if(!state.splitMerchants) state.splitMerchants={};
  const nn=normName(r.merchant);
  if(nn && !GENERIC_MERCHANT.test(r.merchant)) state.splitMerchants[nn]=!!val;
}
/* expenses someone pays back (university, friends): tracked until the refund is marked as received */
const REFUND_CAT='דברים שאני מקבלת עליהם החזר';
function refundCat(){ return (state.settings&&state.settings.refundCat)||REFUND_CAT; }
function awaitsRefund(r){ return r.category===refundCat() && r.amount>0 && !r.pending && r.refund!=='got'; }
function outstandingRefunds(){
  const out=[];
  for(const [k,m] of Object.entries(state.months)) for(const r of m.rows) if(awaitsRefund(r)) out.push({k,r});
  return out.sort((a,b)=>dateNum(a.r.date)-dateNum(b.r.date));
}
function lookupCategory(merchant){
  if(GENERIC_MERCHANT.test(merchant||'')) return '';
  if(state.merchants[merchant]) return state.merchants[merchant];
  const nn=normName(merchant);
  if(nn && normMerchants[nn]) return normMerchants[nn];
  return '';
}
function rememberMerchant(merchant, cat){
  if(!merchant||!cat) return;
  if(GENERIC_MERCHANT.test(merchant)) return;
  state.merchants[merchant]=cat;
  const nn=normName(merchant); if(nn) normMerchants[nn]=cat;
}

/* the bank names a card by company ("מקס") while the card statement names it by its last digits:
   settings.cardAliases = {company name: [last digits,…]} ties the two together for reconciliation */
function cardId(c){ c=String(c==null?'':c).trim(); return /^\d+$/.test(c)? String(Number(c)) : c; }
function sameCard(a,b){
  a=cardId(a); b=cardId(b);
  if(a===b) return true;
  const al=(state.settings&&state.settings.cardAliases)||{};
  for(const [name,ids] of Object.entries(al)){
    const set=[cardId(name)].concat((ids||[]).map(cardId));
    if(set.includes(a) && set.includes(b)) return true;
  }
  return false;
}

/* ---------- merge imported rows into months (voucher dedupe) ---------- */
function allVouchers(){
  const s=new Set();
  for(const m of Object.values(state.months)) for(const r of m.rows) if(r.voucher) s.add(r.voucher);
  return s;
}
function pendingKey(r){ return [r.date,normName(r.merchant),r.txAmount!=null?r.txAmount:r.amount].join('|'); }
function mergeExpenseRows(rows, opts){
  opts=opts||{};
  const vs=allVouchers();
  const pendKeys=new Map(); // pendingKey → {monthKey,row}
  for(const [mk,m] of Object.entries(state.months)) for(const r of m.rows) if(r.pending) pendKeys.set(pendingKey(r),{mk,row:r});
  let added=0, dup=0, upgraded=0, needCat=0, refreshed=0;
  const byVoucher={};
  for(const m of Object.values(state.months)) for(const r of m.rows) if(r.voucher) byVoucher[r.voucher]=r;
  // rows without a voucher (Max, bank transfers): dedupe by date+merchant+amount+card; installments repeat the
  // original purchase date every month, so their billing cycle is part of the key
  const nvKey = r => [r.date,normName(r.merchant),r.amount,r.card||'', /תשלום\s+\d+\s+מתוך/.test(r.details||'')? (r.billedKey||'') : ''].join('|');
  const noVoucherKeys=new Set();
  for(const m of Object.values(state.months)) for(const r of m.rows) if(!r.voucher && !r.pending) noVoucherKeys.add(nvKey(r));
  for(const nr of rows){
    if(nr.voucher && vs.has(nr.voucher)){
      // same transaction seen again: refresh billing facts the statement may have corrected (currency, amount, cycle) — never the user's category/notes
      const ex=byVoucher[nr.voucher];
      if(ex && !nr.pending && nr.amount!=null){
        let ch=false;
        if(nr.currency && ex.currency!==nr.currency){ ex.currency=nr.currency; ch=true; }
        if(ex.amount!==nr.amount){ ex.amount=nr.amount; ex.isRefund=nr.amount<0; ch=true; }
        if(nr.billedKey && ex.billedKey!==nr.billedKey){ ex.billedKey=nr.billedKey; ch=true; }
        if(nr.details && !ex.details){ ex.details=nr.details; ch=true; }
        if(ch) refreshed++;
      }
      dup++; continue;
    }
    if(!nr.voucher && !nr.pending){
      const k=nvKey(nr);
      if(noVoucherKeys.has(k)){ dup++; continue; }
      noVoucherKeys.add(k);
    }
    if(nr.category && !state.categories.includes(nr.category)) state.categories.push(nr.category);
    // upgrade a pending row into a real one
    if(nr.voucher){
      const pk=pendingKey(nr);
      const p=pendKeys.get(pk);
      if(p){
        Object.assign(p.row,{date:nr.date,amount:nr.amount,currency:nr.currency,voucher:nr.voucher,details:nr.details,pending:false,isRefund:nr.isRefund,billedKey:nr.billedKey||null});
        if(p.mk!==nr.monthKey){ // move to correct month
          const arr=state.months[p.mk].rows; arr.splice(arr.indexOf(p.row),1); ensureMonth(nr.monthKey).rows.push(p.row);
        }
        pendKeys.delete(pk); vs.add(nr.voucher); upgraded++; continue;
      }
    } else if(nr.pending){
      if(pendKeys.has(pendingKey(nr))){ dup++; continue; }
    }
    const row={
      id:uid(), date:nr.date, merchant:nr.merchant, amount:nr.amount!=null?nr.amount:(nr.txAmount||0),
      card:nr.card||'', currency:nr.currency||'₪', txCurrency:nr.txCurrency||'₪',
      voucher:nr.voucher||'', txAmount:nr.txAmount!=null?nr.txAmount:nr.amount,
      details:nr.details||'', category:nr.category||'', autoMatched:false,
      notes:nr.notes||'', isRefund:!!nr.isRefund, pending:!!nr.pending,
      billedKey:nr.billedKey||null
    };
    if(!row.category){
      const c=lookupCategory(row.merchant);
      if(c){ row.category=c; row.autoMatched=true; }
    } else if(!opts.noLearn){ rememberMerchant(row.merchant,row.category); }
    if(!row.category) needCat++;
    if(!row.pending && row.amount>0 && isSplitMerchant(row.merchant)) row.split='todo';
    ensureMonth(nr.monthKey).rows.push(row);
    if(nr.pending) pendKeys.set(pendingKey(row),{mk:nr.monthKey,row});
    if(row.voucher) vs.add(row.voucher);
    added++;
  }
  return {added,dup,upgraded,needCat,refreshed};
}
function mergeIncomes(incomes){
  let added=0,dup=0,skippedCardCo=0;
  for(const inc of incomes){
    if(inc.fromCardCo){ skippedCardCo++; continue; } // refunds from card company are already negative rows in the card data
    const m=ensureMonth(inc.monthKey);
    if(m.incomes.some(x=>x.date===inc.date && x.amount===inc.amount && normName(x.source)===normName(inc.source))){ dup++; continue; }
    // a recurring credit (same source, same amount) keeps the category it was given in the most recent earlier month
    let cat=inc.category||'החזרים';
    const prevKeys=Object.keys(state.months).filter(k=>k<inc.monthKey).sort().reverse();
    findPrev: for(const k of prevKeys) for(const x of state.months[k].incomes||[]) if(x.amount===inc.amount && x.category && normName(x.source)===normName(inc.source)){ cat=x.category; break findPrev; }
    m.incomes.push({id:uid(), date:inc.date, source:inc.source, amount:inc.amount, category:cat, notes:''});
    added++;
  }
  return {added,dup,skippedCardCo};
}
function mergeRecon(cardDebits){
  let added=0;
  for(const cd of cardDebits){
    const m=ensureMonth(cd.monthKey);
    if(!m.recon) m.recon=[];
    if(m.recon.some(x=>x.card===cd.card && x.bankAmount===cd.bankAmount && x.date===cd.date)) continue;
    m.recon.push({id:uid(), card:cd.card, bankAmount:cd.bankAmount, date:cd.date});
    added++;
  }
  return {added};
}

/* ---------- backup import (old v1 + new v2) ---------- */
function importBackup(data){
  const notes=[];
  if(data && data.version===2 && data.months){
    state=normalizeState(data); state.ts=Date.now();
    rebuildNormMerchants();
    return ['גיבוי v2 נטען במלואו'];
  }
  if(!(data && Array.isArray(data.categories) && data.merchants)) throw new Error('פורמט גיבוי לא מוכר');
  // v1 backup {categories, merchants, history, investments}
  for(const c of data.categories) if(!state.categories.includes(c)) state.categories.push(c);
  let mCount=0;
  for(const [n,c] of Object.entries(data.merchants)){ if(!state.merchants[n]){ state.merchants[n]=c; mCount++; } }
  rebuildNormMerchants();
  let months=0, rowsN=0, incN=0;
  for(const h of (data.history||[])){
    let mk = keyFromHebLabel(h.label);
    if(!mk){ // derive from row dates
      const counts={};
      for(const r of (h.rows||[])){ const d=parseDMY(r.date); if(d){ const k=monthKeyFromDate(d); counts[k]=(counts[k]||0)+1; } }
      mk = Object.entries(counts).sort((a,b)=>b[1]-a[1]).map(e=>e[0])[0];
    }
    if(!mk) continue;
    const m=ensureMonth(mk); months++;
    const vs=new Set(m.rows.map(r=>r.voucher).filter(Boolean));
    for(const r of (h.rows||[])){
      if(r.voucher && vs.has(String(r.voucher))) continue;
      m.rows.push({id:r.id||uid(), date:r.date, merchant:r.merchant, amount:r.amount, card:r.card||'', currency:r.currency||'₪',
        txCurrency:r.txCurrency||'₪', voucher:r.voucher?String(r.voucher):'', txAmount:r.txAmount!=null?r.txAmount:r.amount,
        details:r.details||'', category:r.category||'', autoMatched:!!r.autoMatched, notes:r.notes||'', isRefund:!!r.isRefund, pending:false});
      rowsN++;
    }
    for(const inc of (h.incomes||[])){
      // place income by its own date (fixes months polluted by a multi-month bank import)
      const d=parseDMY(inc.date); const ik=d? monthKeyFromDate(d) : mk;
      const im=ensureMonth(ik);
      if(im.incomes.some(x=>x.date===inc.date && x.amount===inc.amount && normName(x.source)===normName(inc.source))) continue;
      im.incomes.push({id:inc.id||uid(), date:inc.date, source:inc.source, amount:inc.amount, category:inc.category||'החזרים', notes:inc.notes||''});
      incN++;
    }
    for(const rc of (h.reconciliation||[])){
      if(!m.recon) m.recon=[];
      if(!m.recon.find(x=>x.card===rc.card)) m.recon.push({id:rc.id||uid(), card:rc.card, bankAmount:rc.bankAmount});
    }
  }
  if(Array.isArray(data.investments) && data.investments.length && !state.investments.length) state.investments=data.investments;
  notes.push('נטען גיבוי: '+months+' חודשים, '+rowsN+' הוצאות, '+incN+' הכנסות, '+mCount+' בתי עסק בזיכרון');
  return notes;
}

/* ---------- migration from the old app's localStorage ---------- */
function migrateOldLocalStorage(){
  const notes=[];
  let keys=[];
  try{ keys=Object.keys(localStorage); }catch(e){ return notes; }
  for(const k of keys){
    if(k===LS_KEY) continue;
    let v=null;
    try{ v=JSON.parse(localStorage.getItem(k)); }catch(e){ continue; }
    if(!v || typeof v!=='object') continue;
    try{
      if(v.categories && v.merchants && (v.history||v.months)){
        notes.push(...importBackup(v)); notes.push('(שוחזר אוטומטית מהגרסה הקודמת: '+k+')');
        continue;
      }
      // loose shapes: current-month rows / incomes from the old app
      const rows = Array.isArray(v)? v : (Array.isArray(v.rows)? v.rows : null);
      if(rows && rows.length && rows[0] && rows[0].merchant!==undefined && rows[0].amount!==undefined && rows[0].date){
        const conv=rows.map(r=>{ const d=parseDMY(r.date); return d? Object.assign({},r,{monthKey:monthKeyFromDate(d), voucher:r.voucher?String(r.voucher):''}) : null; }).filter(Boolean);
        const res=mergeExpenseRows(conv,{noLearn:true});
        if(res.added) notes.push('שוחזרו '+res.added+' הוצאות מהגרסה הקודמת ('+k+')');
      }
      const incs = Array.isArray(v)? null : (Array.isArray(v.incomes)? v.incomes : null);
      const incArr = incs || (Array.isArray(v) && v.length && v[0] && v[0].source!==undefined && v[0].amount!==undefined ? v : null);
      if(incArr && incArr.length){
        const conv=incArr.map(x=>{ const d=parseDMY(x.date); return d? {date:x.date, monthKey:monthKeyFromDate(d), source:x.source, amount:x.amount, category:x.category} : null; }).filter(Boolean);
        const res=mergeIncomes(conv);
        if(res.added) notes.push('שוחזרו '+res.added+' הכנסות מהגרסה הקודמת, כל אחת שויכה לחודש הנכון לפי תאריך ('+k+')');
      }
    }catch(e){ /* skip malformed */ }
  }
  return notes;
}

/* ---------- persistence: localStorage instantly, artifact publish debounced ---------- */
function saveLocal(){
  state.ts=Date.now();
  try{ localStorage.setItem(LS_KEY, JSON.stringify(state)); }catch(e){}
}
function markDirty(){
  ui.dirty=true; ui.saveStatus='dirty';
  saveLocal(); renderSaveChip(); schedulePublish();
}
let pubTimer=null;
function schedulePublish(){
  if(!state.settings.autoPublish) return;
  if(pubTimer) clearTimeout(pubTimer);
  pubTimer=setTimeout(publishNow, 15000);
}
function buildDoc(st){
  const css=document.getElementById('APPCSS').textContent;
  const code=document.getElementById('APPCODE').textContent;
  const json=JSON.stringify(st||state).replace(/</g,'\\u003c');
  return '<!doctype html>\n<html lang="he" dir="rtl">\n<head>\n<meta charset="utf-8">\n<meta name="viewport" content="width=device-width, initial-scale=1">\n<title>ניהול הוצאות חכם</title>\n<link rel="preconnect" href="https://fonts.googleapis.com">\n<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Assistant:wght@400;600;700&family=Secular+One&display=swap">\n<style id="APPCSS">'+css+'</style>\n</head>\n<body>\n<script type="application/json" id="APPSTATE">'+json+'</'+'script>\n<div id="app" dir="rtl" lang="he"></div>\n<script id="APPCODE">'+code+'</'+'script>\n</body>\n</html>';
}
async function publishNow(){
  if(pubTimer){ clearTimeout(pubTimer); pubTimer=null; }
  if(ui.artifactNS===undefined) ui.artifactNS = await (window.claude && window.claude.use ? window.claude.use('artifact') : Promise.resolve(null));
  if(!ui.artifactNS || ui.readOnly){ ui.saveStatus = ui.dirty ? 'localonly' : 'saved'; renderSaveChip(); return; }
  ui.saveStatus='saving'; renderSaveChip();
  try{ sessionStorage.setItem('em2_ui', JSON.stringify({month:ui.month, filter:ui.filter, y:window.scrollY})); }catch(e){}
  try{
    await ui.artifactNS.publish(buildDoc());
    ui.dirty=false; ui.saveStatus='saved'; renderSaveChip();
    // page reloads after publish
  }catch(err){
    const code=err && err.code;
    if(code==='conflict'){ /* view reloads to winner */ }
    else if(code==='not_writer' || code==='not_granted' || code==='not_declared' || code==='consent_required'){ ui.readOnly=true; ui.saveStatus='localonly'; renderSaveChip(); }
    else if(code==='rate_limited'){ ui.saveStatus='dirty'; renderSaveChip(); if(state.settings.autoPublish){ pubTimer=setTimeout(publishNow,60000); } }
    else { ui.saveStatus='dirty'; renderSaveChip(); toast('השמירה לענן נכשלה ('+(code||'שגיאה')+'). הנתונים שמורים מקומית בדפדפן.'); }
  }
}

/* ---------- calculations ---------- */
function monthTotals(mk){
  const m=state.months[mk]||{rows:[],incomes:[],recon:[]};
  let exp=0, inc=0, uncat=0, pend=0;
  const byCat={};
  for(const r of m.rows){
    const a = ils(r);
    if(!r.pending) exp+=a; else pend++;
    const c=r.category||'—';
    byCat[c]=(byCat[c]||0)+(r.pending?0:a);
    if(!r.category && !r.pending) uncat++;
  }
  for(const i2 of m.incomes) inc+=i2.amount||0;
  return {exp:round2(exp), inc:round2(inc), net:round2(inc-exp), uncat, pend, byCat};
}
function sortedMonthKeys(){
  return Object.keys(state.months).sort().reverse();
}
function fmtN(x){
  x=Math.round(x||0);
  return (x<0?'-':'')+Math.abs(x).toLocaleString('he-IL',{minimumFractionDigits:0, maximumFractionDigits:0});
}
function fmtN2(x){
  x=round2(x||0); if(Object.is(x,-0)) x=0;
  return x.toLocaleString('he-IL',{minimumFractionDigits:2, maximumFractionDigits:2});
}

/* ---------- UI ---------- */
const app=document.getElementById('app');
function el(tag, attrs, ...kids){
  const e=document.createElement(tag);
  if(attrs) for(const [k,v] of Object.entries(attrs)){
    if(k==='class') e.className=v;
    else if(k.startsWith('on')) e.addEventListener(k.slice(2), v);
    else if(v!==false && v!=null) e.setAttribute(k, v===true?'':v);
  }
  for(const k of kids.flat()){ if(k==null) continue; e.append(k.nodeType? k : document.createTextNode(k)); }
  return e;
}
let toastTimer=null;
function toast(msg, undo){
  let t=document.querySelector('.toast');
  if(t) t.remove();
  t=el('div',{class:'toast', role:'status'}, el('span',null,msg));
  if(undo){
    const b=el('button',{class:'toastbtn'},'ביטול');
    b.addEventListener('click',()=>{ t.remove(); undo(); });
    t.append(b);
  }
  const x=el('button',{class:'toastbtn x', title:'סגירה', 'aria-label':'סגירה'},'✕');
  x.addEventListener('click',()=>t.remove());
  t.append(x);
  document.body.append(t);
  if(toastTimer) clearTimeout(toastTimer);
  toastTimer=setTimeout(()=>t.remove(), undo? 12000 : 7000);
}
/* run a change that is hard to redo by hand, then offer to roll the whole state back */
function snapshot(){ return JSON.stringify(state); }
function restore(snap){ state=normalizeState(JSON.parse(snap)); rebuildNormMerchants(); markDirty(); render(); toast('הפעולה בוטלה'); }
function undoable(msg, fn){ const snap=snapshot(); fn(); markDirty(); render(); toast(msg, ()=>restore(snap)); }
function dateNum(s){ const d=parseDMY(s); return d? d.y*10000+d.m*100+d.d : 0; }
function isCloud(){ return !!(window.claude && window.claude.use); }
function renderSaveChip(){
  const c=document.getElementById('savechip'); if(!c) return;
  c.className='savechip';
  if(!isCloud()){ c.classList.add(ui.dirty?'dirty':'saved'); c.textContent = ui.dirty? 'נשמר בדפדפן · מומלץ להוריד עותק' : 'נשמר בדפדפן'; c.title='הנתונים נשמרים בדפדפן הזה. להעברה/גיבוי: הגדרות ← הורדת האפליקציה עם כל הנתונים'; return; }
  if(ui.saveStatus==='saving'){ c.classList.add('saving'); c.textContent='שומר…'; }
  else if(ui.saveStatus==='dirty'){ c.classList.add('dirty'); c.textContent='שינויים בהמתנה'; }
  else if(ui.saveStatus==='localonly'){ c.classList.add('dirty'); c.textContent='נשמר מקומית בלבד'; }
  else { c.classList.add('saved'); c.textContent='נשמר ✓'; }
}

function render(){
  app.textContent='';
  const mks=sortedMonthKeys();
  if(ui.month==null || !mks.includes(ui.month)) ui.month = mks.find(k=>(state.months[k].rows||[]).length) || mks[0] || null;

  // topbar
  app.append(el('div',{class:'topbar'},
    el('h1',null,'ניהול הוצאות חכם'),
    el('div',{class:'grow'}),
    el('span',{id:'savechip',class:'savechip'},''),
    isCloud()? el('button',{class:'ghost small', onclick:()=>publishNow()},'שמירה עכשיו') : el('button',{class:'ghost small', onclick:()=>{ const b=new Blob([buildDoc()],{type:'text/html;charset=utf-8'}); const a=el('a',{href:URL.createObjectURL(b), download:'ניהול_הוצאות_'+new Date().toISOString().slice(0,10)+'.html'}); document.body.append(a); a.click(); a.remove(); ui.dirty=false; renderSaveChip(); }},'הורדת עותק עם הנתונים')
  ));
  renderSaveChip();

  // dropzone
  const dz=el('div',{class:'dropzone', id:'dz'},
    el('div',null,'גרירת קבצים לכאן או ',
      (()=>{ const b=el('button',{class:'ghost'},'בחירת קבצים'); b.addEventListener('click',()=>fi.click()); return b; })()),
    el('div',{class:'hint'},'דוחות כרטיס אשראי (xlsx) · דוח עו"ש מהבנק (xlsx / xls / html) · קובץ CSV שיוצא מהאפליקציה · גיבוי JSON — אפשר כמה קבצים ביחד, כפילויות מזוהות לבד')
  );
  const fi=el('input',{type:'file', multiple:true, accept:'.xlsx,.xls,.csv,.json,.html,.htm', style:'display:none'});
  fi.addEventListener('change',()=>{ handleFiles([...fi.files]); fi.value=''; });
  dz.append(fi);
  dz.addEventListener('dragover',e=>{e.preventDefault(); dz.classList.add('over');});
  dz.addEventListener('dragleave',()=>dz.classList.remove('over'));
  dz.addEventListener('drop',e=>{e.preventDefault(); dz.classList.remove('over'); handleFiles([...e.dataTransfer.files]);});
  app.append(dz);

  // month tabs
  const tabs=el('div',{class:'monthtabs'});
  for(const mk of mks){
    const t=monthTotals(mk);
    const b=el('button',{class:mk===ui.month?'active':''}, monthLabel(mk), t.uncat? el('span',{class:'flag'},String(t.uncat)) : null);
    b.addEventListener('click',()=>{ ui.month=mk; render(); });
    tabs.append(b);
  }
  app.append(tabs);

  if(!ui.month){
    app.append(el('div',{class:'card muted'},'אין עדיין נתונים. גררו למעלה דוח כרטיס אשראי (xlsx) ודוח תנועות עו"ש מהבנק — ההוצאות, ההכנסות וההתאמה מול הבנק ייבנו לבד. הכול נשמר רק בדפדפן שלך; שום דבר לא נשלח לשום שרת.'));
    renderSettings();
    return;
  }
  const mk=ui.month, m=ensureMonth(mk), t=monthTotals(mk);

  // KPIs
  app.append(el('div',{class:'kpis'},
    el('div',{class:'kpi expense'}, el('div',{class:'lbl'},'הוצאות · '+monthLabel(mk)), el('div',{class:'val num'},fmtN(t.exp)+' ₪')),
    el('div',{class:'kpi income'}, el('div',{class:'lbl'},'הכנסות'), el('div',{class:'val num'},fmtN(t.inc)+' ₪')),
    el('div',{class:'kpi '+(t.net>=0?'income':'expense')}, el('div',{class:'lbl'},'מאזן'), el('div',{class:'val num'},fmtN(t.net)+' ₪')),
    el('div',{class:'kpi'+(t.uncat?' warn':'')}, el('div',{class:'lbl'},'ממתינות לסיווג'), el('div',{class:'val num'},String(t.uncat)))
  ));

  // refunds still owed to me — across all months, so the monthly check catches anything that never came back
  const owed=outstandingRefunds();
  if(owed.length){
    const tot=round2(owed.reduce((s,x)=>s+ils(x.r),0));
    const list=el('div',{class:'owedlist'});
    for(const x of owed){
      const b=el('button',null,'הוחזר ✓');
      b.addEventListener('click',()=>{ x.r.refund='got'; markDirty(); render(); });
      const go=el('button',{class:'ghost'},monthLabel(x.k));
      go.addEventListener('click',()=>{ ui.month=x.k; ui.filter='cat:'+refundCat(); render(); });
      list.append(el('div',{class:'owedrow'}, el('span',{class:'num'},x.r.date), el('span',{class:'owedname'},x.r.merchant.trim()+(x.r.notes?' · '+x.r.notes:'')), el('span',{class:'num'},fmtN2(ils(x.r))+' ₪'), go, b));
    }
    app.append(el('div',{class:'banner warn owed'}, el('div',null, el('b',null,'ממתין להחזר: '), owed.length+' הוצאות · '+fmtN2(tot)+' ₪'), list));
  }

  // monthly dashboard
  app.append(renderDashboard(mk));

  // expenses table
  const filters=el('div',{class:'filterchips'});
  const fdefs=[['all','הכל'],['uncat','לסיווג ('+t.uncat+')'],['abroad','חו"ל'],['refund','זיכויים'],['pending','טרם נקלטו ('+t.pend+')']];
  const splitTodo=m.rows.filter(r=>r.split==='todo'), splitAll=m.rows.filter(r=>r.split);
  if(splitAll.length || ui.filter==='split') fdefs.push(['split','ספליטווייז ('+splitTodo.length+')']);
  if(ui.filter.startsWith('cat:')) fdefs.push([ui.filter, (ui.filter.slice(4)==='—'?'ללא סיווג':ui.filter.slice(4))+' ✕']);
  for(const [k,lbl] of fdefs){
    const b=el('button',{class:ui.filter===k?'active':''},lbl);
    b.addEventListener('click',()=>{ ui.filter = k.startsWith('cat:') ? 'all' : k; render(); });
    filters.append(b);
  }
  let rows=m.rows.slice().sort((a,b)=>dateNum(b.date)-dateNum(a.date));
  if(ui.filter==='uncat') rows=rows.filter(r=>!r.category && !r.pending);
  else if(ui.filter==='abroad') rows=rows.filter(r=>/חו"ל/.test(r.details));
  else if(ui.filter==='refund') rows=rows.filter(r=>r.isRefund);
  else if(ui.filter==='pending') rows=rows.filter(r=>r.pending);
  else if(ui.filter==='split') rows=rows.filter(r=>r.split);
  else if(ui.filter.startsWith('cat:')){ const c=ui.filter.slice(4); rows=rows.filter(r=>(r.category||'—')===c); }

  const tb=el('tbody');
  for(const r of rows){
    const catSel=el('select',{class:'catsel'+(r.autoMatched?' auto':'')});
    catSel.append(el('option',{value:''},'— בחירת קטגוריה —'));
    for(const c of state.categories) catSel.append(el('option',{value:c, selected:r.category===c},c));
    catSel.addEventListener('change',()=>{
      r.category=catSel.value; r.autoMatched=false;
      if(r.category) rememberMerchant(r.merchant, r.category);
      markDirty(); render();
    });
    const noteIn=el('input',{class:'note-in', type:'text', value:r.notes||'', placeholder:'הערה'});
    noteIn.addEventListener('change',()=>{ r.notes=noteIn.value; markDirty(); });
    const splitSel=el('select',{class:'catsel splitsel '+(r.split||''), title:'הוצאה משותפת — להזין בספליטווייז'});
    for(const [v,lbl] of [['','—'],['todo','להזין'],['done','הוזן ✓']]) splitSel.append(el('option',{value:v, selected:(r.split||'')===v},lbl));
    splitSel.addEventListener('change',()=>{ setSplit(r,splitSel.value); markDirty(); render(); });
    let refSel=null;
    if(r.category===refundCat() && r.amount>0 && !r.pending){
      refSel=el('select',{class:'catsel splitsel '+(r.refund==='got'?'done':'todo'), title:'האם ההחזר התקבל'});
      for(const [v,lbl] of [['','ממתין להחזר'],['got','הוחזר ✓']]) refSel.append(el('option',{value:v, selected:(r.refund||'')===v},lbl));
      refSel.addEventListener('change',()=>{ r.refund=refSel.value; markDirty(); render(); });
    }
    const del=el('button',{title:'מחיקה'},'✕');
    del.addEventListener('click',()=>undoable('נמחקה ההוצאה: '+r.merchant.trim(), ()=>{ m.rows.splice(m.rows.indexOf(r),1); }));
    tb.append(el('tr',{class:(!r.category&&!r.pending?'uncat':'')+(r.pending?' pending':''), 'data-q':[r.merchant,r.notes||'',r.category||'',r.details||'',r.card||'',r.amount].join(' ').toLowerCase()},
      el('td',{class:'num'},r.date),
      el('td',{class:'merch', title:r.merchant + (r.details? ' · '+r.details : '')}, r.merchant,
        r.pending? el('span',{class:'flag'},'טרם נקלט') : null,
        /חו"ל/.test(r.details)? el('span',{class:'flag'},'חו"ל') : null,
        /תשלום\s+\d+\s+מתוך/.test(r.details)? el('span',{class:'flag'},'תשלומים') : null),
      el('td',{class:'num'+(r.amount<0?' neg':'')}, r.amount!=null? (r.currency && r.currency!=='₪'? [fmtN2(ils(r)), el('span',{class:'flag', title:'חויב במט"ח, הומר לפי שער '+fxRate(r.currency)}, fmtN2(r.amount)+' '+r.currency)] : fmtN2(r.amount)) : (r.txAmount!=null? fmtN2(r.txAmount)+' '+r.txCurrency : '')),
      el('td',null,catSel),
      el('td',null,noteIn),
      el('td',{class:'trackcell'},refSel,splitSel),
      el('td',{class:'small muted num'},r.card),
      el('td',{class:'row-actions'},del)
    ));
  }
  const table=el('table',null,
    el('thead',null, el('tr',null, el('th',null,'תאריך'), el('th',null,'בית עסק'), el('th',null,'חיוב ₪'), el('th',null,'קטגוריה'), el('th',null,'הערות'), el('th',null,'מעקב'), el('th',null,'כרטיס'), el('th',null,''))),
    tb);
  let splitBanner=null;
  if(splitTodo.length){
    const tot=round2(splitTodo.reduce((s,r)=>s+ils(r),0));
    const doneBtn=el('button',null,'סימון הכל כהוזן');
    doneBtn.addEventListener('click',()=>undoable(splitTodo.length+' הוצאות סומנו כהוזנו', ()=>{ for(const r of splitTodo) r.split='done'; }));
    const showBtn=el('button',null,'הצגה');
    showBtn.addEventListener('click',()=>{ ui.filter='split'; render(); });
    splitBanner=el('div',{class:'banner warn splitbanner'},
      el('div',null, el('b',null,'להזין בספליטווייז: '), splitTodo.length+' הוצאות משותפות · '+fmtN2(tot)+' ₪ (חצי: '+fmtN2(round2(tot/2))+' ₪)'),
      el('div',{class:'small'}, splitTodo.map(r=>r.merchant.trim()+' '+fmtN2(ils(r))).join(' · ')),
      el('div',{class:'splitbtns'}, showBtn, doneBtn));
  } else if(splitAll.length){
    splitBanner=el('div',{class:'banner info'}, 'ספליטווייז: כל '+splitAll.length+' ההוצאות המשותפות של החודש הוזנו ✓');
  }
  // free-text search: hides rows in place (no re-render, so the box keeps focus)
  const search=el('input',{type:'search', class:'search', placeholder:'חיפוש בית עסק / הערה / סכום', value:ui.search||'', 'aria-label':'חיפוש בהוצאות'});
  const shown=el('span',{class:'small muted num'});
  const applySearch=()=>{
    const q=(ui.search||'').trim().toLowerCase(); let n=0, sum=0;
    [...tb.children].forEach((tr,i)=>{ const hit=!q || tr.getAttribute('data-q').includes(q); tr.hidden=!hit; if(hit){ n++; if(!rows[i].pending) sum+=ils(rows[i]); } });
    shown.textContent = rows.length? n+' שורות · '+fmtN2(sum)+' ₪' : '';
  };
  search.addEventListener('input',()=>{ ui.search=search.value; applySearch(); });
  applySearch();
  // manual expense (cash, a transfer that never shows on a statement)
  const xDate=el('input',{type:'text', placeholder:'תאריך (31.07.2026)', style:'width:150px', 'aria-label':'תאריך'});
  const xName=el('input',{type:'text', placeholder:'על מה', style:'width:150px', 'aria-label':'תיאור'});
  const xAmt=el('input',{type:'number', step:'0.01', placeholder:'סכום ₪', style:'width:100px', 'aria-label':'סכום'});
  const xCat=el('select',{class:'catsel', 'aria-label':'קטגוריה'}, el('option',{value:''},'— קטגוריה —'));
  for(const c of state.categories) xCat.append(el('option',{value:c},c));
  const xBtn=el('button',{class:'primary'},'הוספת הוצאה ידנית');
  xBtn.addEventListener('click',()=>{
    const d=parseDMY(xDate.value), a=parseAmount(xAmt.value), nm=xName.value.trim();
    if(!d||a==null||!nm){ toast('צריך תאריך תקין (יום.חודש.שנה), תיאור וסכום'); return; }
    const k=monthKeyFromDate(d);
    ensureMonth(k).rows.push({id:uid(), date:fmtDMY(d), merchant:nm, amount:round2(a), card:MANUAL_CARD, currency:'₪', txCurrency:'₪',
      voucher:'manual:'+uid(), txAmount:round2(a), details:'הוזן ידנית', category:xCat.value, autoMatched:false, notes:'', isRefund:a<0, pending:false, billedKey:k});
    if(xCat.value) rememberMerchant(nm,xCat.value);
    ui.month=k; markDirty(); render();
  });
  app.append(el('div',{class:'card', id:'exptable'}, el('h2',null,'הוצאות'), splitBanner,
    el('div',{class:'tabletools'}, filters, el('div',{class:'grow'}), shown, search),
    rows.length? el('div',{class:'tablewrap'},table) : el('div',{class:'muted'}, m.rows.length? 'אין שורות שמתאימות לסינון הזה.' : 'אין עדיין הוצאות בחודש הזה.'),
    el('div',{class:'inc-add'}, xDate, xName, xAmt, xCat, xBtn)));

  // incomes + reconciliation
  const incBody=el('tbody');
  const incSorted=m.incomes.slice().sort((a,b)=>dateNum(b.date)-dateNum(a.date));
  for(const inc of incSorted){
    const csel=el('select',{class:'catsel'});
    for(const c of state.incomeCategories) csel.append(el('option',{value:c, selected:inc.category===c},c));
    if(inc.category && !state.incomeCategories.includes(inc.category)) csel.append(el('option',{value:inc.category, selected:true},inc.category));
    csel.addEventListener('change',()=>{ inc.category=csel.value; markDirty(); });
    const noteIn=el('input',{class:'note-in', type:'text', value:inc.notes||'', placeholder:'הערה'});
    noteIn.addEventListener('change',()=>{ inc.notes=noteIn.value; markDirty(); });
    const del=el('button',{title:'מחיקה'},'✕');
    del.addEventListener('click',()=>undoable('נמחקה ההכנסה: '+inc.source, ()=>{ m.incomes.splice(m.incomes.indexOf(inc),1); }));
    incBody.append(el('tr',null,
      el('td',{class:'num'},inc.date),
      el('td',null,inc.source),
      el('td',{class:'num'},fmtN2(inc.amount)),
      el('td',null,csel),
      el('td',null,noteIn),
      el('td',{class:'row-actions'},del)));
  }
  const addDate=el('input',{type:'text', placeholder:'תאריך (31.07.2026)', style:'width:150px'});
  const addSrc=el('input',{type:'text', placeholder:'מקור', style:'width:150px'});
  const addAmt=el('input',{type:'number', step:'0.01', placeholder:'סכום', style:'width:100px'});
  const addBtn=el('button',{class:'primary'},'הוספה');
  addBtn.addEventListener('click',()=>{
    const d=parseDMY(addDate.value); const a=parseAmount(addAmt.value);
    if(!d||a==null||!addSrc.value.trim()){ toast('צריך תאריך תקין, מקור וסכום'); return; }
    ensureMonth(monthKeyFromDate(d)).incomes.push({id:uid(), date:fmtDMY(d), source:addSrc.value.trim(), amount:round2(a), category:'החזרים', notes:''});
    markDirty(); render();
  });
  const incCard=el('div',{class:'card'}, el('h2',null,'הכנסות'),
    m.incomes.length? el('div',{class:'tablewrap'}, el('table',null, el('thead',null, el('tr',null, el('th',null,'תאריך'),el('th',null,'מקור'),el('th',null,'סכום ₪'),el('th',null,'סוג'),el('th',null,'הערות'),el('th',null,''))), incBody)) : el('div',{class:'muted'},'אין הכנסות לחודש הזה — ייבוא דוח עו"ש יוסיף אותן אוטומטית (רק של החודש הזה!).'),
    el('div',{class:'inc-add'}, addDate, addSrc, addAmt, addBtn));

  // reconciliation with gap breakdown
  const reconWrap=el('div',{class:'recon'});
  const reconList=m.recon||[];
  const bankByCard={};
  for(const rc of reconList){ const key=Object.keys(bankByCard).find(k=>sameCard(k,rc.card)) || rc.card; bankByCard[key]=round2((bankByCard[key]||0)+(rc.bankAmount||0)); }
  for(const [card,bankAmt] of Object.entries(bankByCard)){
    const inMonth = m.rows.filter(r=>!r.pending && r.amount!=null && sameCard(card, r.card||'?'));
    const appTotal = round2(inMonth.reduce((a,r)=>a+ils(r),0));
    // rows purchased this month but billed in another cycle
    const billedLaterRows = inMonth.filter(r=>r.billedKey && r.billedKey!==mk && r.currency==='₪');
    const billedLater = round2(billedLaterRows.reduce((a,r)=>a+r.amount,0));
    // foreign-currency rows: charged separately in their own currency
    const foreignRows = inMonth.filter(r=>r.currency && r.currency!=='₪');
    const foreign = round2(foreignRows.reduce((a,r)=>a+ils(r),0));
    // rows in other months billed with this month's cycle
    let fromOther=0; const fromOtherRows=[];
    for(const [k2,m2] of Object.entries(state.months)){
      if(k2===mk) continue;
      for(const r of m2.rows) if(!r.pending && r.amount!=null && r.billedKey===mk && sameCard(card, r.card||'?') && r.currency==='₪'){ fromOther+=r.amount; fromOtherRows.push({k2,r}); }
    }
    fromOther=round2(fromOther);
    const expected = round2(appTotal - billedLater - foreign + fromOther);
    const residual = round2(bankAmt - expected);
    let explained = Math.abs(residual)<1;
    const lines=[];
    if(billedLaterRows.length) lines.push(el('div',{class:'small muted num'},'− '+fmtN2(billedLater)+' נרכשו החודש אך יחויבו במחזור אחר ('+billedLaterRows.map(r=>r.merchant.trim()).join(', ')+')'));
    if(foreignRows.length) lines.push(el('div',{class:'small muted num'},'− '+fmtN2(foreign)+' עסקאות במט"ח שמחויבות בנפרד ('+foreignRows.map(r=>r.merchant.trim()+' '+r.currency).join(', ')+')'));
    if(fromOtherRows.length) lines.push(el('div',{class:'small muted num'},'+ '+fmtN2(fromOther)+' חיוב החודש עבור רשומות מחודשים אחרים ('+fromOtherRows.map(x=>x.r.merchant.trim()+' · '+monthLabel(x.k2)).join(', ')+')'));
    if(!explained){
      // hint: an installment/credit row recorded in another month with a matching amount
      let hint=null;
      outer:
      for(const [k2,m2] of Object.entries(state.months)){
        if(k2===mk) continue;
        for(const r of m2.rows){
          if(sameCard(card, r.card||'?') && /תשלום\s+\d+\s+מתוך|קרדיט/.test(r.details||'') && Math.abs(r.amount-residual)<1){ hint={k2,r}; break outer; }
        }
      }
      if(hint){
        explained=true;
        lines.push(el('div',{class:'small muted num'},'+ '+fmtN2(residual)+' חיוב החודש עבור '+hint.r.merchant.trim()+' שרשום ב'+monthLabel(hint.k2)+' ('+(hint.r.details||'')+')'));
      } else {
        lines.push(el('div',{class:'small num', style:'color:var(--expense)'},'פער שנותר ללא הסבר: '+fmtN2(residual)));
      }
    }
    const det=el('details',null,
      el('summary',{class:'small', style:'cursor:pointer;color:var(--muted)'}, explained? 'פירוט ההתאמה' : 'הסבר הפער'),
      el('div',{class:'num small', style:'margin-top:4px'},'נרשם החודש: '+fmtN2(appTotal)), ...lines,
      el('div',{class:'num small'},'= צפוי בבנק: '+fmtN2(expected)+' · בפועל: '+fmtN2(bankAmt)));
    const diff=round2(bankAmt-appTotal);
    reconWrap.append(el('div',{class:'reconbox'},
      el('div',null,'כרטיס '+card),
      el('div',{class:'num'},'חויב בבנק (נטו): '+fmtN2(bankAmt)),
      el('div',{class:'num'},'נרשם כאן: '+fmtN2(appTotal)),
      el('div',{class:explained?'ok':'bad'}, explained? (Math.abs(diff)<1?'תואם ✓':'תואם ✓ (הפער מוסבר)') : 'פער לא מוסבר: '+fmtN2(residual)),
      det));
  }
  const reconCard=el('div',{class:'card'}, el('h2',null,'התאמה מול הבנק'),
    reconList.length? reconWrap : el('div',{class:'muted'},'ייבוא דוח עו"ש ימלא כאן אוטומטית את חיובי כרטיסי האשראי מול מה שרשום באפליקציה.'));
  app.append(el('div',{class:'grid2'}, incCard, reconCard));

  // history overview
  const histBody=el('tbody');
  for(const k of mks){
    const tt=monthTotals(k);
    const tr=el('tr',{style:'cursor:pointer'},
      el('td',null,monthLabel(k)),
      el('td',{class:'num'},fmtN(tt.exp)),
      el('td',{class:'num'},fmtN(tt.inc)),
      el('td',{class:'num', style:'color:'+(tt.net>=0?'var(--income)':'var(--expense)')},fmtN(tt.net)),
      el('td',{class:'num'},String(tt.uncat||'')));
    tr.addEventListener('click',()=>{ ui.month=k; render(); window.scrollTo({top:0}); });
    histBody.append(tr);
  }
  const hist=el('details',{class:'sec'}, el('summary',null,'סקירה חודשית'),
    el('div',{class:'card tablewrap'}, el('table',{class:'histtable'},
      el('thead',null, el('tr',null, el('th',null,'חודש'),el('th',null,'הוצאות ₪'),el('th',null,'הכנסות ₪'),el('th',null,'מאזן ₪'),el('th',null,'לסיווג'))),
      histBody)));
  app.append(hist);

  renderSettings();
}


/* ---------- monthly dashboard ---------- */
const CAT_COLORS_LIGHT=['#2a78d6','#eb6834','#1baf7a','#eda100','#e87ba4','#008300','#4a3aa7','#e34948'];
const CAT_COLORS_DARK =['#3987e5','#d95926','#199e70','#c98500','#d55181','#008300','#9085e9','#e66767'];
const OTHER_CAT='__other__';
function isDark(){ const t=document.documentElement.getAttribute('data-theme'); if(t==='dark') return true; if(t==='light') return false; return window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches; }
function excludedCats(){ return (state.settings.excludeCats||[]); }
/* stable colour slots: categories ranked by all-time spend (changes only slowly), top 8 get a hue, the rest fold into "אחר" */
function catSlots(){
  const tot={};
  for(const m of Object.values(state.months)) for(const r of m.rows) if(!r.pending && r.category) tot[r.category]=(tot[r.category]||0)+Math.abs(ils(r));
  const ranked=Object.keys(tot).sort((a,b)=>tot[b]-tot[a]);
  const slots={}; ranked.slice(0,8).forEach((c,i)=>slots[c]=i);
  return slots;
}
function catColor(cat, slots){ const i=slots[cat]; if(i==null) return 'var(--muted)'; return (isDark()? CAT_COLORS_DARK : CAT_COLORS_LIGHT)[i]; }
function catTotals(mk){ // ₪ by category for one month, excluding pending + excluded categories
  const out={}; const m=state.months[mk]; if(!m) return out;
  const ex=excludedCats();
  for(const r of m.rows){ if(r.pending) continue; const c=r.category||'—'; if(ex.includes(c)) continue; out[c]=(out[c]||0)+ils(r); }
  return out;
}
function monthKeysBack(mk,n){ const ks=[mk]; let k=mk; for(let i=1;i<n;i++){ k=prevMonthKey(k); ks.unshift(k); } return ks; }
function shortMonth(k){ const [y,m]=k.split('-').map(Number); return HEB_MONTHS[m-1].slice(0,3)+"׳ "+String(y).slice(2); }

let tipEl=null;
function showTip(x,y,html){
  if(!tipEl){ tipEl=el('div',{class:'viztip'}); document.body.append(tipEl); }
  tipEl.textContent=''; tipEl.append(...html); tipEl.style.display='block';
  const w=tipEl.offsetWidth, h=tipEl.offsetHeight;
  let left=x+12, top=y-h-10; if(left+w>window.innerWidth-8) left=x-w-12; if(top<8) top=y+14;
  tipEl.style.left=left+'px'; tipEl.style.top=top+'px';
}
function hideTip(){ if(tipEl) tipEl.style.display='none'; }

function renderDashboard(mk){
  const slots=catSlots();
  const cur=catTotals(mk), prev=catTotals(prevMonthKey(mk));
  const total=Object.values(cur).reduce((a,b)=>a+b,0);
  const prevTotal=Object.values(prev).reduce((a,b)=>a+b,0);
  const win=monthKeysBack(mk,6);
  const winTot={}; for(const k of win){ const t=catTotals(k); for(const [c,v] of Object.entries(t)) winTot[c]=(winTot[c]||0)+v; }
  const avg={}; for(const [c,v] of Object.entries(winTot)) avg[c]=v/win.filter(k=>state.months[k]).length;

  /* --- breakdown rows --- */
  const cats=Object.entries(cur).filter(([c,v])=>v!==0).sort((a,b)=>b[1]-a[1]);
  const mx=Math.max(1,...cats.map(c=>Math.abs(c[1])));
  const rows=el('div',{class:'dashrows'});
  const head=el('div',{class:'dashrow head'}, el('div',null,'קטגוריה'), el('div',null,''), el('div',{class:'amt'},'₪'), el('div',{class:'amt'},'חלק'), el('div',{class:'amt'},'מול חודש קודם'), el('div',{class:'amt'},'ממוצע 6 ח׳'));
  rows.append(head);
  for(const [c,v] of cats){
    const p=prev[c]||0; const d=v-p;
    let dEl;
    if(!p && v) dEl=el('span',{class:'muted'},'חדש');
    else { const pct=Math.round(d/Math.abs(p)*100); dEl=el('span',{class:d>0?'up':(d<0?'down':'muted')}, (d>0?'▲ ':(d<0?'▼ ':'')) + fmtN(Math.abs(d)) + (Math.abs(pct)<1000? ' ('+(pct>0?'+':'')+pct+'%)':'')); }
    const label=c==='—'?'ללא סיווג':c;
    const r=el('div',{class:'dashrow', role:'button', tabindex:'0', title:'סינון הטבלה לפי '+label},
      el('div',{class:'catname'}, el('span',{class:'sw', style:'background:'+(c==='—'?'var(--warn)':catColor(c,slots))}), label),
      el('div',{class:'barcell'}, el('div',{class:'bar', style:'width:'+Math.max(1.5,Math.round(Math.abs(v)/mx*100))+'%'+(c==='—'?';background:var(--warn)':'')}),
        avg[c]? el('div',{class:'avgmark', style:'right:'+Math.min(100,Math.round(avg[c]/mx*100))+'%', title:'ממוצע 6 חודשים: '+fmtN(avg[c])}) : null),
      el('div',{class:'amt num'+(v<0?' neg':'')}, fmtN(v)),
      el('div',{class:'amt num muted'}, total? Math.round(v/total*100)+'%' : ''),
      el('div',{class:'amt num small'}, dEl),
      el('div',{class:'amt num muted small'}, avg[c]? fmtN(avg[c]) : '')
    );
    const pick=()=>{ ui.filter = ui.filter==='cat:'+c ? 'all' : 'cat:'+c; render(); const t=document.getElementById('exptable'); if(t) t.scrollIntoView({behavior:'smooth',block:'start'}); };
    r.addEventListener('click',pick);
    r.addEventListener('keydown',e=>{ if(e.key==='Enter'||e.key===' '){ e.preventDefault(); pick(); } });
    rows.append(r);
  }
  const totLine=el('div',{class:'dashtotal num'},
    el('span',null,'סה"כ הוצאות בלוח: '), el('b',null,fmtN(total)+' ₪'),
    prevTotal? el('span',{class:(total-prevTotal)>0?'up':'down', style:'margin-inline-start:10px'}, ((total-prevTotal)>0?'▲ ':'▼ ')+fmtN(Math.abs(total-prevTotal))+' מול '+monthLabel(prevMonthKey(mk))) : null);

  /* --- exclusion chips --- */
  const ex=excludedCats();
  const chips=el('div',{class:'filterchips', style:'margin-top:8px'}, el('span',{class:'small muted', style:'align-self:center'},'לא לספור בלוח: '));
  for(const c of state.categories){
    const b=el('button',{class:ex.includes(c)?'active':'', title:ex.includes(c)?'לחיצה תחזיר את הקטגוריה ללוח':'לחיצה תסתיר את הקטגוריה מהלוח (למשל חיסכון או הוצאות שמוחזרות)'}, c);
    b.addEventListener('click',()=>{ const l=excludedCats().slice(); const i=l.indexOf(c); if(i>=0) l.splice(i,1); else l.push(c); state.settings.excludeCats=l; markDirty(); render(); });
    chips.append(b);
  }

  /* --- 6-month stacked trend --- */
  const trend=renderTrend(win, slots, mk);

  return el('div',{class:'card dash'},
    el('div',{class:'dashhead'}, el('h2',null,'לוח חודשי · '+monthLabel(mk)), el('span',{class:'small muted'},'לחיצה על קטגוריה מסננת את טבלת ההוצאות')),
    cats.length? rows : el('div',{class:'muted'},'אין הוצאות מסווגות החודש.'),
    totLine,
    el('div',{class:'grid2', style:'margin-top:14px'},
      el('div',null, el('h2',null,'מגמה — 6 חודשים אחרונים'), trend),
      el('div',null, el('h2',null,'הגדרות הלוח'), chips,
        el('div',{class:'small muted', style:'margin-top:8px'},'עסקאות במט"ח מומרות לש"ח לפי השער בהגדרות (€ '+fxRate('€')+' · $ '+fxRate('$')+'). ניתן לשנות בהגדרות למטה.')))
  );
}

function renderTrend(keys, slots, activeKey){
  const W=560, H=250, padL=8, padR=58, padT=26, padB=34;
  const data=keys.map(k=>({k, t:catTotals(k)}));
  const colored=Object.keys(slots).sort((a,b)=>slots[a]-slots[b]);
  const series=colored.concat(['—',OTHER_CAT]);
  const stacks=data.map(d=>{ let other=0; const seg={}; for(const [c,v] of Object.entries(d.t)){ if(v<=0) continue; if(slots[c]!=null || c==='—') seg[c]=v; else other+=v; } if(other>0) seg[OTHER_CAT]=other; const tot=Object.values(seg).reduce((a,b)=>a+b,0); return {k:d.k, seg, tot}; });
  const maxT=Math.max(1,...stacks.map(s=>s.tot));
  const niceMax=(()=>{ const p=Math.pow(10,Math.floor(Math.log10(maxT))); const m=maxT/p; const n=m<=1?1:m<=2?2:m<=2.5?2.5:m<=5?5:10; return n*p; })();
  const y=v=>padT+(H-padT-padB)*(1-v/niceMax);
  const n=stacks.length, slotW=(W-padL-padR)/n, barW=Math.min(48,slotW*0.62);
  const svgNS='http://www.w3.org/2000/svg';
  const sv=(tag,attrs)=>{ const e=document.createElementNS(svgNS,tag); for(const [k,v] of Object.entries(attrs||{})) e.setAttribute(k,v); return e; };
  const svg=sv('svg',{viewBox:'0 0 '+W+' '+H, class:'trend', role:'img','aria-label':'הוצאות לפי קטגוריה, 6 חודשים', direction:'ltr'});
  // gridlines (recessive) + y labels on the right (RTL)
  for(let i=0;i<=4;i++){ const v=niceMax*i/4; const yy=y(v); svg.append(sv('line',{x1:padL,x2:W-padR,y1:yy,y2:yy,class:'grid'})); const t=sv('text',{x:W-2,y:yy+4,class:'ylab'}); t.textContent=fmtN(v); svg.append(t); }
  stacks.forEach((s,i)=>{
    const cx=W-padR-slotW*(i+0.5); // newest month on the left in RTL reading order? keep chronological right→left
    const x0=cx-barW/2; let acc=0;
    for(const c of series){
      const v=s.seg[c]; if(!v) continue;
      const y1=y(acc+v), y0=y(acc); const hgt=Math.max(0,y0-y1-1.5);
      const rect=sv('rect',{x:x0,y:y1,width:barW,height:hgt,rx:1.5,fill:c===OTHER_CAT?'var(--muted)':(c==='—'?'var(--warn)':catColor(c,slots)),class:'seg'+(s.k===activeKey?'':' dim')});
      const label=c===OTHER_CAT?'אחר (קטגוריות קטנות)':(c==='—'?'ללא סיווג':c);
      rect.addEventListener('mousemove',e=>showTip(e.clientX,e.clientY,[el('div',{class:'tiptitle'},monthLabel(s.k)), el('div',null,label+': '), el('b',{class:'num'},fmtN(v)+' ₪'), el('div',{class:'muted small'},Math.round(v/s.tot*100)+'% מתוך '+fmtN(s.tot))]));
      rect.addEventListener('mouseleave',hideTip);
      svg.append(rect); acc+=v;
    }
    if(s.tot>0){ const t=sv('text',{x:cx,y:y(s.tot)-6,class:'tot'}); t.textContent=fmtN(s.tot); svg.append(t); }
    const xl=sv('text',{x:cx,y:H-padB+16,class:'xlab'+(s.k===activeKey?' active':'')}); xl.textContent=shortMonth(s.k); svg.append(xl);
    const hit=sv('rect',{x:cx-slotW/2,y:padT,width:slotW,height:H-padT-padB,fill:'transparent',style:'cursor:pointer'});
    hit.addEventListener('click',()=>{ if(state.months[s.k]){ ui.month=s.k; render(); } });
    svg.append(hit);
  });
  svg.append(sv('line',{x1:padL,x2:W-padR,y1:y(0),y2:y(0),class:'axis'}));
  const legend=el('div',{class:'legend'});
  for(const c of series){ if(!stacks.some(s=>s.seg[c])) continue; legend.append(el('span',{class:'lg'}, el('span',{class:'sw', style:'background:'+(c===OTHER_CAT?'var(--muted)':(c==='—'?'var(--warn)':catColor(c,slots)))}), c===OTHER_CAT?'אחר (קטגוריות קטנות)':(c==='—'?'ללא סיווג':c))); }
  return el('div',{class:'trendwrap'}, svg, legend);
}

function renderSettings(){
  const dl=(name, text, type)=>{
    const b=new Blob([text],{type:type||'text/plain;charset=utf-8'});
    const a=el('a',{href:URL.createObjectURL(b), download:name});
    document.body.append(a); a.click(); a.remove();
  };
  const expJson=el('button',null,'ייצוא גיבוי JSON');
  expJson.addEventListener('click',()=>{
    dl('expense_backup_'+new Date().toISOString().slice(0,10)+'.json', JSON.stringify(state,null,1),'application/json');
  });
  const expCsv=el('button',null,'ייצוא CSV לחודש הנוכחי');
  expCsv.addEventListener('click',()=>{
    if(!ui.month) return;
    const m=state.months[ui.month];
    const q=v=>{ v=String(v==null?'':v); return /[",\n]/.test(v)? '"'+v.replace(/"/g,'""')+'"' : v; };
    const lines=['כרטיס,תאריך,הוצאה,עסקה,מטבע,חיוב,מטבע חיוב,מספר שובר,פירוט,סטטוס,הערות'];
    for(const r of m.rows) lines.push([r.card,r.date,r.merchant,r.txAmount,r.txCurrency,r.amount,r.currency,r.voucher,r.category,r.details,r.notes].map(q).join(','));
    dl('הוצאות_'+monthLabel(ui.month).replace(' ','_')+'.csv','﻿'+lines.join('\n'),'text/csv;charset=utf-8');
  });
  const dlApp=el('button',null,'הורדת האפליקציה כקובץ HTML עם כל הנתונים');
  dlApp.addEventListener('click',()=>{ dl('ניהול_הוצאות_'+new Date().toISOString().slice(0,10)+'.html', buildDoc(),'text/html;charset=utf-8'); toast('הקובץ נשמר. אפשר לפתוח אותו בכל דפדפן — הנתונים בתוכו.'); });
  const dlBlank=el('button',null,'הורדת עותק ריק לשיתוף (למשפחה)');
  dlBlank.addEventListener('click',()=>{ const b=blankState(); b.settings.fx=Object.assign({},state.settings.fx); b.categories=DEFAULT_CATS.filter(c=>!/שאני מקבלת/.test(c)).concat([BANK_SAVINGS_CAT]); dl('ניהול_הוצאות_ריק.html', buildDoc(b),'text/html;charset=utf-8'); toast('עותק ריק נשמר — בלי הנתונים והזיכרון שלך. אפשר לשלוח אותו לכל אחד.'); });
  const fxIn={};
  const fxRow=el('div',{class:'fxrow'}, el('span',null,'שערי המרה לש"ח:'));
  for(const cur of ['€','$','£']){
    const inp=el('input',{type:'number', step:'0.0001', value:String(fxRate(cur))}); fxIn[cur]=inp;
    inp.addEventListener('change',()=>{ const v=parseFloat(inp.value); if(v>0){ state.settings.fx[cur]=v; markDirty(); render(); } });
    fxRow.append(el('label',{style:'display:flex;align-items:center;gap:4px'}, cur, inp));
  }
  const autoPub=el('input',{type:'checkbox', id:'autopub'});
  autoPub.checked=!!state.settings.autoPublish;
  autoPub.addEventListener('change',()=>{ state.settings.autoPublish=autoPub.checked; markDirty(); });

  const merBody=el('tbody');
  const merEntries=Object.entries(state.merchants).sort((a,b)=>a[0].localeCompare(b[0],'he'));
  for(const [n,c] of merEntries){
    const del=el('button',null,'✕');
    del.addEventListener('click',()=>undoable('נמחק מהזיכרון: '+n, ()=>{ delete state.merchants[n]; rebuildNormMerchants(); }));
    merBody.append(el('tr',null, el('td',null,n), el('td',null,c), el('td',{class:'row-actions'},del)));
  }
  const usedCats={}; for(const m2 of Object.values(state.months)) for(const r of m2.rows) if(r.category) usedCats[r.category]=(usedCats[r.category]||0)+1;
  const catList=el('div',{class:'filterchips'});
  for(const c of state.categories){
    const n=usedCats[c]||0;
    const b=el('button',{title: n? 'בשימוש ב-'+n+' הוצאות — אי אפשר להסיר' : 'הסרת הקטגוריה', disabled: n>0 || c===refundCat()}, c+(n? ' · '+n : ' ✕'));
    b.addEventListener('click',()=>undoable('הוסרה הקטגוריה: '+c, ()=>{
      state.categories=state.categories.filter(x=>x!==c);
      for(const k of Object.keys(state.merchants)) if(state.merchants[k]===c) delete state.merchants[k];
      state.settings.excludeCats=(state.settings.excludeCats||[]).filter(x=>x!==c);
      rebuildNormMerchants();
    }));
    catList.append(b);
  }
  const kwIn=el('input',{type:'text', style:'flex:1;min-width:220px', value:splitKeywords().join(', '), placeholder:'למשל: חברת החשמל, ארנונה, גז', 'aria-label':'מילות מפתח להוצאות משותפות'});
  kwIn.addEventListener('change',()=>{ state.settings.splitKeywords=kwIn.value.split(',').map(x=>x.trim()).filter(Boolean); markDirty(); });
  const aliasText=Object.entries(state.settings.cardAliases||{}).map(([n,ids])=>n+'='+(ids||[]).join('/')).join(', ');
  const alIn=el('input',{type:'text', style:'flex:1;min-width:220px', value:aliasText, placeholder:'למשל: מקס=1234, ישראכרט=5678', 'aria-label':'שיוך שם חברת אשראי לספרות הכרטיס'});
  alIn.addEventListener('change',()=>{
    const out={};
    for(const part of alIn.value.split(',')){ const i=part.indexOf('='); if(i<0) continue; const n=part.slice(0,i).trim(); const ids=part.slice(i+1).split('/').map(x=>x.trim()).filter(Boolean); if(n && ids.length) out[n]=ids; }
    state.settings.cardAliases=out; markDirty(); render();
  });
  const addCat=el('input',{type:'text', placeholder:'קטגוריה חדשה', style:'width:170px', 'aria-label':'קטגוריה חדשה'});
  const addCatB=el('button',null,'הוספת קטגוריה');
  addCatB.addEventListener('click',()=>{ const v=addCat.value.trim(); if(v && !state.categories.includes(v)){ state.categories.push(v); addCat.value=''; markDirty(); render(); } });
  const incCatList=el('div',{class:'small muted', style:'margin-top:6px'}, state.incomeCategories.join(' · '));
  const addIncCat=el('input',{type:'text', placeholder:'סוג הכנסה חדש', style:'width:170px'});
  const addIncCatB=el('button',null,'הוספת סוג הכנסה');
  addIncCatB.addEventListener('click',()=>{ const v=addIncCat.value.trim(); if(v && !state.incomeCategories.includes(v)){ state.incomeCategories.push(v); addIncCat.value=''; markDirty(); render(); } });

  const sec=el('details',{class:'sec'}, el('summary',null,'הגדרות, גיבוי וזיכרון בתי עסק'),
    el('div',{class:'card'},
      el('div',{style:'display:flex;gap:8px;flex-wrap:wrap;align-items:center'}, expJson, expCsv, dlApp, dlBlank,
        isCloud()? el('label',{style:'display:flex;align-items:center;gap:5px;font-size:.85rem'}, autoPub, 'שמירה אוטומטית לענן (15 שנ׳ אחרי שינוי)') : null),
      fxRow,
      el('div',{class:'banner info'}, isCloud()
        ? 'ייבוא גיבוי: פשוט לגרור את קובץ ה-JSON לאזור הייבוא למעלה. שמירה לענן = הקובץ נשמר בתוך העמוד הזה עצמו וזמין מכל מכשיר.'
        : 'הגרסה הזו רצה כקובץ: הנתונים נשמרים בדפדפן הזה בלבד. כדי לגבות או להעביר למכשיר אחר — "הורדת האפליקציה כקובץ HTML עם כל הנתונים" ופתיחת הקובץ החדש. גיבוי JSON נטען בגרירה לאזור הייבוא.'),
      el('h2',{style:'margin-top:14px'},'קטגוריות הוצאה'), catList,
      el('div',{style:'display:flex;gap:6px;margin-top:6px'}, addCat, addCatB),
      el('h2',{style:'margin-top:14px'},'סוגי הכנסה'), incCatList,
      el('div',{style:'display:flex;gap:6px;margin-top:6px'}, addIncCat, addIncCatB),
      el('h2',{style:'margin-top:14px'},'הוצאות משותפות (ספליטווייז)'),
      el('div',{class:'small muted'},'בית עסק ששמו מכיל אחת מהמילים יסומן אוטומטית "להזין". סימון ידני של שורה מלמד את האפליקציה גם בלי הרשימה.'),
      el('div',{style:'display:flex;gap:6px;margin-top:6px'}, kwIn),
      el('h2',{style:'margin-top:14px'},'התאמת כרטיסים מול הבנק'),
      el('div',{class:'small muted'},'כשהבנק מציג את החיוב בשם חברת האשראי והדוח של הכרטיס מציג 4 ספרות — כאן מחברים ביניהם (שם=ספרות, כמה ספרות מופרדות ב-/).'),
      el('div',{style:'display:flex;gap:6px;margin-top:6px'}, alIn),
      el('h2',{style:'margin-top:14px'},'זיכרון בתי עסק ('+merEntries.length+')'),
      el('div',{class:'tablewrap', style:'max-height:300px;overflow-y:auto'},
        el('table',null, el('thead',null, el('tr',null, el('th',null,'בית עסק'),el('th',null,'קטגוריה'),el('th',null,''))), merBody))
    ));
  app.append(sec);
}

/* ---------- file intake ---------- */
async function handleFiles(files){
  if(!files.length) return;
  const msgs=[]; const snap=snapshot(); const before=snap.length;
  for(const f of files){
    try{
      const res=await parseUpload(f.name, ()=>f.arrayBuffer(), ()=>f.text());
      if(res.kind==='card'){
        const r=mergeExpenseRows(res.data.rows);
        msgs.push(f.name+': כרטיס '+(res.data.card||'?')+' — נוספו '+r.added+' עסקאות'+(r.upgraded?', עודכנו '+r.upgraded+' שהמתינו לקליטה':'')+(r.refreshed?', '+r.refreshed+' עודכנו (מטבע/סכום)':'')+(r.dup?', '+r.dup+' כפולות דולגו':'')+(r.needCat?', '+r.needCat+' ממתינות לסיווג':''));
      } else if(res.kind==='max'){
        const r=mergeExpenseRows(res.data.rows);
        msgs.push(f.name+': נוספו '+r.added+' עסקאות'+(r.dup?', '+r.dup+' כפולות דולגו':''));
      } else if(res.kind==='bank'){
        const ri=mergeIncomes(res.data.incomes);
        const rr=mergeRecon(res.data.cardDebits);
        const rb=mergeExpenseRows(res.data.bankRows||[]);
        msgs.push(f.name+': נוספו '+ri.added+' הכנסות (כל אחת לחודש שלה)'+(ri.dup?', '+ri.dup+' כפולות דולגו':'')+(ri.skippedCardCo?', '+ri.skippedCardCo+' זיכויים מחברת האשראי לא נספרו כהכנסה':'')+(rb.added?', '+rb.added+' חיובים ישירים מהחשבון (שכ"ד, הוראות קבע, העברות)':'')+(rb.needCat?' — '+rb.needCat+' לסיווג':'')+(rr.added?', עודכנה התאמת בנק':''));
      } else if(res.kind==='appcsv'){
        const r=mergeExpenseRows(res.data.rows,{noLearn:false});
        msgs.push(f.name+': נוספו '+r.added+' הוצאות'+(r.dup?', '+r.dup+' כפולות דולגו':''));
      } else if(res.kind==='backup'){
        msgs.push(...importBackup(res.data));
      } else {
        msgs.push(f.name+': פורמט לא מזוהה');
      }
    }catch(err){
      msgs.push(f.name+': שגיאה — '+err.message);
    }
  }
  const changed = JSON.stringify(state).length!==before || JSON.stringify(state)!==snap;
  if(changed){ rebuildNormMerchants(); markDirty(); }
  render();
  toast(msgs.join('\n'), changed? ()=>restore(snap) : null);
}

/* ---------- boot ---------- */
(function boot(){
  let embedded=null;
  try{ embedded=JSON.parse(document.getElementById('APPSTATE').textContent); }catch(e){}
  let local=null;
  try{ local=JSON.parse(localStorage.getItem(LS_KEY)); }catch(e){}
  if(local && embedded && local.version===2 && embedded.version===2){
    const newer = (local.ts||0) >= (embedded.ts||0) ? local : embedded;
    const older = newer===local? embedded : local;
    state = newer;
    // enrich newer state with data the older one has (billing-cycle info, bank recon lines)
    try{
      if(older.months && state.months){
        const bkByVoucher={};
        for(const m2 of Object.values(older.months)) for(const r of m2.rows||[]) if(r.voucher && r.billedKey) bkByVoucher[r.voucher]=r.billedKey;
        for(const m2 of Object.values(state.months)){
          for(const r of m2.rows||[]) if(!r.billedKey && r.voucher && bkByVoucher[r.voucher]) r.billedKey=bkByVoucher[r.voucher];
        }
        for(const [k2,m2] of Object.entries(older.months)){
          if(!m2.recon || !m2.recon.length) continue;
          if(!state.months[k2]) continue;
          const tgt=state.months[k2]; if(!tgt.recon) tgt.recon=[];
          for(const rc of m2.recon) if(!tgt.recon.some(x=>x.card===rc.card && x.bankAmount===rc.bankAmount && x.date===rc.date)) tgt.recon.push(rc);
        }
        // never lose manual work: copy categories/notes the other copy has and this one lacks
        const manByVoucher={};
        for(const m2 of Object.values(older.months)) for(const r of m2.rows||[]) if(r.voucher && (r.category||r.notes||r.split||r.refund)) manByVoucher[r.voucher]={category:r.category,notes:r.notes,split:r.split,refund:r.refund};
        for(const m2 of Object.values(state.months)) for(const r of m2.rows||[]){
          const o=r.voucher && manByVoucher[r.voucher];
          if(!o) continue;
          if(!r.category && o.category){ r.category=o.category; r.autoMatched=false; rememberMerchant(r.merchant,o.category); }
          if(!r.notes && o.notes) r.notes=o.notes;
          if(r.split==null && o.split) r.split=o.split;
          if(r.refund==null && o.refund) r.refund=o.refund;
        }
        if(older.splitMerchants) state.splitMerchants=Object.assign({},older.splitMerchants,state.splitMerchants||{});
        if(older.settings) state.settings=Object.assign({},older.settings,state.settings||{});
        const incNotes={};
        for(const m2 of Object.values(older.months)) for(const i2 of m2.incomes||[]) if(i2.notes) incNotes[[i2.date,i2.amount,normName(i2.source)].join('|')]=i2.notes;
        for(const m2 of Object.values(state.months)) for(const i2 of m2.incomes||[]){
          const n=incNotes[[i2.date,i2.amount,normName(i2.source)].join('|')];
          if(n && !i2.notes) i2.notes=n;
        }
      }
    }catch(e){}
  } else state = (local && local.version===2 && local) || (embedded && embedded.version===2 && embedded) || blankState();
  state=normalizeState(state);
  rebuildNormMerchants();
  let migNotes=[];
  const hasData=Object.keys(state.months).length>0;
  if(!hasData){
    migNotes=migrateOldLocalStorage();
    if(migNotes.length) saveLocal();
  }
  try{
    const u=JSON.parse(sessionStorage.getItem('em2_ui'));
    if(u){ ui.month=u.month; ui.filter=u.filter||'all'; sessionStorage.removeItem('em2_ui'); setTimeout(()=>window.scrollTo({top:u.y||0}),50); }
  }catch(e){}
  render();
  if(migNotes.length){ ui.dirty=true; ui.saveStatus='dirty'; renderSaveChip(); schedulePublish(); toast(migNotes.join('\n')); }
  window.claude && window.claude.use && window.claude.use('artifact').then(ns=>{ ui.artifactNS=ns; if(!ns){ ui.saveStatus = ui.dirty?'localonly':'saved'; renderSaveChip(); } });
})();
