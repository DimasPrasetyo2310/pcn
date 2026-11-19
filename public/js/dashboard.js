// Dashboard Pencapaian Salesperson, MITRA10 Percetakan Negara (10005)
const D = {store:'MITRA10 Percetakan Negara', code:'10005', days:{}, sched:null};
(function(){
const norm = s => String(s).replace(/\s+/g,' ').trim().toLowerCase();
const num = v => {
  if(typeof v==='number') return v;
  const s = String(v).trim().replace(/\s/g,'');
  if(!s) return 0;
  const n = Number(s.replace(/,/g,''));
  return isNaN(n) ? 0 : n;
};
const MON = {jan:1,feb:2,mar:3,apr:4,may:5,mei:5,jun:6,jul:7,aug:8,agu:8,agt:8,sep:9,oct:10,okt:10,nov:11,dec:12,des:12};
const pad = n => String(n).padStart(2,'0');
const ymd = (y,m,d) => { y=+y; if(y<100) y+=2000; m=+m; d=+d; if(m<1||m>12||d<1||d>31) return ''; return y+'-'+pad(m)+'-'+pad(d); };
const isoDate = v => {
  if(v instanceof Date && !isNaN(v)) return ymd(v.getFullYear(), v.getMonth()+1, v.getDate());
  if(typeof v==='number'){
    if(v>20000 && v<80000) return new Date(Date.UTC(1899,11,30)+Math.floor(v)*86400000).toISOString().slice(0,10);
    return '';
  }
  const s = String(v).trim(); let m;
  if((m=/^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s))) return ymd(m[1],m[2],m[3]);
  if((m=/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/.exec(s))) return ymd(m[3],m[2],m[1]);
  if((m=/^(\d{1,2})[\- ]([A-Za-z]{3,})[\- ,]*(\d{2,4})/.exec(s))){ const mo=MON[m[2].slice(0,3).toLowerCase()]; return mo? ymd(m[3],mo,m[1]) : ''; }
  return '';
};
const hourOf = v => {
  if(typeof v==='number') return v<1 ? Math.min(23,Math.floor(v*24+1e-6)) : null;
  const m = /^(\d{1,2}):/.exec(String(v).trim());
  return m ? +m[1] : null;
};
const clean = n => {
  n = String(n).replace(/\s+/g,' ').trim();
  const w = n.split(' '), h = w.length/2;
  if(w.length%2===0 && w.slice(0,h).join(' ')===w.slice(h).join(' ')) n = w.slice(0,h).join(' ');
  return n;
};
const readRows = (XLSX, buf) => {
  const wb = XLSX.read(buf, {type:'array', raw:true});
  const ws = wb.Sheets[wb.SheetNames[0]];
  if(!ws) throw new Error('file kosong');
  return XLSX.utils.sheet_to_json(ws, {header:1, raw:true, defval:''});
};

const tmOf = v => {
  if(typeof v==='number' && v>=0 && v<1){ const t=Math.round(v*86400); return pad(Math.floor(t/3600)%24)+':'+pad(Math.floor(t%3600/60))+':'+pad(t%60); }
  const m=/(\d{1,2}):(\d{2})(?::(\d{2}))?/.exec(String(v)); return m? pad(+m[1])+':'+m[2]+':'+(m[3]||'00') : '';
};
const itemNo10 = v => { const t=String(v==null?'':v).trim(); return /^\d+$/.test(t)? t.padStart(10,'0') : t; };
let mvGet = null;
function mvAdd(map, r){
  const ino=itemNo10(mvGet(r,'item no')); if(!ino) return;
  const np=String(mvGet(r,'np no')).trim(), q=num(mvGet(r,'qty')), tm=tmOf(mvGet(r,'time')), k=ino+'|'+np;
  const e=map[k]||(map[k]=[ino,clean(mvGet(r,'item name')),np,0,tm]); e[3]+=q; if(tm&&(!e[4]||tm<e[4])) e[4]=tm;
}
const mvOut = map => Object.values(map).map(e=>[e[0],e[1],e[2],Math.round(e[3]*100)/100,e[4]]);

function parseReport(XLSX, buf, storeCode){
  const rows = readRows(XLSX, buf);
  let hi = -1;
  for(let i=0;i<Math.min(rows.length,30);i++){ if(rows[i].some(c=>norm(c)==='receipt date')){ hi=i; break; } }
  if(hi<0) throw new Error('baris judul kolom tidak ditemukan (harus ada kolom "Receipt Date")');
  const col = {};
  rows[hi].forEach((c,i)=>{ const k=norm(c); if(k && !(k in col)) col[k]=i; });
  const ALIAS = {'qty':['quantity'],'amount include sales':['amount inculde sales','amount incl sales'],'salescode id':['salesperson id'],'salescode name':['salesperson name'],'dicount amount':['discount amount','discount'],'np no':['nota penjualan no','np no.','nota penjualan no.']};
  for(const k in ALIAS){ if(!(k in col)){ const f=ALIAS[k].find(x=>x in col); if(f) col[k]=col[f]; } }
  const need = ['receipt date','receipt no','qty','amount exclude sales','salescode id','salescode name'];
  const miss = need.filter(k=>!(k in col));
  if(miss.length) throw new Error('kolom tidak ada: '+miss.join(', '));
  const get = (r,k) => (k in col) ? r[col[k]] : '';
  mvGet = get;
  const days = {};
  let skippedRows = 0, skippedStore = 0, used = 0;
  for(let i=hi+1;i<rows.length;i++){
    const r = rows[i];
    const date = isoDate(get(r,'receipt date'));
    if(!date){ if(r.some(c=>String(c).trim()!=='')) skippedRows++; continue; }
    const store = String(get(r,'store no')).trim();
    if(storeCode && store && store!==String(storeCode)){ skippedStore++; continue; }
    const d = days[date] || (days[date]={tot:0,rec:new Set(),lines:0,unS:0,unL:0,unRec:new Set(),people:{},hours:{},prot:{},mb:{},mvx:{}});
    const amt = num(get(r,'amount exclude sales')), inc = num(get(r,'amount include sales')), rn = String(get(r,'receipt no')).trim();
    let sid = String(get(r,'salescode id')).trim(); if(/^(null|none|nan|n\/a|-)$/i.test(sid)) sid='';
    d.tot += amt; d.rec.add(rn); d.lines++; used++;
    { const mbk=String(get(r,'master brand')).trim().toUpperCase()||'LAINNYA'; const mq=d.mb[mbk]||(d.mb[mbk]={q:0,s:0}); mq.q+=num(get(r,'qty')); mq.s+=amt; }
    if(/^bca$/i.test(String(get(r,'master brand')).trim())){
      const pk=sid||'-'; const pr=d.prot[pk]||(d.prot[pk]={id:sid,name:'',q:0,s:0});
      pr.q+=num(get(r,'qty')); pr.s+=amt;
      if(!pr.name && sid){ const nm0=clean(get(r,'salescode name')); if(nm0 && !/^(null|none|nan)$/i.test(nm0)) pr.name=nm0; }
    }
    const hr = hourOf(get(r,'time')); if(hr!==null) d.hours[hr] = (d.hours[hr]||0)+amt;
    if(!sid){ d.unS += amt; d.unL++; d.unRec.add(rn); mvAdd(d.mvx, r); continue; }
    const p = d.people[sid] || (d.people[sid]={id:sid,name:'',sales:0,incl:0,qty:0,rec:new Set(),lines:0,disc:0,cats:{},brands:{},pb:0,pbc:{},it:{},mv:{}});
    if(!p.name){ const nm=clean(get(r,'salescode name')); if(nm && !/^(null|none|nan)$/i.test(nm)) p.name=nm; }
    p.sales += amt; p.incl += inc; p.qty += num(get(r,'qty'));
    p.rec.add(rn); p.lines++; p.disc += num(get(r,'dicount amount'));
    const cat = String(get(r,'category')).trim() || 'Lainnya';
    p.cats[cat] = (p.cats[cat]||0)+amt;
    const mb = String(get(r,'master brand')).trim().toUpperCase() || 'LAINNYA';
    p.brands[mb] = (p.brands[mb]||0)+amt;
    { const ino=String(get(r,'item no')).trim(), idesc=clean(get(r,'item name')), ik=ino||idesc; const ie=p.it[ik]||(p.it[ik]=[cat,mb,ino,idesc,0,0]); ie[4]+=num(get(r,'qty')); ie[5]+=amt; if(/^private brand$/i.test(String(get(r,'master type brand')).trim())) ie[6]=1; }
    mvAdd(p.mv, r);
    if(/^private brand$/i.test(String(get(r,'master type brand')).trim())){ p.pb += amt; p.pbc[cat]=(p.pbc[cat]||0)+amt; }
  }
  const out = {};
  for(const date in days){
    const d = days[date];
    out[date] = {
      basis:'excl2',
      meta:{total:Math.round(d.tot),receipts:d.rec.size,lines:d.lines,unSales:Math.round(d.unS),unLines:d.unL,unReceipts:d.unRec.size},
      mvx:mvOut(d.mvx),
      people:Object.values(d.people).map(p=>{
        const nm = p.name || p.id;
        const cats = {}; for(const c in p.cats) cats[c]=Math.round(p.cats[c]);
        const pbc = {}; for(const c in p.pbc){ const v=Math.round(p.pbc[c]); if(v) pbc[c]=v; }
        const brands = {}; for(const b in p.brands){ const v=Math.round(p.brands[b]); if(v) brands[b]=v; }
        return {id:p.id,name:nm,role:/KASIR/i.test(nm)?'Kasir':'Sales',sales:Math.round(p.sales),incl:Math.round(p.incl),qty:Math.round(p.qty),receipts:p.rec.size,lines:p.lines,disc:Math.round(p.disc),cats,brands,pb:Math.round(p.pb),pbc,it:Object.values(p.it).map(e=>{ const a=[e[0],e[1],e[2],e[3],Math.round(e[4]*100)/100,Math.round(e[5])]; if(e[6]) a.push(1); return a; }),mv:mvOut(p.mv)};
      }),
      mb:Object.fromEntries(Object.entries(d.mb).map(([k,v])=>[k,{q:Math.round(v.q),s:Math.round(v.s)}])),
      prot:Object.values(d.prot).map(x=>({id:x.id,name:x.name||x.id,q:Math.round(x.q),s:Math.round(x.s)})),
      hours:Object.keys(d.hours).map(h=>({h:+h,sales:Math.round(d.hours[h])})).sort((a,b)=>a.h-b.h)
    };
  }
  return {days:out, used, skippedRows, skippedStore};
}

function parseSchedule(XLSX, buf){
  const rows = readRows(XLSX, buf);
  let hi = -1;
  for(let i=0;i<Math.min(rows.length,30);i++){
    if(rows[i].some(c=>norm(c)==='sales id') && rows[i].some(c=>norm(c)==='target/day')){ hi=i; break; }
  }
  if(hi<0) throw new Error('bukan file jadwal (butuh kolom "SALES ID" dan "TARGET/DAY")');
  const col = {};
  rows[hi].forEach((c,i)=>{ const k=norm(c); if(k && !(k in col)) col[k]=i; });
  const dcols = [];
  rows[hi].forEach((c,i)=>{ if(i>col['target/day']){ const d=isoDate(c); if(d) dcols.push([i,d]); } });
  if(!dcols.length) throw new Error('kolom tanggal jadwal tidak ditemukan');
  const g = (r,k) => (k in col) ? String(r[col[k]]).trim() : '';
  const seen = {}, people = [];
  for(let i=hi+1;i<rows.length;i++){
    const r = rows[i];
    const id = g(r,'sales id'); if(!id || seen[id]) continue; seen[id]=1;
    people.push({id, name:clean(g(r,'nama')), brand:g(r,'master brand'), cat:g(r,'category'), zona:g(r,'zonasi'),
      tpd:Math.round(num(r[col['target/day']])), tpm:Math.round(num(r[col['target/month']]||0)),
      c:dcols.map(([ci])=>String(r[ci]).trim().toUpperCase())});
  }
  if(!people.length) throw new Error('tidak ada baris jadwal');
  return {basis:'sched1', dates:dcols.map(x=>x[1]), people};
}
/* ---- Jadwal per zona (diunggah TL zonasi): format ABSENSI ZONA ---- */
const ZKODE = ['S','P','M','OFF','CT','MT','CUTI'];
const ZBLN = ['JANUARI','FEBRUARI','MARET','APRIL','MEI','JUNI','JULI','AGUSTUS','SEPTEMBER','OKTOBER','NOVEMBER','DESEMBER'];
const colName = i => { let s=''; i++; while(i>0){ const r=(i-1)%26; s=String.fromCharCode(65+r)+s; i=Math.floor((i-1)/26); } return s; };
const normZona = z => String(z).replace(/\s+/g,' ').trim().toUpperCase();
function isZonaFile(XLSX, buf){
  try{ const rows=readRows(XLSX, buf); return !!(rows[0] && norm(rows[0][0])==='salesperson id' && norm(rows[0][3]||'')==='zona'); }
  catch(e){ return false; }
}
/* Membaca file jadwal satu zona. Bila format tidak sesuai, melempar Error dengan daftar alasan (satu per baris). master = dokumen jadwal tersimpan (opsional) untuk pemeriksaan silang. */
function parseZona(XLSX, buf, storeCode, master){
  const wb = XLSX.read(buf, {type:'array', raw:true});
  const sn = wb.SheetNames[0], ws = sn && wb.Sheets[sn];
  if(!ws) throw new Error('File kosong, tidak ada sheet yang bisa dibaca.');
  const rows = XLSX.utils.sheet_to_json(ws, {header:1, raw:true, defval:''});
  const err = [], warn = [], more = () => err.length>=12;
  const add = t => { if(err.length<12) err.push(t); else err.length===12 && err.push('…dan kesalahan lain. Perbaiki yang di atas lebih dulu, lalu unggah lagi.'); };
  // 1. bulan dan tahun dari nama sheet
  let y=0, m=0;
  const mm = /([A-Za-z]{3,})\s*[-_.\/ ]*\s*(\d{4})/.exec(String(sn));
  if(mm){ const mo = MON[mm[1].slice(0,3).toLowerCase()]; if(mo){ m=mo; y=+mm[2]; } }
  if(!m) add('Nama sheet "'+sn+'" harus berisi bulan dan tahun, contoh: NOVEMBER 2026.');
  const dim = m ? new Date(y,m,0).getDate() : 0;
  // 2. judul kolom (baris 1)
  const hd = rows[0] || [];
  const want = ['SALESPERSON ID','NAMA','BRAND','ZONA'];
  const hdOk = want.every((w,i)=>norm(hd[i]||'')===w.toLowerCase());
  if(!hdOk) add('Baris 1 harus berisi judul kolom A sampai D: SALESPERSON ID, NAMA, BRAND, ZONA. Ditemukan: '+hd.slice(0,4).map(x=>'"'+String(x).trim()+'"').join(', ')+'.');
  let lastCol = -1; hd.forEach((c,i)=>{ if(String(c).trim()!=='') lastCol=i; });
  if(hdOk && m){
    const ds = []; for(let i=4;i<=lastCol;i++) ds.push(hd[i]);
    const seq = ds.every((c,i)=>Number(String(c).trim())===i+1 && String(c).trim()!=='');
    if(!seq || ds.length!==dim){
      add('Kolom tanggal (mulai kolom E) harus 1 sampai '+dim+' berurutan, karena '+ZBLN[m-1].charAt(0)+ZBLN[m-1].slice(1).toLowerCase()+' '+y+' punya '+dim+' hari. '+(ds.length? 'Ditemukan '+ds.length+' kolom'+(seq?'':' dan urutannya tidak 1, 2, 3, …')+'.' : 'Tidak ditemukan kolom tanggal.'));
    }
  }
  if(err.length) throw Object.assign(new Error(err.join('\n')),{list:err});
  // 3. baris data
  const seen = new Map(), people = [], zonas = new Map();
  const idRe = /^S\d{5}-\d{10}$/;
  for(let i=1;i<rows.length;i++){
    const r = rows[i], rn = i+1;
    if(!r || r.slice(0,4+dim).every(c=>String(c).trim()==='')) continue;
    const id = String(r[0]).trim().toUpperCase(), nm = clean(String(r[1])), br = String(r[2]).trim(), zn = normZona(r[3]);
    let bad = false;
    if(!idRe.test(id)){ add('Baris '+rn+': SALESPERSON ID "'+String(r[0]).trim()+'" tidak valid. Format yang benar contohnya S10005-0000000123.'); bad=true; }
    else if(id.slice(1,6)!==String(storeCode)){ add('Baris '+rn+': SALESPERSON ID '+id+' bukan milik toko '+storeCode+'.'); bad=true; }
    if(!nm){ add('Baris '+rn+': kolom NAMA kosong.'); bad=true; }
    if(!/^ZONA \d+$/.test(zn)){ add('Baris '+rn+': kolom ZONA "'+String(r[3]).trim()+'" tidak valid. Tulis seperti ZONA 5.'); bad=true; }
    else { if(!zonas.has(zn)) zonas.set(zn, rn); }
    if(!bad){
      if(seen.has(id)){ const o=seen.get(id); add('Baris '+rn+': SALESPERSON ID '+id+' sama dengan baris '+o.rn+' ('+o.nm+' dan '+nm+'). Satu ID hanya boleh satu baris, periksa salah satu ID.'); bad=true; }
      else seen.set(id,{rn,nm});
    }
    const c = [];
    for(let k=0;k<dim;k++){
      const raw = String(r[4+k]).trim().toUpperCase();
      if(raw===''){ add('Sel '+colName(4+k)+rn+' ('+(nm||'baris '+rn)+', tanggal '+(k+1)+') kosong. Isi kode jadwal, gunakan OFF untuk hari libur.'); bad=true; c.push(''); continue; }
      if(!ZKODE.includes(raw)){ add('Sel '+colName(4+k)+rn+' ('+(nm||'baris '+rn)+', tanggal '+(k+1)+'): kode "'+String(r[4+k]).trim()+'" tidak dikenal. Kode yang boleh: '+ZKODE.join(', ')+'.'); bad=true; }
      c.push(raw);
    }
    if(!bad) people.push({id, name:nm, brand:br, zona:zn, c});
  }
  if(zonas.size>1) add('Satu file hanya boleh berisi satu zona. Ditemukan: '+[...zonas].map(([z,rn])=>z+' (mulai baris '+rn+')').join(', ')+'.');
  if(!zonas.size && !err.length) add('Tidak ada baris salesperson di file ini.');
  if(err.length) throw Object.assign(new Error(err.join('\n')),{list:err});
  const zona = [...zonas.keys()][0];
  // 4. pemeriksaan silang dengan daftar salesperson yang sudah ada
  const mp = new Map(((master&&master.people)||[]).map(p=>[p.id,p]));
  const diffZ = people.filter(p=>mp.has(p.id) && normZona(mp.get(p.id).zona||'')!=='' && normZona(mp.get(p.id).zona)!==zona);
  if(diffZ.length){
    diffZ.slice(0,8).forEach(p=>add(p.name+' ('+p.id+') terdaftar di '+normZona(mp.get(p.id).zona)+', bukan '+zona+'. Jika memang pindah zona, minta admin memperbarui zonasi di file jadwal & target.'));
    if(diffZ.length>8) add('…dan '+(diffZ.length-8)+' salesperson lain yang zonanya berbeda.');
    throw Object.assign(new Error(err.join('\n')),{list:err});
  }
  const nw = people.filter(p=>!mp.has(p.id));
  if(nw.length) warn.push(nw.length+' salesperson baru ('+nw.slice(0,5).map(p=>p.name).join(', ')+(nw.length>5?', …':'')+') belum punya target. Minta admin melengkapi lewat file jadwal & target.');
  const nmDiff = people.filter(p=>mp.has(p.id) && mp.get(p.id).name && clean(mp.get(p.id).name).toUpperCase()!==p.name.toUpperCase());
  if(nmDiff.length) warn.push(nmDiff.length+' nama berbeda dari data tersimpan ('+nmDiff.slice(0,4).map(p=>mp.get(p.id).name+' → '+p.name).join(', ')+(nmDiff.length>4?', …':'')+'). Nama di file ini dipakai.');
  const gone = ((master&&master.people)||[]).filter(p=>normZona(p.zona||'')===zona && !seen.has(p.id));
  if(gone.length) warn.push(gone.length+' orang di '+zona+' tidak ada di file ini ('+gone.slice(0,5).map(p=>p.name).join(', ')+(gone.length>5?', …':'')+'), jadwalnya tidak diubah.');
  const pre = y+'-'+pad(m)+'-';
  return {zona, y, m, month:y+'-'+pad(m), dates:[...Array(dim)].map((_,i)=>pre+pad(i+1)), people, warnings:warn, newCount:nw.length};
}
/* Menggabungkan jadwal satu zona ke dokumen jadwal: hanya kode jadwal (dan nama, brand) orang di file yang berubah, target tidak disentuh. */
function mergeZona(doc, z, source, now){
  const base = (doc && Array.isArray(doc.people)) ? doc : {basis:'sched1', dates:[], people:[]};
  const od = Array.isArray(base.dates) ? base.dates : [];
  const nd = [...new Set([...od, ...z.dates])].sort();
  const oi = new Map(od.map((d,i)=>[d,i])), zi = new Map(z.dates.map((d,i)=>[d,i]));
  const byId = new Map(z.people.map(p=>[p.id,p]));
  const people = base.people.map(e=>{
    const f = byId.get(e.id), oc = Array.isArray(e.c)?e.c:[];
    const c = nd.map(d=>{ if(f && zi.has(d)) return f.c[zi.get(d)]; return oi.has(d) ? (oc[oi.get(d)]||'') : ''; });
    return f ? Object.assign({}, e, {name:f.name, brand:f.brand||e.brand||'', zona:f.zona, c}) : Object.assign({}, e, {c});
  });
  const have = new Set(base.people.map(e=>e.id));
  z.people.forEach(f=>{ if(have.has(f.id)) return; people.push({id:f.id, name:f.name, brand:f.brand, cat:'', zona:f.zona, tpd:0, tpm:0, c:nd.map(d=>zi.has(d)?f.c[zi.get(d)]:'')}); });
  const zones = Object.assign({}, base.zones||{}, {[z.zona]:{month:z.month, source, uploadedAt:now, n:z.people.length}});
  return Object.assign({}, base, {basis:'sched1', dates:nd, people, zones});
}
function parseNav(XLSX, buf){
  const rows = readRows(XLSX, buf); let hi=-1;
  for(let i=0;i<Math.min(rows.length,30);i++){ const n=rows[i].map(norm); if((n.includes('np no.')||n.includes('np no'))&&n.includes('type')){ hi=i; break; } }
  if(hi<0) throw new Error('bukan file Nota Penjualan NAV (butuh kolom "NP No." dan "Type")');
  const h=rows[hi].map(norm), cn=h.findIndex(x=>x==='np no.'||x==='np no'), ct=h.indexOf('type'), cd=h.indexOf('order date');
  const seen={}, out=[];
  for(let i=hi+1;i<rows.length;i++){ const r=rows[i], np=String(r[cn]==null?'':r[cn]).trim(); if(!np) continue;
    seen[np]={np, type:String(r[ct]==null?'':r[ct]).replace(/\s+/g,' ').trim(), date:cd>=0?isoDate(r[cd]):''}; }
  Object.keys(seen).forEach(k=>out.push(seen[k]));
  if(!out.length) throw new Error('tidak ada baris nota pada file NAV');
  return {nav:out};
}
function parseMoves(XLSX, buf){
  const rows = readRows(XLSX, buf); let hi=-1;
  for(let i=0;i<Math.min(rows.length,30);i++){ const n=rows[i].map(norm); if(n.includes('item no')&&n.includes('moved qty')){ hi=i; break; } }
  if(hi<0) throw new Error('bukan report Stock Movement (butuh kolom "Item No" dan "Moved Qty")');
  const h=rows[hi].map(norm), ci=n=>h.indexOf(n), cdt=ci('date'), cit=ci('item no'), cds=ci('item description'), cu=ci('base uom'), cq=ci('moved qty');
  const moves={}; let used=0;
  for(let i=hi+1;i<rows.length;i++){ const r=rows[i], d=isoDate(r[cdt]), it=itemNo10(r[cit]); if(!d||!it) continue;
    const q=num(r[cq]); const m=moves[d]||(moves[d]={}); const e=m[it]||(m[it]={q:0,uom:'',desc:'',n:0});
    e.q+=q; e.n++; used++; if(!e.uom) e.uom=String(r[cu]==null?'':r[cu]).trim(); if(!e.desc) e.desc=clean(r[cds]==null?'':r[cds]); }
  if(!used) throw new Error('tidak ada baris data pada report Stock Movement');
  Object.keys(moves).forEach(d=>Object.keys(moves[d]).forEach(k=>{ moves[d][k].q=Math.round(moves[d][k].q*100)/100; }));
  return {moves, used};
}
if(typeof module!=='undefined') module.exports = {parseReport, parseSchedule, parseZona, mergeZona, isZonaFile, parseNav, parseMoves};

let P = [], M = {}, H = [], days = {}, dates = [];
const fmtN = new Intl.NumberFormat('id-ID');
const rp = v => 'Rp ' + fmtN.format(Math.round(v));
const jt = v => { const a=Math.abs(v); if(a>=1e9) return 'Rp '+(v/1e9).toFixed(2).replace('.',',')+' M'; if(a>=1e6) return 'Rp '+(v/1e6).toFixed(1).replace('.',',')+' jt'; if(a>=1e3) return 'Rp '+Math.round(v/1e3)+' rb'; return 'Rp '+Math.round(v); };
const pct = v => (v*100).toFixed(1).replace('.',',')+'%';
const title = s => s.toLowerCase().replace(/\b\w/g,c=>c.toUpperCase());
const $ = id => document.getElementById(id);
const esc = s => String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));

const fd = (d,o) => new Date(d+'T00:00:00').toLocaleDateString('id-ID',o||{day:'numeric',month:'long',year:'numeric'});
let SCH = {dates:[],people:[]}, SCHM = {}, curSel = [], lastL = [], lastHasT = false, lastSd = '', lastM = {}, lastDr = null;
const isWork = c => /^[SPM](?:[\/.].*)?$/.test(String(c||''));
// Salesperson yang sudah tidak di toko: tidak ditampilkan (tetap aman bila file jadwal di-upload ulang)
const RESIGNED = new Set(['S10005-0000000706','S10005-0000000517','S10005-0000000704']);
function setSched(s){
  SCH = s || {dates:[],people:[]}; SCHM = {};
  if(SCH.people && SCH.people.some(e=>RESIGNED.has(e.id))) SCH = Object.assign({},SCH,{people:SCH.people.filter(e=>!RESIGNED.has(e.id))});
  SCH.people.forEach(e=>{ const by={}; SCH.dates.forEach((d,i)=>{ by[d]=e.c[i]; }); const wdays=SCH.dates.filter((d,i)=>isWork(e.c[i])).length; SCHM[e.id]=Object.assign({},e,{by,wdays,tpe: e.tpd>0? e.tpd : (e.tpm>0? e.tpm/30 : 0)}); });
  const zs=[...new Set(SCH.people.map(e=>e.zona).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'id',{numeric:true}));
  const sel=$('zona'); if(S.zona && !zs.includes(S.zona)) S.zona='';
  sel.innerHTML='<option value="">Semua zonasi</option>'+zs.map(z=>'<option value="'+esc(z)+'"'+(z===S.zona?' selected':'')+'>'+esc(z)+'</option>').join('');
  $('zona-wrap').hidden = !zs.length;
}
function targetOf(e,sel){
  let target=0, wd=0, cov=0;
  sel.forEach(d=>{ const c=e.by[d]; if(!c) return; cov++; if(isWork(c)){ target+=e.tpe; wd++; } });
  return {target:cov?target:null, wd, cov, code: sel.length===1 ? (e.by[sel[0]]||'') : ''};
}
function applyTargets(sel){
  curSel=sel;
  P.forEach(p=>{
    const e=SCHM[p.id];
    if(e){ Object.assign(p, targetOf(e,sel), {zona:e.zona, brand:e.brand, hasSch:true, tpe:e.tpe, wdays:e.wdays}); }
    else { p.target=null; p.wd=0; p.cov=0; p.code=''; p.zona=''; p.hasSch=false; }
  });
  const have=new Set(P.map(p=>p.id));
  Object.values(SCHM).forEach(e=>{
    if(have.has(e.id)) return;
    const t=targetOf(e,sel);
    if(t.target>0) P.push(Object.assign({id:e.id,name:e.name||e.id,role:'Sales',sales:0,qty:0,receipts:0,lines:0,disc:0,cats:{},cat:'',avg:0,zona:e.zona,brand:e.brand,hasSch:true,tpe:e.tpe,wdays:e.wdays},t));
  });
  P.sort((a,b)=>b.sales-a.sales); P.forEach((p,i)=>{ p.rank=i+1; });
}
function aggregate(from,to){
  const sel = dates.filter(d=>d>=from&&d<=to);
  const m = {total:0,receipts:0,lines:0,unSales:0,unLines:0,unReceipts:0,prot:{},protMiss:0,mb:{},mbMiss:0,days:sel.length,from:sel[0],to:sel[sel.length-1]};
  const map = {}, hrs = {};
  sel.forEach(d=>{ const x=days[d];
    for(const k of ['total','receipts','lines','unSales','unLines','unReceipts']) m[k]+=x.meta[k];
    x.hours.forEach(h=>{hrs[h.h]=(hrs[h.h]||0)+h.sales;});
    if(!x.mb) m.mbMiss++; else for(const k in x.mb){ const a=m.mb[k]||(m.mb[k]={q:0,s:0}); a.q+=x.mb[k].q; a.s+=x.mb[k].s; }
    if(!Array.isArray(x.prot)) m.protMiss++;
    else x.prot.forEach(q=>{ const k=q.id||'-'; const a=m.prot[k]||(m.prot[k]={id:q.id,name:q.name,q:0,s:0}); a.q+=q.q; a.s+=q.s; if(q.name) a.name=q.name; });
    x.people.forEach(p=>{
      if(RESIGNED.has(p.id)) return;
      const a = map[p.id] || (map[p.id]={id:p.id,name:p.name,role:p.role,sales:0,excl:0,qty:0,receipts:0,lines:0,disc:0,cats:{},brands:{},pb:0});
      for(const k of ['sales','excl','qty','receipts','lines','disc','pb']) a[k]+=(p[k]||0);
      for(const b in (p.brands||{})) a.brands[b]=(a.brands[b]||0)+p.brands[b];
      a.name=p.name; for(const c in p.cats) a.cats[c]=(a.cats[c]||0)+p.cats[c];
      if(Array.isArray(p.it)){ const im=a.it||(a.it={}); p.it.forEach(e=>{ const k=e[2]||e[3]; const t=im[k]||(im[k]=[e[0],e[1],e[2],e[3],0,0]); t[4]+=e[4]; t[5]+=e[5]; }); } else a.itMiss=(a.itMiss||0)+1;
    });
  });
  P = Object.values(map).map(a=>{ a.cat=Object.entries(a.cats).sort((x,y)=>y[1]-x[1])[0][0]; return a; }).sort((a,b)=>b.sales-a.sales);
  P.forEach((p,i)=>{ p.rank=i+1; p.avg=p.receipts?p.sales/p.receipts:0; });
  applyTargets(sel);
  M = m;
  H = Object.keys(hrs).map(h=>({h:+h,sales:hrs[h]})).sort((a,b)=>a.h-b.h);
}
function renderSub(){
  const r = M.days? (M.from===M.to? fd(M.from,{weekday:'long',day:'numeric',month:'long',year:'numeric'}) : fd(M.from)+' sampai '+fd(M.to)+' ('+M.days+' hari)') : 'tidak ada data pada rentang ini';
  $('sub').textContent = D.store+' ('+D.code+') · '+r+' · nilai penjualan setelah diskon, tidak termasuk PPN';
}

const S = {dMode:'', dFrom:'', dTo:'', role:'all', zona:'', schDate:'', schSel:[], m:'sales', topn:15, q:'', sort:'sales', dir:-1};

const tip = $('tip');
function showTip(html,e){ tip.innerHTML=html; tip.style.opacity=1; const w=tip.offsetWidth,h=tip.offsetHeight; let x=e.clientX+14,y=e.clientY+14; if(x+w>innerWidth-8)x=e.clientX-w-14; if(y+h>innerHeight-8)y=e.clientY-h-14; tip.style.left=Math.max(8,x)+'px'; tip.style.top=Math.max(8,y)+'px'; }
function hideTip(){ tip.style.opacity=0; }
function tipHTML(p){
  let h='<b>'+esc(p.name)+'</b>Rank #'+p.rank+' dari '+P.length+'<br>Penjualan '+rp(p.sales)+' ('+pct(p.sales/(M.total||1))+' toko)<br>'+p.receipts+' struk · '+p.qty+' qty · '+p.lines+' baris item';
  if(p.lines) h+='<br>Basket size '+rp(p.avg);
  if(P.some(x=>x.disc>0)) h+='<br>Diskon '+rp(p.disc);
  if(p.cat) h+='<br>Kategori utama: '+esc(title(p.cat));
  if(p.hasSch){
    h+='<br>'+esc(p.zona||'-')+' · '+(p.cov? p.wd+' dari '+p.cov+' hari berjadwal kerja' : 'tidak ada jadwal pada periode ini');
    if(p.code) h+=' (jadwal: '+esc(p.code)+')';
    if(p.target>0) h+='<br>Target '+rp(p.target)+' ('+rp(p.tpe)+' × '+p.wd+' hari kerja) · pencapaian '+pct(p.sales/p.target)+' · selisih '+(p.sales-p.target>=0?'+':'-')+rp(Math.abs(p.sales-p.target));
    else if(p.target===0) h+='<br>Tidak ada target (tidak berjadwal kerja)';
  } else h+='<br>Tidak ada di file jadwal, tanpa target';
  return h;
}
const filtered = () => P.filter(p => (S.role==='all' || p.role===S.role));

function renderKPI(){
  const L = filtered(), A = L.filter(p=>p.lines>0), tot = A.reduce((a,p)=>a+p.sales,0);
  $('k-total').textContent = jt(M.total);
  $('k-total-d').textContent = rp(M.total)+' · '+(M.from&&M.to? (M.from===M.to? fd(M.from,{day:'numeric',month:'short'}) : fd(M.from,{day:'numeric',month:'short'})+' – '+fd(M.to,{day:'numeric',month:'short'}))+' · ' : '')+fmtN.format(M.receipts)+' struk';
  $('k-n').textContent = A.length;
  $('k-n-d').textContent = (S.role==='all'?'ada transaksi pada periode ini':'peran '+S.role)+(L.length>A.length?' · '+(L.length-A.length)+' terjadwal tanpa penjualan':'');
  $('k-avg').textContent = jt(tot/(A.length||1));
  const med = A.map(p=>p.sales).sort((a,b)=>a-b); const mid = med.length? (med.length%2? med[(med.length-1)/2] : (med[med.length/2-1]+med[med.length/2])/2):0;
  $('k-avg-d').textContent = 'median '+jt(mid);
  const top = L[0];
  $('k-top').textContent = top? top.name : '-';
  $('k-top-d').textContent = top? jt(top.sales)+' · '+top.lines+' baris item' : '';
  const T = L.filter(p=>p.target>0);
  if(T.length){
    const hit = T.filter(p=>p.sales>=p.target).length, st=T.reduce((a,p)=>a+p.sales,0), tt=T.reduce((a,p)=>a+p.target,0);
    $('k-t-l').textContent='Mencapai target';
    $('k-t').textContent = hit+' / '+T.length;
    $('k-t-d').textContent = 'Total '+pct(st/tt)+' dari target '+jt(tt);
    $('k-t-mw').hidden=false; $('k-t-mw').classList.toggle('full',hit===T.length); $('k-t-m').style.width=Math.round(hit/T.length*100)+'%';
  } else {
    const t10 = A.slice(0,10).reduce((a,p)=>a+p.sales,0);
    $('k-t-l').textContent='Konsentrasi 10 teratas'; $('k-t-mw').hidden=true;
    $('k-t').textContent = pct(t10/(tot||1));
    $('k-t-d').textContent = 'dari penjualan '+(S.role==='all'?'semua orang':S.role.toLowerCase());
  }
}

function renderRank(){
  const L = filtered().slice().sort((a,b)=>b[S.m]-a[S.m]).slice(0,S.topn);
  const key = S.m, mx = Math.max(1,...L.map(p=>p[key]));
  const useT = key==='sales' && L.some(p=>p.target>0);
  const dom = useT? Math.max(mx,...L.map(p=>(p.target||0)*1.05)) : mx;
  const lab = {sales:'penjualan',receipts:'jumlah struk',qty:'qty barang'}[key];
  $('rank-title').textContent = 'Peringkat '+lab;
  $('rank-hint').textContent = L.length+' orang ditampilkan'+(useT?' · garis putus-putus = target sesuai jadwal':'');
  $('rows').innerHTML = L.map(p=>{
    const w = Math.max(0,p[key]/dom*100);
    const v = key==='sales'? jt(p.sales) : fmtN.format(p[key]);
    const low = useT && p.target>0 && p.sales<p.target ? ' low':'';
    const tk = useT && p.target>0 ? '<i class="tick" style="left:'+(p.target/dom*100)+'%"></i>':'';
    return '<div class="row" tabindex="0" data-id="'+esc(p.id)+'"><div class="n" title="'+esc(p.name)+'"><b class="num">'+p.rank+'</b>'+esc(p.name)+'</div><div class="track"><div class="fill'+low+'" style="width:'+w+'%"></div>'+tk+'</div><div class="v num">'+v+'</div></div>';
  }).join('') || '<p class="note">Tidak ada data.</p>';
  $('legend').innerHTML = useT? '<span><i class="sw" style="background:var(--bar)"></i>Mencapai target atau tanpa target</span><span><i class="sw" style="background:var(--muted-bar)"></i>Belum mencapai</span>' : '';
  document.querySelectorAll('#rows .row').forEach(r=>{
    const p = P.find(x=>x.id===r.dataset.id);
    r.addEventListener('mousemove',e=>showTip(tipHTML(p),e));
    r.addEventListener('mouseleave',hideTip);
    r.addEventListener('focus',()=>{const b=r.getBoundingClientRect(); showTip(tipHTML(p),{clientX:b.left+40,clientY:b.bottom-10});});
    r.addEventListener('blur',hideTip);
  });
}

const ZLAB={'1':'Flooring and wall','2':'Bathroom & Plumbing','3':'App, Kitchen & Houseware, Electrical','4':'Building, Doors, Hardware','5':'Paint & sundries, Tools'};
let lstCtx=null;
const lstBar=()=>'<div style="display:flex;gap:8px;flex-wrap:wrap;padding:0 0 12px"><button type="button" class="btn primary sm" data-lst="copy" title="Salin gambar daftar ini, lalu tempel di WhatsApp">Print Screen</button><button type="button" class="btn sm" data-lst="save" title="Simpan sebagai file PNG">Unduh PNG</button></div>';
const lstRow=(p,i,lab,cls)=>[String(i+1),p.id,p.name,p.zona||'–',p.brand?title(p.brand):'–',fmtN.format(Math.round(p.target)),fmtN.format(Math.round(p.sales)),pct(p.sales/p.target),lab,cls];
const MB=[
  {h:'>100%',t:'≥100%',f:a=>a>=1},
  {h:'95% sd 100%',t:'95% sd <100%',f:a=>a>=.95&&a<1},
  {h:'85% sd 95%',t:'85% sd <95%',f:a=>a>=.85&&a<.95},
  {h:'75% sd 85%',t:'75% sd <85%',f:a=>a>=.75&&a<.85},
  {h:'50% sd 75%',t:'50% sd <75%',f:a=>a>=.5&&a<.75},
  {h:'<50%',t:'<50%',f:a=>a<.5}];
const MTD_STORE='Percetakan Negara-PCN';
const isPBp=p=>String(p.brand||'').toUpperCase()==='PRIVATE BRAND';
const mtdBand=p=>{ const a=p.sales/p.target; return MB.findIndex(b=>b.f(a)); };
const mtdPc=v=>(v*100).toFixed(1).replace('.',',')+'%';
function mtdAgg(l){ const n=l.length, c=MB.map(()=>0); l.forEach(p=>c[mtdBand(p)]++); return {n,c,pc:c.map(v=>n?v/n:0)}; }
function mtdSets(){ const mem=P.filter(p=>p.hasSch && p.target>0); return [['SALPER MTD  ALL STAFF EKSTERNAL & PB','all',mem],['SALPER MTD  ALL PRIVATE BRAND ( ONLY PB )','pb',mem.filter(isPBp)]]; }
function renderMtd(){
  const body=$('mtd-body'), sub=$('mtd-sub'), note=$('mtd-note');
  if(!SCH.people.length){ body.innerHTML='<p class="note">Unggah file jadwal &amp; target untuk menampilkan summary ini.</p>'; sub.textContent=''; note.hidden=true; return; }
  note.hidden=false;
  sub.textContent=(M.days? (M.days===1? 'Daily · '+fd(M.from,{day:'numeric',month:'short',year:'numeric'}) : fd(M.from,{day:'numeric',month:'short'})+' sampai '+fd(M.to,{day:'numeric',month:'short',year:'numeric'})+' · '+M.days+' hari') : 'tidak ada data')+' · achievement penjualan terhadap target';
  note.textContent='Hanya salesperson yang ada di jadwal dan punya target. Kelompok memakai batas bawah: ≥100%, 95–<100%, 85–<95%, 75–<85%, 50–<75%, <50%. Klik angka untuk melihat daftar salesperson. Cont = porsi terhadap total salper.';
  const sets=mtdSets().map(s=>[s[0],s[1],mtdAgg(s[2])]);
  const head2=sets.map((s,i)=>'<th class="tt'+(i+1)+'">Total<br>Salper</th>'+MB.map(b=>'<th>'+b.h.replace(' sd ','<br>sd<br>')+'</th><th class="ct">Cont</th>').join('')).join('<th class="sp"></th>');
  const cells=(a,k)=>'<td class="tt cl" data-s="'+k+'" data-b="all">'+a.n+'</td>'+MB.map((b,i)=>'<td class="cl" data-s="'+k+'" data-b="b'+i+'">'+a.c[i]+'</td><td class="c1">'+mtdPc(a.pc[i])+'</td>').join('');
  const line=(first,cls)=>'<tr class="'+cls+'">'+first+sets.map(s=>cells(s[2],s[1])).join('<td class="sp"></td>')+'</tr>';
  body.innerHTML='<table class="mtd"><thead><tr><th class="mn" rowspan="2">No</th><th class="ms" rowspan="2">Store Name</th>'+
    '<th class="g1" colspan="13">'+esc(sets[0][0])+'</th><th class="sp" rowspan="2"></th><th class="g2" colspan="13">'+esc(sets[1][0])+'</th></tr><tr>'+head2+'</tr></thead><tbody>'+
    line('<td>1</td><td class="l">'+esc(MTD_STORE)+'</td>','')+
    line('<td class="l" colspan="2">TOTAL</td>','tot')+'</tbody></table>';
}
function openMtd(set,bk,from){
  let l=P.filter(p=>p.hasSch&&p.target>0); if(set==='pb') l=l.filter(isPBp);
  const cnt=mtdAgg(l).c, bi=(bk&&bk!=='all')?+bk.slice(1):-1, cl=['gd','gd','am','am','bd','bd'];
  const f=(bi>=0? l.filter(p=>mtdBand(p)===bi) : l).slice().sort((a,b)=>b.sales/b.target-a.sales/a.target);
  const rows=f.map((p,i)=>{ const b=mtdBand(p); return kRow(p,['<td class="num mut">'+(i+1)+'</td>','<td>'+esc(p.id)+'</td>','<td>'+esc(p.name)+'</td>','<td>'+esc(p.zona||'–')+'</td>','<td>'+esc(p.brand?title(p.brand):'–')+'</td>',kNum(rp(p.target)),kNum(rp(p.sales)),'<td class="num rt '+cl[b]+'">'+pct(p.sales/p.target)+'</td>','<td class="'+cl[b]+'">'+MB[b].t+'</td>']); });
  const ttl='Daftar salesperson · Summary MTD · '+(set==='pb'?'Private Brand':'Semua staff')+(bi>=0?' · '+MB[bi].t:''), info=[['Total salper',String(l.length)]].concat(MB.map((b,i)=>[b.t,String(cnt[i])]));
  lstCtx={title:ttl,info,rows:f.map((p,i)=>{ const b=mtdBand(p); return lstRow(p,i,MB[b].t,cl[b]); }),name:'daftar-mtd-'+(set==='pb'?'pb':'semua')+(bi>=0?'-b'+bi:''),from};
  kShow(ttl,info,
    rows.length? lstBar()+kTbl([['#'],['Sales ID'],['Salesperson'],['Zonasi'],['Master brand'],['Target',1],['Penjualan',1],['Achievement',1],['Kelompok']],rows)+'<p class="note" style="padding:8px 4px">Klik salah satu baris untuk melihat detail item. Urut dari capaian tertinggi.</p>' : '<p class="note" style="padding:14px 4px">Tidak ada salesperson pada kelompok ini.</p>',from);
}
$('mtd-body').addEventListener('click',e=>{ const c=e.target.closest('td.cl'); if(c) openMtd(c.dataset.s,c.dataset.b,c); });
function mppSummary(list){
  const zs=[...new Set(SCH.people.map(e=>e.zona).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'id',{numeric:true}));
  const agg=l=>{ const n=l.length, s=l.reduce((a,p)=>a+p.sales,0), tg=l.reduce((a,p)=>a+p.target,0), rc=l.reduce((a,p)=>a+p.receipts,0), wd=l.reduce((a,p)=>a+(p.wd||0),0);
    const hi=l.filter(p=>p.sales/p.target>=1).length, lo=l.filter(p=>p.sales/p.target<.75).length, mid=n-hi-lo;
    return {n,hi,mid,lo,hiP:n?hi/n:0,midP:n?mid/n:0,loP:n?lo/n:0,ipd:wd?rc/wd:0,bz:rc?s/rc:0}; };
  const rows=zs.map(z=>{ const num=(z.match(/\d+/)||[''])[0]; const l=list.filter(p=>p.zona===z);
    const cats=[...new Set(SCH.people.filter(e=>e.zona===z).map(e=>title(e.cat||'')).filter(Boolean))];
    return Object.assign({z:num||z, zn:z, cat:ZLAB[num]||cats.join(', ')}, agg(l)); });
  return {rows, tot:agg(list)};
}
function mppTable(name,color,list,key){
  const {rows,tot}=mppSummary(list), pc=v=>Math.round(v*100)+'%', ip=v=>v?v.toFixed(2).replace('.',','):'0', bz=v=>v?fmtN.format(Math.round(v)):'0';
  const tr=r=>'<td class="num" data-b="all">'+r.n+'</td><td class="num '+(r.hi?'ok':'')+'" data-b="hi">'+r.hi+'</td><td class="num g" data-b="hi">'+pc(r.hiP)+'</td><td class="num" data-b="mid">'+r.mid+'</td><td class="num a" data-b="mid">'+pc(r.midP)+'</td><td class="num" data-b="lo">'+r.lo+'</td><td class="num r" data-b="lo">'+pc(r.loP)+'</td><td class="num">'+ip(r.ipd)+'</td><td class="num">'+bz(r.bz)+'</td>';
  return '<div class="mpp"><div class="mpph"><i style="background:'+color+'"></i>'+esc(name)+'</div><div class="scroll" style="max-height:none"><table class="mt"><thead><tr><th class="l">Zona</th><th class="l">Category</th><th>Actual MPP</th><th>Salper ≥100%</th><th>Cont</th><th>Salper 75–99%</th><th>Cont</th><th>Salper &lt;75%</th><th>Cont</th><th>IPD</th><th>BZ</th></tr></thead><tbody>'+
    rows.map(r=>'<tr class="mpp-r" tabindex="0" data-set="'+key+'" data-z="'+esc(r.zn)+'" data-lbl="'+esc('Zona '+r.z+' · '+r.cat)+'" title="Klik untuk melihat daftar salesperson"><td class="l">'+esc(r.z)+'</td><td class="l cat">'+esc(r.cat)+'</td>'+tr(r)+'</tr>').join('')+
    '<tr class="tot mpp-r" tabindex="0" data-set="'+key+'" data-z="" data-lbl="Semua zona" title="Klik untuk melihat daftar salesperson"><td class="l" colspan="2">Total</td>'+tr(tot)+'</tr></tbody></table></div></div>';
}
function renderMPP(){
  const body=$('mpp-body');
  if(!SCH.people.length){ body.innerHTML='<p class="note">Unggah file jadwal &amp; target untuk menampilkan summary ini.</p>'; $('mpp-sub').textContent=''; $('mpp-note').textContent=''; return; }
  const mem=P.filter(p=>p.hasSch && p.target>0);
  $('mpp-sub').textContent = (M.days? (M.days===1? 'Daily · '+fd(M.from,{day:'numeric',month:'short',year:'numeric'}) : fd(M.from,{day:'numeric',month:'short'})+' sampai '+fd(M.to,{day:'numeric',month:'short',year:'numeric'})+' · '+M.days+' hari') : 'tidak ada data')+' · berdasarkan achievement penjualan';
  body.innerHTML = mppTable('Summary all salesperson, all staff (PB & eksternal)','#f0a500',mem,'all')+mppTable('Summary all salesperson, Private Brand only','#2a9de0',mem.filter(p=>String(p.brand||'').toUpperCase()==='PRIVATE BRAND'),'pb');
  }

const COLS = [
  ['rank','#'],['id','Sales ID'],['name','Salesperson'],['zona','Zonasi'],['code','Jadwal'],['cat','Kategori utama'],['mbrand','Master brand'],['sales','Penjualan'],['brand','Penjualan master brand'],['nonbrand','Penjualan non master brand'],['target','Target'],['ach','Pencapaian'],['wd','Hari kerja'],['share','% total'],['receipts','Struk'],['avg','Basket Size']
];
const TXT = new Set(['id','name','cat','zona','code','mbrand']);
const CODE_ORDER = ['S','P','M','P/JKP','P.JKP','S/JKS','S.JKS','OFF','CT','L','MT','OT','OUTING'];
const codeOn = (id,d) => { const e=SCHM[id]; return e? (e.by[d]||'') : ''; };
function schDate(){ const d = S.schDate || M.to; return SCH.dates.includes(d)? d : (SCH.dates[0]||''); }
function schMatch(code){ if(!S.schSel.length) return true; if(!code) return false; return S.schSel.some(v=>v==='WORK'? isWork(code) : code===v); }
const schLabel=()=>S.schSel.map(v=>v==='WORK'?'masuk kerja':v).join(' + ');
function brandSales(p){ if(!p.brand) return -1; const k=String(p.brand).toUpperCase(); return k==='PRIVATE BRAND' ? (p.pb||0) : ((p.brands&&p.brands[k])||0); }
function nonBrand(p){ const b=brandSales(p); return b<0? -1 : Math.max(0,p.sales-b); }
function brandTip(p){ const e=Object.entries(p.brands||{}).sort((a,b)=>b[1]-a[1]); if(!e.length) return 'Belum ada rincian master brand'; return 'Master brand yang dijual: '+e.slice(0,8).map(x=>title(x[0])+' '+rp(x[1])).join(' · ')+(e.length>8?' · +'+(e.length-8)+' lainnya':''); }
const shortD = d => d? fd(d,{day:'numeric',month:'short'}) : '';
function ghost(e){
  return Object.assign({id:e.id,name:e.name||e.id,role:'Sales',sales:0,qty:0,receipts:0,lines:0,disc:0,cats:{},cat:'',avg:0,zona:e.zona,brand:e.brand,hasSch:true,rank:null,tpe:e.tpe,wdays:e.wdays}, targetOf(e,curSel));
}
function renderSchFilter(){
  const box=$('schf');
  if(!SCH.people.length){ box.hidden=true; return; }
  box.hidden=false;
  const d=schDate(), di=$('sch-date');
  di.min=SCH.dates[0]; di.max=SCH.dates[SCH.dates.length-1]; di.value=d;
  $('sch-follow').hidden=!S.schDate;
  const base=SCH.people.filter(e=>!S.zona||e.zona===S.zona);
  const cnt={}; base.forEach(e=>{ const i=SCH.dates.indexOf(d); const c=i>=0? e.c[i] : ''; if(c) cnt[c]=(cnt[c]||0)+1; });
  const codes=Object.keys(cnt).sort((a,b)=>{ const x=CODE_ORDER.indexOf(a), y=CODE_ORDER.indexOf(b); return (x<0?99:x)-(y<0?99:y) || a.localeCompare(b); });
  const work=codes.filter(isWork).reduce((a,c)=>a+cnt[c],0);
  S.schSel=S.schSel.filter(v=>v==='WORK'||cnt[v]);
  const chip=(v,l,n)=>'<button type="button" class="chip2" data-sc="'+esc(v)+'" aria-pressed="'+(v===''?!S.schSel.length:S.schSel.includes(v))+'">'+esc(l)+' <b class="num">'+n+'</b></button>';
  $('sch-chips').innerHTML = chip('','Semua',base.length)+chip('WORK','Masuk kerja',work)+codes.map(c=>chip(c,c,cnt[c])).join('');
}
function detailRange(){
  if(!dates.length || !S.dMode) return null;
  const first=dates[0], last=dates[dates.length-1], cl=d=>d<first?first:d>last?last:d;
  if(S.dMode==='mtd'){ const to=cl(dT.value||last); const ms=to.slice(0,8)+'01'; return {from:ms<first?first:ms, to}; }
  let a=cl(S.dFrom||first), b=cl(S.dTo||last); if(a>b){ const x=a; a=b; b=x; } return {from:a,to:b};
}
function renderTable(){
  const dr=detailRange(); let saved=null;
  if(dr){ saved={P,M,H,curSel}; aggregate(dr.from,dr.to); }
  try{ renderTableInner(dr); }
  finally{ if(saved){ P=saved.P; M=saved.M; H=saved.H; curSel=saved.curSel; } }
}
function syncRange(dr){
  const a=dr?dr.from:dF.value, b=dr?dr.to:dT.value;
  [$('t-from'),$('t-to')].forEach(el=>{ el.min=dates[0]||''; el.max=dates[dates.length-1]||''; });
  $('t-from').value=a||''; $('t-to').value=b||'';
  $('t-mtd').setAttribute('aria-pressed',String(S.dMode==='mtd'));
  $('t-follow').setAttribute('aria-pressed',String(!S.dMode));
}
function renderTableInner(dr){
  lastM=M; lastDr=dr; syncRange(dr);
  const hasT = P.some(p=>p.hasSch);
  renderSchFilter();
  const cols = COLS.filter(c=>hasT || !['zona','code','wd','mbrand','brand','nonbrand','target','ach'].includes(c[0]));
  const sd = schDate();
  $('thead').innerHTML = cols.map(c=>{
    const act = S.sort===c[0];
    const label = c[0]==='code' ? 'Jadwal '+shortD(sd) : c[1];
    return '<th scope="col" class="'+(TXT.has(c[0])?'l':'')+'"'+(act?' aria-sort="'+(S.dir<0?'descending':'ascending')+'"':'')+'><button type="button" data-s="'+c[0]+'">'+label+(act?(S.dir<0?' ↓':' ↑'):'')+'</button></th>';
  }).join('');
  const q = S.q.trim().toLowerCase();
  // Tabel ini hanya menampilkan salesperson aktif: ada di jadwal, bukan kasir (kecuali peran Kasir dipilih)
  let base = filtered().filter(p=>(!S.zona||p.zona===S.zona) && (S.role==='Kasir' || !hasT || (p.hasSch && p.zona && p.role!=='Kasir')));
  if(S.schSel.length){
    const have=new Set(base.map(p=>p.id));
    Object.values(SCHM).forEach(e=>{ if(!have.has(e.id) && (S.role!=='Kasir') && (!S.zona||e.zona===S.zona) && !P.some(p=>p.id===e.id)) base.push(ghost(e)); });
    base = base.filter(p=>schMatch(codeOn(p.id,sd)));
  }
  const ratio = p => p.target>0 ? p.sales/p.target : -1;
  const val = (p,k)=> k==='share'?p.sales : k==='brand'?brandSales(p) : k==='nonbrand'?nonBrand(p) : k==='mbrand'?String(p.brand||'').toLowerCase() : k==='id'?String(p.id||'').toLowerCase() : k==='ach'?ratio(p) : k==='target'?(p.target==null?-1:p.target) : k==='code'?codeOn(p.id,sd) : k==='rank'?(p.rank==null?1e9:p.rank) : TXT.has(k)?String(p[k]||'').toLowerCase() : p[k];
  const L = base.filter(p=>!q||p.name.toLowerCase().includes(q)).sort((a,b)=>{
    const x=val(a,S.sort), y=val(b,S.sort); return (x<y?-1:x>y?1:0)*S.dir; });
  const mx = Math.max(1,...P.map(p=>p.sales));
  lastL=L; lastHasT=hasT; lastSd=sd;
  rowMap=Object.fromEntries(L.map(p=>[p.id,p])); popPer=M.days? (M.from===M.to? shortD(M.from) : shortD(M.from)+' – '+shortD(M.to)) : '';
  $('tbody').innerHTML = L.map(p=>{
    const a=ratio(p), cd=codeOn(p.id,sd);
    let t='<tr data-id="'+esc(p.id)+'" tabindex="0"><td class="rk num">'+(p.rank==null?'–':p.rank)+'</td><td class="l" style="font-size:12px;color:var(--ink2);white-space:nowrap">'+esc(p.id)+'</td><td class="l">'+esc(p.name)+(p.role==='Kasir'?'<span class="chip">Kasir</span>':'')+'</td>'+(hasT?'<td class="l">'+(p.zona?esc(p.zona):'–')+'</td><td class="l">'+(cd?esc(cd):'–')+'</td>':'')+'<td class="l">'+(p.cat?esc(title(p.cat)):'–')+'</td>'+(hasT?'<td class="l">'+(p.brand?esc(title(p.brand)):'–')+'</td>':'')+
      '<td class="num"><div class="cell"><span class="mini"><i style="width:'+Math.max(0,p.sales/mx*100)+'%"></i></span><span style="min-width:100px;display:inline-block;text-align:right">'+rp(p.sales)+'</span></div></td>'+(hasT?'<td class="num" title="'+esc(brandTip(p))+'">'+(p.brand?rp(brandSales(p))+'<span class="bn">'+esc(title(p.brand))+'</span>':'–')+'</td><td class="num">'+(p.brand?rp(nonBrand(p)):'–')+'</td>':'');
    if(hasT){
      t+='<td class="num">'+(p.target==null?'–':p.target===0?'<span style="color:var(--ink3)">tidak ada</span>':rp(p.target))+'</td>'+
        '<td class="num">'+(p.target>0? '<span class="pill '+(a>=1?'hi':a>=.75?'mid':'lo')+'">'+(a>=1?'✓ ':'')+Math.round(a*100)+'%</span>' : '–')+'</td>'+
        '<td class="num">'+(p.hasSch&&p.cov? p.wd : '–')+'</td>';
    }
    t+='<td class="num">'+pct(p.sales/(M.total||1))+'</td><td class="num">'+p.receipts+'</td><td class="num">'+(p.lines?rp(p.avg):'–')+'</td>';
    return t+'</tr>';
  }).join('') || '<tr><td colspan="16" style="text-align:center;color:var(--ink3)">Tidak ada orang yang cocok dengan filter.</td></tr>';
  $('tbl-hint').textContent = L.length+' orang'+(dr? ' · '+(S.dMode==='mtd'?'MTD ':'')+shortD(dr.from)+(dr.to!==dr.from?' sampai '+fd(dr.to,{day:'numeric',month:'short',year:'numeric'}):'')+' ('+M.days+' hari)' : '')+(S.schSel.length? ' · jadwal '+schLabel()+' pada '+shortD(sd) : '')+' · klik judul kolom untuk mengurutkan';
  $('thead').querySelectorAll('button').forEach(bt=>bt.addEventListener('click',()=>{
    const k=bt.dataset.s; if(S.sort===k) S.dir*=-1; else {S.sort=k; S.dir=TXT.has(k)?1:-1;} renderTable(); }));
}
$('sch-chips').addEventListener('click',e=>{
  const b=e.target.closest('[data-sc]'); if(!b) return;
  const v=b.dataset.sc;
  if(v==='') S.schSel=[]; else S.schSel = S.schSel.includes(v)? S.schSel.filter(x=>x!==v) : S.schSel.concat(v);
  renderTable();
  const nb=document.querySelector('#sch-chips [data-sc="'+CSS.escape(v)+'"]'); if(nb) nb.focus();
});
$('t-from').addEventListener('change',e=>{ if(!e.target.value) return; if(S.dMode!=='custom'){ S.dTo=$('t-to').value; } S.dMode='custom'; S.dFrom=e.target.value; renderTable(); });
$('t-to').addEventListener('change',e=>{ if(!e.target.value) return; if(S.dMode!=='custom'){ S.dFrom=$('t-from').value; } S.dMode='custom'; S.dTo=e.target.value; renderTable(); });
$('t-mtd').addEventListener('click',()=>{ S.dMode = S.dMode==='mtd' ? '' : 'mtd'; renderTable(); });
$('t-follow').addEventListener('click',()=>{ S.dMode=''; renderTable(); });
$('sch-date').addEventListener('change',e=>{ if(e.target.value) S.schDate=e.target.value; renderTable(); });
$('sch-follow').addEventListener('click',()=>{ S.schDate=''; renderTable(); });

function renderFoot(){ return; // catatan data disembunyikan
  const noSch = P.filter(p=>p.lines>0 && !p.hasSch).length;
  $('foot').innerHTML = '<b>Catatan data.</b> '+fmtN.format(M.unLines)+' baris item ('+fmtN.format(M.unReceipts)+' struk, '+rp(M.unSales)+', '+pct(M.unSales/(M.total||1))+' dari total) tidak memiliki kode sales, jadi tidak masuk peringkat tetapi tetap dihitung di total toko. Nilai penjualan memakai kolom Amount Exclude Sales. Penjualan master brand adalah jumlah penjualan barang (rincian item) yang dijual salesperson itu dan masuk master brand yang tertulis di file jadwal; untuk Private Brand dipakai penanda Master Type Brand = Private Brand. Satu orang bisa muncul di beberapa struk yang sama dengan orang lain, sehingga jumlah struk per orang tidak dijumlahkan ke total. '+
   '<b>Target.</b> Target harian diambil dari kolom TARGET/DAY di file jadwal, dicocokkan lewat Sales ID. Target MTD atau target satu rentang tanggal adalah target harian dikali jumlah hari kerja yang sudah berjalan dalam rentang itu. Hari kerja adalah tanggal berkode S, P, atau M (termasuk P/JKP dan S/JKS). Kode OFF, CT, L, MT, OT, dan OUTING tidak punya target. Hari setelah tanggal akhir rentang belum dihitung. Penjualan di hari tanpa target tetap dihitung sebagai penjualan.'+(noSch?' '+noSch+' orang bertransaksi tetapi tidak ada di file jadwal, jadi ditampilkan tanpa target.':'');
}

let rowMap={}, popPer='', popLast=null;
function closePop(){ const o=$('pop'); if(o.hidden) return; o.hidden=true; document.documentElement.style.overflow=''; if(popLast&&popLast.focus) popLast.focus(); }
function openPerson(p, from){
  popLast=from||null;
  $('pop-id').textContent=p.id; $('pop-name').textContent=p.name; $('pop-name2').textContent=p.name; $('pop-zona').textContent=p.zona||'–'; $('pop-val').textContent=rp(p.sales);
  $('pop-per').textContent=(popPer?popPer+' · ':'')+'sales exclude PPN';
  const items=Object.values(p.it||{}).sort((a,b)=>b[5]-a[5]);
  const tq=items.reduce((s,e)=>s+e[4],0), ts=items.reduce((s,e)=>s+e[5],0);
  const body=$('pop-body');
  if(!items.length) body.innerHTML='<p class="note" style="padding:14px 4px">'+(p.sales>0?'Detail item belum tersimpan untuk orang ini.':'Belum ada penjualan pada rentang ini.')+'</p>';
  else body.innerHTML='<table class="pt"><thead><tr><th>Category</th><th>Master brand</th><th>No Item</th><th>Deskripsi</th><th class="rt">Qty</th><th class="rt">Total value</th></tr></thead><tbody>'+items.map(e=>'<tr><td>'+esc(title(e[0]))+'</td><td>'+esc(title(e[1]))+'</td><td class="num">'+esc(e[2])+'</td><td>'+esc(e[3])+'</td><td class="num rt">'+fmtN.format(e[4])+'</td><td class="num rt">'+rp(e[5])+'</td></tr>').join('')+'</tbody><tfoot><tr><td colspan="4">Total · '+items.length+' item</td><td class="num rt">'+fmtN.format(tq)+'</td><td class="num rt">'+rp(ts)+'</td></tr></tfoot></table>';
  const n=$('pop-note'); n.hidden=!p.itMiss; n.textContent=p.itMiss? p.itMiss+' tanggal belum memuat detail item. Unggah ulang laporan tanggal tersebut agar daftar item lengkap.':'';
  const o=$('pop'); o.hidden=false; document.documentElement.style.overflow='hidden'; $('pop-x').focus(); $('pop-body').scrollTop=0;
}
$('pop-x').addEventListener('click',closePop);
$('pop').addEventListener('click',e=>{ if(e.target===$('pop')) closePop(); });
document.addEventListener('keydown',e=>{ if(e.key==='Escape') closePop(); });
$('tbody').addEventListener('click',e=>{ const tr=e.target.closest('tr[data-id]'); if(!tr||e.target.closest('button,a')) return; const p=rowMap[tr.dataset.id]; if(p) openPerson(p,tr); });
$('tbody').addEventListener('keydown',e=>{ if(e.key!=='Enter') return; const tr=e.target.closest('tr[data-id]'); if(!tr) return; const p=rowMap[tr.dataset.id]; if(p) openPerson(p,tr); });
/* ---- Kartu ringkasan interaktif: klik untuk rincian ---- */
let kLast=null;
function closeK(){ const o=$('kpop'); if(!o||o.hidden) return; o.hidden=true; document.documentElement.style.overflow=''; if(kLast&&kLast.focus) kLast.focus(); }
function kShow(ttl,info,html,from){
  kLast=from||null; $('kpop-t').textContent=ttl;
  $('kpop-per').textContent=M.days? (M.from===M.to? fd(M.from,{day:'numeric',month:'long',year:'numeric'}) : fd(M.from,{day:'numeric',month:'short'})+' sampai '+fd(M.to,{day:'numeric',month:'short',year:'numeric'})+' ('+M.days+' hari)')+' · sales exclude PPN' : 'Tidak ada data pada rentang ini';
  $('kpop-info').innerHTML=info.map(i=>'<div><span>'+esc(i[0])+'</span><b>'+esc(i[1])+'</b></div>').join('');
  $('kpop-body').innerHTML=html; $('kpop').hidden=false; document.documentElement.style.overflow='hidden'; $('kpop-x').focus(); $('kpop-body').scrollTop=0;
}
const kRow=(p,cells,cls)=>'<tr data-pid="'+esc(p.id)+'" tabindex="0"'+(cls?' class="'+cls+'"':'')+'>'+cells.join('')+'</tr>';
const kTbl=(head,rows)=>'<table class="pt"><thead><tr>'+head.map(h=>'<th'+(h[1]?' class="rt"':'')+'>'+h[0]+'</th>').join('')+'</tr></thead><tbody>'+rows.join('')+'</tbody></table>';
const kNum=v=>'<td class="num rt">'+v+'</td>';
function openK(k, el){
  const L=filtered(), A=L.filter(p=>p.lines>0), tot=A.reduce((a,p)=>a+p.sales,0);
  if(k==='total'){
    const sel=dates.filter(d=>d>=M.from&&d<=M.to), mx=Math.max(1,...sel.map(d=>days[d].meta.total));
    const rows=sel.map(d=>{ const x=days[d], t=x.meta.total, r=x.meta.receipts;
      return '<tr><td>'+esc(fd(d,{weekday:'short',day:'numeric',month:'short'}))+'</td>'+kNum(rp(t))+kNum(fmtN.format(r))+kNum(rp(r?t/r:0))+kNum(x.people.length)+'<td style="width:140px"><span class="bar" style="width:'+Math.round(t/mx*100)+'%"></span></td></tr>'; });
    kShow('Total penjualan per tanggal',[['Total penjualan',rp(M.total)],['Struk',fmtN.format(M.receipts)],['Basket size',rp(M.receipts?M.total/M.receipts:0)],['Rata-rata per hari',rp(M.days?M.total/M.days:0)]],
      sel.length? kTbl([['Tanggal'],['Penjualan',1],['Struk',1],['Basket size',1],['Orang aktif',1],['']],rows) : '<p class="note" style="padding:14px 4px">Belum ada data pada rentang ini.</p>',el);
  } else if(k==='aktif'){
    const ia=L.filter(p=>!(p.lines>0));
    const rows=A.map((p,i)=>kRow(p,['<td class="num mut">'+(i+1)+'</td>','<td>'+esc(p.id)+'</td>','<td>'+esc(p.name)+'</td>','<td>'+esc(p.zona||'–')+'</td>',kNum(rp(p.sales)),kNum(fmtN.format(p.receipts))]))
      .concat(ia.map(p=>kRow(p,['<td class="num mut">–</td>','<td>'+esc(p.id)+'</td>','<td>'+esc(p.name)+'</td>','<td>'+esc(p.zona||'–')+'</td>',kNum('belum ada penjualan'),kNum('–')],'dim')));
    kShow('Orang aktif',[['Aktif (ada transaksi)',String(A.length)],['Terjadwal tanpa penjualan',String(ia.length)],['Total dalam daftar',String(L.length)]],
      rows.length? kTbl([['#'],['Sales ID'],['Salesperson'],['Zonasi'],['Penjualan',1],['Struk',1]],rows)+'<p class="note" style="padding:8px 4px">Klik salah satu baris untuk melihat detail item.</p>' : '<p class="note" style="padding:14px 4px">Belum ada orang pada filter ini.</p>',el);
  } else if(k==='avg'){
    const v=A.map(p=>p.sales).sort((a,b)=>a-b), n=v.length, avg=n?tot/n:0, med=n?(n%2?v[(n-1)/2]:(v[n/2-1]+v[n/2])/2):0;
    const rows=A.map((p,i)=>{ const df=p.sales-avg, c=df>=0?'gd':'bd';
      return kRow(p,['<td class="num mut">'+(i+1)+'</td>','<td>'+esc(p.name)+'</td>','<td>'+esc(p.zona||'–')+'</td>',kNum(rp(p.sales)),'<td class="num rt '+c+'">'+(df>=0?'+':'−')+rp(Math.abs(df))+' ('+(df>=0?'+':'−')+Math.round(Math.abs(df)/(avg||1)*100)+'%)</td>']); });
    kShow('Rata-rata penjualan per orang',[['Rata-rata',rp(avg)],['Median',rp(med)],['Tertinggi',n?rp(v[n-1]):'–'],['Terendah',n?rp(v[0]):'–']],
      rows.length? kTbl([['#'],['Salesperson'],['Zonasi'],['Penjualan',1],['Selisih dari rata-rata',1]],rows) : '<p class="note" style="padding:14px 4px">Belum ada penjualan pada rentang ini.</p>',el);
  } else if(k==='top'){
    if(L[0]) openPerson(L[0], el);
  } else if(k==='target'){
    const T=L.filter(p=>p.target>0);
    if(T.length){
      const hit=T.filter(p=>p.sales>=p.target), st=T.reduce((a,p)=>a+p.sales,0), tt=T.reduce((a,p)=>a+p.target,0), gap=T.reduce((a,p)=>a+Math.max(0,p.target-p.sales),0);
      const rows=T.slice().sort((a,b)=>b.sales/b.target-a.sales/a.target).map(p=>{ const a=p.sales/p.target, ok=a>=1;
        return kRow(p,['<td>'+esc(p.name)+'</td>','<td>'+esc(p.zona||'–')+'</td>',kNum(rp(p.target)),kNum(rp(p.sales)),'<td class="num rt '+(ok?'gd':'')+'">'+Math.round(a*100)+'%</td>','<td class="'+(ok?'gd':'bd')+'">'+(ok?'✓ Tercapai':'Kurang '+rp(p.target-p.sales))+'</td>']); });
      kShow('Pencapaian target',[['Tercapai',hit.length+' / '+T.length],['Achievement total',pct(st/tt)],['Belum tercapai',String(T.length-hit.length)],['Total kekurangan',rp(gap)]],
        kTbl([['Salesperson'],['Zonasi'],['Target',1],['Penjualan',1],['Achievement',1],['Status']],rows)+'<p class="note" style="padding:8px 4px">Klik salah satu baris untuk melihat detail item.</p>',el);
    } else {
      const t10=A.slice(0,10), s10=t10.reduce((a,p)=>a+p.sales,0);
      kShow('Konsentrasi 10 teratas',[['10 teratas',rp(s10)],['Porsi dari total',pct(s10/(tot||1))]],
        t10.length? kTbl([['#'],['Salesperson'],['Penjualan',1],['% total',1]],t10.map((p,i)=>kRow(p,['<td class="num mut">'+(i+1)+'</td>','<td>'+esc(p.name)+'</td>',kNum(rp(p.sales)),kNum(pct(p.sales/(tot||1)))]))) : '<p class="note" style="padding:14px 4px">Belum ada penjualan.</p>',el);
    }
  }
}
['total','aktif','avg','top','target'].forEach((k,i)=>{ const el=document.querySelectorAll('.kpis .kpi')[i]; if(!el) return; el.dataset.k=k; el.tabIndex=0; el.setAttribute('role','button'); el.title='Klik untuk melihat rincian'; });
document.querySelector('.kpis').addEventListener('click',e=>{ const c=e.target.closest('.kpi[data-k]'); if(c) openK(c.dataset.k,c); });
document.querySelector('.kpis').addEventListener('keydown',e=>{ if(e.key!=='Enter'&&e.key!==' ') return; const c=e.target.closest('.kpi[data-k]'); if(!c) return; e.preventDefault(); openK(c.dataset.k,c); });
$('kpop-x').addEventListener('click',closeK);
$('kpop').addEventListener('click',e=>{ if(e.target===$('kpop')){ closeK(); return; } const tr=e.target.closest('tr[data-pid]'); if(!tr) return; const p=P.find(x=>x.id===tr.dataset.pid); if(p){ const f=kLast; closeK(); openPerson(p,f); } });
$('kpop').addEventListener('keydown',e=>{ if(e.key!=='Enter') return; const tr=e.target.closest('tr[data-pid]'); if(!tr) return; const p=P.find(x=>x.id===tr.dataset.pid); if(p){ const f=kLast; closeK(); openPerson(p,f); } });
document.addEventListener('keydown',e=>{ if(e.key==='Escape') closeK(); });
/* ---- Salesperson dengan pencapaian MTD di bawah 60% + schedule sampai akhir bulan ---- */
const LOWT=0.6; let lowInfo=null;
function lowCalc(){
  if(!dates.length) return null;
  const last=dates[dates.length-1], to=(dT.value&&dT.value<=last)?dT.value:last, ms=to.slice(0,8)+'01', from=ms<dates[0]?dates[0]:ms;
  const saved={P,M,H,curSel}; let rows=[], tot=0;
  try{
    aggregate(from,to);
    const T=P.filter(p=>p.hasSch&&p.target>0); tot=T.length;
    rows=T.filter(p=>p.sales/p.target<LOWT).map(p=>{ const codes=(SCH.dates||[]).map(d=>codeOn(p.id,d)); return {id:p.id,name:p.name,zona:p.zona||'',brand:p.brand||'',cat:(SCHM[p.id]&&SCHM[p.id].cat)||p.cat||'',tpd:p.tpe||0,target:p.target,sales:p.sales,pct:p.sales/p.target,codes,full:codes.reduce((a,c)=>a+(isWork(c)?(p.tpe||0):0),0)}; });
  } finally { P=saved.P; M=saved.M; H=saved.H; curSel=saved.curSel; }
  rows.sort((a,b)=>a.pct-b.pct);
  return {rows,tot,from,to};
}
function renderLow(){
  const el=$('low-wrap'); if(!el) return;
  const r=lowInfo=lowCalc(), sd=(SCH&&SCH.dates)||[];
  $('low-xlsx').disabled=$('low-copy').disabled=!(r&&r.rows.length);
  if(!r){ $('low-sub').textContent=''; el.innerHTML='<p class="note" style="padding:14px 4px">Belum ada data.</p>'; $('low-note').textContent=''; return; }
  const multi=sd.length&&sd[0].slice(0,7)!==sd[sd.length-1].slice(0,7), dl=d=>multi? (+d.slice(8))+'/'+(+d.slice(5,7)) : String(+d.slice(8));
  $('low-sub').textContent='MTD '+fd(r.from,{day:'numeric',month:'short'})+' sampai '+fd(r.to,{day:'numeric',month:'short',year:'numeric'})+' · '+r.rows.length+' dari '+r.tot+' orang bertarget · jadwal sampai '+(sd.length? fd(sd[sd.length-1],{day:'numeric',month:'long'}) : '–');
  if(!r.rows.length){ el.innerHTML='<p class="note" style="padding:14px 4px">Tidak ada salesperson di bawah 60% pada rentang ini.</p>'; $('low-note').textContent=''; return; }
  const tl=sd.length? fd(sd[sd.length-1],{day:'numeric',month:'short'}) : 'akhir jadwal';
  const head='<tr><th>#</th><th>Sales ID</th><th>Salesperson</th><th>Zonasi</th><th>Category</th><th class="rt">Target harian</th><th class="rt">Target MTD</th><th class="rt">Actual MTD</th><th class="rt">%</th>'+sd.map((d,i)=>'<th class="sc'+(i===0?' sep':'')+'" title="'+esc(fd(d))+'">'+dl(d)+'</th>').join('')+'<th class="rt sep">Target s/d '+esc(tl)+'</th><th class="rt">Actual</th><th class="rt">Achievement</th></tr>';
  const body=r.rows.map((p,i)=>'<tr><td class="num mut">'+(i+1)+'</td><td>'+esc(p.id)+'</td><td>'+esc(p.name)+'</td><td>'+esc(p.zona||'–')+'</td><td>'+esc(p.cat||'–')+'</td><td class="num rt">'+rp(p.tpd)+'</td><td class="num rt">'+rp(p.target)+'</td><td class="num rt">'+rp(p.sales)+'</td><td class="num rt bd">'+pct(p.pct)+'</td>'+p.codes.map((c,j)=>'<td class="sc'+(j===0?' sep':'')+' '+(isWork(c)?'on':'off')+'">'+esc(c||'')+'</td>').join('')+'<td class="num rt sep">'+rp(p.full)+'</td><td class="num rt lo-in"></td><td class="num rt"></td></tr>').join('');
  el.innerHTML='<table class="pt"><thead>'+head+'</thead><tbody>'+body+'</tbody></table>';
  $('low-note').textContent='Urut dari pencapaian terendah. Setelah kolom jadwal harian (S/P/M = masuk kerja, selain itu libur atau cuti) ada Target sampai akhir jadwal, Actual (dikosongkan, diisi manual di spreadsheet), dan Achievement (otomatis di spreadsheet).';
}
const colL=n=>{ let s=''; n++; while(n>0){ const m=(n-1)%26; s=String.fromCharCode(65+m)+s; n=Math.floor((n-1)/26); } return s; };
const lowTL=()=>{ const sd=(SCH&&SCH.dates)||[]; return sd.length? ckLab(sd[sd.length-1]) : 'akhir jadwal'; };
/* Susunan spreadsheet: SALESPERSON ID | SALESPERSON NAME | ZONA | CATEGORY | per tanggal [SHIFT | BUDGET | ACHIEV (kosong, diisi manual) | %] | TARGET | ACTUAL | ACHIEVEMENT */
const LC=4;
function lowLayout(){
  const sd=(SCH&&SCH.dates)||[], n=sd.length, multi=n&&sd[0].slice(0,7)!==sd[n-1].slice(0,7), dl=d=>multi? (+d.slice(8))+'/'+(+d.slice(5,7)) : String(+d.slice(8));
  const T=LC+4*n, hdr=['SALESPERSON ID','SALESPERSON NAME','ZONA','CATEGORY'], lab=['','','',''];
  sd.forEach(d=>{ hdr.push('SHIFT','BUDGET','ACHIEV','%'); lab.push(dl(d),'','',''); });
  hdr.push('TARGET s/d '+lowTL(),'ACTUAL','ACHIEVEMENT'); lab.push('TOTAL','','');
  const rows=lowInfo.rows.map((p,k)=>{ const R=7+k, x=[p.id,p.name,p.zona||'',p.cat||''], bs=[], as=[];
    sd.forEach((d,i)=>{ const c=LC+4*i, bC=colL(c+1)+R, aC=colL(c+2)+R; x.push(p.codes[i]||'', isWork(p.codes[i])?Math.round(p.tpd):0, null, '=IF('+bC+'>0,N('+aC+')/'+bC+',0)'); bs.push(bC); as.push(aC); });
    x.push('=SUM('+bs.join(',')+')','=SUM('+as.join(',')+')','=IF('+colL(T)+R+'>0,'+colL(T+1)+R+'/'+colL(T)+R+',0)');
    return x; });
  const per='MTD '+fd(lowInfo.from,{day:'numeric',month:'short'})+' sampai '+fd(lowInfo.to,{day:'numeric',month:'short',year:'numeric'});
  const title='Schedule Salesperson pencapaian di bawah 60% (MTD) · '+D.store, sub=per+' · kolom ACHIEV diisi manual, kolom % dan ACHIEVEMENT terisi otomatis';
  return {sd,n,T,hdr,lab,rows,title,sub,aoa:[[title],[sub],[],[],lab,hdr].concat(rows)};
}
const lowLoad=src=>new Promise((ok,no)=>{ if(window.ExcelJS) return ok(); const s=document.createElement('script'); const t=setTimeout(()=>no(new Error('timeout')),10000); s.src=src; s.onload=()=>{ clearTimeout(t); ok(); }; s.onerror=()=>{ clearTimeout(t); no(new Error('gagal memuat')); }; document.head.appendChild(s); });
async function lowStyled(Ly){
  await lowLoad('https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js');
  const wb=new ExcelJS.Workbook(), ws=wb.addWorksheet('Di bawah 60% MTD',{views:[{state:'frozen',xSplit:LC,ySplit:6}]});
  const NC=Ly.T+3, thin={style:'thin',color:{argb:'FF9AA5B8'}}, bd={top:thin,left:thin,bottom:thin,right:thin}, fill=c=>({type:'pattern',pattern:'solid',fgColor:{argb:c}});
  ws.getCell(1,1).value=Ly.title; ws.getCell(1,1).font={bold:true,size:14,color:{argb:'FF0F2547'}};
  ws.getCell(2,1).value=Ly.sub; ws.getCell(2,1).font={italic:true,color:{argb:'FF4A576B'}};
  Ly.aoa.slice(4).forEach((row,ri)=>{ const R=5+ri; row.forEach((v,ci)=>{ const cell=ws.getCell(R,ci+1);
    if(typeof v==='string'&&v[0]==='=') cell.value={formula:v.slice(1),result:0}; else if(v!==null&&v!=='') cell.value=v; }); });
  for(let i=0;i<Ly.n;i++){ const c=LC+1+4*i; ws.mergeCells(5,c,5,c+3); }
  ws.mergeCells(5,Ly.T+1,5,Ly.T+3);
  for(let c=1;c<=NC;c++){ const a=ws.getCell(5,c), h=ws.getCell(6,c);
    if(c>LC){ a.fill=fill('FFFFFF00'); a.font={bold:true,size:13,color:{argb:'FFC00000'}}; a.alignment={horizontal:'center',vertical:'middle'}; a.border=bd; }
    else if(c>1) a.fill=fill('FF00FFFF');
    h.fill=fill('FF0F2547'); h.font={bold:true,color:{argb:'FFFFFFFF'}}; h.alignment={horizontal:'center',vertical:'middle',wrapText:true}; h.border=bd; }
  ws.getRow(6).height=30;
  Ly.rows.forEach((row,k)=>{ const R=7+k; ws.getRow(R).height=20;
    for(let c=1;c<=NC;c++){ const cell=ws.getCell(R,c), inDays=c>LC&&c<=LC+4*Ly.n, k4=(c-LC-1)%4; cell.border=bd; cell.font={bold:c<=LC||(inDays&&k4===0)};
      if(c<=LC){ cell.alignment={horizontal:c===1?'left':'center',vertical:'middle'}; continue; }
      if(inDays){
        if(k4===0) cell.alignment={horizontal:'center',vertical:'middle'};
        else if(k4===3){ cell.numFmt='0.00%'; cell.alignment={horizontal:'right',vertical:'middle'}; }
        else { cell.numFmt='#,##0'; cell.alignment={horizontal:'right',vertical:'middle'}; if(k4===2) cell.fill=fill('FFFFFBEA'); }
      } else { cell.numFmt=c===Ly.T+3?'0.00%':'#,##0'; cell.font={bold:true}; cell.alignment={horizontal:'right',vertical:'middle'}; }
    } });
  [22,28,10,26].forEach((w,i)=>{ ws.getColumn(i+1).width=w; });
  for(let i=0;i<Ly.n;i++){ const c=LC+1+4*i; ws.getColumn(c).width=8; ws.getColumn(c+1).width=13; ws.getColumn(c+2).width=13; ws.getColumn(c+3).width=9; }
  ws.getColumn(Ly.T+1).width=18; ws.getColumn(Ly.T+2).width=15; ws.getColumn(Ly.T+3).width=14;
  ws.autoFilter={from:{row:6,column:1},to:{row:6,column:NC}};
  return new Blob([await wb.xlsx.writeBuffer()],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}
function lowPlain(Ly){
  const ws=XLSX.utils.aoa_to_sheet(Ly.aoa);
  Object.keys(ws).forEach(k=>{ if(k[0]==='!') return; const c=ws[k]; if(c.t==='s'&&c.v[0]==='='){ c.f=c.v.slice(1); c.t='n'; c.v=0; } });
  ws['!ref']=XLSX.utils.encode_range({s:{r:0,c:0},e:{r:Ly.aoa.length-1,c:Ly.T+2}});
  ws['!cols']=[{wch:22},{wch:28},{wch:10},{wch:26}].concat(Array.from({length:Ly.n},()=>[{wch:8},{wch:13},{wch:13},{wch:9}]).flat(),[{wch:18},{wch:15},{wch:14}]);
  const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,'Di bawah 60% MTD');
  return new Blob([XLSX.write(wb,{type:'array',bookType:'xlsx'})],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}
async function lowDownload(){
  if(!lowInfo||!lowInfo.rows.length) return;
  const Ly=lowLayout(), nm='salesperson-di-bawah-60-mtd-'+lowInfo.to+'.xlsx'; let blob=null, styled=true;
  $('low-note').textContent='Menyiapkan spreadsheet…';
  try{ blob=await lowStyled(Ly); }
  catch(e){ styled=false; if(!window.XLSX){ $('low-note').textContent='Spreadsheet gagal dibuat. Periksa koneksi, lalu coba lagi.'; return; } try{ blob=lowPlain(Ly); }catch(x){ $('low-note').textContent='Spreadsheet gagal dibuat.'; return; } }
  try{ await saveBlob(blob,nm); $('low-note').textContent='Spreadsheet diunduh'+(styled?'':' (versi polos tanpa warna, karena pustaka format tidak bisa dimuat)')+'. Kolom ACHIEV dikosongkan untuk diisi manual. Untuk Google Sheets: buka sheets.google.com, pilih File, Impor, lalu unggah file ini.'; }
  catch(e){ $('low-note').textContent='Spreadsheet gagal diunduh.'; }
}
async function lowCopy(){
  if(!lowInfo||!lowInfo.rows.length) return;
  const tsv=lowLayout().aoa.map(r=>r.map(v=>v==null?'':v).join('\t')).join('\n');
  const msg='Tabel disalin. Buka Google Sheets, klik sel A1, lalu tekan Ctrl+V. Warna tidak ikut tertempel, jadi untuk tampilan berwarna gunakan tombol Unduh spreadsheet.';
  try{ await navigator.clipboard.writeText(tsv); $('low-note').textContent=msg; }
  catch(e){ const t=document.createElement('textarea'); t.value=tsv; t.style.position='fixed'; t.style.opacity='0'; document.body.appendChild(t); t.select(); let ok=false; try{ ok=document.execCommand('copy'); }catch(x){} t.remove(); $('low-note').textContent=ok? msg : 'Browser tidak mengizinkan menyalin. Gunakan tombol Unduh spreadsheet.'; }
}
$('low-xlsx').addEventListener('click',lowDownload);
$('low-copy').addEventListener('click',lowCopy);
/* ---- Summary per MPP: klik baris atau kolom untuk daftar salesperson ---- */
function openMpp(tr,bandKey){
  const set=tr.dataset.set, zn=tr.dataset.z;
  let l=P.filter(p=>p.hasSch&&p.target>0); if(set==='pb') l=l.filter(p=>String(p.brand||'').toUpperCase()==='PRIVATE BRAND'); if(zn) l=l.filter(p=>p.zona===zn);
  const bandOf=p=>{ const a=p.sales/p.target; return a>=1?'hi':a<.75?'lo':'mid'; }, cnt={hi:0,mid:0,lo:0}; l.forEach(p=>cnt[bandOf(p)]++);
  const lab={hi:'≥100%',mid:'75–99%',lo:'<75%'}, cls={hi:'gd',mid:'am',lo:'bd'}, bk=(bandKey&&bandKey!=='all')? bandKey : '';
  const f=(bk? l.filter(p=>bandOf(p)===bk) : l).slice().sort((a,b)=>b.sales/b.target-a.sales/a.target);
  const rows=f.map((p,i)=>{ const b=bandOf(p); return kRow(p,['<td class="num mut">'+(i+1)+'</td>','<td>'+esc(p.id)+'</td>','<td>'+esc(p.name)+'</td>','<td>'+esc(p.zona||'–')+'</td>','<td>'+esc(p.brand?title(p.brand):'–')+'</td>',kNum(rp(p.target)),kNum(rp(p.sales)),'<td class="num rt '+cls[b]+'">'+pct(p.sales/p.target)+'</td>','<td class="'+cls[b]+'">'+lab[b]+'</td>']); });
  const ttl='Daftar salesperson · '+tr.dataset.lbl+(set==='pb'?' · Private Brand':'')+(bk?' · '+lab[bk]:''), info=[['Actual MPP',String(l.length)],['≥100%',String(cnt.hi)],['75–99%',String(cnt.mid)],['<75%',String(cnt.lo)]];
  lstCtx={title:ttl,info,rows:f.map((p,i)=>{ const b=bandOf(p); return lstRow(p,i,lab[b],cls[b]); }),name:'daftar-mpp-'+(set==='pb'?'pb':'semua')+(bk?'-'+bk:''),from:tr};
  kShow(ttl,info,
    rows.length? lstBar()+kTbl([['#'],['Sales ID'],['Salesperson'],['Zonasi'],['Master brand'],['Target',1],['Penjualan',1],['Achievement',1],['Kelompok']],rows)+'<p class="note" style="padding:8px 4px">Klik salah satu baris untuk melihat detail item. Urut dari capaian tertinggi.</p>' : '<p class="note" style="padding:14px 4px">Tidak ada salesperson pada kelompok ini.</p>',tr);
}
$('mpp-body').addEventListener('click',e=>{ const tr=e.target.closest('tr.mpp-r'); if(!tr) return; const c=e.target.closest('td[data-b]'); openMpp(tr,c?c.dataset.b:''); });
$('mpp-body').addEventListener('keydown',e=>{ if(e.key!=='Enter') return; const tr=e.target.closest('tr.mpp-r'); if(tr) openMpp(tr,''); });
function renderProt(){
  const rows=Object.values(M.prot||{}).sort((a,b)=>b.s-a.s);
  const tq=rows.reduce((s,r)=>s+r.q,0), ts=rows.reduce((s,r)=>s+r.s,0);
  $('prot-sub').textContent=(M.from&&M.to? (M.from===M.to? fd(M.from,{day:'numeric',month:'short'}) : fd(M.from,{day:'numeric',month:'short'})+' – '+fd(M.to,{day:'numeric',month:'short'}))+' · ' : '')+'master brand BCA · exclude PPN';
  const t=$('prot-body');
  if(!rows.length){ t.innerHTML='<p class="note">Belum ada penjualan Mitra10 Proteksi pada rentang ini.</p>'; }
  else t.innerHTML='<table class="pt"><thead><tr><th>#</th><th>Salesperson</th><th class="rt">Qty</th><th class="rt">Value penjualan</th></tr></thead><tbody>'+rows.map((r,i)=>'<tr><td class="num mut">'+(i+1)+'</td><td>'+(r.id? esc(r.name||r.id):'<i>Tanpa kode sales</i>')+'</td><td class="num rt">'+fmtN.format(r.q)+'</td><td class="num rt">'+rp(r.s)+'</td></tr>').join('')+'</tbody><tfoot><tr><td></td><td>Total</td><td class="num rt">'+fmtN.format(tq)+'</td><td class="num rt">'+rp(ts)+'</td></tr></tfoot></table>';
  const n=$('prot-note'); n.hidden=!M.protMiss; n.textContent=M.protMiss? M.protMiss+' tanggal belum memuat data proteksi. Unggah ulang laporan tanggal tersebut agar ikut terhitung.':'';
}
function renderTopBrand(){
  return; // tabel TOP 10 Master brand dihapus
  const rows=Object.entries(M.mb||{}).map(([k,v])=>({k,q:v.q,s:v.s})).sort((a,b)=>b.s-a.s).slice(0,10);
  $('tb-sub').textContent=(M.from&&M.to? (M.from===M.to? fd(M.from,{day:'numeric',month:'short'}) : fd(M.from,{day:'numeric',month:'short'})+' – '+fd(M.to,{day:'numeric',month:'short'}))+' · ' : '')+'sales exclude PPN per master brand';
  const t=$('tb-body');
  if(!rows.length) t.innerHTML='<p class="note">Belum ada data master brand pada rentang ini.</p>';
  else t.innerHTML='<table class="pt"><thead><tr><th>Master brand</th><th class="rt">Qty</th><th class="rt">Value</th></tr></thead><tbody>'+rows.map(r=>'<tr><td>'+esc(title(r.k))+'</td><td class="num rt">'+fmtN.format(r.q)+'</td><td class="num rt">'+rp(r.s)+'</td></tr>').join('')+'</tbody></table>';
  const n=$('tb-note'); n.hidden=!M.mbMiss; n.textContent=M.mbMiss? M.mbMiss+' tanggal belum memuat data master brand. Unggah ulang laporan tanggal tersebut agar ikut terhitung.':'';
}

/* Tampilan salesperson (validasi TL): sama dengan halaman link kode salesperson */
let SPCTX={isPB:false,BR:'',ZCs:new Set()};
let spSel='', spZ='', spShift=[], spDate='', spAppr={}, spApprT={}, spMgr=false, spJab='';
async function spLoadRole(email){
  if(!SB||!email) return;
  const { data } = await SB.from('sp_appr_managers').select('email,jabatan');
  const m=(data||[]).find(r=>String(r.email).toLowerCase()===String(email).toLowerCase());
  spMgr=!!m; spJab=m?(m.jabatan||'Manager'):'';
  if(m&&$('who-role')) $('who-role').textContent=spJab; else if(canWrite===true&&$('who-role')) $('who-role').textContent='Team Leader';
  try{ const { data:st } = await SB.rpc('sp_pin_status'); pinInfo=st||null; }catch(e){ pinInfo=null; }
  const pb=$('sp-pinbtns'); if(pb){ pb.hidden=!(pinInfo&&(pinInfo.role==='editor'||pinInfo.role==='admin')); const ad=$('pin-adm'); if(ad) ad.hidden=!(pinInfo&&pinInfo.role==='admin'); const nt=$('pin-note'); if(nt) nt.textContent=(pinInfo&&!pinInfo.has_pin&&pinInfo.role!=='viewer')?'PIN approval Anda belum diatur. Minta admin mengatur PIN.':''; }
  renderSP(true);
}
async function spLoadAppr(){
  if(TIM||!SB||!spSel) return; const id=spSel;
  if(spApprT[id]&&Date.now()-spApprT[id]<15000) return; spApprT[id]=Date.now();
  const { data, error } = await SB.from('sp_approvals').select('sale_date,tl_ok,tl_by,tl_at,mod_ok,mod_by,mod_at').eq('store_code',D.code).eq('sales_id',id);
  if(error){ spApprT[id]=0; return; }
  const o={}; (data||[]).forEach(r=>{ o[r.sale_date]=r; }); spAppr[id]=o; if(id===spSel) renderSP(true);
}
let pinInfo=null;
function pinOv(html){ let o=$('pin-ov'); if(!o){ o=document.createElement('div'); o.id='pin-ov'; o.className='pinov'; o.addEventListener('click',ev=>{ if(ev.target===o||(ev.target.closest&&ev.target.closest('[data-pinx]'))) pinClose(); }); document.body.appendChild(o); } o.innerHTML='<div class="pinbox" role="dialog" aria-modal="true">'+html+'</div>'; o.hidden=false; const f=o.querySelector('input'); if(f) f.focus(); }
function pinClose(){ const o=$('pin-ov'); if(o) o.hidden=true; }
document.addEventListener('keydown',ev=>{ if(ev.key==='Escape') pinClose(); });
document.addEventListener('input',ev=>{ const t=ev.target; if(t&&t.closest&&t.closest('#pin-ov')&&t.tagName==='INPUT'){ const v=t.value.replace(/\D/g,'').slice(0,6); if(v!==t.value) t.value=v; } });
function pinErr(r){ const m={nopin:'PIN Anda belum diatur. Minta admin mengatur PIN.',locked:'Terlalu banyak salah. Coba lagi dalam '+(r.wait||5)+' menit.',bad:'PIN salah. Sisa percobaan: '+r.left+'.',role:'Akun Anda tidak punya izin.',role_tl:'Approval TL hanya untuk Team Leader.',role_mod:'Approval MOD hanya untuk Store Manager / Asst Manager.',need_tl:'Approval TL harus dilakukan lebih dulu.',mod_first:'Batalkan Approval MOD lebih dulu.',format:'PIN harus 6 digit angka.',nouser:'Akun tidak ditemukan.',kind:'Permintaan tidak valid.'}; return m[r.err]||'Gagal ('+(r.err||'error')+').'; }
const pinIn=(id,ph)=>'<input id="'+id+'" type="text" inputmode="numeric" maxlength="6" autocomplete="off" autocapitalize="off" spellcheck="false" data-lpignore="true" name="pin-'+Math.random().toString(36).slice(2,7)+'" placeholder="'+ph+'">';
document.addEventListener('click',ev=>{ const b=ev.target.closest&&ev.target.closest('#sp-list .spap'); if(!b) return;
  const k=b.dataset.k, id=b.dataset.id, d=b.dataset.d, on=b.dataset.on==='1', e=SCHM[id];
  pinOv('<h3>'+(on?'Approve ':'Batalkan ')+k.toUpperCase()+'</h3><p class="note">'+esc((e&&e.name)||id)+' · '+esc(fd(d,{weekday:'short',day:'numeric',month:'short'}))+'</p>'+pinIn('pin-in','PIN 6 digit')+'<div class="pinerr" id="pin-err"></div><div class="pinbt"><button class="btn sm" data-pinx>Batal</button><button class="btn sm pri" id="pin-ok">Konfirmasi</button></div>');
  const go=async()=>{ const v=$('pin-in').value.trim(); if(!/^\d{6}$/.test(v)){ $('pin-err').textContent='Masukkan PIN 6 digit.'; return; }
    $('pin-ok').disabled=true; const { data, error } = await SB.rpc('sp_approve',{p_store:D.code,p_sales_id:id,p_date:d,p_kind:k,p_pin:v,p_on:on}); $('pin-ok').disabled=false;
    if(error||!data){ $('pin-err').textContent='Gagal menghubungi server. Coba lagi.'; return; }
    if(!data.ok){ $('pin-err').textContent=pinErr(data); $('pin-in').value=''; $('pin-in').focus(); return; }
    (spAppr[id]||(spAppr[id]={}))[d]=data.row; pinClose(); renderSP(true); };
  $('pin-ok').onclick=go; $('pin-in').onkeydown=x=>{ if(x.key==='Enter') go(); };
});
function pinChange(){
  pinOv('<h3>Ganti PIN saya</h3><p class="note">PIN 6 digit angka. Jangan dibagikan ke orang lain.</p>'+pinIn('pc-old','PIN lama')+pinIn('pc-new','PIN baru (6 digit)')+pinIn('pc-new2','Ulangi PIN baru')+'<div class="pinerr" id="pin-err"></div><div class="pinbt"><button class="btn sm" data-pinx>Batal</button><button class="btn sm pri" id="pin-ok">Simpan</button></div>');
  $('pin-ok').onclick=async()=>{ const o=$('pc-old').value.trim(), n=$('pc-new').value.trim(), n2=$('pc-new2').value.trim(), er=$('pin-err');
    if(!/^\d{6}$/.test(o)||!/^\d{6}$/.test(n)){ er.textContent='PIN harus 6 digit angka.'; return; } if(n!==n2){ er.textContent='PIN baru dan ulangannya tidak sama.'; return; }
    $('pin-ok').disabled=true; const { data, error } = await SB.rpc('sp_pin_change',{p_old:o,p_new:n}); $('pin-ok').disabled=false;
    if(error||!data){ er.textContent='Gagal menghubungi server.'; return; } if(!data.ok){ er.textContent=pinErr(data); return; }
    er.style.color='#0e6b49'; er.textContent='PIN berhasil diganti.'; setTimeout(pinClose,900); };
}
async function pinAdmin(msg){
  pinOv('<h3>Atur PIN tim</h3><p class="note">Memuat…</p>'); const { data } = await SB.rpc('sp_pin_admin_list');
  if(!data||!data.ok){ pinOv('<h3>Atur PIN tim</h3><p class="note">Gagal memuat daftar.</p><div class="pinbt"><button class="btn sm" data-pinx>Tutup</button></div>'); return; }
  pinOv('<h3>Atur PIN tim</h3><p class="note">Isi PIN 6 digit lalu Simpan. Salah 5 kali = terkunci 5 menit; menyimpan PIN baru membuka kunci. TL bisa mengganti PIN sendiri lewat "Ganti PIN saya".</p><div class="pinls">'+data.rows.map(r=>'<div class="pinrow"><div><b>'+esc(r.email)+'</b><br><small>'+(r.mgr?'Manager (MOD)':'Team Leader')+' · '+(r.has_pin?'PIN sudah diatur':'<span style="color:#b4271c">Belum ada PIN</span>')+'</small></div><input type="text" inputmode="numeric" maxlength="6" autocomplete="off" autocapitalize="off" spellcheck="false" data-lpignore="true" name="pn-'+Math.random().toString(36).slice(2,7)+'" placeholder="6 digit" data-em="'+esc(r.email)+'"><button class="btn sm" data-pinset="'+esc(r.email)+'">Simpan</button></div>').join('')+'</div><div class="pinerr" id="pin-err" style="'+(msg?'color:#0e6b49':'')+'">'+esc(msg||'')+'</div><div class="pinbt"><button class="btn sm" data-pinx>Tutup</button></div>');
}
document.addEventListener('click',async ev=>{ const t=ev.target; if(!t||!t.closest) return;
  if(t.closest('#pin-chg')) return pinChange(); if(t.closest('#pin-adm')) return pinAdmin();
  const b=t.closest('[data-pinset]'); if(!b) return; const em=b.dataset.pinset, inp=[...document.querySelectorAll('#pin-ov input[data-em]')].find(x=>x.dataset.em===em), v=(inp&&inp.value||'').trim();
  if(!/^\d{6}$/.test(v)){ $('pin-err').style.color=''; $('pin-err').textContent='PIN harus 6 digit angka.'; return; }
  b.disabled=true; const { data, error } = await SB.rpc('sp_pin_admin_set',{p_email:em,p_pin:v}); b.disabled=false;
  if(error||!data||!data.ok){ $('pin-err').style.color=''; $('pin-err').textContent=data?pinErr(data):'Gagal menghubungi server.'; return; }
  pinAdmin('PIN '+em+' tersimpan.'); });
let spApprDT=0, spApprReady=false;
async function spLoadApprDate(){ // memuat semua approval toko (untuk summary & notice hari tidak lengkap)
  if(TIM||!SB) return;
  if(spApprDT&&Date.now()-spApprDT<15000) return; spApprDT=Date.now();
  const all=[]; for(let from=0;;from+=1000){ const { data, error } = await SB.from('sp_approvals').select('sales_id,sale_date,tl_ok,tl_by,tl_at,mod_ok,mod_by,mod_at').eq('store_code',D.code).range(from,from+999); if(error){ spApprDT=0; return; } all.push(...(data||[])); if(!data||data.length<1000) break; }
  const m={}; all.forEach(r=>{ (m[r.sales_id]||(m[r.sales_id]={}))[r.sale_date]=r; }); spAppr=m; spApprReady=true; renderSP(true);
}
document.addEventListener('change',async ev=>{
  const c=ev.target; if(!c||!c.classList||!c.classList.contains('spc')) return;
  const k=c.dataset.k, d=c.dataset.d, id=c.dataset.id||spSel, msg=$('sp-msg'); c.disabled=true;
  const row={store_code:D.code, sales_id:id, sale_date:d}; row[k+'_ok']=c.checked;
  const { data, error } = await SB.from('sp_approvals').upsert(row,{onConflict:'store_code,sales_id,sale_date'}).select('sale_date,tl_ok,tl_by,tl_at,mod_ok,mod_by,mod_at').single();
  if(error){ c.checked=!c.checked; c.disabled=false; if(msg) msg.textContent='Gagal menyimpan approval: '+((/row-level|permission/i.test(error.message||''))?'akun Anda tidak punya izin (hanya editor/admin).':(error.message||'error')); return; }
  (spAppr[id]||(spAppr[id]={}))[d]=data; renderSP(true);
});
function spOpen(k){
  const e=SCHM[spSel], last=dates[dates.length-1]; if(!e||!last) return;
  let dts,ttl,per,mode='',kind=''; const BRQ=String(e.brand||'').trim().toUpperCase(); const lf=d=>fd(d,{day:'numeric',month:'short',year:'numeric'});
  if(k==='D'){ dts=[last]; ttl='Item terjual · Daily'; per=lf(last); }
  else if(k==='M'){ dts=SCH.dates.filter(d=>d<=last); ttl='Item terjual · MTD'; per='1 s.d. '+lf(last)+' (semua hari)'; }
  else if(/^K:/.test(k)){ kind=k.slice(2); dts=SCH.dates.filter(d=>d<=last); ttl=(kind==='mb'?(SPCTX.isPB?'Private Brand (category zonasi)':'Master brand '+(e.brand||'')):kind==='non'?(SPCTX.isPB?'Non Private Brand':'Non master brand'):'Total penjualan')+' · MTD'; per='1 s.d. '+lf(last); }
  else { const pr=String(k).split('|'); mode=pr[1]||''; k=pr[0]; dts=[k]; ttl='Item terjual · '+fd(k,{weekday:'short',day:'numeric',month:'short'}); per=lf(k); }
  const m={}; dts.forEach(d=>{ const q=days[d]&&(days[d].people||[]).find(x=>x.id===spSel); if(q&&Array.isArray(q.it)) q.it.forEach(r=>{ if(kind){ const isMb=SPCTX.isPB?(r[6]===1&&SPCTX.ZCs.has(String(r[0]==null?'':r[0]).trim().toUpperCase())):(SPCTX.BR&&String(r[1]||'').trim().toUpperCase()===SPCTX.BR); if(kind==='mb'&&!isMb) return; if(kind==='non'&&isMb) return; } const key=r[2]||r[3]; const t=m[key]||(m[key]=[r[0],r[1],r[2],r[3],0,0]); t[4]+=r[4]; t[5]+=r[5]; }); });
  const it=Object.values(m).sort((a,b)=>b[5]-a[5]), tq=it.reduce((s,r)=>s+r[4],0), ts=it.reduce((s,r)=>s+r[5],0);
  const oldD=kind==='mb'&&SPCTX.isPB&&!it.length&&dts.some(d=>{ const q=days[d]&&(days[d].people||[]).find(x=>x.id===spSel); return q&&(+q.pb||0)>0&&!q.pbc; });
  const html=oldD? '<p class="note" style="padding:14px 4px">Data penjualan yang tersimpan belum memuat penanda Master type brand per item, jadi detail Private Brand belum bisa ditampilkan. Minta Team Leader upload ulang file penjualan (semua hari bulan ini) di dashboard, lalu detail ini akan terisi.</p>' : it.length? '<table class="pt"><thead><tr><th>Category</th><th>No Item</th><th>Description</th><th class="rt">Qty</th><th class="rt">Value</th></tr></thead><tbody>'+it.map(r=>'<tr><td>'+esc(title(r[0]))+'</td><td class="num">'+esc(r[2])+'</td><td>'+esc(r[3])+'</td><td class="num rt">'+fmtN.format(r[4])+'</td><td class="num rt">'+fmtN.format(r[5])+'</td></tr>').join('')+'<tr class="tot"><td colspan="3"><b>TOTAL ('+it.length+' item)</b></td><td class="num rt"><b>'+fmtN.format(tq)+'</b></td><td class="num rt"><b>'+fmtN.format(ts)+'</b></td></tr></tbody></table>' : '<p class="note" style="padding:14px 4px">Belum ada rincian item untuk periode ini.</p>';
  kShow(ttl,[['Salesperson',e.name||e.id],['Sales ID',e.id],['Zonasi',e.zona||'–'],['Total value','Rp '+fmtN.format(ts)]],html,null);
  $('kpop-per').textContent=per+' · sales exclude PPN';
}
document.addEventListener('click',ev=>{ const c=ev.target.closest&&ev.target.closest('#sp-body [data-spd]'); if(!c) return; spOpen(c.dataset.spd); });
document.addEventListener('keydown',ev=>{ if(ev.key==='Enter'&&ev.target.matches&&ev.target.matches('#sp-body [data-spd][role=button]')) spOpen(ev.target.dataset.spd); });
function spList(all0,last,noFetch){
  const box=$('sp-list'); if(!box) return;
  const ld=(spDate&&SCH.dates.includes(spDate))?spDate:last; if(!ld){ box.innerHTML=''; return; }
  if(!noFetch) spLoadApprDate(ld);
  const noAp=!!TIM; const ready=(typeof spApprReady==='undefined')||spApprReady;
  const nn=v=>fmtN.format(Math.round(v)), pc=p=>p==null?'–':(Math.round(p*10)/10).toString().replace('.',',')+'%', sd=d=>fd(d,{day:'numeric',month:'short'});
  const scope=all0.filter(e=>!spZ||e.zona===spZ).map(e0=>SCHM[e0.id]||e0);
  const dayStat=(d,shifts)=>{ const i=SCH.dates.indexOf(d), dd=days[d], idx={}; if(dd) (dd.people||[]).forEach(q=>{ idx[q.id]=q; });
    let n=0,tl=0,md=0,b=0,a=0;
    scope.forEach(e=>{ const c=e.c[i]; if(!isWork(c)) return; if(shifts&&shifts.length&&!shifts.includes(c)) return;
      const bb=e.tpe||0, q=idx[e.id], aa=q?(+q.sales||0):0, ap=((spAppr[e.id]||{})[d])||{};
      n++; b+=bb; a+=aa; if(ap.tl_ok) tl++; const auto=ap.tl_ok&&bb>0&&aa/bb*100>=100; if(ap.mod_ok||auto) md++; });
    return {n,tl,md,b,a,has:!!dd}; };
  const st=dayStat(ld,spShift);
  SPV={scope,last,ld,shifts:spShift,noAp};
  const tile=(lbl,val,sub,cls,f,so)=>'<div class="'+(cls||'')+'" data-spp="'+(f||'all')+'" data-sps="'+(so||'')+'" role="button" tabindex="0" title="Klik untuk melihat detail"><small>'+lbl+'</small><b>'+val+'</b><span>'+sub+'</span></div>';
  const okc=(x)=>!ready?'':(st.n>0&&x===st.n?'apok':'apno');
  let tiles=tile('Salesperson kerja',st.n,'shift kerja pada '+esc(sd(ld)))+
    (noAp?'':tile('Approval TL',ready?st.tl+' / '+st.n:'…',ready?(st.n?pc(st.tl/st.n*100)+' sudah approve':'–'):'memuat',okc(st.tl),'tl')+
             tile('Approval MOD',ready?st.md+' / '+st.n:'…',ready?(st.n?pc(st.md/st.n*100)+' sudah approve (termasuk auto)':'–'):'memuat',okc(st.md),'md'))+
    tile('Pencapaian hari itu',st.has&&st.b>0?pc(st.a/st.b*100):'–',st.has?'Rp '+nn(st.a)+' dari budget Rp '+nn(st.b):'belum ada data penjualan','','all','ach');
  let notice='';
  if(!noAp&&ready){
    const inc=SCH.dates.filter(d=>d<=last).map(d=>Object.assign({d},dayStat(d))).filter(x=>x.n>0&&(x.tl<x.n||x.md<x.n));
    notice=inc.length?'<div class="spwarn">⚠ <b>Validasi belum lengkap pada '+inc.length+' hari</b> (s.d. '+esc(sd(last))+'). Klik tanggal untuk melihat detail:<div class="spchips">'+inc.map(x=>'<button type="button" class="spchip'+(x.d===ld?' on':'')+'" data-pickd="'+x.d+'">'+esc(sd(x.d))+' · TL '+x.tl+'/'+x.n+' · MOD '+x.md+'/'+x.n+'</button>').join('')+'</div></div>'
      :'<div class="spokn">✓ Validasi lengkap untuk semua hari s.d. '+esc(sd(last))+'.</div>';
  }
  box.innerHTML='<div class="spv"><b>Summary validasi</b> <span class="note">Tanggal '+esc(fd(ld,{weekday:'short',day:'numeric',month:'short',year:'numeric'}))+(st.has?'':' · belum ada data penjualan')+' · klik kartu atau tanggal untuk melihat detail. Approval dilakukan oleh TL dan MOD di HP salesperson.</span><div class="spsum">'+tiles+'</div>'+notice+'</div>';
}
let SPV=null, SPPOP=null;
const spTime=iso=>{ if(!iso) return ''; try{ return new Date(iso).toLocaleTimeString('id-ID',{timeZone:'Asia/Jakarta',hour:'2-digit',minute:'2-digit'}).replace('.',':'); }catch(e){ return ''; } };
function spValRows(d){ const c=SPV, i=SCH.dates.indexOf(d), dd=days[d], idx={}; if(dd) (dd.people||[]).forEach(q=>{ idx[q.id]=q; }); const out=[];
  c.scope.forEach(e=>{ const sh=e.c[i]; if(!isWork(sh)) return; if(c.shifts&&c.shifts.length&&!c.shifts.includes(sh)) return;
    const bb=e.tpe||0, q=idx[e.id], aa=q?(+q.sales||0):0, ap=((spAppr[e.id]||{})[d])||{}, auto=!ap.mod_ok&&!!ap.tl_ok&&bb>0&&aa/bb*100>=100;
    out.push({e,sh,bb,aa,pct:bb>0?aa/bb*100:null,tl:!!ap.tl_ok,tlb:ap.tl_by,tla:ap.tl_at,md:!!ap.mod_ok||auto,auto,mdb:ap.mod_by,mda:ap.mod_at}); });
  return out; }
function spValShow(d,f,so,open){
  const c=SPV; if(!c) return; SPPOP={d,f,so}; if(open) kShow('Detail validasi & pencapaian',[],'',null);
  const nn=v=>fmtN.format(Math.round(v)), pc=p=>p==null?'–':(Math.round(p*10)/10).toString().replace('.',',')+'%', sd=x=>fd(x,{day:'numeric',month:'short'});
  const who=v=>v?esc(String(v).split('@')[0]):'';
  const rows=spValRows(d), n=rows.length, cTl=rows.filter(r=>r.tl).length, cMd=rows.filter(r=>r.md).length, tb=rows.reduce((s,r)=>s+r.bb,0), ta=rows.reduce((s,r)=>s+r.aa,0);
  const noAp=c.noAp; if(noAp&&(f==='tl'||f==='md'||f==='inc'||f==='ok')) f='all';
  const flt={all:()=>true,tl:r=>!r.tl,md:r=>!r.md,inc:r=>!r.tl||!r.md,ok:r=>r.tl&&r.md}[f]||(()=>true);
  const sh=rows.filter(flt); const cnt=g=>rows.filter(g).length;
  sh.sort(so==='ach'?(a,b)=>(b.pct==null)-(a.pct==null)||(b.pct||0)-(a.pct||0):so==='low'?(a,b)=>(a.pct==null)-(b.pct==null)||(a.pct||0)-(b.pct||0):(a,b)=>String(a.e.name||a.e.id).localeCompare(String(b.e.name||b.e.id),'id'));
  $('kpop-t').textContent='Detail validasi & pencapaian';
  $('kpop-per').textContent=fd(d,{weekday:'long',day:'numeric',month:'long',year:'numeric'})+' · sales exclude PPN';
  const info=[['Salesperson kerja',n]].concat(noAp?[]:[['Approval TL',cTl+' / '+n],['Approval MOD',cMd+' / '+n]]).concat([['Pencapaian',tb>0?pc(ta/tb*100):'–'],['Actual',    'Rp '+nn(ta)],['Budget','Rp '+nn(tb)]]);
  $('kpop-info').innerHTML=info.map(i=>'<div><span>'+esc(i[0])+'</span><b>'+esc(String(i[1]))+'</b></div>').join('');
  const dopt=SCH.dates.filter(x=>x<=c.last).map(x=>{ let w=''; if(!noAp){ const bk=SPV.ld; const rr=spValRows(x); if(rr.length&&rr.some(r=>!r.tl||!r.md)) w=' ⚠'; } return '<option value="'+x+'"'+(x===d?' selected':'')+'>'+esc(fd(x,{weekday:'short',day:'numeric',month:'short'}))+w+'</option>'; }).join('');
  const tabs=[['all','Semua',n]].concat(noAp?[]:[['tl','Belum TL',cnt(r=>!r.tl)],['md','Belum MOD',cnt(r=>!r.md)],['inc','Belum lengkap',cnt(r=>!r.tl||!r.md)],['ok','Lengkap',cnt(r=>r.tl&&r.md)]]);
  const ctl='<div class="spvc"><label>Tanggal <select data-spvd>'+dopt+'</select></label><div class="spvt">'+tabs.map(t=>'<button type="button" data-spvf="'+t[0]+'"'+(f===t[0]?' class="on"':'')+'>'+t[1]+' <i>'+t[2]+'</i></button>').join('')+'</div><label>Urutkan <select data-spvs><option value=""'+(!so?' selected':'')+'>Nama</option><option value="ach"'+(so==='ach'?' selected':'')+'>Pencapaian tertinggi</option><option value="low"'+(so==='low'?' selected':'')+'>Pencapaian terendah</option></select></label></div>';
  const stc=(on,by,at,auto)=>on?'<td class="spok">✓ '+(auto?'Auto':who(by))+(at&&!auto?'<small>'+esc(spTime(at))+'</small>':'')+'</td>':'<td class="spno">Belum</td>';
  const col=p=>p==null?'':p>=100?'#0e6b49':p>=75?'#a86a00':'#b4271c';
  const body=sh.length?'<table class="pt spvtb"><thead><tr><th>#</th><th>Sales ID</th><th>Nama</th><th>Zonasi</th><th>Shift</th><th class="rt">Budget</th><th class="rt">Actual</th><th class="rt">%</th>'+(noAp?'':'<th>TL</th><th>MOD</th>')+'</tr></thead><tbody>'+sh.map((r,k)=>'<tr><td class="mut">'+(k+1)+'</td><td class="num">'+esc(String(r.e.id).replace(/^S10005-0*/,''))+'</td><td><b>'+esc(r.e.name||r.e.id)+'</b></td><td>'+esc(r.e.zona||'–')+'</td><td>'+esc(r.sh)+'</td>'+kNum(nn(r.bb))+kNum(nn(r.aa))+'<td class="num rt" style="color:'+col(r.pct)+';font-weight:700">'+pc(r.pct)+'</td>'+(noAp?'':stc(r.tl,r.tlb,r.tla)+stc(r.md,r.mdb,r.mda,r.auto))+'</tr>').join('')+'<tr class="tot"><td colspan="5"><b>TOTAL ('+sh.length+' orang)</b></td>'+kNum('<b>'+nn(sh.reduce((s,r)=>s+r.bb,0))+'</b>')+kNum('<b>'+nn(sh.reduce((s,r)=>s+r.aa,0))+'</b>')+'<td class="num rt"><b>'+pc((x=>x.b>0?x.a/x.b*100:null)({b:sh.reduce((s,r)=>s+r.bb,0),a:sh.reduce((s,r)=>s+r.aa,0)}))+'</b></td>'+(noAp?'':'<td>'+sh.filter(r=>r.tl).length+'/'+sh.length+'</td><td>'+sh.filter(r=>r.md).length+'/'+sh.length+'</td>')+'</tr></tbody></table>':'<p class="note" style="padding:14px 4px">'+(n?'Tidak ada salesperson pada filter ini.':'Tidak ada salesperson kerja pada tanggal ini.')+'</p>';
  $('kpop-body').innerHTML=ctl+body; if(open) $('kpop-body').scrollTop=0; }
document.addEventListener('click',ev=>{ const t=ev.target.closest&&ev.target.closest('#sp-list [data-pickd], #sp-list [data-spp], #kpop-body [data-spvf]'); if(!t) return;
  if(t.dataset.spvf){ if(SPPOP) spValShow(SPPOP.d,t.dataset.spvf,SPPOP.so); return; }
  if(!SPV) return; if(t.dataset.pickd) spValShow(t.dataset.pickd,'inc','',true); else spValShow(SPV.ld,t.dataset.spp,t.dataset.sps,true); });
document.addEventListener('keydown',ev=>{ if(ev.key==='Enter'&&ev.target.matches&&ev.target.matches('#sp-list [data-spp]')&&SPV) spValShow(SPV.ld,ev.target.dataset.spp,ev.target.dataset.sps,true); });
document.addEventListener('change',ev=>{ const t=ev.target; if(!t||!t.matches||!SPPOP) return; if(t.matches('[data-spvd]')) spValShow(t.value,SPPOP.f,SPPOP.so); else if(t.matches('[data-spvs]')) spValShow(SPPOP.d,SPPOP.f,t.value); });
function renderSP(noFetch){
  const box=$('sp-body'), sel=$('sp-sel'); if(!box) return;
  const all0=SCH.people.filter(e=>e.zona && e.role!=='Kasir').sort((a,b)=>String(a.name||a.id).localeCompare(String(b.name||b.id),'id'));
  if(!all0.length){ sel.innerHTML=''; box.innerHTML='<p class="note">Unggah file jadwal &amp; target untuk menampilkan tampilan ini.</p>'; return; }
  const zs=[...new Set(all0.map(e=>e.zona))].sort((a,b)=>String(a).localeCompare(String(b),'id',{numeric:true}));
  if(spZ&&!zs.includes(spZ)) spZ='';
  const ppl=all0.filter(e=>!spZ||e.zona===spZ);
  if(!ppl.some(e=>e.id===spSel)) spSel=ppl[0].id;
  const shs=[...new Set(all0.flatMap(e=>e.c||[]).filter(Boolean))].sort((a,b)=>String(a).localeCompare(String(b),'id'));
  spShift=spShift.filter(x=>shs.includes(x));
  if(spDate&&!SCH.dates.includes(spDate)) spDate='';
  const op=(v,t,cur)=>'<option value="'+esc(v)+'"'+(String(v)===String(cur)?' selected':'')+'>'+esc(t)+'</option>';
  $('sp-zona').innerHTML=op('','Semua zonasi',spZ)+zs.map(z=>op(z,z,spZ)).join('');
  { const L=$('sp-shiftl'), sig=shs.join('|'); if(L.dataset.sig!==sig){ L.dataset.sig=sig; L.innerHTML='<button type="button" class="btn sm" data-spshreset="1">Semua shift</button>'+shs.map(c=>'<label><input type="checkbox" class="spshc" value="'+esc(c)+'"> '+esc(c)+'</label>').join(''); }
    L.querySelectorAll('.spshc').forEach(x=>{ x.checked=spShift.includes(x.value); });
    $('sp-shifts').textContent=spShift.length? 'Shift: '+spShift.join(', ') : 'Semua shift'; }
  $('sp-date').innerHTML=op('','Semua tanggal',spDate)+SCH.dates.map(d=>op(d,fd(d,{weekday:'short',day:'numeric',month:'short'}),spDate)).join('');
  sel.innerHTML=ppl.map(e=>'<option value="'+esc(e.id)+'"'+(e.id===spSel?' selected':'')+'>'+esc((e.name||e.id)+' · '+e.id.replace(/^S\d+-0*/,''))+'</option>').join('');
  if(!noFetch) spLoadAppr();
  spList(all0,dates[dates.length-1]||'',noFetch);
  const nn=v=>fmtN.format(Math.round(v)); const e=SCHM[spSel], last=dates[dates.length-1]||'';
  const cls=p=>p==null?'spn':p>=100?'spg':p>=75?'spa':'spb', pc=p=>p==null?'–':(Math.round(p*10)/10).toString().replace('.',',')+'%';
  const sales={}; dates.forEach(d=>{ const q=(days[d].people||[]).find(x=>x.id===spSel); sales[d]=q?(+q.sales||0):0; });
  const BR=String(e.brand||'').trim().toUpperCase(), noAp=!!TIM;
  const isPB=BR==='PRIVATE BRAND', UPc=x=>String(x==null?'':x).trim().toUpperCase();
  const ZCs=new Set(SCH.people.filter(x=>x.zona&&e.zona&&UPc(x.zona)===UPc(e.zona)&&x.cat).map(x=>UPc(x.cat))); if(!ZCs.size&&e.cat) ZCs.add(UPc(e.cat)); // category yang termasuk di zonasi salesperson
  const sumC=o=>{ let v=0; Object.keys(o||{}).forEach(k=>{ if(ZCs.has(UPc(k))) v+=+o[k]||0; }); return v; };
  let pbFb=false; SPCTX={isPB,BR,ZCs};
  const dq=d=>{ const q=days[d]&&(days[d].people||[]).find(x=>x.id===spSel); const it=(q&&q.it)||[]; let mb=0,bca=0; it.forEach(r=>{ const v=+r[5]||0, m=UPc(r[1]); if(m==='BCA') bca+=+r[4]||0; if(!isPB&&BR&&m===BR) mb+=v; }); if(isPB){ if(q&&q.pbc&&typeof q.pbc==='object') mb=sumC(q.pbc); else { mb=+(q&&q.pb)||0; if(mb>0) pbFb=true; } } return {mb,bca,cat:isPB?sumC(q&&q.cats):0,inv:q?(+q.receipts||0):0}; };
  const apr=spAppr[spSel]||{};
  const apWho=(ap,k)=>{ if(!ap[k+'_ok']) return ''; const t=ap[k+'_at']?new Date(ap[k+'_at']).toLocaleString('id-ID',{day:'numeric',month:'short',hour:'2-digit',minute:'2-digit'}):''; return 'Disetujui oleh '+(ap[k+'_by']||'?')+(t?' · '+t:''); };
  const ab=(k,d,ap,on)=>{ const ok=!!ap[k+'_ok'], who=apWho(ap,k);
    if(on) return '<td class="c"><input type="checkbox" class="spc" data-k="'+k+'" data-d="'+d+'"'+(ok?' checked':'')+' title="'+esc(who||'Klik untuk menyetujui')+'" aria-label="Approval '+k.toUpperCase()+' '+d+'"></td>';
    return '<td class="c" title="'+esc(who)+'">'+(ok?'<span class="spg">✓</span>':'–')+'</td>'; };
  const canAp=canWrite===true&&!noAp;
  let mB=0,mA=0,mMB=0,mCat=0,tInv=0,tBca=0,monthB=0,rows='';
  SCH.dates.forEach((d,i)=>{ const c=e.c[i], w=isWork(c), fut=!last||d>last, b=w?e.tpe:0, a=sales[d]||0, m=dq(d);
    if(w) monthB+=b; if(!fut){ tInv+=m.inv; tBca+=m.bca; mA+=a; mMB+=m.mb; mCat+=m.cat; if(w) mB+=b; }
    if((spShift.length&&!spShift.includes(c))||(spDate&&d!==spDate)) return;
    const pp=(w&&b>0&&!fut)?a/b*100:null;
    const ap=apr[d]||{};
    rows+='<tr class="'+(!w?'off':'')+'"><td>'+esc(fd(d,{weekday:'short',day:'numeric',month:'short'}))+'</td><td style="text-align:center">'+esc(c||'–')+'</td><td>'+(w?nn(b):'–')+'</td>'+(fut?'<td></td>':'<td'+(a?' class="spk" data-spd="'+d+'" title="Lihat item"':'')+'>'+nn(a)+'</td>')+'<td class="'+cls(pp)+'"><b>'+(pp==null?'':pc(pp))+'</b></td><td>'+(fut?'':m.inv)+'</td><td>'+(fut?'':m.bca)+'</td>'+(noAp?'':(fut?'<td></td><td></td>':ab('tl',d,ap,canAp&&!spMgr)+ab('mod',d,ap,canAp&&spMgr)))+'</tr>'; });
  const mp=mB>0?mA/mB*100:null;
  rows+='<tr class="tt"><td>TOTAL s.d. '+esc(last?fd(last,{day:'numeric',month:'short'}):'-')+'</td><td></td><td>'+nn(mB)+'</td><td>'+nn(mA)+'</td><td>'+pc(mp)+'</td><td>'+tInv+'</td><td>'+tBca+'</td>'+(noAp?'':'<td></td><td></td>')+'</tr>';
  const li=SCH.dates.indexOf(last), lw=li>=0&&isWork(e.c[li]), da=last?(sales[last]||0):0, dp=(lw&&e.tpe>0)?da/e.tpe*100:null;
  const nonv=Math.max(0,mA-mMB), sh=v=>mA>0?pc(v/mA*100):'–', cn=String(e.cat||'–');
  let cmp=''; if(isPB){ const pc2=mCat>0?mMB/mCat*100:null, zl=[...ZCs].map(title).join(', '); cmp='<div class="spcmp"><b>Perbandingan Private Brand dengan Category zonasi</b> <span class="note">'+esc(e.zona||'–')+': '+esc(zl)+'</span><div class="cbar"><i style="width:'+(mCat>0?Math.min(100,mMB/mCat*100):0)+'%"></i></div><div class="note">Private Brand (Master type, hanya category zonasi) Rp '+nn(mMB)+' · Total penjualan category zonasi Rp '+nn(mCat)+' → <b>'+(pc2==null?'–':pc(pc2))+'</b> dari penjualan category zonasi.</div>'+(pbFb?'<div class="note" style="color:#a86a00">Sebagian data lama belum memuat rincian Private Brand per category, hari itu dihitung dari semua category. Upload ulang file penjualan.</div>':'')+'</div>'; }
  const mbc='<div class="spmb"><b>Total penjualan master brand</b> <span class="note">'+(last?'1 s.d. '+esc(fd(last,{day:'numeric',month:'short'})):'Belum ada data')+' · semua hari ber-penjualan dihitung, termasuk OFF/CT/RO'+(isPB?' · Private Brand = Master type brand, hanya category zonasi':'')+' · klik kartu untuk melihat item</span><div class="spmbs"><div class="spk" data-spd="K:mb" tabindex="0" role="button"><small>'+(isPB?'Private Brand (category zonasi)':'Master brand '+esc(e.brand||''))+'</small><b>Rp '+nn(mMB)+'</b><span>'+sh(mMB)+' dari total</span></div><div class="spk" data-spd="K:non" tabindex="0" role="button"><small>'+(isPB?'Non Private Brand':'Non master brand')+'</small><b>Rp '+nn(nonv)+'</b><span>'+sh(nonv)+' dari total</span></div><div class="tot spk" data-spd="K:all" tabindex="0" role="button"><small>Total penjualan</small><b>Rp '+nn(mA)+'</b><span>'+(mp==null?'':'Pencapaian '+pc(mp)+' dari budget')+'</span></div></div>'+cmp+'</div>';
  box.innerHTML='<div class="spv"><b style="font-size:18px">'+esc(e.name||e.id)+'</b> <span class="note">'+esc(e.id)+'</span>'+
    '<div class="spi"><div><small>Zonasi</small><b>'+esc(e.zona||'–')+'</b></div><div><small>Category</small><b>'+esc(e.cat||'–')+'</b></div><div><small>Master brand</small><b>'+esc(e.brand||'–')+'</b></div></div>'+
    '<div class="sps"><div><small>Achievement Daily</small><div class="note">'+(last?esc(fd(last,{weekday:'short',day:'numeric',month:'short'})):'Belum ada data')+'</div><div class="bg spk '+cls(dp)+'" data-spd="D" tabindex="0" role="button">'+pc(dp)+'</div><div class="note">'+(lw?'Actual Rp '+nn(da)+' dari budget Rp '+nn(e.tpe):'Hari non-shift (OFF/CT/RO dll), tidak ada budget'+(da?' · penjualan Rp '+nn(da)+' tetap dihitung di MTD':''))+'</div></div>'+
    '<div><small>Achievement MTD</small><div class="note">'+(last?'1 s.d. '+esc(fd(last,{day:'numeric',month:'short',year:'numeric'})):'Belum ada data')+'</div><div class="bg spk '+cls(mp)+'" data-spd="M" tabindex="0" role="button">'+pc(mp)+'</div><div class="note">Actual Rp '+nn(mA)+' dari budget Rp '+nn(mB)+'</div></div></div>'+
    mbc+'<div class="note">Target satu bulan penuh: Rp '+nn(monthB)+'</div><div class="spw"><table><thead><tr><th>Tanggal</th><th>Shift</th><th>Budget</th><th>Total Actual Sales</th><th>%</th><th>Invoice</th><th>Proteksi Qty</th>'+(noAp?'':'<th>Approval TL</th><th>Approval MOD</th>')+'</tr></thead><tbody>'+rows+'</tbody></table></div><p class="note" id="sp-msg">'+(noAp?'':(canAp?(spMgr?'Anda ('+spJab+') dapat mencentang Approval MOD. Approval TL diisi oleh Team Leader.':'Anda (Team Leader) dapat mencentang Approval TL. Approval MOD diisi oleh Store Manager / Asst Manager.')+' Centang lagi untuk membatalkan.':'Peran Anda hanya bisa melihat approval.'))+(spShift.length||spDate?' Tabel disaring ('+[spShift.length?'shift '+spShift.join(', '):'',spDate?fd(spDate,{day:"numeric",month:"short"}):''].filter(Boolean).join(', ')+'); baris TOTAL tetap MTD seluruh bulan.':'')+' Total Actual Sales = seluruh penjualan hari itu, termasuk hari OFF/CT/RO; budget hanya pada hari shift. % = Total Actual Sales ÷ Budget. Proteksi Qty = qty master brand BCA.</p></div>';
}
document.addEventListener('click',ev=>{ const t=ev.target; if(t&&t.closest&&t.closest('[data-spshreset]')){ spShift=[]; renderSP(true); return; } const bx=$('sp-shiftbox'); if(bx&&bx.open&&!(t.closest&&t.closest('#sp-shiftbox'))) bx.open=false; });
document.addEventListener('change',ev=>{ const id=ev.target&&ev.target.id; if(id==='sp-sel'){ spSel=ev.target.value; renderSP(); } else if(id==='sp-zona'){ spZ=ev.target.value; renderSP(); } else if(ev.target.classList&&ev.target.classList.contains('spshc')){ spShift=[...document.querySelectorAll('#sp-shiftl .spshc:checked')].map(x=>x.value); renderSP(true); } else if(id==='sp-date'){ spDate=ev.target.value; renderSP(true); } });

/* ---- Task Order Movement harian ---- */
let MV={d:'',data:null,key:'',ver:0,busy:false,err:''}, MVP=null, mvReady=false;
const mvName=(d,id)=>{ if(!id) return '(Tanpa salesperson)'; const q=days[d]&&(days[d].people||[]).find(x=>x.id===id); return (q&&q.name)||(SCHM[id]&&SCHM[id].name)||id; };
const mvShort=id=>id?String(id).replace(/^S10005-0*/,''):'–';
const mvQ=v=>fmtN.format(Math.round(v*100)/100);
function mvStat(L){ const s={n:0,q:0,tn:0,tq:0,dn:0,dq:0,an:0,unk:0}; L.forEach(l=>{ s.n++; s.q+=l[5]; if(l[7]){ s.tn++; s.tq+=l[5]; if((l[9]||0)>=l[5]) s.dn++; s.dq+=Math.min(l[5],l[9]||0); } else { s.an++; if(!l[8]) s.unk++; } }); return s; }
async function mvLoad(key){
  MV.key=key; MV.busy=true; MV.err=''; const d=MV.d;
  const { data, error } = await SB.rpc('mv_report',{p_store:D.code,p_date:d});
  MV.busy=false; if(MV.d!==d||MV.again){ MV.again=false; MV.key=''; mvRender(); return; }
  if(error||!data||!data.ok){ MV.data=null; MV.err='Data movement belum bisa dimuat'+(error?' ('+(error.message||error.code)+')':'')+'.'; MV.key=''; } else MV.data=data;
  mvDraw(); if(MVP&&!$('kpop').hidden) mvPopDraw();
}
function mvRender(){
  const card=$('mv-card'); if(!card||!mvReady) return;
  if(!SB||TIM||!dates.length){ card.hidden=true; return; }
  card.hidden=false;
  if(!MV.d||!dates.includes(MV.d)) MV.d=dates[dates.length-1];
  $('mv-date').innerHTML=dates.slice().reverse().map(x=>'<option value="'+x+'"'+(x===MV.d?' selected':'')+'>'+esc(fd(x,{weekday:'short',day:'numeric',month:'short'}))+'</option>').join('');
  const key=MV.d+'|'+MV.ver+'|'+((days[MV.d]&&days[MV.d].uploadedAt)||'');
  if(MV.key!==key&&!MV.busy){ MV.data=null; mvDraw(); mvLoad(key); } else { if(MV.busy&&MV.key!==key) MV.again=true; mvDraw(); }
}
function mvDraw(){
  const b=$('mv-body'); if(!b) return;
  if(!MV.data){ b.innerHTML='<p class="note">'+esc(MV.err||'Memuat data movement…')+'</p>'; return; }
  const R=MV.data, L=R.lines||[], S=mvStat(L), d=MV.d;
  const warn=[];
  if(!R.has_mv) warn.push('<b>Sales Detail tanggal ini belum memuat No Nota per item.</b> Upload ulang file Sales Detail tanggal ini supaya task order bisa dihitung.');
  if(!R.nav_n) warn.push('<b>File Nota Penjualan NAV belum pernah diupload.</b> Semua nota bernomor tidak bisa ditentukan Type-nya.');
  if(R.has_mv&&!R.moves_n) warn.push('<b>Report Stock Movement tanggal ini belum diupload.</b> Semua task order tampil belum movement.');
  if(S.unk) warn.push(S.unk+' item memakai No Nota yang tidak ada di file NAV (tidak dihitung sebagai task order). Upload file NAV terbaru.');
  const tile=(lbl,val,sub,cls,f)=>'<div class="'+(cls||'')+'" data-mvf="'+f+'" role="button" tabindex="0" title="Klik untuk melihat detail"><small>'+lbl+'</small><b>'+val+'</b><span>'+sub+'</span></div>';
  const belum=S.tn-S.dn;
  const tiles=tile('Item terjual',S.n,mvQ(S.q)+' qty','','all')+
    tile('Task order movement',S.tn,mvQ(S.tq)+' qty wajib movement','','task')+
    tile('Sudah movement',S.dn+' / '+S.tn,S.tn?Math.round(S.dn/S.tn*1000)/10+'% selesai':'tidak ada task',S.tn&&!belum?'apok':(S.tn?'':'apok'),'done')+
    tile('Belum movement',belum,mvQ(S.tq-S.dq)+' qty belum dipindah',belum?'apno':'apok','belum')+
    tile('Auto checklist',S.an,'type lain, tidak perlu movement','','auto');
  const g={}; L.forEach(l=>{ (g[l[0]]||(g[l[0]]=[])).push(l); });
  const rows=Object.keys(g).map(id=>{ const st=mvStat(g[id]); return {id,name:mvName(d,id),st,ok:st.dn===st.tn}; })
    .sort((a,b)=>(a.ok-b.ok)||String(a.name).localeCompare(String(b.name),'id'));
  const tb=rows.length?'<div class="spw" style="max-height:560px;overflow:auto"><table class="mvtb"><thead><tr><th>Sales ID</th><th>Nama sales</th><th class="rt">Item terjual</th><th>Movement</th></tr></thead><tbody>'+rows.map(r=>'<tr data-mvr="'+esc(r.id)+'" tabindex="0" title="Klik untuk melihat detail item"><td class="num">'+esc(mvShort(r.id))+'</td><td><b>'+esc(r.name)+'</b></td><td class="num rt">'+r.st.n+' <small>'+mvQ(r.st.q)+' qty</small></td><td class="'+(r.ok?'mvok':'mvno')+'">'+(r.st.tn===0?'✓ Done <small>tidak ada task (auto)</small>':r.ok?'✓ Done <small>'+r.st.dn+'/'+r.st.tn+' task</small>':'✕ Belum <small>'+r.st.dn+'/'+r.st.tn+' task · sisa '+mvQ(r.st.tq-r.st.dq)+' qty</small>')+'</td></tr>').join('')+'<tr class="tot"><td colspan="2"><b>TOTAL ('+rows.length+' orang)</b></td><td class="num rt"><b>'+S.n+'</b> <small>'+mvQ(S.q)+' qty</small></td><td><b>'+S.dn+'/'+S.tn+' task</b></td></tr></tbody></table></div>':'<p class="note">Tidak ada item terjual pada tanggal ini.</p>';
  b.innerHTML='<div class="spv">'+(warn.length?'<div class="spwarn">'+warn.map(w=>'<div>⚠ '+w+'</div>').join('')+'</div>':'')+'<div class="spsum mvsum">'+tiles+'</div></div>'+tb+'<p class="note">Wajib movement: Cash&amp;Carry via Pickup &amp; Delivery, Cash&amp;Carry via Gudang Store, dan Bawa Langsung (kolom NP No kosong). Type lain otomatis ✓. Qty movement dibagi ke salesperson berdasarkan urutan jam penjualan; item yang sama dijumlahkan per tanggal. Klik kartu atau nama untuk melihat detail item.'+(R.moves_n?' Stock Movement terakhir diupload '+esc(new Date(R.moves_at).toLocaleString('id-ID',{timeZone:'Asia/Jakarta',day:'numeric',month:'short',hour:'2-digit',minute:'2-digit'}))+' WIB.':'')+'</p>';
}
function mvPop(who,f){ MVP={who,f}; mvPopDraw(true); }
function mvPopDraw(open){
  if(!MV.data||!MVP) return; const d=MV.d, all=(MV.data.lines||[]).filter(l=>MVP.who===null||l[0]===MVP.who), S=mvStat(all);
  const flt={all:()=>true,task:l=>l[7],belum:l=>l[7]&&(l[9]||0)<l[5],done:l=>l[7]&&(l[9]||0)>=l[5],auto:l=>!l[7]}, f=flt[MVP.f]?MVP.f:'all';
  const cnt=k=>all.filter(flt[k]).length, sh=all.filter(flt[f]);
  sh.sort((a,b)=>(((b[7]&&(b[9]||0)<b[5])?1:0)-((a[7]&&(a[9]||0)<a[5])?1:0))||String(a[1]).localeCompare(String(b[1])));
  const ttl=MVP.who===null?'Task Order Movement · semua salesperson':mvName(d,MVP.who);
  if(open) kShow(ttl,[],'',null);
  $('kpop-t').textContent=ttl; $('kpop-per').textContent=fd(d,{weekday:'long',day:'numeric',month:'long',year:'numeric'})+' · item terjual vs Stock Movement';
  const info=[['Item terjual',S.n],['Task order',S.tn],['Sudah movement',S.dn+' / '+S.tn],['Belum movement',S.tn-S.dn],['Auto checklist',S.an]];
  $('kpop-info').innerHTML=info.map(i=>'<div><span>'+esc(i[0])+'</span><b>'+esc(String(i[1]))+'</b></div>').join('');
  const tabs=[['all','Semua'],['task','Task order'],['belum','Belum'],['done','Sudah'],['auto','Auto']];
  const ctl='<div class="spvc"><div class="spvt">'+tabs.map(t=>'<button type="button" data-mvt="'+t[0]+'"'+(f===t[0]?' class="on"':'')+'>'+t[1]+' <i>'+cnt(t[0])+'</i></button>').join('')+'</div></div>';
  const stc=l=>{ const q=l[5], a=l[9]||0; if(!l[7]) return l[8]?'<td class="spok">✓ Auto<small>'+esc(l[5]>0?l[6]:'')+'</small></td>':'<td class="mvwarn">? Type belum diketahui<small>NP tidak ada di NAV</small></td>'; return a>=q?'<td class="spok">✓ Done<small>'+mvQ(a)+' / '+mvQ(q)+'</small></td>':'<td class="spno">✕ '+(a>0?'Sebagian':'Belum')+'<small>'+mvQ(a)+' / '+mvQ(q)+'</small></td>'; };
  const body=sh.length?'<table class="pt spvtb"><thead><tr>'+(MVP.who===null?'<th>Salesperson</th>':'')+'<th>Item No</th><th>Description</th><th>No Nota</th><th>Type</th><th class="rt">Qty jual</th><th>Status movement</th></tr></thead><tbody>'+sh.map(l=>'<tr>'+(MVP.who===null?'<td>'+esc(mvName(d,l[0]))+'</td>':'')+'<td class="num">'+esc(l[1])+'</td><td>'+esc(l[2])+'</td><td class="num">'+esc(l[3]||'–')+'</td><td>'+esc(l[6])+'</td><td class="num rt">'+mvQ(l[5])+'</td>'+stc(l)+'</tr>').join('')+'</tbody></table>':'<p class="note" style="padding:14px 4px">Tidak ada item pada filter ini.</p>';
  $('kpop-body').innerHTML=ctl+body; if(open) $('kpop-body').scrollTop=0;
}
document.addEventListener('click',ev=>{ const t=ev.target.closest&&ev.target.closest('#mv-body [data-mvf], #mv-body tr[data-mvr], #kpop-body [data-mvt]'); if(!t) return;
  if(t.dataset.mvt){ if(MVP){ MVP.f=t.dataset.mvt; mvPopDraw(); } return; }
  if(!MV.data) return; if(t.dataset.mvf) mvPop(null,t.dataset.mvf); else mvPop(t.dataset.mvr,'all'); });
document.addEventListener('keydown',ev=>{ if(ev.key!=='Enter') return; const t=ev.target.closest&&ev.target.closest('#mv-body [data-mvf], #mv-body tr[data-mvr]'); if(!t||!MV.data) return; if(t.dataset.mvf) mvPop(null,t.dataset.mvf); else mvPop(t.dataset.mvr,'all'); });
document.addEventListener('change',ev=>{ if(ev.target&&ev.target.id==='mv-date'){ MV.d=ev.target.value; mvRender(); } });
function all(){ renderSub(); renderKPI(); renderRank(); renderProt(); renderTopBrand(); renderMPP(); renderMtd(); renderSP(); mvRender(); renderTable(); renderLow(); renderFoot(); }
document.querySelectorAll('[data-role]').forEach(b=>b.addEventListener('click',()=>{ S.role=b.dataset.role; document.querySelectorAll('[data-role]').forEach(x=>x.setAttribute('aria-pressed',x===b)); all(); }));
document.querySelectorAll('[data-m]').forEach(b=>b.addEventListener('click',()=>{ S.m=b.dataset.m; document.querySelectorAll('[data-m]').forEach(x=>x.setAttribute('aria-pressed',x===b)); renderRank(); }));
$('topn').addEventListener('change',e=>{ S.topn=parseInt(e.target.value,10); renderRank(); });
$('q').addEventListener('input',e=>{ S.q=e.target.value; renderTable(); });
// ---- Screenshot tabel detail (PNG untuk dibagikan) ----
function shotPages(){
  const rowsAll=lastL, hasT=lastHasT, sd=lastSd, PAGE=30, FONT="'Public Sans',system-ui,-apple-system,'Segoe UI',sans-serif", DISP="'Schibsted Grotesk','Public Sans',system-ui,sans-serif", padX=28;
  const q = S.q.trim();
  const f = []; f.push(S.role==='all'?'Semua peran':S.role==='Kasir'?'Kasir':'Sales'); if(S.zona) f.push('Zonasi '+S.zona);
  if(lastHasT && !S.schSel.length) f.push('Tanggal jadwal '+shortD(sd)); if(S.schSel.length) f.push('Jadwal '+schLabel()+' ('+shortD(sd)+')'); if(lastDr) f.push((S.dMode==='mtd'?'MTD':'Rentang tabel')+' '+lastM.days+' hari'); if(q) f.push('Cari "'+q+'"');
  const period = lastM.from? (lastM.from===lastM.to? fd(lastM.from) : fd(lastM.from,{day:'numeric',month:'short'})+' sampai '+fd(lastM.to)) : '';
  const cellsOf=p=>{ const cd=codeOn(p.id,sd); return hasT
      ? [p.rank==null?'–':p.rank, p.id, p.name, p.zona||'–', cd||'–', p.brand?title(p.brand):'–', rp(p.sales), p.brand?brandSales(p):null, p.brand?rp(nonBrand(p)):'–', p.target==null?'–':p.target===0?'tidak ada':rp(p.target), null]
      : [p.rank==null?'–':p.rank, p.id, p.name, p.cat?title(p.cat):'–', rp(p.sales), pct(p.sales/(lastM.total||1)), p.receipts]; };
  const mc=document.createElement('canvas').getContext('2d'), mw=(t,f)=>{ mc.font=f; return mc.measureText(String(t)).width; };
  const HF='700 15px '+FONT, bodyF=i=>(i===1?'400 18px ':(i===2?'500 24px ':'400 24px '))+FONT;
  const spec = hasT
    ? [['#','r',44,60],['Sales ID','l',150,230],['Salesperson','l',190,300],['Zonasi','l',92,130],['Jadwal '+shortD(sd),'l',92,150],['Master brand','l',140,220],['Penjualan (Rp)','r',130,230],['Penjualan master brand (Rp)','bt',150,240],['Non master brand (Rp)','r',130,230],['Target (Rp)','r',130,230],['Pencapaian','bar',190,190]]
    : [['#','r',44,60],['Sales ID','l',150,230],['Salesperson','l',190,300],['Kategori utama','l',180,300],['Penjualan (Rp)','r',130,230],['% total','r',100,160],['Struk','r',80,120]];
  const hl=h=>{ const t=h.toUpperCase(), w=t.split(' '); if(w.length<2||mw(t,HF)<=118) return [t]; const m=Math.ceil(w.length/2); return [w.slice(0,m).join(' '),w.slice(m).join(' ')]; };
  const rn=v=>typeof v==='string'? v.replace(/^Rp\s?/,'') : v, RI=hasT? [6,8,9] : [4];
  const allCells=rowsAll.map(p=>cellsOf(p).map((v,i)=>RI.includes(i)? rn(v) : v));
  const cols=spec.map((k,i)=>{ let w=Math.max(...hl(k[0]).map(t=>mw(t,HF)));
    allCells.forEach((c,r)=>{ if(k[1]==='bar') return; const v=c[i]; if(v==null) return; const t=k[1]==='bt'? rn(rp(v)) : v; w=Math.max(w,mw(t,bodyF(i))); });
    w=k[1]==='bar'? k[2] : Math.min(k[3],Math.max(k[2],Math.ceil(w)+24)); return {h:k[0],al:k[1],w,lines:hl(k[0])}; });
  let tw=cols.reduce((a,c)=>a+c.w,0); const W=Math.max(1320,Math.round(tw+padX*2)); const sc=(W-padX*2)/tw; cols.forEach(c=>c.w=c.w*sc); tw=W-padX*2;
  const pages=Math.max(1,Math.ceil(rowsAll.length/PAGE)), headH=124, thH=64, rowH=58, footH=64, out=[];
  for(let pg=0;pg<pages;pg++){
    const rows=rowsAll.slice(pg*PAGE,(pg+1)*PAGE), H=headH+thH+rowH*Math.max(rows.length,1)+footH, S2=Math.max(1,Math.min(2,12000/H,9000/W));
    const c=document.createElement('canvas'); c.width=Math.round(W*S2); c.height=Math.round(H*S2); const g=c.getContext('2d'); g.scale(S2,S2);
    g.fillStyle='#fff'; g.fillRect(0,0,W,H); g.textBaseline='middle';
    const fit=(t,w)=>{ t=String(t); if(g.measureText(t).width<=w) return t; while(t.length>1 && g.measureText(t+'…').width>w) t=t.slice(0,-1); return t+'…'; };
    g.fillStyle='#0b0b0b'; g.font='700 34px '+DISP; g.textAlign='left'; g.fillText('Pencapaian Salesperson',padX,36);
    if(pages>1){ g.textAlign='right'; g.font='700 22px '+FONT; g.fillStyle='#0f2547'; g.fillText('Halaman '+(pg+1)+' dari '+pages+' · nomor '+(pg*PAGE+1)+'–'+(pg*PAGE+rows.length),W-padX,36); g.textAlign='left'; }
    g.fillStyle='#52514e'; g.font='400 20px '+FONT;
    g.fillText(fit(D.store+' · '+period+' · penjualan tanpa PPN',W-padX*2),padX,74);
    g.fillStyle='#7a7975'; g.fillText(fit('Filter: '+f.join(' · ')+' · '+rowsAll.length+' orang',W-padX*2),padX,102);
    let y=headH; g.fillStyle='#0f2547'; g.fillRect(padX,y,tw,thH); g.fillStyle='#f5a800'; g.fillRect(padX,y+thH-3,tw,3);
    g.font=HF; g.fillStyle='#ffffff';
    const cx=i=>padX+cols.slice(0,i).reduce((a,x)=>a+x.w,0);
    cols.forEach((k,i)=>{ const x=cx(i), l=k.al==='l'; g.textAlign=l?'left':'right'; k.lines.forEach((t,j)=>g.fillText(t,l?x+12:x+k.w-12,y+(thH-3)/2+(j-(k.lines.length-1)/2)*18)); });
    y+=thH;
    rows.forEach((p,r)=>{
      const a=p.target>0? p.sales/p.target : -1, cells=allCells[pg*PAGE+r];
      if(r%2){ g.fillStyle='#fafafa'; g.fillRect(padX,y,tw,rowH); }
      g.fillStyle='#ecedeb'; g.fillRect(padX,y+rowH-1,tw,1);
      cols.forEach((k,i)=>{
        const x=cx(i), w=k.w, mid=y+rowH/2;
        if(k.al==='bar'){
          const bw=w-24-74, bx=x+12;
          if(a>=0){ g.fillStyle='#ecedeb'; g.fillRect(bx,mid-5,bw,10); g.fillStyle=a>=1?'#1f8a4c':'#2a78d6'; g.fillRect(bx,mid-5,Math.min(1,a)*bw,10); g.font='700 22px '+FONT; g.textAlign='right'; g.fillStyle=a>=1?'#1f8a4c':'#0b0b0b'; g.fillText(Math.round(a*100)+'%',x+w-12,mid); }
          else { g.font='400 24px '+FONT; g.textAlign='right'; g.fillStyle='#7a7975'; g.fillText('–',x+w-12,mid); }
          return;
        }
        if(k.al==='bt'){ g.textAlign='right'; if(cells[i]==null){ g.font=bodyF(i); g.fillStyle='#7a7975'; g.fillText('–',x+w-12,mid); } else { g.font=bodyF(i); g.fillStyle='#0b0b0b'; g.fillText(rn(rp(cells[i])),x+w-12,mid-9); g.font='400 15px '+FONT; g.fillStyle='#7a7975'; g.fillText(fit(title(p.brand),w-24),x+w-12,mid+13); } return; }
        const l=k.al==='l'; g.textAlign=l?'left':'right';
        g.font=bodyF(i); g.fillStyle=(i===0||cells[i]==='–'||cells[i]==='tidak ada')?'#7a7975':'#0b0b0b';
        g.fillText(fit(cells[i],w-24),l?x+12:x+w-12,mid);
      });
      y+=rowH;
    });
    if(!rows.length){ g.fillStyle='#7a7975'; g.textAlign='center'; g.font='400 22px '+FONT; g.fillText('Tidak ada orang yang cocok dengan filter.',W/2,y+rowH/2); y+=rowH; }
    g.fillStyle='#7a7975'; g.font='400 16px '+FONT; g.textAlign='left';
    const now=new Date().toLocaleString('id-ID',{day:'numeric',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit'});
    g.fillText('Dibuat '+now+(hasT?' · target harian dari file jadwal':'')+(pages>1?' · halaman '+(pg+1)+'/'+pages:''),padX,y+footH/2+4);
    out.push(c);
  }
  return out;
}
let toastT=0;
const shotMsg=(t,k)=>{ const s=$('shot-msg'); clearTimeout(toastT); s.textContent=t||''; s.hidden=!t; if(t&&k) toastT=setTimeout(()=>{ s.hidden=true; },4500); };
const shotName=()=>'print-screen-'+(lastM.from||'data')+(lastM.to&&lastM.to!==lastM.from?'_'+lastM.to:'')+(S.schSel.length?'_'+S.schSel.join('-').replace(/[^A-Za-z0-9-]/g,''):'')+'.png';
const toBlob=c=>new Promise((ok,no)=>c.toBlob(b=>b?ok(b):no(new Error('gagal')),'image/png'));
// ---- Gambar summary MPP
function mppCanvas(){
  const mem=P.filter(p=>p.hasSch && p.target>0), sets=[['Summary all salesperson, all staff (PB & eksternal)','#f5a800','#2a1d00',mem],['Summary all salesperson, Private Brand only','#2a9de0','#ffffff',mem.filter(p=>String(p.brand||'').toUpperCase()==='PRIVATE BRAND')]];
  const FONT="'Public Sans',system-ui,-apple-system,'Segoe UI',sans-serif", DISP="'Schibsted Grotesk','Public Sans',system-ui,sans-serif";
  const W=1200, px=32, S2=2, cols=[['Zona',56,'c'],['Category',280,'l'],['Actual\nMPP',90,'r'],['Salper\n≥100%',96,'r'],['Cont',70,'r'],['Salper\n75–99%',96,'r'],['Cont',70,'r'],['Salper\n<75%',96,'r'],['Cont',70,'r'],['IPD',70,'r'],['BZ',138,'r']];
  const pc=v=>Math.round(v*100)+'%', ip=v=>v?v.toFixed(2).replace('.',','):'0', bz=v=>v?fmtN.format(Math.round(v)):'0';
  const data=sets.map(s=>mppSummary(s[3]));
  const tblH=30+40+30*data[0].rows.length+34, H=100+tblH*2+16+44;
  const c=document.createElement('canvas'); c.width=W*S2; c.height=H*S2; const g=c.getContext('2d'); g.scale(S2,S2);
  g.fillStyle='#fff'; g.fillRect(0,0,W,H); g.textBaseline='middle';
  g.fillStyle='#0e1a2b'; g.font='700 22px '+DISP; g.textAlign='left'; g.fillText('Summary Pencapaian per MPP',px,30);
  const per = M.days? (M.days===1? 'Daily · '+fd(M.from,{day:'numeric',month:'long',year:'numeric'}) : fd(M.from,{day:'numeric',month:'short'})+' sampai '+fd(M.to,{day:'numeric',month:'short',year:'numeric'})+' ('+M.days+' hari)') : '';
  g.fillStyle='#4a576b'; g.font='400 13px '+FONT; g.fillText(D.store+' · '+per+' · berdasarkan achievement penjualan tanpa PPN',px,56);
  let y=76;
  sets.forEach((s,si)=>{
    const {rows,tot}=data[si];
    g.fillStyle=s[1]; g.fillRect(px,y,W-px*2,30); g.fillStyle=s[2]; g.font='700 14px '+FONT; g.textAlign='left'; g.fillText(s[0],px+12,y+15); y+=30;
    g.fillStyle='#0f2547'; g.fillRect(px,y,W-px*2,40);
    let x=px; cols.forEach(k=>{ g.fillStyle='#fff'; g.font='600 11.5px '+FONT; const ls=k[0].split('\n'); g.textAlign=k[2]==='l'?'left':k[2]==='c'?'center':'right'; const tx=k[2]==='l'?x+10:k[2]==='c'?x+k[1]/2:x+k[1]-10; ls.forEach((l,i)=>g.fillText(l,tx,y+20+(i-(ls.length-1)/2)*14)); x+=k[1]; });
    y+=40;
    const line=(r,isTot,label)=>{
      if(isTot){ g.fillStyle='#ffe29a'; g.fillRect(px,y,W-px*2,34); } else { g.fillStyle='#e3e8f1'; g.fillRect(px,y+29,W-px*2,1); }
      const h=isTot?34:30, mid=y+h/2;
      const vals=isTot?['Total','',r.n,r.hi,pc(r.hiP),r.mid,pc(r.midP),r.lo,pc(r.loP),ip(r.ipd),bz(r.bz)]:[r.z,r.cat,r.n,r.hi,pc(r.hiP),r.mid,pc(r.midP),r.lo,pc(r.loP),ip(r.ipd),bz(r.bz)];
      let x=px; cols.forEach((k,i)=>{
        g.font=(isTot?'700 ':'400 ')+'13px '+FONT; g.fillStyle='#0e1a2b';
        if(i===4&&!isTot) g.fillStyle='#0e6b49'; if(i===6&&!isTot) g.fillStyle='#a86a00'; if(i===8&&!isTot) g.fillStyle='#b4271c';
        g.textAlign=k[2]==='l'?'left':k[2]==='c'?'center':'right';
        if(isTot&&i===0){ g.textAlign='left'; g.fillText('Total',x+10,mid); }
        else if(!(isTot&&i===1)) g.fillText(String(vals[i]),k[2]==='l'?x+10:k[2]==='c'?x+k[1]/2:x+k[1]-10,mid);
        x+=k[1]; });
      y+=h; };
    rows.forEach(r=>line(r,false)); line(tot,true);
    y+=16;
  });
  g.fillStyle='#7a8699'; g.font='400 11px '+FONT; g.textAlign='left';
  g.fillText('Dibuat '+new Date().toLocaleString('id-ID',{day:'numeric',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit'})+' · IPD = rata-rata invoice per salesperson per hari kerja · BZ = rata-rata penjualan per struk',px,y+8);
  return c;
}
const mppBlob=async()=>{ try{ await document.fonts.ready; }catch(e){} return toBlob(mppCanvas()); };
const mppName=()=>'summary-mpp-'+(M.from||'data')+(M.to&&M.to!==M.from?'_'+M.to:'')+'.png';
// ---- Salin ke clipboard (untuk ditempel di WhatsApp) atau simpan sebagai file
async function saveBlob(blob,name){
  const dl = window.claude && window.claude.use ? await window.claude.use('downloads') : null;
  if(dl){ await dl.save({filename:name, data:blob}); return; }
  const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=name; document.body.appendChild(a); a.click(); a.remove(); setTimeout(()=>URL.revokeObjectURL(a.href),4000);
}
async function shotAction(make,name,mode,btns){
  btns.forEach(b=>b.disabled=true); shotMsg('Membuat gambar…');
  try{
    const bp=make();
    if(mode==='copy'){
      let copied=false;
      try{ if(navigator.clipboard&&window.ClipboardItem){ await navigator.clipboard.write([new ClipboardItem({'image/png':bp})]); copied=true; } }catch(e){ copied=false; }
      if(copied){ shotMsg('Gambar disalin. Buka WhatsApp lalu tempel dengan Ctrl+V (di ponsel: tekan lama lalu Tempel).','ok'); }
      else {
        const blob=await bp;
        try{ await saveBlob(blob,name); shotMsg('Browser tidak mengizinkan menyalin, jadi gambar disimpan sebagai file. Kirim file itu di WhatsApp.','ok'); }
        catch(e){ shotMsg(e&&e.code==='declined'?'Dibatalkan.':'Gambar tidak bisa disalin atau disimpan di sini.','warn'); }
      }
    } else {
      try{ await saveBlob(await bp,name); shotMsg('Gambar disimpan.','ok'); }
      catch(e){ shotMsg(e&&e.code==='declined'?'Dibatalkan.':'Penyimpanan gambar tidak tersedia di sini.','warn'); }
    }
  }catch(e){ shotMsg('Gambar gagal dibuat.','warn'); }
  btns.forEach(b=>b.disabled=false);
}
const bDet=[$('ps'),$('ps-dl')], bMpp=[$('mpp-copy'),$('mpp-dl')];
const shotPageBlobs=async()=>{ try{ await document.fonts.ready; }catch(e){} return Promise.all(shotPages().map(toBlob)); };
const shotBlob=async()=>(await shotPageBlobs())[0];
const shotPgName=(nm,i,n)=>n>1? nm+'-hal'+(i+1) : nm;
let shotURLs=[];
function shotPreview(arr,nm,meta){
  shotURLs.forEach(u=>URL.revokeObjectURL(u)); shotURLs=arr.map(b=>URL.createObjectURL(b));
  const html='<p class="note" style="margin:0 0 12px;font-size:14px">Klik <b>Salin halaman</b>, lalu tempel (Ctrl+V) di WhatsApp. Ulangi untuk halaman berikutnya. Tidak perlu mengunduh.</p>'+arr.map((b,i)=>'<figure style="margin:0 0 16px"><figcaption style="font-weight:700;margin:0 0 6px;display:flex;gap:8px;align-items:center;flex-wrap:wrap">Halaman '+(i+1)+' dari '+arr.length+'<button type="button" class="btn primary sm" data-shot="copy" data-i="'+i+'">Salin halaman '+(i+1)+'</button><button type="button" class="btn sm" data-shot="dl" data-i="'+i+'" title="Simpan sebagai file jika salin tidak berhasil">Simpan file</button></figcaption><img src="'+shotURLs[i]+'" alt="Print Screen halaman '+(i+1)+'" style="width:100%;border:1px solid var(--line);border-radius:8px;display:block"></figure>').join('')+'<p class="note" style="font-size:13px">Jika salin tidak berhasil di perangkat Anda: <button type="button" class="btn sm" data-shot="all">Simpan semua halaman sebagai file</button></p>';
  kShow('Print Screen · '+arr.length+' halaman',[['Jumlah halaman',String(arr.length)],[meta?meta.lbl:'Orang',String(meta?meta.n:lastL.length)],['Per halaman',meta?'30 baris':'30 orang']],html,(meta&&meta.from)||$('ps'));
  shotLast={arr,nm};
}
let shotLast=null;
$('kpop-body').addEventListener('click',async e=>{ const b=e.target.closest('[data-shot]'); if(!b||!shotLast) return; const {arr,nm}=shotLast, i=+b.dataset.i;
  try{
    if(b.dataset.shot==='all'){ for(let k=0;k<arr.length;k++) await saveBlob(arr[k],shotPgName(nm,k,arr.length)+'.png'); shotMsg('Semua halaman disimpan.','ok'); }
    else if(b.dataset.shot==='dl'){ await saveBlob(arr[i],shotPgName(nm,i,arr.length)+'.png'); shotMsg('Halaman '+(i+1)+' disimpan.','ok'); }
    else { await navigator.clipboard.write([new ClipboardItem({'image/png':arr[i]})]); b.textContent='✓ Halaman '+(i+1)+' tersalin'; shotMsg('Halaman '+(i+1)+' disalin. Tempel di WhatsApp dengan Ctrl+V.','ok'); }
  }catch(x){ shotMsg(x&&x.code==='declined'?'Dibatalkan.':'Tidak bisa menyalin di sini. Gunakan tombol Unduh.','warn'); }
});
async function shotDetail(mode){
  bDet.forEach(b=>b.disabled=true); shotMsg('Membuat gambar…');
  const pp=shotPageBlobs(), nm=shotName();
  try{
    let copied=false;
    if(mode==='copy'){ try{ if(navigator.clipboard&&window.ClipboardItem){ await navigator.clipboard.write([new ClipboardItem({'image/png':pp.then(a=>a[0])})]); copied=true; } }catch(e){ copied=false; } }
    const arr=await pp, n=arr.length;
    if(mode==='copy'&&copied){
      if(n>1) shotPreview(arr,nm);
      shotMsg(n>1? 'Halaman 1 dari '+n+' sudah disalin. Tempel di WhatsApp, lalu salin halaman berikutnya dari jendela yang terbuka.' : 'Gambar disalin. Buka WhatsApp lalu tempel dengan Ctrl+V (di ponsel: tekan lama lalu Tempel).','ok');
    } else {
      try{ for(let i=0;i<n;i++) await saveBlob(arr[i],shotPgName(nm,i,n)+'.png');
        shotMsg(mode==='copy'? 'Browser tidak mengizinkan menyalin, jadi '+(n>1?n+' halaman disimpan sebagai file':'gambar disimpan sebagai file')+'. Kirim file itu di WhatsApp.' : (n>1? n+' halaman disimpan.' : 'Gambar disimpan.'),'ok'); }
      catch(e){ shotMsg(e&&e.code==='declined'?'Dibatalkan.':'Penyimpanan gambar tidak tersedia di sini.','warn'); }
      if(n>1&&mode==='copy') shotPreview(arr,nm);
    }
  }catch(e){ shotMsg('Gambar gagal dibuat.','warn'); }
  bDet.forEach(b=>b.disabled=false);
}
$('ps').addEventListener('click',()=>shotDetail('copy'));
$('ps-dl').addEventListener('click',()=>shotDetail('save'));
$('mpp-copy').addEventListener('click',()=>shotAction(mppBlob,mppName(),'copy',bMpp));
$('mpp-dl').addEventListener('click',()=>shotAction(mppBlob,mppName(),'save',bMpp));
// ---- Gambar summary MTD (dua blok ditumpuk agar tetap terbaca di WhatsApp)
function mtdCanvas(){
  const sets=mtdSets().map((s,i)=>[s[0].replace(/\s+/g,' '),i?'#1b9ad2':'#b8860b',mtdAgg(s[2])]);
  const FONT="'Public Sans',system-ui,-apple-system,'Segoe UI',sans-serif", DISP="'Schibsted Grotesk','Public Sans',system-ui,sans-serif";
  const px=28, S2=2, cols=[['No',44,'c'],['Store Name',230,'l'],['Total\nSalper',86,'c']];
  MB.forEach(b=>{ cols.push([b.h.replace(' sd ','\nsd\n'),92,'c']); cols.push(['Cont',78,'c']); });
  const W=px*2+cols.reduce((a,k)=>a+k[1],0), blockH=34+62+38+40+22, H=84+blockH*2+30;
  const c=document.createElement('canvas'); c.width=W*S2; c.height=H*S2; const g=c.getContext('2d'); g.scale(S2,S2);
  g.fillStyle='#fff'; g.fillRect(0,0,W,H); g.textBaseline='middle';
  g.fillStyle='#0e1a2b'; g.font='700 24px '+DISP; g.textAlign='left'; g.fillText('Summary MTD Salesperson',px,30);
  const per = M.days? (M.days===1? 'Daily · '+fd(M.from,{day:'numeric',month:'long',year:'numeric'}) : fd(M.from,{day:'numeric',month:'short'})+' sampai '+fd(M.to,{day:'numeric',month:'short',year:'numeric'})+' ('+M.days+' hari)') : '';
  g.fillStyle='#4a576b'; g.font='400 14px '+FONT; g.fillText(MTD_STORE+' · '+per+' · achievement penjualan terhadap target',px,58);
  let y=78;
  const al=k=>k[2]==='l'?'left':'center', tx=(k,x)=>k[2]==='l'?x+12:x+k[1]/2;
  sets.forEach(s=>{
    g.fillStyle=s[1]; g.fillRect(px,y,W-px*2,34); g.fillStyle='#fff'; g.font='700 16px '+FONT; g.textAlign='center'; g.fillText(s[0],W/2,y+17); y+=34;
    let x=px; cols.forEach((k,i)=>{
      g.fillStyle= i<2?'#0f2547' : i===2? s[1] : (i%2===1?'#e4e7ec':'#f3e27a'); g.fillRect(x,y,k[1],62);
      g.fillStyle= (i<3)?'#fff':'#0e1a2b'; g.font='700 13px '+FONT; g.textAlign=al(k);
      const ls=k[0].split('\n'), lh=15, y0=y+31-(ls.length-1)*lh/2; ls.forEach((t,j)=>g.fillText(t,tx(k,x),y0+j*lh)); x+=k[1]; });
    y+=62;
    const a=s[2];
    [['1',false],['TOTAL',true]].forEach(([lab,isTot])=>{
      const h=isTot?40:38; x=px;
      const vals=[isTot?'TOTAL':lab, isTot?'':MTD_STORE, a.n]; MB.forEach((b,i)=>{ vals.push(a.c[i]); vals.push(mtdPc(a.pc[i])); });
      cols.forEach((k,i)=>{
        g.fillStyle= isTot? (i<2?'#0f2547': (i>2&&i%2===0?'#f3e27a':'#ffe29a')) : (i>2&&i%2===0?'#fbf3c4':(i===2?'#f3e9c8':'#fff')); g.fillRect(x,y,k[1],h);
        g.fillStyle= (isTot&&i<2)?'#fff':'#0e1a2b'; g.font=((isTot||i===2)?'700 ':'400 ')+'15px '+FONT; g.textAlign=al(k);
        if(isTot&&i===1){ g.textAlign='left'; g.fillText('TOTAL',px+12,y+h/2); }
        else if(!(isTot&&i===0)) g.fillText(String(vals[i]),tx(k,x),y+h/2);
        x+=k[1]; });
      g.fillStyle='#d5d9e2'; g.fillRect(px,y+h-1,W-px*2,1); y+=h; });
    y+=22;
  });
  g.fillStyle='#7a8699'; g.font='400 12px '+FONT; g.textAlign='left';
  g.fillText('Dibuat '+new Date().toLocaleString('id-ID',{day:'numeric',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit'})+' · batas bawah kelompok: ≥100%, 95%, 85%, 75%, 50% · Cont = porsi terhadap total salper',px,y+4);
  return c;
}
const mtdBlob=async()=>{ try{ await document.fonts.ready; }catch(e){} return toBlob(mtdCanvas()); };
const mtdName=()=>'summary-mtd-'+(M.from||'data')+(M.to&&M.to!==M.from?'_'+M.to:'')+'.png';
const bMtd=[$('mtd-copy'),$('mtd-dl')];
$('mtd-copy').addEventListener('click',()=>shotAction(mtdBlob,mtdName(),'copy',bMtd));
$('mtd-dl').addEventListener('click',()=>shotAction(mtdBlob,mtdName(),'save',bMtd));
// ---- Print Screen untuk jendela daftar salesperson (30 baris per halaman, header sama di tiap halaman)
function lstPages(ctx){
  const FONT="'Public Sans',system-ui,-apple-system,'Segoe UI',sans-serif", DISP="'Schibsted Grotesk','Public Sans',system-ui,sans-serif";
  const cols=[['#',54,'c'],['Sales ID',240,'l'],['Salesperson',262,'l'],['Zonasi',100,'l'],['Master brand',180,'l'],['Target (Rp)',140,'r'],['Penjualan (Rp)',150,'r'],['Achievement',136,'r'],['Kelompok',160,'l']];
  const px=28, S2=2, PER=30, W=px*2+cols.reduce((a,k)=>a+k[1],0), pages=Math.max(1,Math.ceil(ctx.rows.length/PER));
  const chipH=56, headY=96+chipH+14, th=46, rh=38, H=headY+th+rh*Math.min(PER,Math.max(ctx.rows.length,1))+54;
  const per = M.days? (M.days===1? 'Daily · '+fd(M.from,{day:'numeric',month:'long',year:'numeric'}) : fd(M.from,{day:'numeric',month:'short'})+' sampai '+fd(M.to,{day:'numeric',month:'short',year:'numeric'})+' ('+M.days+' hari)') : '';
  const out=[];
  for(let pg=0;pg<pages;pg++){
    const rows=ctx.rows.slice(pg*PER,pg*PER+PER), hh=headY+th+rh*rows.length+54;
    const c=document.createElement('canvas'); c.width=W*S2; c.height=hh*S2; const g=c.getContext('2d'); g.scale(S2,S2);
    g.fillStyle='#fff'; g.fillRect(0,0,W,hh); g.textBaseline='middle'; g.textAlign='left';
    g.fillStyle='#0e1a2b'; g.font='700 25px '+DISP; g.fillText(ctx.title+(pages>1?'  ·  halaman '+(pg+1)+'/'+pages:''),px,32);
    g.fillStyle='#4a576b'; g.font='400 15px '+FONT; g.fillText(per+' · sales exclude PPN',px,62);
    const cw=Math.min(170,(W-px*2-(ctx.info.length-1)*10)/ctx.info.length); let cx=px;
    ctx.info.forEach(([lb,v])=>{ g.fillStyle='#f1f3f8'; g.fillRect(cx,84,cw,chipH); g.fillStyle='#f0a500'; g.fillRect(cx,84,4,chipH);
      g.fillStyle='#5b6678'; g.font='600 12px '+FONT; g.fillText(String(lb).toUpperCase(),cx+14,84+18); g.fillStyle='#0e1a2b'; g.font='700 22px '+FONT; g.fillText(String(v),cx+14,84+40); cx+=cw+10; });
    g.fillStyle='#0f2547'; g.fillRect(px,headY,W-px*2,th);
    let x=px; cols.forEach(k=>{ g.fillStyle='#fff'; g.font='700 14px '+FONT; g.textAlign=k[2]==='l'?'left':k[2]==='c'?'center':'right'; g.fillText(k[0].toUpperCase(),k[2]==='l'?x+10:k[2]==='c'?x+k[1]/2:x+k[1]-12,headY+th/2); x+=k[1]; });
    const colr={gd:'#0e6b49',am:'#a86a00',bd:'#b4271c'};
    rows.forEach((r,i)=>{ const y=headY+th+i*rh; if(i%2) { g.fillStyle='#f6f8fb'; g.fillRect(px,y,W-px*2,rh); } g.fillStyle='#e3e8f1'; g.fillRect(px,y+rh-1,W-px*2,1);
      let x2=px; cols.forEach((k,j)=>{ g.font=(j===2?'600 ':'400 ')+'17px '+FONT; g.fillStyle=(j===7||j===8)?(colr[r[9]]||'#0e1a2b'):(j===0?'#7a8699':'#0e1a2b'); if(j===7) g.font='700 17px '+FONT;
        g.textAlign=k[2]==='l'?'left':k[2]==='c'?'center':'right'; let t=String(r[j]); const mw=k[1]-22; while(t.length>3&&g.measureText(t).width>mw) t=t.slice(0,-2)+'…';
        g.fillText(t,k[2]==='l'?x2+10:k[2]==='c'?x2+k[1]/2:x2+k[1]-12,y+rh/2); x2+=k[1]; }); });
    g.fillStyle='#7a8699'; g.font='400 13px '+FONT; g.textAlign='left';
    g.fillText('Dibuat '+new Date().toLocaleString('id-ID',{day:'numeric',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit'})+' · urut dari capaian tertinggi'+(pages>1?' · halaman '+(pg+1)+'/'+pages:''),px,headY+th+rh*rows.length+28);
    out.push(c);
  }
  return out;
}
async function lstShot(mode){
  const ctx=lstCtx; if(!ctx) return; const btns=[...document.querySelectorAll('#kpop-body [data-lst]')]; btns.forEach(b=>b.disabled=true); shotMsg('Membuat gambar…');
  const pp=(async()=>{ try{ await document.fonts.ready; }catch(e){} return Promise.all(lstPages(ctx).map(toBlob)); })(), nm=ctx.name, meta={n:ctx.rows.length,lbl:'Salesperson',from:ctx.from};
  try{
    let copied=false;
    if(mode==='copy'){ try{ if(navigator.clipboard&&window.ClipboardItem){ await navigator.clipboard.write([new ClipboardItem({'image/png':pp.then(a=>a[0])})]); copied=true; } }catch(e){ copied=false; } }
    const arr=await pp, n=arr.length;
    if(mode==='copy'&&copied){
      if(n>1) shotPreview(arr,nm,meta);
      shotMsg(n>1? 'Halaman 1 dari '+n+' sudah disalin. Tempel di WhatsApp, lalu salin halaman berikutnya dari jendela yang terbuka.' : 'Gambar disalin. Buka WhatsApp lalu tempel dengan Ctrl+V (di ponsel: tekan lama lalu Tempel).','ok');
    } else {
      try{ for(let i=0;i<n;i++) await saveBlob(arr[i],shotPgName(nm,i,n)+'.png');
        shotMsg(mode==='copy'? 'Browser tidak mengizinkan menyalin, jadi '+(n>1?n+' halaman disimpan sebagai file':'gambar disimpan sebagai file')+'. Kirim file itu di WhatsApp.' : (n>1? n+' halaman disimpan.' : 'Gambar disimpan.'),'ok'); }
      catch(e){ shotMsg(e&&e.code==='declined'?'Dibatalkan.':'Penyimpanan gambar tidak tersedia di sini.','warn'); }
      if(n>1&&mode==='copy') shotPreview(arr,nm,meta);
    }
  }catch(e){ shotMsg('Gambar gagal dibuat.','warn'); }
  btns.forEach(b=>b.disabled=false);
}
$('kpop-body').addEventListener('click',e=>{ const b=e.target.closest('[data-lst]'); if(b) lstShot(b.dataset.lst); });



let canWrite=false;
const dF=$('d-from'), dT=$('d-to');
let dbOverlay={}, localOverlay={}, dbRef=null, dbReady=false;
let dbSchedDoc=null, localSched=null, dbSchedRef=null;
const setStatus=(msg,kind)=>{ const s=$('status'); s.textContent=msg||''; s.className='status'+(kind?' '+kind:''); };
let staleCount=0;
function effectiveDays(){
  const out={}; staleCount=0; for(const d in D.days) out[d]=D.days[d];
  for(const src of [dbOverlay,localOverlay]) for(const d in src){
    if(src[d].deleted) delete out[d];
    else if(src[d].basis==='excl2') out[d]=src[d];
    else staleCount++;
  }
  return out;
}
function effectiveSched(){
  if(localSched) return localSched;
  if(dbSchedDoc && dbSchedDoc.basis==='sched1' && Array.isArray(dbSchedDoc.people)) return dbSchedDoc;
  return D.sched || null;
}
function setRange(a,b){ if(a>b) b=a; dF.value=a; dT.value=b; aggregate(a,b); all(); }
function rebuild(fullRange){
  const pf=dF.value, pt=dT.value;
  const oldLast = dates[dates.length-1];
  const wasAll = !pf || (pf===dates[0] && pt===oldLast);
  const followEnd = !!pt && pt===oldLast;
  days=effectiveDays(); dates=Object.keys(days).sort();
  setSched(effectiveSched());
  [dF,dT].forEach(el=>{ el.min=dates[0]||''; el.max=dates[dates.length-1]||''; });
  renderDataBar();
  if(staleCount && !$('status').textContent) setStatus(staleCount+' hari data tersimpan memakai format lama (belum ada rincian master brand) dan diabaikan. Unggah ulang file untuk menggantinya.','warn');
  if(!dates.length){ dF.value=''; dT.value=''; aggregate('9','0'); all(); return; }
  const first=dates[0], last=dates[dates.length-1];
  if(fullRange||wasAll) setRange(first,last);
  else setRange(pf<first?first:pf, followEnd?last:(pt>last?last:pt));
}
/* Checklist database wajib unggah + template standar (format sama dengan halaman Monitoring) */
const CKOPEN={};
const CKMON=['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];
const ckLab=d=>(+d.slice(8))+' '+CKMON[+d.slice(5,7)-1];
const ckDates=a=>a.length? a.slice(0,12).map(ckLab).join(', ')+(a.length>12?' dan '+(a.length-12)+' tanggal lain':'') : '';
function ckPeriod(){
  const z=n=>String(n).padStart(2,'0'), t=new Date(), today=t.getFullYear()+'-'+z(t.getMonth()+1)+'-'+z(t.getDate());
  const ref=$('d-to').value||dates[dates.length-1]||today;
  return {y:+ref.slice(0,4),m:+ref.slice(5,7),day:+ref.slice(8,10),ref};
}
function ckItems(){
  const P=ckPeriod(), z=n=>String(n).padStart(2,'0'), ym=P.y+'-'+z(P.m), exp=[]; for(let i=1;i<=P.day;i++) exp.push(ym+'-'+z(i));
  const sd=new Set((SCH&&SCH.dates)||[]), hasS=!!(SCH&&SCH.people&&SCH.people.length);
  const out=[], mk=(id,title,file,bad,warn,okTxt,badTxt,lines,tpl)=>out.push({id,title,file,state:bad.length?'bad':(warn&&warn.length?'warn':'ok'),summary:bad.length?badTxt:(warn&&warn.length?warn[0]:okTxt),lines:lines.filter(Boolean),tpl});
  const miss=exp.filter(d=>!days[d]);
  const nojam=exp.filter(d=>days[d]&&!(days[d].hours&&days[d].hours.length));
  mk('sales','Sales Detail Report (per transaksi)','File: Sales Detail Report (kolom Receipt Date)',miss,nojam.length?[nojam.length+' hari belum ada rincian jam transaksi']:null,'Lengkap sampai '+ckLab(P.ref),miss.length+' tanggal belum ada',
    [miss.length?'Belum ada: '+ckDates(miss):'Semua tanggal 1 sampai '+ckLab(P.ref)+' sudah ada.',nojam.length?'Belum ada rincian jam: '+ckDates(nojam):'','Isi: penjualan exclude PPN per salesperson, struk, dan jam transaksi. Tarik setelah hari selesai.'],['sales']);
  const old=(f)=>exp.filter(d=>days[d]&&!f(days[d]));
  const mbOk=x=>!!(x.mb&&Object.keys(x.mb).length)&&Array.isArray(x.prot), itOk=x=>(x.people||[]).some(p=>Array.isArray(p.it)&&p.it.length);
  const mbOld=old(mbOk), mbBad=miss.concat(mbOld).sort();
  mk('mb','Master brand dan Mitra10 Proteksi (BCA)','File: Sales Detail Report (kolom Master Brand dan Item)',mbBad,null,'Lengkap sampai '+ckLab(P.ref),mbBad.length+' tanggal belum ada atau perlu unggah ulang',
    [miss.length?'Belum ada: '+ckDates(miss):'',mbOld.length?'Perlu diunggah ulang (data lama belum memuat master brand atau proteksi): '+ckDates(mbOld):'',!mbBad.length?'Semua tanggal memuat master brand dan proteksi.':'','Isi: penjualan per master brand dan qty serta value Mitra10 Proteksi per salesperson.'],['sales']);
  const itOld=old(itOk), itBad=miss.concat(itOld).sort();
  mk('item','Detail item per salesperson (pop up)','File: Sales Detail Report (kolom Item No, Item Name, Qty)',itBad,null,'Lengkap sampai '+ckLab(P.ref),itBad.length+' tanggal belum ada atau perlu unggah ulang',
    [miss.length?'Belum ada: '+ckDates(miss):'',itOld.length?'Perlu diunggah ulang (data lama belum memuat detail item): '+ckDates(itOld):'',!itBad.length?'Semua tanggal memuat detail item.':'','Isi: category, master brand, no item, deskripsi, qty, dan value per salesperson.'],['sales']);
  const ms=hasS? exp.filter(d=>!sd.has(d)) : exp;
  mk('sched','Target dan daftar salesperson (admin, bulanan)','File: jadwal & target (kolom SALES ID, TARGET/DAY, lalu kolom tanggal)',ms,null,hasS? SCH.people.length+' orang, sampai '+ckLab(SCH.dates[SCH.dates.length-1]) : '',hasS? ms.length+' tanggal belum ada di jadwal' : 'Belum ada file jadwal',
    [hasS? (ms.length?'Belum ada di jadwal: '+ckDates(ms):'Semua tanggal 1 sampai '+ckLab(P.ref)+' ada di jadwal.') : 'Belum ada jadwal & target.','Isi: kode jadwal harian (S/P/M/OFF dst), zonasi, master brand, category, TARGET/DAY dan TARGET/MONTH.'],['sched']);
  return out;
}
function abItems(){
  const P=ckPeriod(), z=n=>String(n).padStart(2,'0'), out=[];
  const mk=(id,title,file,bad,okTxt,badTxt,lines,tpl)=>out.push({id,title,file,state:bad.length?'bad':'ok',summary:bad.length?badTxt:okTxt,lines:lines.filter(Boolean),tpl});
  const exp=[]; for(let i=1;i<=P.day;i++) exp.push(P.y+'-'+z(P.m)+'-'+z(i));
  const sd=new Set((SCH&&SCH.dates)||[]), covered=exp.length&&exp.every(d=>sd.has(d)), lastI=covered?SCH.dates.indexOf(exp[exp.length-1]):-1;
  const zk=new Set(); ((SCH&&SCH.people)||[]).forEach(e=>{ if(e.zona) zk.add(normZona(e.zona)); }); Object.keys((SCH&&SCH.zones)||{}).forEach(k=>zk.add(k));
  const zs=[...zk].sort((a,b)=>(parseInt(a.replace(/\D/g,''))||0)-(parseInt(b.replace(/\D/g,''))||0)||a.localeCompare(b)), ymk=P.y+'-'+z(P.m);
  zs.forEach(zn=>{
    const zi=(SCH&&SCH.zones&&SCH.zones[zn])||null, up=!!(zi&&zi.month===ymk), ppl=((SCH&&SCH.people)||[]).filter(e=>normZona(e.zona||'')===zn);
    // Zona yang jadwalnya sudah ada di file jadwal & target untuk periode ini dianggap sudah input
    const filled=lastI>=0&&ppl.length&&ppl.filter(e=>String((e.c||[])[lastI]||'').trim()!=='').length>=Math.ceil(ppl.length/2);
    const okM=up||filled;
    const when=zi&&zi.uploadedAt? new Date(zi.uploadedAt).toLocaleString('id-ID',{day:'numeric',month:'short',hour:'2-digit',minute:'2-digit'}) : '';
    const ml=zi&&zi.month? CKMON[+zi.month.slice(5,7)-1]+' '+zi.month.slice(0,4) : '';
    mk('z:'+zn,'Absensi '+zn+' (TL zonasi)','File: ABSENSI '+zn+' (kolom SALESPERSON ID, NAMA, BRAND, ZONA, lalu tanggal 1 sampai akhir bulan)',okM?[]:[1],
      up? 'Diunggah TL '+when+' · '+zi.n+' orang' : 'Sudah ada di jadwal (dianggap sudah input)',
      zi&&zi.month? 'Terakhir untuk '+ml+', belum ada untuk '+CKMON[P.m-1]+' '+P.y : 'Belum ada jadwal '+CKMON[P.m-1]+' '+P.y,
      [up? 'Sumber: '+(zi.source||'-')+' ('+zi.n+' orang). '+ppl.length+' orang tercatat di '+zn+'.' : (okM? ppl.length+' orang di '+zn+' sudah punya jadwal sampai '+ckLab(P.ref)+'. TL boleh memperbarui dengan format ABSENSI.' : 'Jadwal '+zn+' untuk '+CKMON[P.m-1]+' '+P.y+' belum ada. TL zonasi mengunggah file ABSENSI '+zn+'.'),
       'Format baku: baris 1 = SALESPERSON ID, NAMA, BRAND, ZONA, 1, 2, 3 … sampai akhir bulan. Nama sheet berisi bulan dan tahun (contoh NOVEMBER 2026). Kode: '+ZKODE.join(', ')+'. File yang tidak sesuai ditolak beserta alasannya.'],['zona:'+zn]);
  });
  return out;
}
function tplSheet(kind,y,m){
  const z=n=>String(n).padStart(2,'0'), dim=new Date(y,m,0).getDate(), ds=[...Array(dim)].map((_,i)=>y+'-'+z(m)+'-'+z(i+1));
  if(kind.startsWith('zona:')){
    const zn=kind.slice(5), idx=new Map(((SCH&&SCH.dates)||[]).map((d,i)=>[d,i]));
    const ppl=((SCH&&SCH.people)||[]).filter(e=>normZona(e.zona||'')===zn);
    return {name:ZBLN[m-1]+' '+y, rows:[['SALESPERSON ID','NAMA','BRAND','ZONA',...ds.map((_,i)=>i+1)],...ppl.map(e=>[e.id,e.name,e.brand||'',zn,...ds.map(d=>idx.has(d)?(e.c[idx.get(d)]||''):'')])]};
  }
  if(kind==='sales') return {name:'Sales Detail',rows:[['Receipt Date','Receipt No','Store No','Time','Salescode ID','Salescode Name','Category','Master Brand','Master Type Brand','Item No','Item Name','Qty','Amount Include Sales','Dicount Amount','Amount Exclude Sales']]};
  return {name:'Jadwal '+CKMON[m-1]+' '+y,rows:[['No','Sales ID','Nama','Master Brand','Category','Zonasi','Target/Month','Target/Day',...ds]]};
}
const TPLGUIDE=[['PETUNJUK TEMPLATE UNGGAH DASHBOARD PENCAPAIAN SALESPERSON'],[''],['1. Isi hanya sheet data. Jangan ubah nama kolom di baris judul.'],['2. Sheet Petunjuk ini dilewati saat diunggah (hanya sheet pertama yang dibaca).'],['3. Sales Detail: tempel data laporan Sales Detail dari sistem di bawah baris judul. Kolom wajib: Receipt Date, Receipt No, Qty, Amount Exclude Sales, Salescode ID, Salescode Name. Store No harus 10005.'],['4. Jadwal: satu baris per salesperson. Isi kode jadwal harian (S, P, M, OFF, dst) pada kolom tanggal, lalu TARGET/DAY dan TARGET/MONTH.'],['5. Unggah lewat tombol Unggah file. Tanggal yang sudah ada akan diganti.']];
async function dlTpl(kinds){
  if(!window.XLSX){ setStatus('Pembaca Excel belum termuat. Coba lagi sebentar.','err'); return; }
  const P=ckPeriod(), wb=XLSX.utils.book_new();
  kinds.forEach(k=>{ const t=tplSheet(k,P.y,P.m); XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(t.rows),t.name); });
  if(!kinds[0].startsWith('zona:')) XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(TPLGUIDE),'Petunjuk');
  const out=XLSX.write(wb,{type:'array',bookType:'xlsx'});
  const blob=new Blob([out],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  const nm={sales:'sales-detail',sched:'jadwal'}, k0=kinds[0], zf=k0.startsWith('zona:')? 'jadwal-'+k0.slice(5).toLowerCase().replace(/\s+/g,'-') : null;
  try{ await saveBlob(blob,(zf? zf : 'template-'+(kinds.length>1?'semua':nm[k0]))+'-'+P.y+'-'+String(P.m).padStart(2,'0')+'.xlsx'); }catch(e){ setStatus('Template gagal diunduh.','warn'); }
}
function renderCk(pf,it,label,sub){
  if(!$(pf+'Btn')) return;
  const nOk=it.filter(x=>x.state==='ok').length;
  const worst=it.some(x=>x.state==='bad')?'bad':it.some(x=>x.state==='warn')?'warn':'ok';
  $(pf+'Btn').className='ckbtn '+worst; $(pf+'Btn').innerHTML='<span class="ckdot"></span>'+label+' <b>'+nOk+'/'+it.length+'</b>';
  const ic={ok:'✓',warn:'!',bad:'✕'}, esc2=s=>String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;');
  $(pf+'List').innerHTML=it.map((x,i)=>'<li class="ckrow '+x.state+(CKOPEN[x.id]?' open':'')+'"><button type="button" class="ckhead" data-ck="'+x.id+'" aria-expanded="'+!!CKOPEN[x.id]+'"><span class="ckic">'+ic[x.state]+'</span><span class="cktx"><b>'+(i+1)+'. '+x.title+'</b><small>'+esc2(x.summary)+'</small></span><span class="ckch">▾</span></button>'
    +'<div class="ckdet"><p>'+x.lines.map(esc2).join('</p><p>')+'</p><p class="ckf">'+x.file+'</p><div class="ckacts">'+(canWrite!==false?'<button type="button" class="ckchip" data-up="'+(x.id.startsWith('z:')?x.id.slice(2):'1')+'">'+(x.id.startsWith('z:')?'Unggah jadwal zona':'Unggah file')+'</button>':'')+'<button type="button" class="ckchip" data-tpl="'+x.id+'">Unduh '+(x.id.startsWith('z:')?'format zona (terisi)':'template')+'</button></div></div></li>').join('');
  $(pf+'Sub').textContent=sub;
}
function renderChecklist(){
  const P=ckPeriod(), sub='Periode '+CKMON[P.m-1]+' '+P.y+', sampai '+ckLab(P.ref)+' (mengikuti tanggal "Sampai")';
  renderCk('ck',ckItems(),'Checklist database',sub);
  renderCk('ab',abItems(),'Checklist absensi',sub+'. Zona yang jadwalnya sudah ada dianggap sudah input.');
}
document.addEventListener('click',e=>{
  for(const pf of ['ck','ab']){
    const b=e.target.closest('#'+pf+'Btn');
    if(b){ const o=$(pf+'Panel').hidden; ['ck','ab'].forEach(q=>{ if($(q+'Panel')){ $(q+'Panel').hidden=true; $(q+'Btn').setAttribute('aria-expanded','false'); } }); $(pf+'Panel').hidden=!o; b.setAttribute('aria-expanded',String(o)); return; }
  }
  const h=e.target.closest('.ckhead'); if(h){ CKOPEN[h.dataset.ck]=!CKOPEN[h.dataset.ck]; renderChecklist(); return; }
  const u=e.target.closest('[data-up]'); if(u){ ZEXP=u.dataset.up==='1'?null:u.dataset.up; $('file').click(); return; }
  const t=e.target.closest('[data-tpl]'); if(t){ const it=ckItems().concat(abItems()).find(x=>x.id===t.dataset.tpl); if(it) dlTpl(it.tpl); return; }
  if(e.target.closest('#ckAll')){ dlTpl(['sales','sched']); return; }
  ['ck','ab'].forEach(pf=>{ if(!e.target.closest('#'+pf+'Panel') && $(pf+'Panel') && !$(pf+'Panel').hidden){ $(pf+'Panel').hidden=true; $(pf+'Btn').setAttribute('aria-expanded','false'); } });
});
document.addEventListener('keydown',e=>{ if(e.key==='Escape') ['ck','ab'].forEach(pf=>{ if($(pf+'Panel')&&!$(pf+'Panel').hidden){ $(pf+'Panel').hidden=true; $(pf+'Btn').focus(); } }); });
function renderDataBar(){
  const n=dates.length, f=o=>fd(o,{day:'numeric',month:'short',year:'numeric'});
  $('db-title').textContent = n? n+' hari data · '+f(dates[0])+(n>1?' sampai '+f(dates[n-1]):'') : 'Belum ada data';
  $('db-sub').textContent = 'Unggah laporan Sales Detail, jadwal & target, Nota Penjualan NAV, atau Report Stock Movement (.xlsx, .xls, .csv), bisa lebih dari satu file. Jadwal per zona (format ABSENSI ZONA) diunggah TL zonasi lewat Checklist database. Tanggal penjualan yang sudah ada akan diganti.'+(dbReady?' Data tersimpan di Supabase dan langsung tampil di semua perangkat tim.':' Data belum tersambung ke penyimpanan.');
  if(canWrite===false) $('db-sub').textContent='Peran Anda viewer, jadi hanya bisa melihat data. Untuk memperbarui data, minta admin mengubah peran Anda menjadi editor.';
  const s=SCH;
  $('sch-info').textContent = s.people.length? 'Jadwal & target: '+s.people.length+' orang · '+f(s.dates[0])+(s.dates.length>1?' sampai '+f(s.dates[s.dates.length-1]):'')+(s.source?' · '+s.source:'') : 'Belum ada jadwal & target. Unggah file jadwal untuk menampilkan kolom target.';
  $('mg-body').innerHTML = dates.slice().reverse().map(d=>{
    const x=days[d]; return '<tr><td>'+fd(d)+'</td><td class="num">'+rp(x.meta.total)+'</td><td class="num">'+x.people.length+'</td><td>'+esc(x.source||'File awal')+'</td><td><button type="button" class="btn sm" data-del="'+d+'">Hapus</button></td></tr>';
  }).join('') || '<tr><td colspan="5" style="text-align:center;color:var(--ink3)">Belum ada data. Unggah file untuk mulai.</td></tr>';
  renderChecklist();
}
const errText = e => !e? '' : e.code==='quota_exceeded' ? 'Penyimpanan penuh.' : (e.code==='revoked'||e.code==='not_granted') ? 'Akses penyimpanan tidak tersedia.' : 'Penyimpanan gagal ('+(e.code||'error')+').';
const dbErr = e => !e? '' : (e.code==='42501'||/row-level security|permission/i.test(e.message||'')) ? 'Akun Anda tidak punya izin mengubah data.' : 'Penyimpanan gagal ('+(e.message||e.code||'error')+').';
async function persistMv(nav, moves){
  if(!SB) return 'Penyimpanan belum tersedia.';
  let fail=0, last=null; const now=new Date().toISOString();
  if(nav&&nav.length){
    for(let i=0;i<nav.length;i+=400){
      const chunk=nav.slice(i,i+400).map(x=>({store_code:D.code,np_no:x.np,type:x.type,order_date:x.date||null,uploaded_at:now}));
      const { error } = await SB.from('mv_nav').upsert(chunk,{onConflict:'store_code,np_no'}); if(error){ fail++; last=error; break; }
    }
  }
  for(const d of Object.keys(moves||{})){
    const del=await SB.from('mv_moves').delete().eq('store_code',D.code).eq('move_date',d); if(del.error){ fail++; last=del.error; continue; }
    const rows=Object.entries(moves[d]).map(([k,v])=>({store_code:D.code,move_date:d,item_no:k,item_name:v.desc||'',uom:v.uom||'',qty:v.q,n:v.n,uploaded_at:now}));
    const { error } = await SB.from('mv_moves').insert(rows); if(error){ fail++; last=error; }
  }
  return fail? dbErr(last)+' Data NAV / Stock Movement belum tersimpan.' : '';
}
async function persist(ds, sched){
  if(!SB) return 'Penyimpanan belum tersedia.';
  let fail=0, last=null;
  for(const d of ds){
    const doc=localOverlay[d];
    const { error } = await SB.from('dashboard_days').upsert({store_code:D.code, sale_date:d, doc, source:doc.source||null, uploaded_at:new Date().toISOString()},{onConflict:'store_code,sale_date'});
    if(error){ fail++; last=error; } else delete localOverlay[d];
  }
  if(sched){
    const { error } = await SB.from('dashboard_schedule').upsert({store_code:D.code, doc:sched, source:sched.source||null, uploaded_at:new Date().toISOString()},{onConflict:'store_code'});
    if(error){ fail++; last=error; } else if(localSched===sched) localSched=null;
  }
  if(!fail) await loadAll();
  return fail? dbErr(last)+' '+fail+' bagian belum tersimpan.' : '';
}
let ZEXP=null;
async function saveZona(z, source){
  const code=D.code;
  for(let t=0;t<6;t++){
    const {data:row,error:e0}=await SB.from('dashboard_schedule').select('doc,uploaded_at').eq('store_code',code).maybeSingle();
    if(e0) throw e0;
    const now=new Date().toISOString(), doc=mergeZona(row&&row.doc,z,source,now);
    if(!row){
      const {error}=await SB.from('dashboard_schedule').insert({store_code:code,doc,source:null,uploaded_at:now});
      if(!error) return doc; if(error.code!=='23505') throw error; continue;
    }
    const {data,error}=await SB.from('dashboard_schedule').update({doc,uploaded_at:now}).eq('store_code',code).eq('uploaded_at',row.uploaded_at).select('store_code');
    if(error) throw error;
    if(data&&data.length) return doc;
    await new Promise(r=>setTimeout(r,150+Math.random()*250));
  }
  throw new Error('Jadwal sedang diperbarui pengguna lain. Coba unggah lagi sebentar lagi.');
}
async function handleZona(f, buf){
  let z;
  try{ z=parseZona(window.XLSX, buf, D.code, effectiveSched()); }
  catch(e){ return {ok:false, msg:'DITOLAK: '+f.name+'\n'+(e&&e.message||'format tidak sesuai')+'\nFile tidak disimpan. Perbaiki file, lalu unggah lagi.'}; }
  if(ZEXP && z.zona!==ZEXP) return {ok:false, msg:'DITOLAK: '+f.name+'\nFile ini berisi '+z.zona+', tetapi Anda mengunggah dari baris '+ZEXP+'. Pilih file yang benar.\nFile tidak disimpan.'};
  if(!SB) return {ok:false, msg:'DITOLAK: '+f.name+'\nPenyimpanan belum tersedia.'};
  try{ await saveZona(z, f.name); }
  catch(e){ return {ok:false, msg:'GAGAL MENYIMPAN: '+f.name+'\n'+((e&&e.code)?dbErr(e):(e&&e.message)||'coba lagi')}; }
  return {ok:true, msg:'Jadwal '+z.zona+' '+CKMON[z.m-1]+' '+z.y+' tersimpan: '+z.people.length+' orang ('+f.name+').'+(z.warnings.length?'\nCatatan:\n- '+z.warnings.join('\n- '):'')};
}
let busy=false;
async function handleFiles(list){
  const files=[...list]; if(!files.length||busy) return;
  if(!canWrite){ setStatus('Peran Anda hanya bisa melihat data.','warn'); return; }
  if(!window.XLSX){ setStatus('Pembaca Excel gagal dimuat. Periksa koneksi, lalu muat ulang halaman.','err'); return; }
  busy=true; $('upload').disabled=true;
  const found={}, notes=[], zres=[]; let used=0, newSched=null, newNav=null, newMoves=null, movRows=0;
  for(const f of files){
    try{
      setStatus('Membaca '+f.name+'…');
      const buf=await f.arrayBuffer();
      if(isZonaFile(window.XLSX,buf)){ zres.push(await handleZona(f,buf)); continue; }
      let r=null;
      try{ r=parseReport(window.XLSX, buf, D.code); }
      catch(e){
        if(!/Receipt Date/.test(e&&e.message||'')) throw e;
        try{ const sc=parseSchedule(window.XLSX, buf); sc.source=f.name; sc.uploadedAt=new Date().toISOString(); const cz=effectiveSched(); if(cz&&cz.zones) sc.zones=cz.zones; newSched=sc; continue; }catch(e2){}
        try{ const nv=parseNav(window.XLSX, buf); newNav=(newNav||[]).concat(nv.nav); continue; }catch(e3){}
        try{ const mm=parseMoves(window.XLSX, buf); newMoves=newMoves||{}; for(const d in mm.moves){ const t=newMoves[d]||(newMoves[d]={}); for(const k in mm.moves[d]){ const v=mm.moves[d][k]; if(t[k]){ t[k].q+=v.q; t[k].n+=v.n; } else t[k]=v; } } movRows+=mm.used; continue; }catch(e4){}
        throw new Error('format tidak dikenali (bukan Sales Detail, jadwal & target, Nota Penjualan NAV, atau Stock Movement)');
      }
      for(const d in r.days) found[d]=Object.assign(r.days[d],{source:f.name,uploadedAt:new Date().toISOString()});
      used+=r.used;
      if(r.skippedStore) notes.push(f.name+': '+r.skippedStore+' baris dari toko lain dilewati.');
      if(r.skippedRows) notes.push(f.name+': '+r.skippedRows+' baris tanpa tanggal valid dilewati.');
      if(!r.used) notes.push(f.name+': tidak ada baris data untuk toko '+D.code+'.');
    }catch(e){ notes.push(f.name+': '+(e&&e.message||'gagal dibaca')); }
  }
  const ds=Object.keys(found).sort();
  const zmsg=zres.map(x=>x.msg).join('\n\n'), zbad=zres.some(x=>!x.ok);
  if(zres.some(x=>x.ok)) await loadAll();
  if(!ds.length && !newSched && !newNav && !newMoves){
    if(zres.length){ setStatus(zmsg+(notes.length?'\n'+notes.join(' '):''), zbad?'err':'ok'); }
    else setStatus('Tidak ada data yang bisa dimuat. '+notes.join(' '),'err');
    ZEXP=null; busy=false; $('upload').disabled=false; return;
  }
  ds.forEach(d=>{ localOverlay[d]=found[d]; });
  if(newSched) localSched=newSched;
  rebuild(true);
  const parts=[];
  if(ds.length){ const rng = ds.length>1? fd(ds[0],{day:'numeric',month:'short'})+' sampai '+fd(ds[ds.length-1],{day:'numeric',month:'short',year:'numeric'}) : fd(ds[0]); parts.push(fmtN.format(used)+' baris penjualan, '+ds.length+' hari ('+rng+')'); }
  if(newSched) parts.push('jadwal & target '+newSched.people.length+' orang');
  if(newNav) parts.push('Nota Penjualan NAV '+newNav.length+' nota');
  if(newMoves) parts.push('Stock Movement '+movRows+' baris ('+Object.keys(newMoves).sort().map(d=>fd(d,{day:'numeric',month:'short'})).join(', ')+')');
  setStatus('Dashboard diperbarui: '+parts.join('; ')+'. Menyimpan…');
  let err=await persist(ds, newSched);
  if(newNav||newMoves){ const e2=await persistMv(newNav,newMoves); if(e2) err=(err?err+' ':'')+e2; MV.ver++; mvRender(); }
  setStatus('Dashboard diperbarui: '+parts.join('; ')+'.'+(err? ' '+err : (dbRef?' Tersimpan.':''))+(notes.length?' '+notes.join(' '):'')+(zmsg?'\n\n'+zmsg:''), err||notes.length||zbad?'warn':'ok');
  ZEXP=null; busy=false; $('upload').disabled=false;
}
async function removeDate(d){
  delete localOverlay[d];
  if(!SB||!canWrite){ setStatus('Peran Anda tidak bisa menghapus data.','warn'); return; }
  const { data, error } = await SB.from('dashboard_days').delete().eq('store_code',D.code).eq('sale_date',d).select('sale_date');
  if(error){ setStatus(dbErr(error),'warn'); return; }
  if(!data||!data.length){ setStatus('Data '+fd(d)+' tidak terhapus. Akun Anda mungkin tidak punya izin.','warn'); return; }
  setStatus('Data '+fd(d)+' dihapus.','ok');
  await loadAll();
}
$('upload').addEventListener('click',()=>{ ZEXP=null; $('file').click(); });
$('file').addEventListener('change',e=>{ handleFiles(e.target.files); e.target.value=''; });
$('manage').addEventListener('click',()=>{ const p=$('manage-panel'); p.hidden=!p.hidden; $('manage').setAttribute('aria-expanded',String(!p.hidden)); });
let armed=null, armT=0;
$('mg-body').addEventListener('click',e=>{
  const b=e.target.closest('[data-del]'); if(!b) return;
  if(armed===b){ clearTimeout(armT); armed=null; removeDate(b.dataset.del); return; }
  if(armed){ armed.textContent='Hapus'; } clearTimeout(armT);
  armed=b; b.textContent='Klik lagi untuk hapus';
  armT=setTimeout(()=>{ if(armed===b){ b.textContent='Hapus'; armed=null; } },4000);
});
const bar=$('databar');
addEventListener('dragover',e=>{ if(e.dataTransfer&&[...e.dataTransfer.types].includes('Files')){ e.preventDefault(); bar.classList.add('drag'); } });
addEventListener('dragleave',e=>{ if(!e.relatedTarget) bar.classList.remove('drag'); });
addEventListener('drop',e=>{ if(e.dataTransfer&&e.dataTransfer.files.length){ e.preventDefault(); bar.classList.remove('drag'); handleFiles(e.dataTransfer.files); } });
dF.addEventListener('change',()=>{ if(dF.value) setRange(dF.value, dT.value||dF.value); });
dT.addEventListener('change',()=>{ if(dT.value) setRange(dF.value||dT.value, dT.value); });
$('d-all').addEventListener('click',()=>{ if(dates.length) setRange(dates[0],dates[dates.length-1]); });
$('zona').addEventListener('change',e=>{ S.zona=e.target.value; all(); });
rebuild(true);
// ---- Supabase: login, peran, dan data bersama ----
const CFG = window.APP_CONFIG || {};
const SB = (window.supabase && CFG.SUPABASE_URL && CFG.SUPABASE_PUBLISHABLE_KEY && !/YOUR-/.test(CFG.SUPABASE_URL)) ? window.supabase.createClient(CFG.SUPABASE_URL, CFG.SUPABASE_PUBLISHABLE_KEY) : null;
const TIM = new URLSearchParams(location.search).get('t');  // link tim sales: hanya baca, tanpa login
if(TIM) document.body.classList.add('tim');
mvReady=true;
const bootEl=$('boot'), appEl=$('app');
const bootMsg=(t,showOut)=>{ $('boot-msg').textContent=t; $('boot-out').hidden=!showOut; bootEl.hidden=false; appEl.hidden=true; $('psdock').hidden=true; };
const toLogin=()=>{ try{sessionStorage.setItem('mitra10_next','dashboard.html')}catch(e){} location.replace('login.html?next=dashboard.html'); };
async function loadAll(){
  if(!SB) return;
  if(TIM){
    const { data, error } = await SB.rpc('tim_view_data',{p_token:TIM});
    if(error){ if(!dbReady) bootMsg('Data tidak bisa dibaca: '+error.message,false); return; }
    if(!data){ bootMsg('Link tidak valid atau sudah dinonaktifkan. Minta link terbaru ke admin.',false); return; }
    const o={}; (data.days||[]).forEach(r=>{ if(r.doc) o[r.sale_date]=r.doc; });
    dbOverlay=o; dbSchedDoc=data.sched||null; dbReady=true; rebuild(false); return;
  }
  const [d,s]=await Promise.all([
    SB.from('dashboard_days').select('sale_date,doc').eq('store_code',D.code).order('sale_date'),
    SB.from('dashboard_schedule').select('doc').eq('store_code',D.code).maybeSingle()
  ]);
  if(d.error||s.error){ dbReady=false; setStatus('Data tidak bisa dibaca. '+dbErr(d.error||s.error),'warn'); renderDataBar(); return; }
  const o={}; (d.data||[]).forEach(r=>{ if(r.doc) o[r.sale_date]=r.doc; });
  dbOverlay=o; dbSchedDoc=(s.data&&s.data.doc)||null; dbReady=true; rebuild(false);
}
let reloadT=0; const reloadSoon=()=>{ clearTimeout(reloadT); reloadT=setTimeout(()=>{ if(!busy) loadAll(); },500); };
function applyWrite(){ $('upload').hidden=!canWrite; $('manage').hidden=!canWrite; renderDataBar(); }
const ROLE_LABEL={admin:'Admin',editor:'Editor',viewer:'Viewer'};
async function boot(){
  rebuild(true);
  if(!SB){ bootMsg('Konfigurasi belum lengkap. Isi SUPABASE_URL dan SUPABASE_PUBLISHABLE_KEY di config.js.',false); return; }
  if(TIM){
    canWrite=false; await loadAll();
    if(dbReady){ bootEl.hidden=true; appEl.hidden=false; $('psdock').hidden=false; applyWrite(); }
    const refreshT=()=>{ if(!busy && !document.hidden) loadAll(); };
    setInterval(refreshT,45000); document.addEventListener('visibilitychange',refreshT); addEventListener('focus',refreshT);
    return;
  }
  const { data:{ session } } = await SB.auth.getSession();
  if(!session){ toLogin(); return; }
  SB.auth.onAuthStateChange(ev=>{ if(ev==='SIGNED_OUT') toLogin(); });
  const { data:roleRow, error:roleErr } = await SB.from('user_roles').select('role').eq('user_id',session.user.id).maybeSingle();
  if(roleErr){ bootMsg('Peran akun tidak bisa dibaca: '+roleErr.message,true); return; }
  if(!roleRow){ bootMsg('Akun '+session.user.email+' belum punya peran. Hubungi admin.',true); return; }
  canWrite = roleRow.role==='admin'||roleRow.role==='editor';
  $('who-name').textContent=session.user.email; spLoadRole(session.user.email); $('who-role').textContent=ROLE_LABEL[roleRow.role]||roleRow.role; if(window.startPresence)startPresence(SB,session.user.email,ROLE_LABEL[roleRow.role]||roleRow.role,'Salesperson');
  bootEl.hidden=true; appEl.hidden=false; $('psdock').hidden=false;
  applyWrite();
  await loadAll();
  const ch=SB.channel('dashboard-data');
  ['dashboard_days','dashboard_schedule'].forEach(t=>ch.on('postgres_changes',{event:'*',schema:'public',table:t},reloadSoon));
  ch.subscribe();
  const refresh=()=>{ if(!busy && !document.hidden) loadAll(); };
  setInterval(refresh,60000);
  document.addEventListener('visibilitychange',refresh);
  addEventListener('focus',refresh);
}
const signOut=async()=>{ try{ if(SB) await SB.auth.signOut(); }catch(e){} location.replace('index.html'); };
$('logout').addEventListener('click',signOut); $('boot-out').addEventListener('click',signOut);
boot().catch(e=>bootMsg('Gagal memulai: '+(e&&e.message||e),true));

})();
