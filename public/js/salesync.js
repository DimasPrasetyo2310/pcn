// Pembaca Sales Detail per salesperson yang dipakai bersama oleh halaman Monitoring Sales Bulanan.
// SALIN dari parseReport di dashboard.js: jika aturan hitung di sana diubah, ubah juga di sini.
// Satu unggahan Sales Detail Report By Salesperson di Monitoring juga memperbarui dashboard salesperson (tabel dashboard_days).
window.SalesSync=(function(){
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

async function save(SB, store, days, source){
  let n=0, fail=null;
  for(const d of Object.keys(days).sort()){
    const doc=Object.assign({},days[d],{source:source,uploadedAt:new Date().toISOString()});
    const { error } = await SB.from('dashboard_days').upsert({store_code:store, sale_date:d, doc, source:source||null, uploaded_at:new Date().toISOString()},{onConflict:'store_code,sale_date'});
    if(error){ fail=error; break; } n++;
  }
  return {saved:n, error:fail};
}
function hasSalesperson(XLSX, buf){
  try{ const rows=readRows(XLSX,buf); for(let i=0;i<Math.min(rows.length,30);i++){ if(rows[i].some(c=>norm(c)==='receipt date')){ const h=rows[i].map(norm); return (h.includes('salescode id')||h.includes('salesperson id'))&&(h.includes('salescode name')||h.includes('salesperson name'))&&h.includes('receipt no'); } } }catch(e){}
  return false;
}
return {parseReport, save, hasSalesperson};
})();
