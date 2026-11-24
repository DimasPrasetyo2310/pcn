(function(){
'use strict';
const CFG=window.APP_CONFIG||{};
const SB=(window.supabase&&CFG.SUPABASE_URL&&CFG.SUPABASE_PUBLISHABLE_KEY&&!/YOUR-/.test(CFG.SUPABASE_URL))?window.supabase.createClient(CFG.SUPABASE_URL,CFG.SUPABASE_PUBLISHABLE_KEY,{auth:{persistSession:false,autoRefreshToken:false}}):null;
let APR=null; // daftar approver (nama) untuk dropdown PIN
const KEY=(new URLSearchParams(location.search).get('k')||'').trim();
const $=id=>document.getElementById(id);
const isWork=c=>/^[SPM](?:[\/.].*)?$/.test(String(c||''));
const MON=['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];
const DOW=['Min','Sen','Sel','Rab','Kam','Jum','Sab'];
const nf=new Intl.NumberFormat('id-ID',{maximumFractionDigits:0});
const rp=n=>nf.format(Math.round(n||0));
const fd=iso=>{const d=new Date(iso+'T00:00:00');return DOW[d.getDay()]+', '+d.getDate()+' '+MON[d.getMonth()]};
const fdl=iso=>{const d=new Date(iso+'T00:00:00');return d.getDate()+' '+MON[d.getMonth()]+' '+d.getFullYear()};
const cls=p=>p==null?'na':p>=100?'gd':p>=75?'am':'bd';
const pct=p=>p==null?'–':(Math.round(p*10)/10).toString().replace('.',',')+'%';
const SKEY='sp_code_'+KEY;
let CODE='',timer=null,RD=null;

function esc(s){return String(s==null?'':s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]))}
function showErr(m){const e=$('err');e.textContent=m;e.hidden=!m}
function showLogin(){clearInterval(timer);$('view').hidden=true;$('login').hidden=false;$('code').focus()}

async function fetchData(code){
  const {data,error}=await SB.rpc('sp_view_data',{p_key:KEY,p_code:code});
  if(error) throw error;
  return data;
}

function render(r){
  RD=r;
  const p=r.p,dates=r.dates||[],sales={};
  (r.sales||[]).forEach(x=>{sales[x.d]=Number(x.s)||0});
  const last=r.last||null;
  const tpe=p.tpd>0?p.tpd:(p.tpm>0?p.tpm/30:0);
  $('v-name').textContent=p.name||'';
  $('v-code').textContent=p.id||'';
  $('v-zona').textContent=p.zona||'–';
  $('v-cat').textContent=p.cat||'–';
  $('v-brand').textContent=p.brand||'–';

  const BR=String(p.brand||'').trim().toUpperCase(), isPB=BR==='PRIVATE BRAND', CAT=String(p.cat||'').trim().toUpperCase();
  const UP=x=>String(x==null?'':x).trim().toUpperCase();
  const ZC=new Set((r.zcats&&r.zcats.length?r.zcats:[p.cat]).map(UP).filter(Boolean)); // category yang termasuk di zonasi salesperson
  PCTX={isPB,BR,ZC}; let pbFallback=false;
  const sumCats=(o)=>{ let v=0; Object.keys(o||{}).forEach(k=>{ if(ZC.has(UP(k))) v+=Number(o[k])||0; }); return v; };
  const dm=x=>{ // proteksi (qty BCA), penjualan master brand dan total category (zonasi) per hari
    const it=x&&x.it||[]; let mb=0,bca=0;
    it.forEach(e=>{const m=UP(e[1]); if(m==='BCA')bca+=Number(e[4])||0; if(!isPB&&BR&&m===BR)mb+=Number(e[5])||0;});
    if(isPB){ if(x&&x.pbc&&typeof x.pbc==='object') mb=sumCats(x.pbc); else { mb=Number(x&&x.pb)||0; if(mb>0) pbFallback=true; } }
    return {mb,bca,cat:isPB?sumCats(x&&x.cats):0};
  };
  const byD={},apr={};(r.sales||[]).forEach(x=>{byD[x.d]=x});(r.appr||[]).forEach(x=>{apr[x.d]=x});
  const ck=(on,at,by)=>on?'<span class="ok" title="'+esc((by?'Disetujui oleh '+by:'Disetujui')+(at?' · '+fdl(String(at).slice(0,10)):''))+'">✓</span>'+(by?'<div class="by">'+esc(by)+'</div>':''):'<span class="na">–</span>';
  // Penjualan pada hari non-shift (OFF, CT, RO, dll) tetap dihitung; budget hanya pada hari kerja (S, P, M)
  let mB=0,mA=0,mMB=0,mCat=0,tInv=0,tBca=0,monthB=0,rows='';
  dates.forEach((d,i)=>{
    const c=p.c?p.c[i]:'';
    const w=isWork(c);
    const fut=!last||d>last;
    const b=w?tpe:0;
    const x=byD[d]||{},a=sales[d]||0,m=dm(x),inv=Number(x.r)||0;
    if(w) monthB+=b;
    if(!fut){ tInv+=inv; tBca+=m.bca; mA+=a; mMB+=m.mb; mCat+=m.cat; if(w) mB+=b; }
    const pp=(w&&b>0&&!fut)?a/b*100:null;
    const today=last&&d===last;
    const ap=apr[d]||{}, autoM=!ap.mod&&ap.tl&&pp!=null&&pp>=AUTO_MOD;
    rows+='<tr class="'+(!w?'off ':'')+(fut?'fut ':'')+(today?'today':'')+'"><td>'+esc(fd(d))+'</td><td class="c"><span class="pill'+(w?' w':'')+'">'+esc(c||'–')+'</span></td><td>'+(w?rp(b):'–')+'</td><td'+(fut||!a?'':' class="clk" data-d="'+d+'" title="Lihat item"')+'>'+(fut?'':rp(a))+'</td><td class="p '+cls(pp)+'">'+(pp==null?'':pct(pp))+'</td><td>'+(fut?'':inv)+'</td><td>'+(fut?'':m.bca)+'</td><td class="c">'+(fut?'':ck(ap.tl,ap.tl_at,ap.tl_by))+'</td><td class="c">'+(fut?'':(autoM?'<span class="ok" title="Auto Approved: pencapaian harian '+esc(pct(pp))+' (≥ '+AUTO_MOD+'%)">✓</span><div class="by">Auto</div>':ck(ap.mod,ap.mod_at,ap.mod_by)))+'</td></tr>';
  });
  const mp=mB>0?mA/mB*100:null;
  rows+='<tr class="tot"><td>TOTAL s.d. '+(last?esc(fdl(last)):'-')+'</td><td class="c"></td><td>'+rp(mB)+'</td><td>'+rp(mA)+'</td><td class="'+cls(mp)+'">'+pct(mp)+'</td><td>'+tInv+'</td><td>'+tBca+'</td><td></td><td></td></tr>';
  $('t-body').innerHTML=rows;
  $('t-sub').textContent='Target satu bulan penuh: Rp '+rp(monthB)+' · Master brand: '+(p.brand||'–')+' · Data terakhir: '+(last?fdl(last):'belum ada');

  // Kartu total penjualan master brand vs non master brand
  const nonv=Math.max(0,mA-mMB), sh=v=>mA>0?pct(v/mA*100):'–';
  $('mb-sub').textContent=(last?'1 s.d. '+fdl(last):'Belum ada data')+' · semua hari ber-penjualan dihitung, termasuk OFF/CT/RO · tanpa PPN'+(isPB?' · Private Brand = Master type brand, hanya category zonasi '+(p.zona||''):'')+' · klik kartu untuk melihat item';
  $('mb-l1').textContent=isPB?'Private Brand (category zonasi)':'Master brand '+(p.brand||'');
  $('mb-v1').textContent='Rp '+rp(mMB); $('mb-p1').textContent=sh(mMB)+' dari total';
  $('mb-l2').textContent=isPB?'Non Private Brand':'Non master brand';
  $('mb-v2').textContent='Rp '+rp(nonv); $('mb-p2').textContent=sh(nonv)+' dari total';
  $('mb-v3').textContent='Rp '+rp(mA); $('mb-p3').textContent=mp==null?'':'Pencapaian '+pct(mp)+' dari budget';
  const cmp=$('mb-cmp');
  if(isPB){
    const zl=[...ZC].map(tc).join(', '), pc2=mCat>0?mMB/mCat*100:null, w1=mCat>0?Math.min(100,mMB/mCat*100):0;
    cmp.hidden=false;
    cmp.innerHTML='<h3>Perbandingan Private Brand dengan Category zonasi</h3>'+
      '<div class="note" style="margin:-4px 0 8px">'+esc(p.zona||'–')+': '+esc(zl)+'</div>'+
      '<div class="cmp"><div class="cr"><span>Private Brand (Master type), hanya category zonasi</span><b>Rp '+rp(mMB)+'</b></div>'+
      '<div class="cbar"><i style="width:'+w1+'%"></i></div>'+
      '<div class="cr"><span>Total penjualan category zonasi</span><b>Rp '+rp(mCat)+'</b></div>'+
      '<p class="note" style="margin:6px 0 0">Private Brand = <b>'+(pc2==null?'–':pct(pc2))+'</b> dari penjualan category zonasi.</p>'+
      (pbFallback?'<p class="note" style="margin:6px 0 0;color:#a86a00">Sebagian data lama belum memuat rincian Private Brand per category, jadi hari itu dihitung dari semua category. Minta Team Leader upload ulang file penjualan.</p>':'')+'</div>';
  } else { cmp.hidden=true; cmp.innerHTML=''; }

  // Daily
  const li=last?dates.indexOf(last):-1;
  const lw=li>=0&&isWork(p.c[li]);
  const da=last?(sales[last]||0):0;
  const dp=(lw&&tpe>0)?da/tpe*100:null;
  $('s-dd').textContent=last?fd(last):'Belum ada data';
  $('s-dp').textContent=dp==null?'–':pct(dp);
  $('s-dp').className='big clk '+cls(dp);
  $('s-db').style.width=Math.min(100,Math.max(0,dp||0))+'%';
  $('s-db').className=cls(dp);
  $('s-dn').textContent=last?(lw?('Actual Rp '+rp(da)+' dari budget Rp '+rp(tpe)):'Hari non-shift (OFF/CT/RO dll), tidak ada budget · penjualan Rp '+rp(da)+' tetap dihitung di MTD'):'';
  // MTD
  $('s-md').textContent=last?('1 s.d. '+fdl(last)):'Belum ada data';
  $('s-mp').textContent=pct(mp);
  $('s-mp').className='big clk '+cls(mp);
  $('s-mb').style.width=Math.min(100,Math.max(0,mp||0))+'%';
  $('s-mb').className=cls(mp);
  $('s-mn').textContent='Actual Rp '+rp(mA)+' dari budget Rp '+rp(mB);

  const st=t=>{if(!t)return '–';const d=new Date(t);return isNaN(d)?'–':d.getDate()+' '+MON[d.getMonth()]+' '+d.getFullYear()+', '+String(d.getHours()).padStart(2,'0')+'.'+String(d.getMinutes()).padStart(2,'0')};
  $('stamp').textContent='Jadwal diunggah: '+st(r.sched_at)+' · Data penjualan diunggah: '+st(r.data_at)+' · Dimuat: '+st(new Date().toISOString());
  $('login').hidden=true;$('view').hidden=false;
  renderApprover();
}


let PCTX={isPB:false,BR:'',ZC:new Set()};
function inMB(e){ const U=x=>String(x==null?'':x).trim().toUpperCase(); return PCTX.isPB?(e[6]===1&&PCTX.ZC.has(U(e[0]))):(PCTX.BR&&U(e[1])===PCTX.BR); }
function itemsFor(dts,kind){
  const set=new Set(dts),m={};
  (RD.sales||[]).forEach(x=>{ if(!set.has(x.d)) return; (x.it||[]).forEach(e=>{ if(kind==='mb'&&!inMB(e)) return; if(kind==='non'&&inMB(e)) return; const k=e[2]||e[3];const t=m[k]||(m[k]=[e[0],e[1],e[2],e[3],0,0]);t[4]+=Number(e[4])||0;t[5]+=Number(e[5])||0}); });
  return Object.values(m).sort((a,b)=>b[5]-a[5]);
}
const tc=s=>String(s||'').toLowerCase().replace(/\b\w/g,c=>c.toUpperCase());
function msTypes(dts){ const m={}; (MVD||[]).forEach(x=>{ if(dts.indexOf(x.d)<0) return; (x.l||[]).forEach(l=>{ const k=String(l[0]).replace(/^0+/,''), t=String(l[5]||'').trim(); if(!t) return; const a=m[k]||(m[k]=[]); if(a.indexOf(t)<0) a.push(t); }); }); return m; }
function openPop(title,per,dts,kind){
  const MT=msTypes(dts), it=itemsFor(dts,kind),tq=it.reduce((s,e)=>s+e[4],0),ts=it.reduce((s,e)=>s+e[5],0);
  $('pop-t').textContent=title;$('pop-per').textContent=per+' · penjualan tanpa PPN';
  const old=kind==='mb'&&PCTX.isPB&&!it.length&&(RD.sales||[]).some(x=>dts.includes(x.d)&&(Number(x.pb)||0)>0&&!x.pbc);
  $('pop-body').innerHTML=old?'<p class="note" style="padding:16px">Data penjualan yang tersimpan belum memuat penanda Master type brand per item, jadi detail Private Brand belum bisa ditampilkan. Minta Team Leader upload ulang file penjualan (semua hari bulan ini) di dashboard, lalu detail ini akan terisi.</p>':it.length?'<table><thead><tr><th class="l">Category</th><th class="l">No Item</th><th class="l">Description</th><th class="l">Type MS</th><th>Qty</th><th>Value (Rp)</th></tr></thead><tbody>'+it.map(e=>'<tr><td class="l">'+esc(tc(e[0]))+'</td><td class="l">'+esc(e[2])+'</td><td class="w">'+esc(e[3])+'</td><td class="l">'+esc((MT[String(e[2]).replace(/^0+/,'')]||['–']).join(', '))+'</td><td>'+rp(e[4])+'</td><td>'+rp(e[5])+'</td></tr>').join('')+'<tr class="tot"><td class="l" colspan="4">TOTAL ('+it.length+' item)</td><td>'+rp(tq)+'</td><td>'+rp(ts)+'</td></tr></tbody></table>':'<p class="note" style="padding:16px">Belum ada rincian item untuk periode ini.</p>';
  $('pop').hidden=false;$('pop-body').scrollTop=0;$('pop-x').focus();
}
function allDates(upTo){ return (RD.dates||[]).filter(d=>d<=upTo); }
function closePop(){$('pop').hidden=true}
async function load(code,fromForm){
  const btn=$('go');
  if(fromForm){btn.disabled=true;btn.textContent='Memuat...';showErr('')}
  try{
    const r=await fetchData(code);
    if(!r||r.ok!==true){
      if(fromForm||$('login').hidden) {showLogin();showErr(r&&r.err==='key'?'Link tidak valid atau sudah dicabut. Minta link terbaru ke Team Leader.':'Terjadi kesalahan, coba lagi.')}
      return;
    }
    if(!r.found){
      try{sessionStorage.removeItem(SKEY)}catch(e){}
      showLogin();showErr('Kode tidak ditemukan di jadwal bulan ini. Periksa kembali kode Anda.');
      return;
    }
    CODE=code;
    try{sessionStorage.setItem(SKEY,code)}catch(e){}
    render(r);
    mvLoad(fromForm);
    clearInterval(timer);
    timer=setInterval(()=>{if(!document.hidden)load(CODE,false)},30000);
  }catch(e){
    try{console.error('sp',e)}catch(_){}
    if(fromForm||$('login').hidden){showLogin();showErr('Gagal terhubung. Periksa internet Anda lalu coba lagi.')}
  }finally{
    btn.disabled=false;btn.textContent='Lihat pencapaian';
  }
}


// ---- Validasi TL/MOD: TL/MOD memilih nama dan memasukkan PIN di HP salesperson ----
async function initApprover(){
  if(!SB||APR) return;
  try{ const {data}=await SB.rpc('sp_approvers',{p_key:KEY}); if(data&&data.ok){ APR={tl:data.tl||[],mod:data.mod||[]}; renderApprover(); } }catch(e){}
}
const AUTO_MOD=100; // MOD otomatis disetujui bila pencapaian harian >= 100%
function dayPct(d){ if(!RD||!RD.p) return null; const p=RD.p, i=(RD.dates||[]).indexOf(d); if(i<0) return null; if(!isWork(p.c?p.c[i]:'')) return null; const t=p.tpd>0?p.tpd:(p.tpm>0?p.tpm/30:0); if(!(t>0)) return null; const x=(RD.sales||[]).find(y=>y.d===d); return (x?Number(x.s)||0:0)/t*100; }
function todayWIB(){ return new Intl.DateTimeFormat('en-CA',{timeZone:'Asia/Jakarta',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date()); }
function apDay(){ return (RD&&RD.today)||todayWIB(); } // hari H
function renderApprover(){
  const card=$('ap-card'); if(!card) return;
  if(!APR||!RD||!RD.p){ card.hidden=true; return; }
  const d=apDay(), ap=(RD.appr||[]).find(x=>x.d===d)||{};
  const ts=t=>{if(!t)return '';const x=new Date(t);return isNaN(x)?'':' · '+x.getDate()+' '+MON[x.getMonth()]+', '+String(x.getHours()).padStart(2,'0')+'.'+String(x.getMinutes()).padStart(2,'0')};
  $('ap-who').textContent='Hari H: '+fd(d)+'. Team Leader / Manager: tekan tombol, pilih nama Anda, lalu masukkan PIN Anda sendiri di HP ini.';
  const dp=dayPct(d), auto=!ap.mod&&ap.tl&&dp!=null&&dp>=AUTO_MOD, pend=!ap.mod&&!ap.tl&&dp!=null&&dp>=AUTO_MOD;
  const col=(k)=>{ const ok=!!ap[k]||(k==='mod'&&auto), lock=(k==='tl'&&ap.mod), has=(APR[k]||[]).length>0;
    if(k==='mod'&&auto) return '<div class="apok"><small>Approval MOD</small><b>✓ Auto Approved</b><span class="note">Pencapaian harian '+esc(pct(dp))+' (≥ '+AUTO_MOD+'%), tidak perlu PIN.</span></div>';
    const st=ok?'✓ Disetujui oleh '+esc(ap[k+'_by']||'?')+esc(ts(ap[k+'_at'])):(k==='mod'&&!ap.tl?'Menunggu Approval TL':'Belum');
    const hint=(k==='mod'&&pend)?'<span class="note">Pencapaian harian '+esc(pct(dp))+' (≥ '+AUTO_MOD+'%): otomatis disetujui setelah TL approve.</span>':'';
    let btn;
    if(ok) btn=lock?'':'<button class="btn ghost" data-ap="0" data-k="'+k+'">Batalkan Approval '+k.toUpperCase()+'</button>';
    else if(!has) btn='<button class="btn" disabled>Belum ada PIN '+k.toUpperCase()+'</button>';
    else if(k==='mod'&&!ap.tl) btn='<button class="btn" disabled>Approve MOD (menunggu TL)</button>';
    else btn='<button class="btn" data-ap="1" data-k="'+k+'">Approve '+k.toUpperCase()+'</button>';
    return '<div class="'+(ok?'apok':'apno')+'"><small>Approval '+(k==='tl'?'TL':'MOD')+'</small><b>'+st+'</b>'+hint+btn+'</div>'; };
  $('ap-grid').innerHTML=col('tl')+col('mod');
  { // berapa kali validasi (TL/MOD) tidak lengkap, hari berjalan tidak dihitung
    const have=new Set((RD.sales||[]).map(y=>y.d)), miss=[];
    (RD.dates||[]).forEach((x,i)=>{ if(x>=d||!have.has(x)||!isWork(RD.p.c?RD.p.c[i]:'')) return; const a=(RD.appr||[]).find(y=>y.d===x)||{}, pc=dayPct(x), t=!!a.tl, m=!!a.mod||(t&&pc!=null&&pc>=AUTO_MOD); if(!t||!m) miss.push(x); });
    const sdt=x=>(+x.slice(8))+' '+MON[+x.slice(5,7)-1];
    $('ap-miss').innerHTML=miss.length?'<div style="border:1.5px solid #e8b4af;background:#fdeae8;color:#b4271c;border-radius:10px;padding:10px 12px"><b>Validasi tidak lengkap: '+miss.length+' kali bulan ini</b><div class="note" style="margin-top:4px;color:#7a2b25">'+miss.map(sdt).join(', ')+'</div></div>'
      :'<div style="border:1.5px solid #b7dcc9;background:#e4f4ec;color:#0e6b49;border-radius:10px;padding:10px 12px"><b>✓ Validasi selalu lengkap bulan ini</b></div>';
  }
  card.hidden=false;
}
function apErr(r){ const m={nopin:'PIN akun ini belum diatur. Minta admin mengatur PIN.',locked:'Terlalu banyak salah. Akun ini terkunci, coba lagi dalam '+(r.wait||5)+' menit.',bad:'PIN salah. Sisa percobaan: '+r.left+'.',role_tl:'Nama ini bukan Team Leader.',role_mod:'Nama ini bukan Manager (MOD).',need_tl:'Approval TL harus dilakukan lebih dulu.',mod_first:'Batalkan Approval MOD lebih dulu.',not_owner:'Approval hanya bisa dibatalkan oleh orang yang menyetujuinya.',format:'PIN harus 6 digit angka.',code:'Kode salesperson tidak dikenali.',key:'Link tidak valid.',nouser:'Nama tidak ditemukan.'}; return m[r.err]||'Gagal ('+(r.err||'error')+').'; }
function apPin(on,k){
  const d=apDay(); if(!APR) return;
  const list=APR[k]||[];
  let o=$('pin-ov'); if(!o){ o=document.createElement('div'); o.id='pin-ov'; o.className='pinov'; o.addEventListener('click',e=>{ if(e.target===o||(e.target.closest&&e.target.closest('[data-pinx]'))) o.hidden=true; }); document.body.appendChild(o); }
  o.innerHTML='<div class="pinbox" role="dialog" aria-modal="true"><h3>'+(on?'Approve ':'Batalkan ')+k.toUpperCase()+'</h3><p class="note">'+esc(RD.p.name||RD.p.id)+' · '+esc(fd(d))+'</p><label class="note">Nama '+(k==='tl'?'Team Leader':'Manager')+'</label><select id="pin-who">'+list.map(x=>'<option value="'+esc(x.id)+'">'+esc(x.name)+'</option>').join('')+'</select><input id="pin-in" type="text" inputmode="numeric" maxlength="6" autocomplete="off" spellcheck="false" data-lpignore="true" name="pin-'+Math.random().toString(36).slice(2,7)+'" placeholder="PIN 6 digit"><div class="pinerr" id="pin-err"></div><div class="pinbt"><button class="btn ghost" data-pinx type="button">Batal</button><button class="btn" id="pin-ok" type="button">Konfirmasi</button></div></div>';
  o.hidden=false; const inp=$('pin-in'); inp.addEventListener('input',()=>{ inp.value=inp.value.replace(/\D/g,'').slice(0,6); }); $('pin-who').focus();
  const go=async()=>{ const v=inp.value.trim(); if(!/^\d{6}$/.test(v)){ $('pin-err').textContent='Masukkan PIN 6 digit.'; return; }
    $('pin-ok').disabled=true; const {data,error}=await SB.rpc('sp_approve_anon',{p_key:KEY,p_code:CODE,p_kind:k,p_approver:$('pin-who').value,p_pin:v,p_on:on}); $('pin-ok').disabled=false;
    if(error||!data){ $('pin-err').textContent='Gagal menghubungi server. Coba lagi.'; return; }
    if(!data.ok){ $('pin-err').textContent=apErr(data); inp.value=''; inp.focus(); return; }
    const r=data.row, i=(RD.appr||[]).findIndex(x=>x.d===r.d); RD.appr=RD.appr||[]; if(i>=0) RD.appr[i]=r; else RD.appr.push(r);
    o.hidden=true; render(RD); $('ap-msg').textContent=(on?'Approval ':'Pembatalan approval ')+k.toUpperCase()+' tersimpan.'; };
  $('pin-ok').onclick=go; inp.onkeydown=e=>{ if(e.key==='Enter') go(); };
}

// ---- Task Order Movement: item yang dijual vs Stock Movement gudang ----
let MVD=null, MVSEL='', MVF='all', mvAt=0, mvBusy=false;
const mvq=v=>nf.format(Math.round((Number(v)||0)*100)/100);
async function mvLoad(force){
  if(!SB||!CODE||mvBusy) return; if(!force&&Date.now()-mvAt<120000) return;
  mvBusy=true; mvAt=Date.now();
  try{ const {data}=await SB.rpc('sp_mv_view',{p_key:KEY,p_code:CODE}); if(data&&data.ok){ MVD=data.days||[]; } }catch(e){}
  mvBusy=false; mvRender();
}
// ---- Suggest Location (dari Nota Penjualan): W*/OV* = Gudang, C*/D* + angka = Area Selling, kosong = Tidak input Lokasi, lainnya apa adanya ----
const locKind=t=>{ t=String(t==null?'':t).trim(); return !t?'nil':/^(w\d|ov)/i.test(t)?'wh':/^[cd]\d/i.test(t)?'sel':'oth'; };
function locChips(l,hasLoc){
  if(!hasLoc) return '';
  const ls=(Array.isArray(l[9])&&l[9].length)?l[9]:[['',Number(l[10]||l[4])]], multi=ls.length>1;
  return '<div class="chips">'+ls.map(x=>{ const t=String(x[0]||'').trim(), k=locKind(t), lab=k==='nil'?'Tidak input Lokasi':k==='wh'?'Gudang':k==='sel'?'Area Selling':t;
    return '<span class="chip '+k+'"><b>'+esc(lab)+'</b>'+((k==='wh'||k==='sel')?'<span>'+esc(t.toUpperCase())+'</span>':'')+(multi?'<span>· '+mvq(x[1])+' qty</span>':'')+'</span>'; }).join('')+'</div>';
}
function mvRender(){
  const card=$('mv-card'); if(!card) return;
  if(!MVD||!MVD.length){ card.hidden=true; return; }
  const today=apDay(); const ds=MVD.map(x=>x.d);
  if(!MVSEL||!ds.includes(MVSEL)) MVSEL=ds.includes(today)?today:ds[ds.length-1];
  const day=MVD.find(x=>x.d===MVSEL)||MVD[0], L=day.l||[];
  const hasLoc=L.some(l=>Array.isArray(l[9])&&l[9].length);
  const tasks=L.filter(l=>l[6]), done=tasks.filter(l=>(Number(l[8])||0)>=Number(l[4])), belum=tasks.length-done.length, auto=L.length-tasks.length;
  const flt={all:()=>true,belum:l=>l[6]&&(Number(l[8])||0)<Number(l[4]),done:l=>l[6]&&(Number(l[8])||0)>=Number(l[4]),auto:l=>!l[6]};
  const f=flt[MVF]?MVF:'all', cnt=k=>L.filter(flt[k]).length;
  const bar=!tasks.length?'<div class="mvbar ok">✓ Tidak ada task order movement<small>Semua item Anda tidak memerlukan movement.</small></div>'
    :belum===0?'<div class="mvbar ok">✓ Done · semua task order sudah di-movement<small>'+done.length+' dari '+tasks.length+' task order</small></div>'
    :'<div class="mvbar no">✕ Belum selesai · '+belum+' task order belum di-movement<small>'+done.length+' dari '+tasks.length+' sudah di-movement'+(day.moves_n?'':' · report Stock Movement tanggal ini belum diupload')+'</small></div>';
  const rows=L.slice().sort((a,b)=>{ const ra=a[6]?((Number(a[8])||0)<Number(a[4])?0:1):2, rb=b[6]?((Number(b[8])||0)<Number(b[4])?0:1):2; return ra-rb||String(a[1]).localeCompare(String(b[1])); }).filter(flt[f]).map(l=>{
    const q=Number(l[4]), a=Number(l[8])||0; let cl,st;
    if(!l[6]){ if(l[7]){ cl='ok'; st=(Array.isArray(l[9])&&l[9].some(x=>locKind(x[0])==='wh')?'✓ Auto (Gudang)':'✓ Auto'); } else { cl='wr'; st='? Type belum diketahui'; } }
    else if(a>=q){ cl='ok'; st='✓ Done'; } else { cl='no'; st=(a>0?'✕ Sebagian':'✕ Belum'); }
    return '<div class="mvrow '+cl+'"><div><b>'+esc(l[1])+'</b><small>'+esc(l[0])+' · '+esc(l[2]||'Tanpa No Nota')+'</small><small>'+esc(l[5])+'</small>'+locChips(l,hasLoc)+'</div><div class="q">'+mvq(q)+' qty<small>'+(l[6]?'Tersedia di X333: '+mvq(Math.min(a,q))+' / '+mvq(q)+' · ':'')+st+'</small></div></div>'; }).join('')||'<p class="note">Tidak ada item pada filter ini.</p>';
  const tabs=[['all','Semua'],['belum','Belum'],['done','Sudah'],['auto','Auto']];
  $('mv-inner').innerHTML='<div class="mvhd"><label class="note">Tanggal <select id="mv-d">'+MVD.slice().reverse().map(x=>'<option value="'+x.d+'"'+(x.d===MVSEL?' selected':'')+'>'+esc(fd(x.d))+(x.d===today?' (hari ini)':'')+'</option>').join('')+'</select></label></div>'+bar+
    '<div class="mvsum"><div><small>Item terjual</small><b>'+L.length+'</b></div><div><small>Task order</small><b>'+tasks.length+'</b></div><div><small>Sudah movement</small><b>'+done.length+'</b></div><div><small>Belum movement</small><b>'+belum+'</b></div></div>'+
    '<div class="mvt">'+tabs.map(t=>'<button type="button" data-mvt="'+t[0]+'"'+(f===t[0]?' class="on"':'')+'>'+t[1]+' <i>'+cnt(t[0])+'</i></button>').join('')+'</div>'+rows;
  card.hidden=false;
}
function init(){
  if(!SB||!KEY){
    $('login').hidden=false;
    $('form').hidden=true;
    $('login').querySelector('.sub').textContent='Link tidak lengkap. Minta link yang benar ke Team Leader Anda.';
    return;
  }
  $('login').hidden=false;
  $('form').addEventListener('submit',ev=>{
    ev.preventDefault();
    const v=$('code').value.trim();
    if(!v){showErr('Masukkan kode salesperson Anda.');return}
    load(v,true);
  });
  initApprover();
  $('mv-shot').addEventListener('click',()=>{ if(window.DomShot) DomShot($('mv-card'),'task-order-movement-'+(MVSEL||'')+'.png',{btn:$('mv-shot')}); });
  $('pop-shot').addEventListener('click',()=>{ if(window.DomShot) DomShot(document.querySelector('#pop .dlg'),DomShot.slug($('pop-t').textContent)+'.png'); });
  $('mv-card').addEventListener('click',e=>{ const b=e.target.closest('button[data-mvt]'); if(b){ MVF=b.dataset.mvt; mvRender(); } });
  $('mv-card').addEventListener('change',e=>{ if(e.target&&e.target.id==='mv-d'){ MVSEL=e.target.value; mvRender(); } });
  $('ap-grid').addEventListener('click',e=>{ const b=e.target.closest('button[data-ap]'); if(b&&!b.disabled) apPin(b.dataset.ap==='1',b.dataset.k); });
  document.addEventListener('keydown',e=>{ if(e.key==='Escape'){ const o=$('pin-ov'); if(o) o.hidden=true; } });
  $('pop-x').addEventListener('click',closePop);
  $('pop').addEventListener('click',e=>{if(e.target===$('pop'))closePop()});
  document.addEventListener('keydown',e=>{if(e.key==='Escape')closePop()});
  $('s-dp').addEventListener('click',()=>{if(RD&&RD.last)openPop('Item terjual · Daily',fdl(RD.last),[RD.last])});
  $('s-mp').addEventListener('click',()=>{if(RD&&RD.last)openPop('Item terjual · MTD','1 s.d. '+fdl(RD.last),allDates(RD.last))});
  $('mb-card').addEventListener('click',e=>{const c=e.target.closest('.mbs>div[data-k]');if(!c||!RD||!RD.last)return;const k=c.dataset.k;openPop((k==='mb'?$('mb-l1').textContent:k==='non'?$('mb-l2').textContent:'Total penjualan')+' · MTD','1 s.d. '+fdl(RD.last),allDates(RD.last),k)});
  $('mb-card').addEventListener('keydown',e=>{if(e.key==='Enter'&&e.target.matches&&e.target.matches('.mbs>div[data-k]'))e.target.click()});
  $('t-body').addEventListener('click',e=>{const c=e.target.closest('td[data-d]');if(c)openPop('Item terjual · '+fd(c.dataset.d),fdl(c.dataset.d),[c.dataset.d])});
  $('switch').addEventListener('click',()=>{try{sessionStorage.removeItem(SKEY)}catch(e){}CODE='';$('code').value='';showErr('');showLogin()});
  $('reload').addEventListener('click',()=>load(CODE,false));
  document.addEventListener('visibilitychange',()=>{if(!document.hidden&&CODE&&!$('view').hidden)load(CODE,false)});
  let saved='';try{saved=sessionStorage.getItem(SKEY)||''}catch(e){}
  if(saved){$('code').value=saved;load(saved,true)} else $('code').focus();
}
init();
})();
