(function(){
'use strict';
const CFG=window.APP_CONFIG||{};
const SB=(window.supabase&&CFG.SUPABASE_URL&&CFG.SUPABASE_PUBLISHABLE_KEY&&!/YOUR-/.test(CFG.SUPABASE_URL))?window.supabase.createClient(CFG.SUPABASE_URL,CFG.SUPABASE_PUBLISHABLE_KEY,{auth:{persistSession:false,autoRefreshToken:false}}):null;
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

  let tB=0,tA=0,mB=0,mA=0,monthB=0,rows='';
  dates.forEach((d,i)=>{
    const c=p.c?p.c[i]:'';
    const w=isWork(c);
    const fut=!last||d>last;
    const b=w?tpe:0;
    const a=sales[d]||0;
    if(w) monthB+=b;
    if(!fut){ tB+=b; tA+=a; if(w){mB+=b;mA+=a;} }
    const pp=(w&&b>0&&!fut)?a/b*100:null;
    const today=last&&d===last;
    rows+='<tr class="'+(!w?'off ':'')+(fut?'fut ':'')+(today?'today':'')+'"><td>'+esc(fd(d))+'</td><td class="c"><span class="pill'+(w?' w':'')+'">'+esc(c||'–')+'</span></td><td>'+(w?rp(b):'–')+'</td><td'+(fut||!a?'':' class="clk" data-d="'+d+'" title="Lihat item"')+'>'+(fut?'':rp(a))+'</td><td class="p '+cls(pp)+'">'+(pp==null?'':pct(pp))+'</td></tr>';
  });
  const mp=mB>0?mA/mB*100:null;
  rows+='<tr class="tot"><td>TOTAL s.d. '+(last?esc(fdl(last)):'-')+'</td><td class="c"></td><td>'+rp(mB)+'</td><td>'+rp(mA)+'</td><td class="'+cls(mp)+'">'+pct(mp)+'</td></tr>';
  $('t-body').innerHTML=rows;
  $('t-sub').textContent='Target satu bulan penuh: Rp '+rp(monthB)+' · Data terakhir: '+(last?fdl(last):'belum ada');

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
  $('s-dn').textContent=last?(lw?('Actual Rp '+rp(da)+' dari budget Rp '+rp(tpe)):'Hari ini libur/cuti, tidak ada budget'):'';
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
}


function itemsFor(dts){
  const set=new Set(dts),m={};
  (RD.sales||[]).forEach(x=>{ if(!set.has(x.d)) return; (x.it||[]).forEach(e=>{const k=e[2]||e[3];const t=m[k]||(m[k]=[e[0],e[1],e[2],e[3],0,0]);t[4]+=Number(e[4])||0;t[5]+=Number(e[5])||0}); });
  return Object.values(m).sort((a,b)=>b[5]-a[5]);
}
const tc=s=>String(s||'').toLowerCase().replace(/\b\w/g,c=>c.toUpperCase());
function openPop(title,per,dts){
  const it=itemsFor(dts),tq=it.reduce((s,e)=>s+e[4],0),ts=it.reduce((s,e)=>s+e[5],0);
  $('pop-t').textContent=title;$('pop-per').textContent=per+' · penjualan tanpa PPN';
  $('pop-body').innerHTML=it.length?'<table><thead><tr><th class="l">Category</th><th class="l">Master brand</th><th class="l">No Item</th><th class="l">Description</th><th>Qty</th><th>Value (Rp)</th></tr></thead><tbody>'+it.map(e=>'<tr><td class="l">'+esc(tc(e[0]))+'</td><td class="l">'+esc(tc(e[1]))+'</td><td class="l">'+esc(e[2])+'</td><td class="w">'+esc(e[3])+'</td><td>'+rp(e[4])+'</td><td>'+rp(e[5])+'</td></tr>').join('')+'<tr class="tot"><td class="l" colspan="4">TOTAL ('+it.length+' item)</td><td>'+rp(tq)+'</td><td>'+rp(ts)+'</td></tr></tbody></table>':'<p class="note" style="padding:16px">Belum ada rincian item untuk periode ini.</p>';
  $('pop').hidden=false;$('pop-body').scrollTop=0;$('pop-x').focus();
}
function workDates(upTo){ const p=RD.p; return (RD.dates||[]).filter((d,i)=>d<=upTo&&isWork(p.c[i])); }
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
    clearInterval(timer);
    timer=setInterval(()=>{if(!document.hidden)load(CODE,false)},30000);
  }catch(e){
    if(fromForm||$('login').hidden){showLogin();showErr('Gagal terhubung. Periksa internet Anda lalu coba lagi.')}
  }finally{
    btn.disabled=false;btn.textContent='Lihat pencapaian';
  }
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
  $('pop-x').addEventListener('click',closePop);
  $('pop').addEventListener('click',e=>{if(e.target===$('pop'))closePop()});
  document.addEventListener('keydown',e=>{if(e.key==='Escape')closePop()});
  $('s-dp').addEventListener('click',()=>{if(RD&&RD.last)openPop('Item terjual · Daily',fdl(RD.last),[RD.last])});
  $('s-mp').addEventListener('click',()=>{if(RD&&RD.last)openPop('Item terjual · MTD','1 s.d. '+fdl(RD.last)+' (hari kerja)',workDates(RD.last))});
  $('t-body').addEventListener('click',e=>{const c=e.target.closest('td[data-d]');if(c)openPop('Item terjual · '+fd(c.dataset.d),fdl(c.dataset.d),[c.dataset.d])});
  $('switch').addEventListener('click',()=>{try{sessionStorage.removeItem(SKEY)}catch(e){}CODE='';$('code').value='';showErr('');showLogin()});
  $('reload').addEventListener('click',()=>load(CODE,false));
  document.addEventListener('visibilitychange',()=>{if(!document.hidden&&CODE&&!$('view').hidden)load(CODE,false)});
  let saved='';try{saved=sessionStorage.getItem(SKEY)||''}catch(e){}
  if(saved){$('code').value=saved;load(saved,true)} else $('code').focus();
}
init();
})();
