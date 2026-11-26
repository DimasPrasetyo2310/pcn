// Task Follow Up Customer Daily di halaman salesperson (tanpa login).
// Data & batas harian dikunci di server lewat RPC fu_sp_tasks / fu_sp_send / fu_sp_status (key link + kode salesperson).
(function(){
'use strict';
const CFG=window.APP_CONFIG||{};
const SB=(window.supabase&&CFG.SUPABASE_URL&&CFG.SUPABASE_PUBLISHABLE_KEY)?window.supabase.createClient(CFG.SUPABASE_URL,CFG.SUPABASE_PUBLISHABLE_KEY,{auth:{persistSession:false,autoRefreshToken:false}}):null;
const $=id=>document.getElementById(id);
const esc=s=>String(s==null?'':s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const nf=new Intl.NumberFormat('id-ID');
const digits=s=>String(s==null?'':s).replace(/\D/g,'');
const waNum=hp=>{let d=digits(hp);if(!d)return '';if(d.startsWith('0'))d='62'+d.slice(1);else if(d.startsWith('8'))d='62'+d;return d.length>=10?d:''};
const today=()=>new Date().toLocaleDateString('sv-SE',{timeZone:'Asia/Jakarta'});
const jk=ts=>ts?new Date(ts).toLocaleDateString('sv-SE',{timeZone:'Asia/Jakarta'}):'';
const fdt=ts=>{if(!ts)return '';const d=new Date(ts);return d.toLocaleDateString('id-ID',{day:'numeric',month:'short',timeZone:'Asia/Jakarta'})+' '+d.toLocaleTimeString('id-ID',{hour:'2-digit',minute:'2-digit',timeZone:'Asia/Jakarta'})};
const titleCase=s=>String(s||'').toLowerCase().replace(/(^|[\s.'-])([a-z])/g,(m,a,b)=>a+b.toUpperCase()).trim();
const contacted=c=>c.st!=='Baru'||!!c.lc;
const sentToday=c=>!!c.lc&&jk(c.lc)===today();
const isMobile=()=>/Android|iPhone|iPad|iPod|Mobile/i.test(navigator.userAgent);
const GAP=[8,15];

let K='',C='',SPNAME='',D=null,TAB='todo',Q='',SHOW=60,QQ=null,RR=0;

/* ---------- Pesan ---------- */
function salam(){const h=+new Date().toLocaleString('en-US',{hour:'numeric',hour12:false,timeZone:'Asia/Jakarta'})%24;return h<11?'Selamat pagi':h<15?'Selamat siang':h<18?'Selamat sore':'Selamat malam'}
function firstName(full){const w=String(full||'').replace(/\./g,'. ').trim().split(/\s+/).filter(Boolean);const f=w.find(x=>x.replace(/\W/g,'').length>2)||w[0]||'';return titleCase(f.replace(/\.$/,''))}
function renderMsg(isi,c){
  let t=String(isi||'');
  // Kode uji MASTER memakai database salesperson lain: nama di pesan memakai nama salesperson pemilik database
  const full=(/^MASTER\b/i.test(SPNAME)?'':SPNAME)||(D&&D.name)||'';
  const v={salam:salam(),nama:titleCase(c.n),sales:firstName(full),sales_lengkap:titleCase(full),member:c.m||'',toko:'MITRA10 Percetakan Negara'};
  t=t.replace(/\{(salam|nama|sales_lengkap|sales|member|toko)\}/gi,(m,k)=>v[k.toLowerCase()]);
  for(let i=0;i<5&&/\{[^{}]*\|[^{}]*\}/.test(t);i++)t=t.replace(/\{([^{}]*\|[^{}]*)\}/g,(m,o)=>{const a=o.split('|');return a[Math.floor(Math.random()*a.length)]});
  return t.replace(/[ \t]+\n/g,'\n').trim();
}
function pickTpl(){
  const T=(D&&D.templates)||[];if(!T.length)return null;
  const v=($('fu-tpl')||{}).value;
  if(v==='rr'||!v){const t=T[RR%T.length];RR++;return t}
  return T.find(t=>String(t.id)===v)||T[0];
}
function waMode(){let m='';try{m=localStorage.getItem('fu_wa_mode')||''}catch(e){}return m==='web'||m==='app'?m:(isMobile()?'app':'web')}

/* ---------- Gambar ke clipboard ---------- */
const IMGC=new Map();
function pngBlob(url){
  if(IMGC.has(url))return Promise.resolve(IMGC.get(url));
  const p=fetch(url,{mode:'cors'}).then(r=>{if(!r.ok)throw new Error('img');return r.blob()}).then(b=>b.type==='image/png'?b:new Promise((res,rej)=>{
    const im=new Image();im.crossOrigin='anonymous';im.onload=()=>{const cv=document.createElement('canvas');cv.width=im.naturalWidth;cv.height=im.naturalHeight;cv.getContext('2d').drawImage(im,0,0);cv.toBlob(x=>x?res(x):rej(new Error('png')),'image/png')};im.onerror=()=>rej(new Error('img'));im.src=URL.createObjectURL(b)}))
    .then(b=>{IMGC.set(url,b);return b});
  p.catch(()=>IMGC.delete(url));IMGC.set(url,p);return p;
}
function copyImg(url){
  if(!(navigator.clipboard&&navigator.clipboard.write&&window.ClipboardItem))return Promise.reject(new Error('x'));
  const c=IMGC.get(url);return navigator.clipboard.write([new ClipboardItem({'image/png':c instanceof Blob?c:pngBlob(url)})]);
}

/* ---------- Data ---------- */
async function load(){
  if(!SB||!K||!C)return;
  const card=$('fu-card');
  try{
    const {data,error}=await SB.rpc('fu_sp_tasks',{p_key:K,p_code:C});
    if(error)throw error;
    if(!data||!data.ok||!data.found){card.hidden=true;return}
    card.hidden=false;
    if(data.maintenance){D=null;$('fu-sub').textContent='Sedang dilakukan pemeliharaan berkala. Fitur ini akan segera tersedia.';$('fu-badge').textContent='Segera hadir';$('fu-badge').className='fu-badge';$('fu-body').innerHTML='';return}
    D=data;(D.templates||[]).forEach(t=>{if(t.img)pngBlob(t.img).catch(()=>{})});
    renderCard();if($('fu-ovl'))renderPop();
  }catch(e){card.hidden=false;$('fu-sub').textContent='Gagal memuat data follow up. Tekan "Muat ulang".';$('fu-body').innerHTML=''}
}
function stats(){
  const L=D.customers||[],done=L.filter(contacted).length;
  return {total:L.length,done,todo:L.length-done,deal:L.filter(c=>c.st==='Deal').length,bat:L.filter(c=>c.st==='Batal').length,today:L.filter(sentToday).length};
}
function left(){return Math.max(0,(D.limit||40)-(D.sent_today||0))}

/* ---------- Kartu ---------- */
function renderCard(){
  const s=stats(),lim=D.limit||40,sent=D.sent_today||0,pct=Math.min(100,sent/lim*100);
  $('fu-sub').textContent=`Database customer ${D.sid||''} · batas ${lim} pesan WhatsApp per hari`;
  const b=$('fu-badge');b.textContent=`Hari ini ${sent}/${lim}`;b.className='fu-badge'+(sent>=lim?' full':'');
  $('fu-body').innerHTML=`
    <div class="fu-sum">
      <div class="hero"><small>Total database</small><b>${nf.format(s.total)}</b></div>
      <div><small>Sudah terkirim</small><b>${nf.format(s.done)}</b></div>
      <div><small>Belum dihubungi</small><b>${nf.format(s.todo)}</b></div>
      <div><small>Terkirim hari ini</small><b>${nf.format(s.today)}</b></div>
    </div>
    <div class="fu-prog"><div class="row"><span>Target blasting hari ini</span><b>${sent} dari ${lim}</b></div><div class="bar"><i style="width:${pct}%"></i></div></div>
    <div class="fu-act">
      <button class="btn wa" type="button" id="fu-go"${!left()||!s.todo?' disabled':''}>${!left()?'Kuota hari ini sudah habis':`Mulai blasting hari ini (sisa ${left()})`}</button>
      <button class="btn ghost" type="button" id="fu-det">Detail database</button>
    </div>
    ${!(D.templates||[]).length?'<p class="note">Belum ada template pesan aktif. Hubungi admin.</p>':''}`;
  $('fu-go').onclick=()=>{openPop('todo');startQueue()};
  $('fu-det').onclick=()=>openPop('done');
}

/* ---------- Popup ---------- */
function openPop(tab){
  TAB=tab||'todo';Q='';SHOW=60;
  let o=$('fu-ovl');
  if(!o){o=document.createElement('div');o.className='ovl';o.id='fu-ovl';document.body.appendChild(o);
    o.addEventListener('click',e=>{if(e.target===o)closePop()});}
  const T=D.templates||[];
  o.innerHTML=`<div class="dlg" role="dialog" aria-modal="true" aria-labelledby="fu-pt">
    <div class="hd"><div><h2 id="fu-pt">Task Follow Up · ${esc(SPNAME||D.name)}</h2><div class="per" id="fu-pper"></div></div><button class="x" type="button" id="fu-x" aria-label="Tutup">×</button></div>
    <div class="bd">
      <div class="fu-ctl">
        <label>Template pesan<select id="fu-tpl">${T.length>1?`<option value="rr">Acak bergilir (${T.length} template)</option>`:''}${T.map(t=>`<option value="${t.id}">${esc(t.judul)}</option>`).join('')}</select></label>
        <label>Buka WhatsApp lewat<select id="fu-mode"><option value="web">WhatsApp Web (laptop/PC)</option><option value="app">Aplikasi WhatsApp (HP)</option></select></label>
        <div class="fu-img" id="fu-imgrow" hidden><img id="fu-img" alt=""><div>Gambar ikut dikirim: setelah chat terbuka, <b id="fu-paste"></b>, lalu Kirim. <button type="button" class="btn ghost" id="fu-imgcopy" style="margin-left:6px">Salin gambar</button></div></div>
      </div>
      <div class="fu-q" id="fu-q" hidden></div>
      <div class="fu-msg" id="fu-msg" hidden></div>
      <div class="fu-tabs" id="fu-tabs"></div>
      <div style="padding:10px 16px 0"><input type="search" id="fu-search" placeholder="Cari nama, no. HP, no. member" style="padding:9px 10px;font-size:14px"></div>
      <div class="fu-list" id="fu-list"></div>
    </div></div>`;
  $('fu-x').onclick=closePop;
  $('fu-mode').value=waMode();
  $('fu-mode').onchange=e=>{try{localStorage.setItem('fu_wa_mode',e.target.value)}catch(x){}imgRow()};
  $('fu-tpl').onchange=imgRow;
  $('fu-imgcopy').onclick=()=>{const u=$('fu-imgrow').dataset.url;if(!u)return;copyImg(u).then(()=>msg('Gambar tersalin. Tempel di chat WhatsApp.','ok'),()=>msg('Perangkat ini tidak mengizinkan menyalin gambar otomatis.','err'))};
  let st;$('fu-search').oninput=e=>{clearTimeout(st);st=setTimeout(()=>{Q=e.target.value;SHOW=60;renderList()},150)};
  document.addEventListener('keydown',escClose);
  imgRow();renderPop();
}
function escClose(e){if(e.key==='Escape')closePop()}
function closePop(){const o=$('fu-ovl');if(o)o.remove();if(QQ)clearInterval(QQ.tm);QQ=null;document.removeEventListener('keydown',escClose);if(D)renderCard()}
function imgRow(){
  const T=D.templates||[],v=$('fu-tpl').value,t=v==='rr'?T.find(x=>x.img):T.find(x=>String(x.id)===v);
  const r=$('fu-imgrow');r.hidden=!(t&&t.img);
  if(t&&t.img){$('fu-img').src=t.img;r.dataset.url=t.img}
  $('fu-paste').textContent=$('fu-mode').value==='web'?'tekan Ctrl+V di kolom chat':'tahan kolom pesan → Tempel';
}
function msg(t,k){const m=$('fu-msg');if(!m)return;m.hidden=!t;m.className='fu-msg '+(k||'info');m.textContent=t||''}
function renderPop(){
  if(!$('fu-ovl'))return;
  const s=stats();
  $('fu-pper').textContent=`${D.sid||''} · ${nf.format(s.total)} customer · terkirim hari ini ${D.sent_today||0}/${D.limit||40}`;
  const TB=[['todo','Belum dihubungi',s.todo],['done','Sudah terkirim',s.done],['today','Terkirim hari ini',s.today]];
  $('fu-tabs').innerHTML=TB.map(([v,l,n])=>`<button type="button" data-t="${v}" aria-pressed="${TAB===v}">${l} (${nf.format(n)})</button>`).join('');
  $('fu-tabs').querySelectorAll('[data-t]').forEach(b=>b.onclick=()=>{TAB=b.dataset.t;SHOW=60;renderPop()});
  renderList();renderQueue();
}
function filtered(){
  const q=Q.trim().toLowerCase(),qd=digits(Q);
  return (D.customers||[]).filter(c=>{
    if(TAB==='todo'&&contacted(c))return false;if(TAB==='done'&&!contacted(c))return false;
    if(TAB==='today'&&!sentToday(c))return false;if((TAB==='Deal'||TAB==='Batal')&&c.st!==TAB)return false;
    if(q&&!String(c.n).toLowerCase().includes(q)&&!(qd.length>=3&&(digits(c.hp).includes(qd)||String(c.m||'').includes(qd))))return false;
    return true;
  }).sort((a,b)=>TAB==='todo'?String(a.n).localeCompare(String(b.n),'id'):String(b.lc||'').localeCompare(String(a.lc||'')));
}
function renderList(){
  const L=filtered(),part=L.slice(0,SHOW),full=!left();
  $('fu-list').innerHTML=(part.length?part.map(c=>{const wa=waNum(c.hp),st=sentToday(c);
    const done=contacted(c);
    return `<div class="fu-row${done?' done':''}" data-id="${c.id}"><div class="who"><b>${esc(c.n)}</b>
      <small>${c.m?'No. member '+esc(c.m)+' · ':''}${wa?esc(c.hp):'tanpa nomor'}</small></div>
      <div class="acts">${done?`<span class="fu-sent">✓ Sudah terkirim${c.lc?' · '+esc(fdt(c.lc)):''}</span>`:(wa?`<button type="button" class="wa" data-wa="${c.id}"${full?' disabled title="Kuota hari ini habis"':''}>Kirim WA</button>`:'<span class="note" style="margin:0">Tanpa nomor HP</span>')}</div></div>`}).join('')
    :'<p class="note" style="text-align:center;padding:16px">Tidak ada customer di daftar ini.</p>')
    +(L.length>SHOW?`<button type="button" class="btn ghost" id="fu-more" style="align-self:center">Tampilkan lebih banyak (${nf.format(L.length-SHOW)} lagi)</button>`:'');
  $('fu-list').querySelectorAll('[data-wa]').forEach(b=>b.onclick=()=>send(+b.dataset.wa,b));
  const m=$('fu-more');if(m)m.onclick=()=>{SHOW+=60;renderList()};
}

/* ---------- Kirim ---------- */
async function send(id,btn){
  const c=(D.customers||[]).find(x=>x.id===id);if(!c)return false;
  const num=waNum(c.hp);if(!num){msg('Customer ini tidak punya nomor HP yang valid.','err');return false}
  if(!left()){msg(`Kuota ${D.limit} pesan hari ini sudah habis. Lanjutkan besok supaya nomor WhatsApp Anda tidak diblokir.`,'err');return false}
  const t=pickTpl();if(!t){msg('Belum ada template pesan aktif. Hubungi admin.','err');return false}
  const text=renderMsg(t.isi,c);
  const ph=text.match(/\[[A-Z0-9 /.\-–]{2,}\]/g);if(ph){msg(`Template "${t.judul}" belum lengkap (${[...new Set(ph)].join(', ')}). Hubungi admin.`,'err');return false}
  const mode=$('fu-mode')?$('fu-mode').value:waMode();
  if(btn)btn.disabled=true;
  let copied=null,copyP=null;
  if(t.img){if(mode==='web')copied=await Promise.race([copyImg(t.img).then(()=>true,()=>false),new Promise(r=>setTimeout(()=>r(false),2500))]);else copyP=copyImg(t.img).then(()=>true,()=>false)}
  const w=window.open('',mode==='web'?'wa_blast':'_blank');
  let r;
  try{const {data,error}=await SB.rpc('fu_sp_send',{p_key:K,p_code:C,p_customer_id:c.id,p_template_id:t.id,p_text:text});if(error)throw error;r=data}
  catch(e){r={ok:false,err:'net'}}
  if(!r||!r.ok){
    try{if(w&&w.location.href==='about:blank')w.close()}catch(e){}
    if(btn)btn.disabled=false;
    const why=r&&r.err==='limit'?`Kuota ${r.limit||D.limit} pesan hari ini sudah habis. Lanjutkan besok.`:r&&r.err==='maintenance'?'Fitur sedang dalam pemeliharaan.':'Gagal mencatat, coba lagi.';
    if(r&&r.err==='limit'){D.sent_today=r.sent_today||D.limit}
    msg(why,'err');renderPop();return false;
  }
  const url=(mode==='web'?'https://web.whatsapp.com/send?phone=':'https://api.whatsapp.com/send?phone=')+num+'&text='+encodeURIComponent(text);
  if(w)w.location.href=url;else location.href=url;
  if(copyP)copied=await copyP;
  D.sent_today=r.sent_today;c.lc=new Date().toISOString();if(c.st==='Baru')c.st='Dihubungi';
  const paste=mode==='web'?'tekan Ctrl+V di kolom chat':'tahan kolom pesan → Tempel';
  msg(`Pesan ke ${c.n} siap dikirim (${r.sent_today}/${r.limit}). `+(copied===true?`Gambar sudah tersalin: ${paste}, lalu tekan Kirim.`:copied===false?'Gambar tidak bisa disalin otomatis, tekan "Salin gambar" lalu tempel manual.':'Tekan Kirim di WhatsApp.'),copied===false?'err':'ok');
  renderPop();renderCard();
  return true;
}
async function setStatus(id,st,btn){
  const c=(D.customers||[]).find(x=>x.id===id);if(!c)return;
  btn.disabled=true;
  const {data,error}=await SB.rpc('fu_sp_status',{p_key:K,p_code:C,p_customer_id:id,p_status:st,p_note:null});
  if(error||!data||!data.ok){btn.disabled=false;msg('Gagal mengubah status.','err');return}
  c.st=st;msg(`${c.n} ditandai ${st}.`,'ok');renderPop();
}

/* ---------- Antrean ---------- */
function startQueue(){
  const n=Math.min(left(),40);
  if(!n){msg(`Kuota ${D.limit} pesan hari ini sudah habis.`,'err');return}
  const ids=(D.customers||[]).filter(c=>!contacted(c)&&waNum(c.hp)).sort((a,b)=>String(a.n).localeCompare(String(b.n),'id')).slice(0,n).map(c=>c.id);
  if(!ids.length){msg('Semua customer sudah dihubungi.','info');return}
  QQ={ids,i:0,sent:0,skip:0,wait:0,busy:false,tm:null};
  msg(`Antrean ${ids.length} customer siap. Klik "Buka chat pertama". Setelah menekan Kirim di WhatsApp, kembali ke halaman ini untuk chat berikutnya.`,'info');
  renderQueue();
}
function renderQueue(){
  const b=$('fu-q');if(!b)return;if(!QQ){b.hidden=true;return}
  b.hidden=false;const n=QQ.ids.length,i=QQ.i,done=i>=n,c=done?null:(D.customers||[]).find(x=>x.id===QQ.ids[i]);
  b.innerHTML=`<div class="qrow"><div>${done?`<b>Selesai.</b> ${QQ.sent} pesan dibuka, ${QQ.skip} dilewati.`:`<b>${i+1} dari ${n}</b> · berikutnya: <b>${esc(c?c.n:'')}</b>`}</div>
    <div class="row">${done?'':`<button type="button" class="btn wa" id="fq-next"${QQ.wait>0||QQ.busy?' disabled':''}>${QQ.wait>0?`Jeda aman… ${QQ.wait} dtk`:(i===0?'Buka chat pertama':'Buka chat berikutnya')}</button><button type="button" class="btn ghost" id="fq-skip">Lewati</button>`}<button type="button" class="btn ghost" id="fq-stop">${done?'Tutup':'Berhenti'}</button></div></div>
    <div class="bar"><i style="width:${n?Math.round(i/n*100):0}%"></i></div>`;
  const nx=$('fq-next');if(nx)nx.onclick=async()=>{
    if(!QQ||QQ.busy||QQ.wait>0)return;QQ.busy=true;renderQueue();
    const ok=await send(QQ.ids[QQ.i]);if(!QQ)return;QQ.busy=false;
    if(!ok){if(!left()){QQ.i=QQ.ids.length}renderQueue();return}
    QQ.sent++;QQ.i++;
    if(QQ.i<QQ.ids.length&&left()){QQ.wait=GAP[0]+Math.floor(Math.random()*(GAP[1]-GAP[0]+1));renderQueue();
      QQ.tm=setInterval(()=>{if(!QQ){return}QQ.wait--;if(QQ.wait<=0){clearInterval(QQ.tm);QQ.wait=0}renderQueue()},1000)}
    else{QQ.i=QQ.ids.length;renderQueue()}
  };
  const sk=$('fq-skip');if(sk)sk.onclick=()=>{QQ.skip++;QQ.i++;renderQueue()};
  $('fq-stop').onclick=()=>{if(QQ)clearInterval(QQ.tm);QQ=null;renderQueue()};
}

window.addEventListener('sp:render',e=>{const d=e.detail||{};const same=K===d.key&&C===d.code;K=d.key||'';C=d.code||'';SPNAME=d.name||'';if(!same||!D)load()});
const sw=$('switch');if(sw)sw.addEventListener('click',()=>{D=null;K='';C='';const c=$('fu-card');if(c)c.hidden=true;closePop()});
const rl=$('reload');if(rl)rl.addEventListener('click',()=>setTimeout(load,300));
})();
