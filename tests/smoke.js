// Smoke test: memuat setiap halaman di Chromium dengan Supabase palsu (data kosong, peran admin) dan gagal bila ada error JS.
const http=require('http'),fs=require('fs'),path=require('path');
const req=n=>{try{return require(n)}catch(e){return require('/opt/npm-tools/node_modules/'+n)}};
let pw; try{pw=require('playwright')}catch(e){pw=require('playwright-core')}
const root=path.resolve(__dirname,'../public');
const mime={'.html':'text/html','.js':'text/javascript','.css':'text/css'};
const srv=http.createServer((q,r)=>{const f=path.join(root,decodeURIComponent(q.url.split('?')[0]).replace(/^\/+$/,'/index.html'));fs.readFile(f,(e,d)=>{if(e){r.statusCode=404;return r.end()}r.setHeader('content-type',mime[path.extname(f)]||'application/octet-stream');r.end(d)})});
const STUB=`(function(){
 const user={id:'u1',email:'test@example.com'};
 function qb(table){const st={table,count:null};
  const res=()=>({data:(st.single?null:[]),error:null,count:0});
  const p=new Proxy(function(){},{get(_,k){
    if(k==='then')return (ok,ko)=>Promise.resolve(st.single==='one'||st.single==='maybe'?(table==='user_roles'?{data:{role:'admin'},error:null}:{data:null,error:null}):res()).then(ok,ko);
    if(k==='maybeSingle')return ()=>{st.single='maybe';return p};
    if(k==='single')return ()=>{st.single='one';return p};
    return ()=>p;},apply(){return p}});
  return p;}
 window.supabase={createClient:()=>({
  auth:{getSession:async()=>({data:{session:{user,access_token:'x'}},error:null}),getUser:async()=>({data:{user},error:null}),onAuthStateChange:()=>({data:{subscription:{unsubscribe(){}}}}),signOut:async()=>({}),signInWithPassword:async()=>({data:{},error:{message:'stub'}})},
  from:qb,rpc:async()=>({data:null,error:null}),
  channel:()=>{const c={on:()=>c,subscribe:()=>c,track:async()=>{},unsubscribe(){},presenceState:()=>({})};return c},removeChannel(){}})};
})();`;
(async()=>{
  await new Promise(r=>srv.listen(0,r)); const port=srv.address().port;
  const browser=await pw.chromium.launch(); let bad=0;
  for(const pg of ['index.html','login.html','dashboard.html','monitoring.html','salesperson.html','followup.html']){
    const ctx=await browser.newContext({viewport:{width:1280,height:900}}), page=await ctx.newPage(), errs=[];
    page.on('pageerror',e=>errs.push('pageerror: '+e.message));
    page.on('console',m=>{ if(m.type()==='error'&&!/Failed to load resource|ERR_|net::/.test(m.text())) errs.push('console: '+m.text()); });
    await page.route('**/*',rt=>{const u=rt.request().url();
      if(u.startsWith('http://127.0.0.1')) return rt.continue();
      if(/supabase/.test(u)&&/\.js/.test(u)) return rt.fulfill({contentType:'text/javascript',body:STUB});
      if(/xlsx|html2canvas|exceljs/i.test(u)) return rt.fulfill({contentType:'text/javascript',body:'window.XLSX=window.XLSX||{utils:{},read(){throw new Error("stub")}};'});
      return rt.abort();});
    await page.goto('http://127.0.0.1:'+port+'/'+pg,{waitUntil:'load'}).catch(e=>errs.push('goto: '+e.message));
    await page.waitForTimeout(1500);
    const t=await page.evaluate(()=>document.body.innerText.length);
    if(t<20) errs.push('halaman kosong');
    console.log((errs.length?'GAGAL ':'OK    ')+pg); errs.forEach(e=>console.log('   - '+e)); if(errs.length) bad++;
    await ctx.close();
  }
  await browser.close(); srv.close(); process.exit(bad?1:0);
})();
