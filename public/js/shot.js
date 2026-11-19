/* Print screen HD (PNG) untuk kartu & pop-up — tanpa pustaka luar.
   window.DomShot(el, namaFile, {scale, btn}) : menyalin elemen (gaya terhitung disalin inline), memperluas tabel
   yang ter-scroll agar tampil utuh, lalu merender ke PNG lewat SVG foreignObject pada skala HD (default 3x). */
(function(){
  var SKIP={'width':1,'height':1,'inline-size':1,'block-size':1,'min-width':1,'min-height':1,'max-width':1,'max-height':1,'perspective-origin':1,'transform-origin':1};
  var MON=['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];
  function stamp(){ var d=new Date(), p=function(n){ return String(n).padStart(2,'0'); }; return d.getDate()+' '+MON[d.getMonth()]+' '+d.getFullYear()+' '+p(d.getHours())+':'+p(d.getMinutes()); }
  async function save(blob,name){
    var dl=null; try{ dl=(window.claude&&window.claude.use)?await window.claude.use('downloads'):null; }catch(e){}
    if(dl){ await dl.save({filename:name,data:blob}); return; }
    var a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=name; document.body.appendChild(a); a.click(); a.remove(); setTimeout(function(){ URL.revokeObjectURL(a.href); },4000);
  }
  var FR=null, DEF={};
  function defs(tag){
    if(DEF[tag]) return DEF[tag];
    if(!FR){ FR=document.createElement('iframe'); FR.setAttribute('aria-hidden','true'); FR.style.cssText='position:absolute;left:-100000px;top:0;width:800px;height:600px;border:0;visibility:hidden'; document.body.appendChild(FR); FR.contentDocument.open(); FR.contentDocument.write('<!doctype html><html><body></body></html>'); FR.contentDocument.close(); }
    var d=FR.contentDocument, host=d.body, e=d.createElement(tag); host.appendChild(e);
    var cs=FR.contentWindow.getComputedStyle(e), m={}; for(var i=0;i<cs.length;i++){ m[cs[i]]=cs.getPropertyValue(cs[i]); } host.removeChild(e); return DEF[tag]=m;
  }
  function inline(src,dst){
    var cs=getComputedStyle(src), df=defs(src.tagName.toLowerCase()), css='';
    for(var i=0;i<cs.length;i++){ var k=cs[i]; if(SKIP[k]) continue; var v=cs.getPropertyValue(k); if(df[k]===v) continue; css+=k+':'+v+';'; }
    dst.setAttribute('style',css);
    if(cs.position==='sticky'||cs.position==='fixed') dst.style.position='static';
    if(cs.overflowY!=='visible'||cs.maxHeight!=='none'){ dst.style.overflow='visible'; dst.style.maxHeight='none'; dst.style.height='auto'; }
    if(src.tagName==='TABLE'&&src.parentElement&&src.offsetWidth>=src.parentElement.clientWidth-2) dst.style.width='100%';
    if(src.tagName==='SELECT'){ var t=document.createElement('b'); t.textContent=src.options[src.selectedIndex]?src.options[src.selectedIndex].text:''; t.setAttribute('style','padding:6px 0;font-weight:600'); return t; }
    var sk=src.children, dk=dst.children;
    for(var j=0;j<sk.length&&j<dk.length;j++){ var r=inline(sk[j],dk[j]); if(r&&r!==dk[j]) dst.replaceChild(r,dk[j]); }
    return dst;
  }
  window.DomShot=async function(el,name,o){
    o=o||{}; var btn=o.btn, old=btn?btn.textContent:'', host=null;
    if(btn){ btn.disabled=true; btn.textContent='Membuat gambar…'; }
    try{
      try{ await document.fonts.ready; }catch(e){}
      var need=el.offsetWidth; el.querySelectorAll('table').forEach(function(t){ need=Math.max(need,t.scrollWidth+44); });
      var W=Math.ceil(Math.max(el.offsetWidth,need)), SC=o.scale||3;
      var clone=el.cloneNode(true); var top=inline(el,clone); if(top!==clone) clone=top;
      clone.style.width=W+'px'; clone.style.maxWidth='none'; clone.style.margin='0'; clone.style.boxShadow='none'; clone.style.borderRadius='0'; clone.style.position='static'; clone.style.transform='none'; clone.style.overflow='visible'; clone.style.height='auto'; clone.style.background=clone.style.backgroundColor&&clone.style.backgroundColor!=='rgba(0, 0, 0, 0)'?clone.style.backgroundColor:'#fff';
      clone.querySelectorAll('[data-noshot]').forEach(function(n){ n.remove(); });
      clone.querySelectorAll('[hidden]').forEach(function(n){ n.remove(); });
      var f=document.createElement('div'); f.setAttribute('style','padding:10px 16px;font:12px/1.4 system-ui,Segoe UI,Arial,sans-serif;color:#5b6577;border-top:1px solid #e2e6ee;background:#fff;box-sizing:border-box'); f.textContent='MITRA10 Percetakan Negara · dicetak '+stamp(); clone.appendChild(f);
      host=document.createElement('div'); host.setAttribute('aria-hidden','true'); host.style.cssText='position:absolute;left:-100000px;top:0;width:'+W+'px;background:#fff;pointer-events:none'; host.appendChild(clone); document.body.appendChild(host);
      var H=Math.ceil(Math.max(clone.scrollHeight,clone.offsetHeight));
      if(W*SC>16000||H*SC>16000) SC=Math.max(1,Math.floor(16000/Math.max(W,H)));
      clone.setAttribute('xmlns','http://www.w3.org/1999/xhtml');
      var xml=new XMLSerializer().serializeToString(clone);
      var svg='<svg xmlns="http://www.w3.org/2000/svg" width="'+W+'" height="'+H+'" viewBox="0 0 '+W+' '+H+'"><foreignObject x="0" y="0" width="100%" height="100%">'+xml+'</foreignObject></svg>';
      var img=new Image(); await new Promise(function(ok,no){ img.onload=ok; img.onerror=function(){ no(new Error('browser tidak bisa merender gambar')); }; img.src='data:image/svg+xml;charset=utf-8,'+encodeURIComponent(svg); });
      var cv=document.createElement('canvas'); cv.width=Math.round(W*SC); cv.height=Math.round(H*SC); var cx=cv.getContext('2d'); cx.fillStyle='#fff'; cx.fillRect(0,0,cv.width,cv.height); cx.imageSmoothingQuality='high'; cx.drawImage(img,0,0,cv.width,cv.height);
      var blob=await new Promise(function(ok,no){ cv.toBlob(function(b){ b?ok(b):no(new Error('gagal membuat gambar')); },'image/png'); });
      await save(blob,name);
      return {w:cv.width,h:cv.height,bytes:blob.size};
    }catch(err){ console.error(err); try{ window.alert('Print screen gagal: '+(err&&err.message||err)); }catch(e){} }
    finally{ if(host&&host.parentNode) host.parentNode.removeChild(host); if(FR&&FR.parentNode){ FR.parentNode.removeChild(FR); FR=null; DEF={}; } if(btn){ btn.disabled=false; btn.textContent=old; } }
  };
  window.DomShot.slug=function(s){ return String(s||'gambar').toLowerCase().replace(/[^a-z0-9]+/g,'-').replace(/^-+|-+$/g,'').slice(0,60)||'gambar'; };
})();
