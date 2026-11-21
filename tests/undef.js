// Pemeriksa sederhana: identifier yang dipakai tetapi tidak pernah dideklarasikan di file mana pun (kecuali global browser)
const req=n=>{try{return require(n)}catch(e){return require('/opt/npm-tools/node_modules/'+n)}};
const acorn=req('acorn'),walk=req('acorn-walk'),fs=require('fs'),path=require('path');
const root=process.argv[2]||'.';
const G=new Set(('window document console Math JSON Date Array Object String Number Boolean RegExp Map Set WeakMap WeakSet Promise Symbol Error TypeError RangeError parseInt parseFloat isNaN isFinite undefined NaN Infinity setTimeout clearTimeout setInterval clearInterval requestAnimationFrame cancelAnimationFrame fetch URL URLSearchParams Blob File FileReader FormData Image Intl navigator location localStorage sessionStorage history screen performance atob btoa encodeURIComponent decodeURIComponent encodeURI decodeURI escape unescape Uint8Array Uint16Array Uint32Array Int32Array Float64Array Float32Array ArrayBuffer DataView TextEncoder TextDecoder AbortController Event CustomEvent MutationObserver ResizeObserver IntersectionObserver DOMParser XMLSerializer HTMLElement Node alert confirm prompt getComputedStyle matchMedia open XLSX supabase html2canvas CSS crypto queueMicrotask structuredClone BroadcastChannel Notification Reflect Proxy BigInt globalThis module require exports arguments self top parent frames name event OffscreenCanvas createImageBitmap ClipboardItem ExcelJS DomShot startPresence addEventListener Headers Request Response WebSocket EventSource Worker SharedWorker ImageData Path2D SVGElement Element Document getSelection scrollTo scrollBy innerWidth innerHeight devicePixelRatio pageYOffset pageXOffset visualViewport APP_CONFIG showSaveFilePicker caches indexedDB').split(' '));
const files=[];(function w(d){for(const f of fs.readdirSync(d)){const p=path.join(d,f);if(fs.statSync(p).isDirectory()){if(!/node_modules|\.git$/.test(f))w(p)}else if(/\.js$/.test(f)&&!/\.min\./.test(f))files.push(p)}})(root);
// skrip inline html
const inl=[];for(const f of fs.readdirSync(root).filter(f=>f.endsWith('.html'))){const t=fs.readFileSync(path.join(root,f),'utf8');let m,re=/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g;while(m=re.exec(t)){if(m[1].trim())inl.push({f:f+'#inline',code:m[1]})}}
const srcs=files.map(f=>({f,code:fs.readFileSync(f,'utf8')})).concat(inl);
let bad=0;
for(const {f,code} of srcs){
  let ast;try{ast=acorn.parse(code,{ecmaVersion:'latest',sourceType:'script',allowReturnOutsideFunction:true,locations:true})}catch(e){console.log('SINTAKS',f,e.message);bad++;continue}
  const decl=new Set(),refs=[];
  const addPat=p=>{if(!p)return;if(p.type==='Identifier')decl.add(p.name);else if(p.type==='ObjectPattern')p.properties.forEach(q=>addPat(q.value||q.argument));else if(p.type==='ArrayPattern')p.elements.forEach(addPat);else if(p.type==='RestElement')addPat(p.argument);else if(p.type==='AssignmentPattern')addPat(p.left)};
  walk.full(ast,n=>{
    if(n.type==='VariableDeclarator')addPat(n.id);
    if(n.type==='FunctionDeclaration'||n.type==='FunctionExpression'||n.type==='ArrowFunctionExpression'){if(n.id)decl.add(n.id.name);n.params.forEach(addPat)}
    if(n.type==='ClassDeclaration'&&n.id)decl.add(n.id.name);
    if(n.type==='CatchClause'&&n.param)addPat(n.param);
  });
  walk.fullAncestor(ast,(n,_s,anc)=>{
    if(n.type!=='Identifier')return;const p=anc[anc.length-2];if(!p)return;
    if(p.type==='MemberExpression'&&p.property===n&&!p.computed)return;
    if(p.type==='Property'&&p.key===n&&!p.computed&&!p.shorthand)return;
    if((p.type==='MethodDefinition'||p.type==='PropertyDefinition')&&p.key===n)return;
    if(p.type==='LabeledStatement'||p.type==='BreakStatement'||p.type==='ContinueStatement')return;
    if(p.type==='VariableDeclarator'&&p.id===n)return;
    if((p.type==='FunctionDeclaration'||p.type==='FunctionExpression')&&p.id===n)return;
    refs.push(n);
  });
  const seen=new Set();
  for(const r of refs){if(decl.has(r.name)||G.has(r.name)||seen.has(r.name))continue;seen.add(r.name);console.log('TAK-DIDEKLARASI',f+':'+r.loc.start.line,r.name);bad++}
}
console.log(bad?('Total temuan: '+bad):'OK: tidak ada identifier tak dikenal / kesalahan sintaks ('+srcs.length+' skrip)');
process.exit(bad?1:0);
