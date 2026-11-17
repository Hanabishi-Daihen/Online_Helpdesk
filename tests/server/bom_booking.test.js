// Sign-in: register with an e-mail code, approval, sign in, forgot password, sessions, the api() door
const vm=require('vm'),fs=require('fs'),path=require('path'),crypto=require('crypto');
const ROOT=path.resolve(__dirname,'../..');
const src=fs.readdirSync(ROOT).filter(f=>f.endsWith('.gs')).map(f=>fs.readFileSync(path.join(ROOT,f),'utf8')).join('\n');
const mkSheet=()=>{let rows=[];const sh={rows,getLastRow:()=>rows.length,getMaxColumns:()=>26,appendRow:r=>{rows.push(r.map(String))},deleteRow:r=>{rows.splice(r-1,1)},setFrozenRows(){},
 getRange(r,c,nr,nc){nr=nr||1;nc=nc||1;const R={setValues(v){v.forEach((row,i)=>{while(rows.length<r+i)rows.push([]);row.forEach((x,j)=>rows[r+i-1][c+j-1]=String(x))});return R},
  getValues:()=>Array.from({length:nr},(_,i)=>Array.from({length:nc},(_,j)=>(rows[r+i-1]||[])[c+j-1]==null?'':(rows[r+i-1]||[])[c+j-1])),getDisplayValues(){return R.getValues()},
  clearContent(){for(let i=0;i<nr;i++)for(let j=0;j<nc;j++)if(rows[r+i-1])rows[r+i-1][c+j-1]='';while(rows.length&&rows[rows.length-1].every(x=>x===''||x==null))rows.pop();return R},setNumberFormat:()=>R,setFontWeight:()=>R,setBackground:()=>R,setValue(v){return R.setValues([[v]])}};return R}};return sh};
const props={ADMIN_PASSCODE:'control'},cache={},sheets={},mails=[];
const SS={getSheetByName:n=>sheets[n]||null,insertSheet:n=>sheets[n]=mkSheet()};
const dig=(alg,s)=>Array.from(crypto.createHash(alg).update(typeof s==='string'?Buffer.from(s,'utf8'):Buffer.from(s)).digest()).map(b=>b>127?b-256:b);
const b64=a=>Buffer.from(a.map(b=>b&255)).toString('base64');
const mk=()=>{const ctx={console:{log(){},warn(){}},JSON,Math,Date,String,Number,Array,Object,Error,RegExp,
 PropertiesService:{getScriptProperties:()=>({getProperty:k=>k in props?props[k]:null,setProperty:(k,v)=>{props[k]=String(v)},deleteProperty:k=>{delete props[k]},getProperties:()=>({...props})})},
 Utilities:{formatDate:()=>'14/11/2026 09:00:00',getUuid:()=>crypto.randomUUID(),sleep(){},computeDigest:(a,s)=>dig(a,s),base64Encode:b64,base64EncodeWebSafe:a=>b64(a).replace(/\+/g,'-').replace(/\//g,'_'),DigestAlgorithm:{SHA_256:'sha256',MD5:'md5'},Charset:{UTF_8:'utf8'}},
 SpreadsheetApp:{openById:()=>SS,flush(){}},LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){},tryLock:()=>true})},
 MailApp:{sendEmail:o=>mails.push(o)},
 CacheService:{getScriptCache:()=>({get:k=>k in cache?cache[k]:null,put:(k,v)=>{cache[k]=String(v)},putAll:o=>Object.assign(cache,o),getAll:ks=>{const o={};ks.forEach(k=>{if(k in cache)o[k]=cache[k]});return o},remove:k=>{delete cache[k]},removeAll:ks=>ks.forEach(k=>delete cache[k])})}};
 ctx.globalThis=ctx;vm.createContext(ctx);vm.runInContext(src,ctx);ctx.audit_=()=>{};return ctx};
let ok=0,bad=0;const T=(n,c,x)=>{c?ok++:bad++;console.log((c?'PASS ':'FAIL ')+n+(c?'':' → '+JSON.stringify(x)))};
const E=f=>{try{const r=f();return r&&r.err?r.err:''}catch(e){return e.message}};
// ── Booking BOM: edit a project's CODE list · book in the system (cart) ──
// the shared sheet fake lacks a few calls Import BOM makes (grid size, bold header): no-ops here
const _ins=SS.insertSheet;SS.insertSheet=n=>{const sh=_ins(n),gr=sh.getRange.bind(sh);sh.getMaxRows=()=>Math.max(1000,sh.rows.length+1);sh.insertRowsAfter=()=>{};
  sh.getRange=(...a)=>{const R=gr(...a);return new Proxy(R,{get:(t,k)=>k in t?t[k]:()=>R})};return sh};
const c=mk(),audits=[];c.audit_=r=>audits.push(...r);c.__trusted=true;c.__user={e:'admin:a',n:'A.Admin',a:1};
c.bomSnapBuild_=()=>({projects:[],codes:[]});                                     // the calculation itself is not under test here
c.stockData_=()=>['42040008\tMiniature circuit breaker\tPLS6-C32/3\t5\tPC\tSTORE1','42046001\tMiniature circuit breaker\tPLS6-C6/2-DC\t3\tPCS\tSTORE1','43401027\tAUXILIARY RELAY\tRH4B-U\t10\tPCS\tSTORE1'];
const items=()=>c.bomItems_(),proj=k=>c.bomProjects_().filter(p=>p.project===k)[0];
// a project as Import BOM leaves it
c.importBom('session',[{project:'PP1039D01',title:'RCC',rev:'A',file:'bom.xlsx',items:[{code:'42040008',qty:1,material:'MCB 32A'},{code:'42046001',qty:3,material:'MCB 6A'},{code:'99999999',qty:2,material:'Old part'}]}]);
const sig0=c.bomSig_(items().PP1039D01);
// edit: qty change, swap a CODE, cancel a line
let r=c.bomSaveItems('session','PP1039D01',[{code:'42040008',qty:2,material:'MCB 32A'},{code:'43401027',qty:3},{code:'99999999',qty:0}],sig0,{'43401027':'42046001'},'เปลี่ยนรุ่นตามลูกค้า');
const it=items().PP1039D01||[];
T('edit: the new list replaces the old one',it.length===2&&it.find(x=>x.code==='42040008').qty===2&&it.find(x=>x.code==='43401027').qty===3&&!it.find(x=>x.code==='99999999'),it);
T('edit: a new CODE gets its name from Item stock',it.find(x=>x.code==='43401027').material==='AUXILIARY RELAY',it.find(x=>x.code==='43401027'));
T('edit: the changes are listed (qty, swap, cancelled)',r.changes.some(x=>/42040008 1→2/.test(x))&&r.changes.some(x=>/42046001 ⇄ 43401027/.test(x))&&r.changes.some(x=>/−99999999/.test(x)),r.changes);
T('edit: audit log + project note + item count',/แก้ไขรายการ/.test(audits.map(a=>a[1]).join())&&/เปลี่ยนรุ่นตามลูกค้า/.test(proj('PP1039D01').note)&&proj('PP1039D01').items===2,[proj('PP1039D01'),audits.slice(-1)]);
T('edit: someone else saved in between → refused',/ถูกแก้ไขโดยคนอื่น/.test(E(()=>c.bomSaveItems('session','PP1039D01',[{code:'42040008',qty:9}],sig0,null,''))));
T('edit: cannot empty the list (close / delete instead)',/อย่างน้อย 1 รายการ/.test(E(()=>c.bomSaveItems('session','PP1039D01',[{code:'42040008',qty:0}],null,null,''))));
{const a0=audits.length;let x=null,er='';try{x=c.bomSaveItems('session','PP1039D01',items().PP1039D01,c.bomSig_(items().PP1039D01),null,'')}catch(e){er=e.message}T('edit: nothing changed → no write, no Booking BOM audit',x&&!x.changes.length&&!audits.slice(a0).some(r=>/BOM|รายการ/.test(r[1])),[er,x&&x.changes,audits.slice(a0)])}
// book in the system
r=c.bomBook('session',{project:'pp1050b01',title:'LCC',needDate:'05/12/2026',mode:'new',items:[{code:'42040008',qty:2},{code:'42040008',qty:1},{code:'43401027',qty:4,remark:'สำรอง'}],note:'ทดสอบตะกร้า'});
const nb=items().PP1050B01||[],pn=proj('PP1050B01');
T('book new: project made (upper-case, from the cart, need date, "จองในระบบ")',r.created&&pn&&pn.file==='จองในระบบ'&&pn.needManual==='05/12/2026'&&pn.status==='active'&&pn.title==='LCC',pn);
T('book new: same CODE in the cart adds up',nb.length===2&&nb.find(x=>x.code==='42040008').qty===3,nb);
T('book new: a project that exists is refused (use "add")',/มีโปรเจกต์ PP1050B01 อยู่แล้ว/.test(E(()=>c.bomBook('session',{project:'PP1050B01',mode:'new',items:[{code:'42040008',qty:1}]}))));
r=c.bomBook('session',{project:'PP1039D01',mode:'add',items:[{code:'42040008',qty:1},{code:'42046001',qty:2}]});
const na=items().PP1039D01;
T('book add: quantities add to the project, new CODEs join',!r.created&&na.find(x=>x.code==='42040008').qty===3&&na.find(x=>x.code==='42046001').qty===2&&na.length===3,na);
T('book: other projects untouched',(items().PP1050B01||[]).length===2);
T('book: empty cart / bad project number / bad date refused',/ตะกร้ายังว่าง/.test(E(()=>c.bomBook('session',{project:'PP1',mode:'new',items:[]})))||/เลขโปรเจกต์/.test(E(()=>c.bomBook('session',{project:'PP1',mode:'new',items:[]}))));
T('book: bad need date refused',/วว\/ดด\/ปปปป/.test(E(()=>c.bomBook('session',{project:'PP2000A01',mode:'new',needDate:'2026-12-01',items:[{code:'42040008',qty:1}]}))));
T('Admin only: a signed-in user cannot book or edit',(()=>{const u=mk();u.__trusted=true;u.__user={e:'u@x',n:'U',a:0};return /Admin|admin|สิทธิ์/.test(E(()=>u.bomBook('',{project:'PP2000A01',mode:'new',items:[{code:'1',qty:1}]})))&&/Admin|admin|สิทธิ์/.test(E(()=>u.bomSaveItems('','PP1039D01',[{code:'1',qty:1}])))})());
c.bomProjectAction('session','PP1050B01','close');r=c.bomBook('session',{project:'PP1050B01',mode:'add',items:[{code:'42046001',qty:1}]});
T('book add to a closed project: reopened, closed time cleared',r.reopened&&proj('PP1050B01').status==='active'&&!proj('PP1050B01').closed,proj('PP1050B01'));
T('page may call both (allow-list)',!!c.API_FNS.bomBook&&!!c.API_FNS.bomSaveItems);
console.log(ok+' passed, '+bad+' failed');process.exit(bad?1:0);
