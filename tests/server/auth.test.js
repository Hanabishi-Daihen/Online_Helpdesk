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
const lastCode=()=>{const m=mails[mails.length-1];return m&&(/(\d{6})/.exec(m.subject)||[])[1]};
const skipWait=()=>Object.keys(cache).filter(k=>/^otpn_.*_t$/.test(k)).forEach(k=>delete cache[k]);
let c=mk();
// ── the door ──
T('unknown function refused',/not allowed/.test(E(()=>c.api('', 'eval',['1+1']))));
T('private function refused',/not allowed/.test(E(()=>c.api('', 'admHash_',['a','b']))));
T('page function needs a session',/AUTH_REQUIRED/.test(E(()=>c.api('', 'getPulse',[]))));
c=mk();T('direct call (console) refused',/AUTH_REQUIRED/.test(E(()=>c.getPulse())));
c=mk();T('auth functions are open',c.api('', 'authInfo',[]).on===true);
// ── register: weak passwords ──
T('short password',/8–64/.test(E(()=>c.api('', 'authRegisterStart',['somchai@example.co.th','Somchai','DS13','abc12']))));
T('letters + digits',/ตัวอักษรและตัวเลข/.test(E(()=>c.api('', 'authRegisterStart',['somchai@example.co.th','Somchai','DS13','abcdefgh']))));
T('bad e-mail',/อีเมล/.test(E(()=>c.api('', 'authRegisterStart',['somchai@','Somchai','DS13','Abcd1234x']))));
// ── register → pending (approval on, no auto domain) ──
let r=c.api('', 'authRegisterStart',['Somchai@Example.co.th','Somchai K.','DS13','Somchai2026']);
T('code sent',r.sent&&mails.length===1&&mails[0].to==='somchai@example.co.th'&&/^\d{6}$/.test(lastCode()),r);
T('resend within a minute refused',/รอ 1 นาที/.test(E(()=>c.api('', 'authRegisterStart',['somchai@example.co.th','Somchai K.','DS13','Somchai2026']))));
T('wrong code',/ไม่ถูกต้อง/.test(E(()=>c.api('', 'authRegisterVerify',['somchai@example.co.th','000000']))));
r=c.api('', 'authRegisterVerify',['somchai@example.co.th',lastCode(),true,'test']);
T('needs Admin approval',r.pending===true&&sheets.Users.rows[1][3]==='pending',r);
T('pending cannot sign in',/รอ Admin/.test(E(()=>c.api('', 'authLogin',['somchai@example.co.th','Somchai2026',true]))));
// ── Admin approves ──
cache['adm_ADM']=JSON.stringify({u:'admin1',n:'A.Admin'});
T('users list needs Admin',/Admin only/.test(E(()=>c.authUsers('x'))));
r=c.authUserAction('ADM','somchai@example.co.th','approve');T('approve',r.users[0].status==='active',r);
r=c.api('', 'authLogin',['somchai@example.co.th','Somchai2026',true,'pc']);
T('sign-in answer carries the start-up data (settings, templates, schedules, pulse, versions)',!!(r.boot&&r.boot.r&&r.boot.r.getSettings&&r.boot.r.getPublicTemplates&&'getSchedules' in r.boot.r&&r.boot.r.getPulse&&r.boot.v),r.boot&&Object.keys(r.boot.r||{}));
T('sign in',r.token&&r.user.name==='Somchai K.'&&r.user.email==='somchai@example.co.th',r);
const tok=r.token;
T('session hash only in the sheet',!JSON.stringify(sheets.Sessions.rows).includes(tok));
T('page function with session',(()=>{const x=mk();x.getPulse=()=>({ok:1});const v=x.api(tok,'getPulse',[]);return v.ok===1&&x.__user.e==='somchai@example.co.th'})());
T('session survives a cache flush (read from the sheet)',(()=>{Object.keys(cache).filter(k=>k.startsWith('ses_')).forEach(k=>delete cache[k]);return mk().authMe(tok).user.name==='Somchai K.'})());
// ── lockout ──
for(let i=0;i<5;i++)c.api('', 'authLogin',['somchai@example.co.th','wrong'+i]);
T('5 wrong passwords lock the e-mail',/5 ครั้ง/.test(E(()=>c.api('', 'authLogin',['somchai@example.co.th','Somchai2026']))));
// ── forgot password ──
const n0=mails.length;r=c.api('', 'authForgotStart',['nobody@example.co.th']);T('unknown e-mail: same answer, no mail',r.sent&&mails.length===n0,r);
skipWait();r=c.api('', 'authForgotStart',['somchai@example.co.th']);T('forgot sends a code',r.sent&&mails.length===n0+1&&/รหัสผ่านใหม่/.test(mails[mails.length-1].subject));
T('weak new password',/8–64/.test(E(()=>c.api('', 'authForgotReset',['somchai@example.co.th',lastCode(),'x1']))));
r=c.api('', 'authForgotReset',['somchai@example.co.th',lastCode(),'NewPass2026',false]);
T('reset gives a new session',!!r.token,r);
T('old session signed out',mk().authMe(tok).none===true);
T('lock cleared, new password works',!!c.api('', 'authLogin',['somchai@example.co.th','NewPass2026']).token);
T('old password refused',/ไม่ถูกต้อง/.test(E(()=>c.api('', 'authLogin',['somchai@example.co.th','Somchai2026']))));
// ── auto-approve domain ──
cache['adm_ADM']=JSON.stringify({u:'admin1',n:'A.Admin'});c.authSaveCfg('ADM',{on:true,domains:'@example.co.th',approve:true});
skipWait();c.api('', 'authRegisterStart',['malee@example.co.th','Malee','DS11','Malee2026x']);r=c.api('', 'authRegisterVerify',['malee@example.co.th',lastCode()]);
T('company domain is active at once',!!r.token&&r.user.status==='active',r);
skipWait();c.api('', 'authRegisterStart',['guest@gmail.com','Guest','','Guest2026x']);r=c.api('', 'authRegisterVerify',['guest@gmail.com',lastCode()]);
T('other domain waits for approval',r.pending===true,r);
// ── change password / sign out ──
const t2=c.api('', 'authLogin',['malee@example.co.th','Malee2026x',true]).token,t3=c.api('', 'authLogin',['malee@example.co.th','Malee2026x',true]).token;
T('change password needs the old one',/เดิมไม่ถูกต้อง/.test(E(()=>c.api(t2,'authChangePassword',[t2,'nope','Malee2027x']))));
r=c.api(t2,'authChangePassword',[t2,'Malee2026x','Malee2027x']);T('change password',r.ok===true,r);
T('other devices signed out, this one kept',mk().authMe(t3).none===true&&!!mk().authMe(t2).user);
c.api(t2,'authLogout',[t2]);T('sign out',mk().authMe(t2).none===true);
// ── Admin account signs in on the same form ──
c=mk();c.adminLogin=()=>({token:'AT',user:'admin1',name:'A.Admin',names:['A.Admin']});c.admUsers_=()=>[{u:'admin1',n:'A.Admin'}];
r=c.api('', 'authLogin',['admin1','whatever',true]);T('Admin user name signs in and unlocks Admin',!!r.token&&r.admin.token==='AT'&&r.user.admin===true,r);
// ── Admin actions ──
cache['adm_ADM']=JSON.stringify({u:'admin1',n:'A.Admin'});
r=c.authUserAction('ADM','malee@example.co.th','reset');T('reset gives a temporary password',/^Dh/.test(r.temp)&&!!c.api('', 'authLogin',['malee@example.co.th',r.temp]).token,r);
c.authUserAction('ADM','malee@example.co.th','disable');T('disabled cannot sign in',/ระงับ/.test(E(()=>c.api('', 'authLogin',['malee@example.co.th',r.temp]))));
// ── role: a registered user made Admin ──
c.authUserAction('ADM','malee@example.co.th','enable');r=c.authUserAction('ADM','malee@example.co.th','admin');
T('role set to admin',r.users.find(u=>u.email==='malee@example.co.th').role==='admin');
const pwM=c.authUserAction('ADM','malee@example.co.th','reset').temp;r=c.api('', 'authLogin',['malee@example.co.th',pwM,true]);
T('Admin role signs in with Admin mode',!!r.token&&r.user.admin===true&&r.admin&&r.admin.token==='session'&&r.admin.names.includes('Malee'),r);
const tA=r.token;T('Admin role: Admin functions work without an Admin token',(()=>{const x=mk();return Array.isArray(x.api(tA,'authUsers',[''] ).users)})());
T('authMe gives the role back after reload',(()=>{const m=mk().authMe(tA);return m.user.admin===true&&m.admin.name==='Malee'})());
const tU=c.api('', 'authLogin',['somchai@example.co.th','NewPass2026',true]).token;
T('normal user is refused Admin functions',/Admin only/.test(E(()=>mk().api(tU,'authUsers',['']))));
c.authUserAction('ADM','malee@example.co.th','user');T('back to user: old Admin session signed out',mk().authMe(tA).none===true);
// ── sign-in off ──
c.authSaveCfg('ADM',{on:false});c=mk();c.getPulse=()=>({ok:2});T('sign-in off: pages work without a session',c.api('', 'getPulse',[]).ok===2);
T('sign-in off: console call still refused',/AUTH_REQUIRED/.test(E(()=>mk().getSchedules())));
// ── every page call is on the list ──
const html=fs.readFileSync(path.join(ROOT,'index.html'),'utf8'),pub=new Set();fs.readdirSync(ROOT).filter(f=>f.endsWith('.gs')).forEach(f=>{for(const m of fs.readFileSync(path.join(ROOT,f),'utf8').matchAll(/^function ([A-Za-z0-9]+)\(/gm))pub.add(m[1])});
const lits=new Set([...html.matchAll(/["']([A-Za-z][A-Za-z0-9]+)["']/g)].map(m=>m[1]));
const miss=[...pub].filter(f=>lits.has(f)&&!c.API_FNS[f]&&!c.AUTH_OPEN[f]&&!c.AUTH_FNS[f]&&!['doGet','api','getOrderApprovalView','actOrderApproval'].includes(f));
T('every function the page names is on API_FNS',!miss.length,miss);
console.log(ok+' passed, '+bad+' failed');process.exit(bad?1:0);
