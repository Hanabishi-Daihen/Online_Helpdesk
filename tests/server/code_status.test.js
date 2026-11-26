// ทะเบียน CODE เลิกใช้: Admin writes reasons; Stock Alert / Stock ช่วยใช้ use them; Order Item refuses a newly added ⛔ CODE unless Admin confirms
const vm=require('vm'),fs=require('fs'),path=require('path'),crypto=require('crypto'),zlib=require('zlib');
const ROOT=path.resolve(__dirname,'../..');
const src=fs.readdirSync(ROOT).filter(f=>f.endsWith('.gs')).map(f=>fs.readFileSync(path.join(ROOT,f),'utf8')).join('\n');
const signed=b=>Array.from(b).map(x=>x>127?x-256:x),buf=a=>Buffer.from(a.map(x=>x&255));
const blob=(b,ty)=>({b,getBytes:()=>signed(b),getDataAsString:e=>!e||/utf/i.test(e)?b.toString('utf8'):/874/.test(e)?new TextDecoder('windows-874').decode(b):b.toString('latin1'),setContentType(){return this},getContentType:()=>ty});
const mkSheet=name=>{let rows=[];const sh={name,rows,getName:()=>name,setName(n){delete book.sheets[name];name=n;book.sheets[n]=sh},getLastRow:()=>rows.length,getMaxRows:()=>Math.max(1000,rows.length),insertRowsAfter(){},clearContents(){rows.length=0},
 getRange(r,c,nr,nc){nr=nr||1;nc=nc||1;const R={setValues(v){v.forEach((row,i)=>{while(rows.length<r+i)rows.push([]);row.forEach((x,j)=>rows[r+i-1][c+j-1]=x)});return R},setNumberFormat:()=>R,
  getValues:()=>Array.from({length:nr},(_,i)=>Array.from({length:nc},(_,j)=>{const v=(rows[r+i-1]||[])[c+j-1];return v==null?'':v}))};return R}};return sh};
const book={sheets:{},id:'SHIPSS'};const SS={getId:()=>book.id,getSheetByName:n=>book.sheets[n]||null,insertSheet:n=>book.sheets[n]=mkSheet(n),deleteSheet:sh=>{delete book.sheets[sh.getName()]},getSheets:()=>Object.values(book.sheets)};
let FILE=Buffer.from(''),ETAG='"1"',fetches=[],ranged=true;const props={},cache={},trig=[];
const mk=()=>{const ctx={console:{log(){},warn(){}},JSON,Math,Date,String,Number,Array,Object,Error,RegExp,isFinite,
 PropertiesService:{getScriptProperties:()=>({getProperty:k=>k in props?props[k]:null,setProperty:(k,v)=>{props[k]=String(v)},deleteProperty:k=>{delete props[k]},getProperties:()=>({...props})})},
 Utilities:{formatDate:()=>'18/11/2026 09:00:00',getUuid:()=>crypto.randomUUID(),sleep(){},DigestAlgorithm:{MD5:'md5',SHA_256:'sha256'},Charset:{UTF_8:'utf8'},
  computeDigest:(a,s)=>signed(crypto.createHash(a).update(typeof s==='string'?Buffer.from(s,'utf8'):buf(s)).digest()),base64Encode:a=>buf(a).toString('base64'),base64Decode:s=>signed(Buffer.from(s,'base64')),
  newBlob:(d,ty)=>blob(typeof d==='string'?Buffer.from(d,'utf8'):buf(d),ty),gzip:b=>blob(zlib.gzipSync(b.b)),ungzip:b=>blob(zlib.gunzipSync(b.b))},
 UrlFetchApp:{fetch(url,o){const m=/bytes=(\d+)-(\d+)/.exec(o.headers.Range||'');fetches.push(m?[+m[1],+m[2]]:null);
  if(!ranged||!m){return{getResponseCode:()=>200,getAllHeaders:()=>({ETag:ETAG}),getBlob:()=>blob(FILE)}}
  const a=+m[1],z=Math.min(+m[2],FILE.length-1);return{getResponseCode:()=>206,getAllHeaders:()=>({'Content-Range':'bytes '+a+'-'+z+'/'+FILE.length,ETag:ETAG}),getBlob:()=>blob(FILE.subarray(a,z+1))}}},
 SpreadsheetApp:{openById:()=>SS,create:()=>{book.sheets['Sheet1']=mkSheet('Sheet1');return SS},flush(){}},DriveApp:{getFileById:()=>({moveTo(){}}),getFolderById:()=>({})},
 ScriptApp:{getProjectTriggers:()=>trig.slice(),deleteTrigger:t=>trig.splice(trig.indexOf(t),1),newTrigger:h=>{const b={timeBased:()=>b,everyDays:()=>b,everyHours:()=>b,atHour:()=>b,nearMinute:()=>b,create:()=>{const t={getHandlerFunction:()=>h};trig.push(t);return t}};return b}},
 LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){},tryLock:()=>true})},
 CacheService:{getScriptCache:()=>({get:k=>k in cache?cache[k]:null,put:(k,v)=>{cache[k]=String(v)},putAll:o=>Object.assign(cache,o),getAll:ks=>{const o={};ks.forEach(k=>{if(k in cache)o[k]=cache[k]});return o},remove:k=>{delete cache[k]}})}};
 ctx.globalThis=ctx;vm.createContext(ctx);vm.runInContext(src,ctx);ctx.audit_=()=>{};ctx.__trusted=true;ctx.__user={e:'admin:a',n:'A.Admin',a:1};return ctx};
const _ins=SS.insertSheet;SS.insertSheet=n=>{const sh=_ins(n);sh.setFrozenRows=()=>{};sh.appendRow=r=>sh.rows.push(r.map(String));sh.deleteRow=r=>sh.rows.splice(r-1,1);sh.getMaxColumns=()=>26;const gr=sh.getRange.bind(sh);sh.getRange=(...a)=>{const R=gr(...a);R.getDisplayValues=()=>R.getValues().map(r=>r.map(String));const P=new Proxy(R,{get:(t,k)=>k in t?(typeof t[k]==='function'?(...x)=>{const v=t[k](...x);return v===t?P:v}:t[k]):()=>P});return P};return sh};
let ok=0,bad=0;const T=(n,c,x)=>{c?ok++:bad++;console.log((c?'PASS ':'FAIL ')+n+(c?'':' → '+JSON.stringify(x)))};
const c=mk();c.__trusted=true;c.__user={e:'admin:a',n:'A.Admin',a:1};const audits=[];c.audit_=r=>audits.push(...r);
const E=f=>{try{f();return''}catch(e){return e.message}};
// Item stock: 46001008 ACTIVE 0 · 46001009 active · 43400001 active (ticked เลิกใช้ in Stock Alert)
c.stockData_=()=>[['46001008','PILOT LAMP APW-116DD','',5,'PC','S1','','DS13',100,'','','0'],['46001009','PILOT LAMP APW-199DS','',120,'PC','S1','','DS13',120,'','','1'],['43400001','TRIP RELAY','',32,'PC','S1','','DS13',1000,'','','1'],['42046001','MCB C6','',10,'PC','S1','','DS13',100,'','','1']].map(r=>r.join('\t'));
c.saSettings_=()=>({'43400001':{code:'43400001',notUse:true,note:'ไม่ใช้แล้ว'}});
let g=c.getCodeStatus('tok');
T('without rows: Stock Alert tick shows up as retired, reason "not given yet"',g.map['43400001']&&g.map['43400001'].auto==='sa'&&g.reasons.length===9&&g.admin);
T('ACTIVE 0 is found when Item stock is scanned (orders)',c.csAll_(true)['46001008'].reason==='official');
g=c.saveCodeStatus('tok',{code:'46001008',level:'off',reason:'eol',detail:'ผู้ผลิตเลิกผลิตรุ่น DD',replace:'46001009, 46001008'});
const r=g.map['46001008'];
T('save: reason, detail, replacement (not itself), date today, by Admin',r.reason==='eol'&&r.replace==='46001009'&&/^\d{2}\/\d{2}\/\d{4}$/.test(r.since)&&r.by==='A.Admin'&&!r.auto,r);
T('save: audit names the CODE and reason',audits.some(a=>/บันทึกเหตุผล CODE เลิกใช้/.test(a[1])&&a[2]==='46001008'&&/ผู้ผลิตเลิกผลิต/.test(a[3])));
T('save: checks (reason, "other" needs detail, date, https link)',/เลือกเหตุผล/.test(E(()=>c.saveCodeStatus('tok',{code:'1',reason:'official'})))&&/รายละเอียด/.test(E(()=>c.saveCodeStatus('tok',{code:'1',reason:'other'})))&&/วว\/ดด/.test(E(()=>c.saveCodeStatus('tok',{code:'1',reason:'eol',since:'2026-01-01'})))&&/https/.test(E(()=>c.saveCodeStatus('tok',{code:'1',reason:'eol',doc:'http://x'}))));
c.saveCodeStatus('tok',{code:'42046001',level:'nonew',reason:'model',replace:'42046002'});
T('one row per CODE (saving again replaces)',(()=>{c.saveCodeStatus('tok',{code:'46001008',level:'off',reason:'model',replace:'46001009'});return c.getCodeStatus('tok').map['46001008'].reason==='model'&&book.sheets.CodeStatus.rows.filter(x=>x[0]==='46001008').length===1})());
// Order Item check
const it=codes=>codes.map(k=>({code:k,qty:1}));
T('order: a newly added ⛔ CODE is refused for a user, with the reason and the replacement',(()=>{c.__user={e:'u:s',n:'S.Somchai',a:0};const m=E(()=>c.csOrderCheck_(it(['46001008']),[],null,''));c.__user={e:'admin:a',n:'A.Admin',a:1};return /CODE เลิกใช้แล้ว: 46001008/.test(m)&&/ใช้ 46001009 แทน/.test(m)&&/ให้ Admin ยืนยัน/.test(m)})());
T('order: ACTIVE 0 without a row and a Stock Alert tick are refused too',(()=>{c.__user={e:'u:s',n:'S',a:0};c.isAdmin_=()=>false;const m=E(()=>c.csOrderCheck_(it(['43400001']),[],null,''));return /43400001/.test(m)})());
c.isAdmin_=t=>t==='ADM';
T('order: ⚠ ห้ามสั่งใหม่ is not refused (warning only)',E(()=>c.csOrderCheck_(it(['42046001']),[],null,''))==='');
T('order: a CODE already on the saved order is not refused again (old orders can still be edited)',E(()=>c.csOrderCheck_(it(['46001008']),it(['46001008']),null,''))==='');
T('order: Admin without a reason is refused; with a reason it goes through and is noted',/Admin ยืนยัน/.test(E(()=>c.csOrderCheck_(it(['46001008']),[],null,'ADM')))&&c.csOrderCheck_(it(['46001008']),[],{'46001008':'ลูกค้ากำหนดรุ่นเดิม'},'ADM')[0]==='46001008: ลูกค้ากำหนดรุ่นเดิม');
T('order: a user cannot use an Admin ack',/46001008/.test(E(()=>c.csOrderCheck_(it(['46001008']),[],{'46001008':'x'},''))));
T('saveOrder runs the check (Order Project is left out)',/csOrderCheck_\(x\.items, cur \? cur\.items : \[\], csAck, csTok\)/.test(fs.readFileSync(path.join(ROOT,'OrderService.gs'),'utf8'))&&/x\.otype !== 'project'/.test(fs.readFileSync(path.join(ROOT,'OrderService.gs'),'utf8')));
g=c.deleteCodeStatus('ADM','42046001');
T('delete: the row goes',!g.map['42046001']);
c.__user={e:'u:s',n:'S.Somchai',a:0};c.isAdmin_=()=>false;
T('user cannot write',/Admin only/.test(E(()=>c.saveCodeStatus('tok',{code:'1',reason:'eol'})))&&/Admin only/.test(E(()=>c.deleteCodeStatus('tok','46001008'))));
T('page may call them (allow-list)',['getCodeStatus','saveCodeStatus','deleteCodeStatus'].every(f=>!!c.API_FNS[f]));
console.log(ok+' passed, '+bad+' failed');process.exit(bad?1:0);
