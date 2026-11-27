// Booking BOM: a failed Item stock read never turns every CODE into "ไม่พบใน Stock"
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
const c=mk();c.getSchedules=()=>[];c.pendingBundle_=()=>null;
const STOCK=['42040008\tMiniature circuit breaker\tPLS6-C32/3\t5\tPC\tSTORE1','42046001\tMiniature circuit breaker\tPLS6-C6/2-DC\t3\tPCS\tSTORE1'];
let META={rows:2};
let mode='ok',reads=0;c.stockData_=()=>{reads++;if(mode==='throw')throw new Error('Service Spreadsheets timed out');return mode==='empty'?[]:STOCK};
c.stockMeta_=()=>META;
const _sb=c.bomSnapBuild_;c.bomSnapBuild_=()=>({projects:[],codes:[]});   // import only stores the project
c.importBom('session',[{project:'PP1039D01',title:'RCC',rev:'A',file:'bom.xlsx',items:[{code:'42040008',qty:1,material:'MCB 32A'},{code:'42046001',qty:10,material:'MCB 6A'}]}]);
c.bomSnapBuild_=_sb;
const code=(s,k)=>s.codes.find(x=>x.code===k)||{};
let s=c.bomSnapBuild_();
T('normal: CODEs found in Item stock with Onhand',code(s,'42040008').known&&code(s,'42040008').onhand===5&&code(s,'42046001').short&&s.stockRows===2,s.codes);
T('normal: no error on the snapshot',!s.err);
mode='throw';reads=0;s=c.bomSnapBuild_();
T('read fails → retried once',reads===2,reads);
T('read fails → the last good snapshot stays (still found, Onhand 5)',code(s,'42040008').known&&code(s,'42040008').onhand===5,s.codes);
T('read fails → the reason is on the snapshot for the page',/อ่าน Item stock ไม่สำเร็จ/.test(s.err)&&/timed out/.test(s.err),s.err);
T('…and stored, so the page shows it',/อ่าน Item stock/.test((c.bomSnapGet_()||{}).err||''));
mode='empty';s=c.bomSnapBuild_();
T('0 rows while Item stock has data → not booked against an empty stock',code(s,'42040008').known&&/0 แถว/.test(s.err),[s.err,s.codes]);
mode='ok';let g=c.getBom('session');
T('getBom: error under 5 min old → no rebuild on every open',!!g.err);
const sv=c.bomSnapGet_();sv.errT=Date.now()-6e5;c.bomSnapPut_(sv);g=c.getBom('session');
T('getBom: older error → rebuilt and cleared once stock reads again',!g.err&&code(g,'42040008').known,g.err);
mode='throw';META={rows:0};c.bomSnapPut_(Object.assign(c.bomSnapGet_(),{v:3}));
T('Item stock really empty (never synced) → compute still runs (all not found)',(()=>{mode='empty';const x=c.bomSnapBuild_();return !x.err&&code(x,'42040008').known===false})());
// a By Project CODE booked into a project → marked pj (the page shows it under Order Project, not the stock list)
mode='ok';META={rows:3};STOCK.push('4556PP1034C00701\tTerminal block\tUK5N\t4\tPCS\tSTORE2');
c.importBom('session',[{project:'PP1045A01',title:'MCC',rev:'A',file:'bom.xlsx',items:[{code:'42040008',qty:1},{code:'4556PP1034C00701',qty:2}]}]);
{const x=c.bomSnapBuild_(),pp=x.projects.find(p=>p.project==='PP1045A01')||{items:[]},a=pp.items.find(i=>i.code==='4556PP1034C00701')||{},b=pp.items.find(i=>i.code==='42040008')||{};
T('By Project CODE: item and CODE row marked pj, stock CODE not · letter sent for the cart',a.pj===1&&b.pj===0&&code(x,'4556PP1034C00701').pj===1&&x.pjL==='C'&&a.fromStock===2,[a,b,x.pjL])}
console.log(ok+' passed, '+bad+' failed');process.exit(bad?1:0);
