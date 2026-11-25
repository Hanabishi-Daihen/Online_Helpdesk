// Booking BOM × Shipment: CODEs MF issued (CONTROL WIRING) vs the BOM → alert at 80 %, closed after 14 days; units sharing a CATEGORY
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
const c=mk();c.wsCached_=(k,v,x,f)=>f();c.getSchedules=()=>[];c.pendingBundle_=()=>null;c.__trusted=true;c.__user={e:'admin:a',n:'A.Admin',a:1};
const audits=[];c.audit_=r=>audits.push(...r);
c.stockData_=()=>['42040008\tMCB\tC32\t50\tPC\tS1','42046001\tMCB\tC6\t50\tPC\tS1'];c.stockMeta_=()=>({rows:2});
const _sb=c.bomSnapBuild_;c.bomSnapBuild_=()=>({projects:[],codes:[]});
const it=(code,qty)=>({code,qty});
c.importBom('session',[{project:'PP1048B02',title:'LCC',rev:'A',file:'b.xlsx',items:[it('10000001',4),it('10000002',3),it('10000003',2),it('10000004',1),it('10000005',1)]}]);
c.bomSnapBuild_=_sb;
// Shipment lines: CODE, DESC, yyyymmdd, DOCNUM, CUSTNAME, LOC, QTY, UNIT, UNITCOST, COST, USER (issue = negative cost)
const L=(code,d8,qty,doc,dept,cost)=>[code,'x',d8,doc||'154850/3090',dept||'CONTROL WIRING','WH1',qty,'PC',10,cost==null?-10*qty:cost,'U'];
let LINES={};c.shipMeta_=()=>({ver:'1',idx:{'1048P2':[2],'1050P1':[3]}});c.shipLines_=(m,cat)=>(LINES[cat]||[]).map(x=>x.slice());
const proj=(s,k)=>s.projects.find(p=>p.project===k)||{};
LINES['1048P2']=[L('10000001','20261120',4),L('10000002','20261120',3),L('10000003','20261121',2),L('10000004','20261121',1),
  L('10000005','20261121',1,'','ASSEMBLY'),L('10000005','20261101',1)];   // other dept · before the BOM was imported
let s=c.bomSnapBuild_(),sh=proj(s,'PP1048B02').ship;
T('4 of 5 CODEs issued to CONTROL WIRING after import → 80 %, alert, closes in 14 days',sh&&sh.full===4&&sh.n===5&&sh.pct===0.8&&sh.alert&&sh.closeOn&&sh.last==='21/11/2026',sh);
T('another department and lines before the BOM import are not counted',sh.by['10000005']===undefined,sh.by);
T('summary: one project issued and still open',c.getBomSummary().ship===1);
LINES['1048P2'].push(L('10000001','20261122',1,'','CONTROL WIRING',10));   // return 1 (positive cost)
s=c.bomSnapBuild_();sh=proj(s,'PP1048B02').ship;
T('a return takes the CODE back below the BOM qty → 3/5, no alert',sh.by['10000001']===3&&sh.full===3&&sh.part===1&&!sh.alert,sh);
T('below 80 % the first-alert date is dropped',!JSON.parse(c.schedProp_('BOM_SHIP_SEEN')||'{}').PP1048B02);
LINES['1048P2'].pop();
// two units of one project number share CATEGORY 1048P2
c.bomSnapBuild_=()=>({projects:[],codes:[]});
c.importBom('session',[{project:'PP1048A02',title:'RCC',rev:'A',file:'a.xlsx',items:[it('10000001',2),it('20000001',5)]}]);
c.bomSnapBuild_=_sb;
LINES['1048P2']=[L('10000001','20261120',4),L('10000002','20261120',3),L('10000003','20261121',2),L('10000004','20261121',1),
  L('20000001','20261120',5,'MF241A/PP1048A02-02'),L('10000001','20261120',2,'MF241A/PP1048A02-02'),L('10000005','20261120',1,'MF250/PP1048C02-01')];
s=c.bomSnapBuild_();const a=proj(s,'PP1048A02').ship,b=proj(s,'PP1048B02').ship;
T('DOCNUM with the full number goes to that unit only (A02 gets its 5 + 2)',a.by['20000001']===5&&a.full===2&&a.tagged===2,a);
T('untagged lines are shared by need order: A02 already has its 10000001, so B02 gets all 4',b.by['10000001']===4&&b.full===4&&b.shared,b);
T('a line tagged to a unit with no open BOM (C02) counts for nobody',!a.by['10000005']&&!b.by['10000005']);
// 14 days after the first alert → closed automatically
const seen=JSON.parse(c.schedProp_('BOM_SHIP_SEEN')||'{}');T('first-alert dates kept for both units',seen.PP1048A02&&seen.PP1048B02,seen);
seen.PP1048B02=Date.now()-15*864e5;c.schedProp_('BOM_SHIP_SEEN',JSON.stringify(seen));
s=c.bomSnapBuild_();
T('15 days after the alert, B02 is closed automatically with the reason, A02 still open',proj(s,'PP1048B02').status==='closed'&&/อัตโนมัติ: Shipment เบิกแล้ว 80%/.test(proj(s,'PP1048B02').closed)&&proj(s,'PP1048A02').status==='active',[proj(s,'PP1048B02').status,proj(s,'PP1048B02').closed]);
T('audit names the closed project',audits.some(x=>/ปิดโปรเจกต์ BOM อัตโนมัติ \(Shipment/.test(x[1])&&/PP1048B02/.test(x[3])));
T('a closed project leaves the first-alert list',!JSON.parse(c.schedProp_('BOM_SHIP_SEEN')||'{}').PP1048B02);
T('closed B02 no longer books: A02 now takes the shared 10000001 lines it needs',proj(s,'PP1048A02').ship.by['10000001']>=2);
c.shipMeta_=()=>null;s=c.bomSnapBuild_();
T('no Shipment data → no ship info, nothing closed, the calculation still runs',proj(s,'PP1048A02').ship===null&&proj(s,'PP1048A02').status==='active');
console.log(ok+' passed, '+bad+' failed');process.exit(bad?1:0);
