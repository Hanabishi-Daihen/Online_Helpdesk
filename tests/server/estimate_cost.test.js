// Estimate Cost: Shipment CSV read in byte ranges (> 50 MB in real life) · projects packed per CATEGORY · the page's calls
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
let ok=0,bad=0;const T=(n,c,x)=>{c?ok++:bad++;console.log((c?'PASS ':'FAIL ')+n+(c?'':' → '+JSON.stringify(x)))};
const E=f=>{try{f();return''}catch(e){return e.message}};
const c=mk();c.SHIP_CHUNK=256;   // tiny ranges: records, quotes and Thai characters fall across range edges
// project number ⇄ CATEGORY
T('PP103001 / PP1030D01 / PP1030-01 / 1030P01 → 1030P1',['PP103001','PP1030D01','pp1030-01','1030P01','1030p1'].every(s=>c.shipCat_(s)==='1030P1'),['PP103001','PP1030D01'].map(c.shipCat_));
T('old style PP099901 → P99901 (as the Access form)',c.shipCat_('PP099901')==='P99901'&&c.shipCat_('P99901')==='P99901');
T('not a project → empty',['','PH0024','1043T1','PP10','ABC'].every(s=>c.shipCat_(s)===''));
T('CATEGORY → project number on the report',c.shipProj_('1030P1')==='PP103001'&&c.shipProj_('1024P12')==='PP102412'&&c.shipProj_('P99901')==='PP099901');
// the file (newest first, as the query)
const H='FMTITEMNO,DESCRIPTION,DateShip,DOCNUM,CUSTNAME,CATEGORY,LOCATION,SHIPQTY,SHIPUNIT,UNITCOST,SHIPCOST,AUDTUSER';
const L=[
 '37000017,Winding Thermometer,05/11/2026,MF241A/PP1030D01-02,CONTROL WIRING,1030P1,WH1,2,EA,33949.07,-67898.15,USR1',
 '4556PP1030C00201,"Annunciator Type: ESPAN02E-24 DC.110V., I/P: No Contacts,",04/11/2026,154862/3098,CONTROL WIRING,1030P1,WH1,1,EA,37800,37800,USR1',
 '4556PP1030C00201,"Annunciator Type: ESPAN02E-24 DC.110V., I/P: No Contacts,",03/11/2026,154860/3098,CONTROL WIRING,1030P1,WH1,1,EA,37800,-37800,USR1',
 '46420024,"รีเลย์ซ็อกเก็ต ""PYFZ-14"" (FOR MY4N)",02/11/2026,154850/3090,CONTROL WIRING,1030P1,WH1,67,PC,190.575,-12768.5,USR2',
 '46420024,"SOCKET RELAY PYFZ-14\nline two",01/11/2026,154849/3090,CONTROL WIRING,1030P1,WH1,2,PC,190.575,-381.15,USR2',
 '10190005,HEX BOLT M6x20,01/11/2026,154000/1,TANKING,1030P1,WH2,50,PC,1.5,-75,USR3',
 '10190005,HEX BOLT M6x20,30/10/2026,153000/1,TANKING,1043T1,WH2,50,PC,1.5,-75,USR3',
 '10190006,BOLT,29/10/2026,153001/1,CORE SHEARING,PH0024,WH2,5,PC,1,-5,USR3',
 '20000001,ไม้ อัด,28/10/2026,152000/1,CONSTRUCTION UNIT,1024P2,WH3,3,SH,500,-1500,USR4',
 '20000002,OLD JOB PART,01/03/2024,100000/1,CONSTRUCTION UNIT,P99901,WH3,1,EA,10,-10,USR4'];
FILE=Buffer.from('\uFEFF'+H+'\r\n'+L.join('\r\n')+'\r\n','utf8');
props.SHIP_SRC_URL='https://example.sharepoint.com/:x:/g/personal/s_somchai/ABCDEF';
let r=c.syncShipNow('');
T('sync: ok, only project CATEGORY rows kept (1043T1 / PH0024 dropped)',r.ok&&r.rows===8&&r.skipped===2&&r.projects===3&&r.depts===3,r);
T('sync: the file was read in several ranges, each ≤ the chunk size',fetches.length>3&&fetches.every(x=>x&&x[1]-x[0]===255),fetches.slice(0,4));
T('sync: own spreadsheet with Ship + ShipMeta, no temp sheet / Sheet1 left',!!book.sheets.Ship&&!!book.sheets.ShipMeta&&!book.sheets.Ship_new&&!book.sheets.Sheet1,Object.keys(book.sheets));
let m=c.getEstMeta('');
const s1=m.sum.find(s=>s[0]==='CONTROL WIRING'&&s[1]==='1030P1');
T('meta: department × project with lines, total, last ship date',s1&&s1[2]==='PP103001'&&s1[3]===5&&Math.abs(s1[4]-(-67898.15+37800-37800-12768.5-381.15))<.01&&s1[5]==='05/11/2026',s1);
T('meta: date range of the data',m.range==='01/03/2024 – 05/11/2026',m.range);
let p=c.getEstCost('','PP1030D01');
T('project: every line of 1030P1 (all departments), newest first',p.cat==='1030P1'&&p.project==='PP103001'&&p.lines.length===6&&p.lines[0][2]==='20261105'&&p.lines.every((x,i)=>!i||x[2]<=p.lines[i-1][2]),p.lines.map(x=>x[2]));
T('project: quotes, commas, Thai text and a line break inside a field survive range edges',p.lines.some(x=>x[1]==='รีเลย์ซ็อกเก็ต "PYFZ-14" (FOR MY4N)')&&p.lines.some(x=>x[1]==='Annunciator Type: ESPAN02E-24 DC.110V., I/P: No Contacts,')&&p.lines.some(x=>x[1]==='SOCKET RELAY PYFZ-14 line two'),p.lines.map(x=>x[1]));
T('project: numbers (qty, unit cost, cost) and the reversal (+)',(()=>{const x=p.lines.find(l=>l[9]===37800);return x&&x[6]===1&&x[8]===37800&&p.lines.find(l=>l[0]==='46420024'&&l[6]===67)[9]===-12768.5})());
T('project: old style P99901',c.getEstCost('','PP099901').lines.length===1);
T('project: bad number refused',/เลขโปรเจกต์ไม่ถูกต้อง/.test(E(()=>c.getEstCost('','PP10'))));
// unchanged file → skipped; changed → rebuilt
fetches=[];r=c.syncShipRun_('อัตโนมัติ',false);
T('auto sync: the same file → skipped after the first range',r.ok&&r.unchanged&&fetches.length===1,[r,fetches.length]);
FILE=Buffer.from(H+'\n'+L.slice(0,6).join('\n')+'\n','utf8');ETAG='"2"';c.SHIP_CELL=120;   // tiny cells: a project is split over many rows
r=c.syncShipRun_('อัตโนมัติ',false);p=c.getEstCost('','1030P1');
T('changed file → rebuilt; a project split over several cells reads back whole',r.ok&&!r.unchanged&&r.rows===6&&p.lines.length===6&&book.sheets.Ship.rows.length>3,[r,p.lines.length,book.sheets.Ship.rows.length]);
T('rebuilt: the new data replaces the old (no P99901 any more)',c.getEstCost('','P99901').lines.length===0&&!c.getEstMeta('').sum.some(s=>s[1]==='P99901'));
// Thai Windows-874 CSV (Access "ANSI" export) and a server that ignores Range
c.SHIP_CELL=45000;FILE=Buffer.from(H+'\r\n20000001,'+'\xe4\xc1\xe9 \xcd\xd1\xb4'+',28/10/2026,152000/1,CONSTRUCTION UNIT,1024P2,WH3,3,SH,500,-1500,USR4\r\n','latin1');ETAG='"3"';ranged=false;
r=c.syncShipNow('');p=c.getEstCost('','PP102402');
T('windows-874 file, sent whole (no Range) → Thai text right',r.ok&&p.lines.length===1&&p.lines[0][1]==='ไม้ อัด',[r,p.lines]);
ranged=true;FILE=Buffer.concat([Buffer.from('PK\x03\x04','latin1'),Buffer.alloc(300)]);ETAG='"4"';
r=c.syncShipNow('');
T('an .xlsx link → a clear message, old data kept',!r.ok&&/\.csv/.test(r.msg)&&c.getEstCost('','PP102402').lines.length===1,r);
FILE=Buffer.from('A,B,C\n1,2,3\n');ETAG='"5"';r=c.syncShipNow('');
T('wrong columns → a clear message',!r.ok&&/FMTITEMNO/.test(r.msg),r);
// settings + report settings
T('source: save link + daily trigger, read back',(()=>{const s=c.setShipSource('','https://example.sharepoint.com/:x:/g/personal/x/Y',105);return s.hours===105&&trig.filter(t=>t.getHandlerFunction()==='shipAutoSync').length===1})());
T('source: a folder link is refused',/โฟลเดอร์/.test(E(()=>c.setShipSource('','https://example.sharepoint.com/personal/x/_layouts/15/onedrive.aspx?id=1',24))));
let cfg=c.saveEstCfg('','control wiring',{title:'ESTIMATE COST OF CONTROL CABINET',sign:['P.Approver','C.Checker','Y.Preparer']});
T('report settings per department (title + 3 names)',cfg['CONTROL WIRING'].title==='ESTIMATE COST OF CONTROL CABINET'&&cfg['CONTROL WIRING'].sign[2]==='Y.Preparer'&&c.getEstMeta('').cfg['CONTROL WIRING'].sign[0]==='P.Approver');
T('report settings: empty → back to the default',!('CONTROL WIRING' in c.saveEstCfg('','CONTROL WIRING',{title:'',sign:['','','']})));
T('Admin only: a signed-in user is refused',(()=>{const u=mk();u.__user={e:'u@x',n:'U',a:0};return ['getEstMeta','getEstCost','saveEstCfg','getShipSource','setShipSource','syncShipNow'].every(f=>/Admin/.test(E(()=>u[f]('','PP103001',{}))))})());
T('page may call them (allow-list)',['getEstMeta','getEstCost','saveEstCfg','getShipSource','setShipSource','syncShipNow'].every(f=>!!c.API_FNS[f]));
console.log(ok+' passed, '+bad+' failed');process.exit(bad?1:0);
