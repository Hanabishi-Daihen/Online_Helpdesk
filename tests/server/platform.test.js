// CODE index (Item stock in 256 buckets by CODE) and the background error log
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
const rows=[];for(let i=0;i<3000;i++)rows.push([String(40000000+i),'ITEM '+i,'',i%7,'PC','S'+(i%3),'','DS13','10','','','1'].join('\t'));
rows.push(['42046001','MCB','',5,'PC','STORE2','','DS13','100','','','0'].join('\t'),['42046001','MCB','',3,'PC','STORE1','','DS13','100','','','0'].join('\t'),['4556PP1030C00201','ANNUNCIATOR','',1,'EA','S1','','DS13','5','','','1'].join('\t'));
let calls=0;const c=mk();c.stockData_=()=>{calls++;return rows};props.STOCK_HASH='abcdef1234567890';
c.stockIdxWarm_(rows,props.STOCK_HASH);
const ks=Object.keys(cache).filter(k=>/^sidx_abcdef1234_/.test(k));
T('index: 256 buckets in the cache, none over 100 KB',ks.length===256&&ks.every(k=>cache[k].length<100000),[ks.length,Math.max(...ks.map(k=>cache[k].length))]);
const d=mk();let dc=0;d.stockData_=()=>{dc++;return rows};
const r=d.stockRowsFor_({'42046001':1,' 4556pp1030c00201 ':1});
T('lookup: both locations of a CODE and a 16-digit CODE (any spelling), without unpacking all of Item stock',r.length===3&&r.filter(x=>x.startsWith('42046001')).length===2&&dc===0,[r.length,dc]);
T('lookup: unknown CODE → nothing',d.stockRowsFor_({'99999999':1}).length===0&&dc===0);
const b=d.sidxOf_('40000001');delete cache['sidx_abcdef1234_'+b];
const e2=mk();let ec=0;e2.stockData_=()=>{ec++;return rows};
T('a bucket evicted from the cache → rebuilt from Item stock once, then served',e2.stockRowsFor_({'40000001':1}).length===1&&ec===1&&cache['sidx_abcdef1234_'+b]!=null&&e2.stockRowsFor_({'40000008':1}).length===1&&ec===1,ec);
delete props.STOCK_HASH;const f=mk();f.stockData_=()=>rows;
T('no sync yet (no hash) → plain scan still answers',f.stockRowsFor_({'42046001':1}).length===2);
// error log
const g=mk();for(let i=0;i<20;i++)g.errLog_(new Error('boom '+i),'Test:'+i);
const L=JSON.parse(props.ERR_LOG);
T('error log: newest first, at most 15 per run, count per day',L.l.length===15&&L.l[0][2]==='boom 14'&&L.l[0][1]==='Test:14'&&Object.values(L.d)[0]===15,L.l.slice(0,2));
const h=mk();for(let i=0;i<60;i++){h.__errN=0;h.errLog_(new Error('x'.repeat(300)),'Long')}
T('error log: stays under the property size limit',props.ERR_LOG.length<=9000&&JSON.parse(props.ERR_LOG).l.length<=40,props.ERR_LOG.length);
T('errors that used to vanish are logged now (the calculation and sync code)',(()=>{const src=fs.readdirSync(ROOT).filter(f=>f.endsWith('.gs')).map(f=>fs.readFileSync(path.join(ROOT,f),'utf8')).join('\n');return (src.match(/errLog_\(e, '/g)||[]).length>=60})());
// หน่วยงาน (UnitCfg.gs)
const u=mk();u.__user={e:'admin:a',n:'A.Admin',a:1};
T('unit: defaults DS13 · CONTROL WIRING · SV · C · 8',JSON.stringify(u.unitCfg_())===JSON.stringify({by:'DS13',mfDept:'CONTROL WIRING',svPrefix:'SV',projLetter:'C',codeLen:8}),u.unitCfg_());
T('unit: rules follow the defaults',u.unitDeptOk_('Control  Wiring')&&!u.unitDeptOk_('TANKING')&&u.unitSvSkip_('sv024269')&&u.unitProjC_('4556PP1030C00201')&&!u.unitProjC_('4556PP1030F00201')&&u.saCodeOk_('42046001')&&!u.saCodeOk_('SV024269'));
u.saveUnitCfg('',{by:'ds11',mfDept:'assembly',svPrefix:'SV, SR',projLetter:'F',codeLen:'10'});
const u2=mk();
T('unit: saved and used by every rule',u2.unitCfg_().by==='DS11'&&u2.unitDeptOk_('ASSEMBLY')&&!u2.unitDeptOk_('CONTROL WIRING')&&u2.unitSvSkip_('SR0001')&&u2.unitProjC_('4556PP1030F00201')&&!u2.unitProjC_('4556PP1030C00201')&&u2.saCodeOk_('4204600112')&&!u2.saCodeOk_('42046001'),u2.unitCfg_());
T('unit: the orderer is copied to Booking BOM and Stock Alert settings',u2.bomCfg_().orderBy==='DS11'&&u2.saCfg_().by==='DS11');
T('unit: checks',/รหัสผู้สั่ง/.test(E(()=>u.saveUnitCfg('',{mfDept:'X',projLetter:'C',codeLen:8})))&&/4–20/.test(E(()=>u.saveUnitCfg('',{by:'A',mfDept:'X',projLetter:'C',codeLen:2}))));
const u3=mk();u3.__user={e:'u:s',n:'S',a:0};u3.isAdmin_=()=>false;
T('unit: users read, only Admin saves',!!u3.getUnitCfg('').cfg&&/Admin only/.test(E(()=>u3.saveUnitCfg('',{by:'X'}))));
// ติดตามใบสั่ง showed 7 while the badge said 8: the version moved before the sheet write was committed
{const seq=[],x=mk();x.SpreadsheetApp={flush:()=>seq.push('flush')};const sp=x.PropertiesService.getScriptProperties;
 x.PropertiesService={getScriptProperties:()=>{const P=sp();return{...P,setProperty:(k,v)=>{seq.push('set '+k);P.setProperty(k,v)}}}};
 x.dataBump_('ORD_VER');T('dataBump_: sheet writes committed (flush) before the version moves',seq.join(',')==='flush,set ORD_VER',seq);
 seq.length=0;x.withLock_(()=>seq.push('fn'));T('withLock_: flush before the lock opens',seq.join(',')==='fn,flush',seq)}
{const x=mk();x.Utilities.base64EncodeWebSafe=a=>Buffer.from(a.map(b=>b&255)).toString('base64');let n=0;props.ORD_VER='1';
 const v1=x.wsCached_('t_race',['ord'],'',()=>{n++;props.ORD_VER='2';return{rows:7}});
 const v2=x.wsCached_('t_race',['ord'],'',()=>{n++;return{rows:8}});
 T('wsCached_: data changed while computing → answered but not kept (next call computes again)',v1.rows===7&&v2.rows===8&&n===2,[v1,v2,n]);
 const v3=x.wsCached_('t_race',['ord'],'',()=>{n++;return{rows:9}});T('wsCached_: a stable answer is kept',v3.rows===8&&v3.cached&&n===2,[v3,n])}
console.log(ok+' passed, '+bad+' failed');process.exit(bad?1:0);
