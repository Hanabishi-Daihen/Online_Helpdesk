// หางปลา × อุปกรณ์ (LugGuide.gs): catalog data until the first Admin edit, then the sheet · checks on save · delete guard
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
let g=c.getLugGuide('tok');
T('no sheet rows → the checked catalog data (91 lugs, groups, 231 devices)',!g.custom&&g.T.length===91&&g.D.length===231&&g.F.PYF14A&&g.F.PYF14A_N&&g.admin,{custom:g.custom,T:g.T.length});
T('catalog data: rechecked values (JST R5.5-5 L 19.8, BW50 M5 from Fuji, TP48X M3.5 = A)',g.T.find(t=>t.id==='47180005').L===19.8&&g.F.MCCB_BW50.cf==='B'&&g.F.TP4.cf==='A');
T('catalog data: the MY4N relay lists its three sockets',g.D.find(x=>x.it==='43401016').f==='PYF14A,PYF14A_N,PYFZ');
T('catalog data: every device points at existing groups',g.D.every(x=>x.f.split(',').every(k=>g.F[k])),g.D.filter(x=>!x.f.split(',').every(k=>g.F[k])).slice(0,3));
g=c.saveLug('tok','F',{key:'PYF14A',n:'Omron Socket PYF14A (MY4N)',cn:'captive',sc:'M3',a:'0.5',b:'1.5',w:'6',cf:'A',ov:{'47280021':['0','แคบ'],'x':[7,'']}},'PYF14A');
const rows=book.sheets.LugGuide.rows;
T('first edit copies the catalog data into the sheet, then saves the change',g.custom&&rows.length===1+91+Object.keys(g.F).length+231&&g.F.PYF14A.ov['47280021'][0]===0&&!g.F.PYF14A.ov.X&&g.F.PYF14A.a===0.5,rows.length);
T('save: one row replaced, not added',rows.filter(r=>r[0]==='F'&&r[1]==='PYF14A').length===1);
T('save: audit names the group',audits.some(a=>/บันทึกหางปลา/.test(a[1])&&/PYF14A/.test(a[2])));
g=c.saveLug('tok','T',{id:'47999999',m:'RF2-4',k:'RING_I',wire:'1.5, 2.5',sc:'M4',W:'8.5',cf:'B'},'');
T('new lug saved with numbers as numbers',g.T.find(t=>t.id==='47999999').W===8.5);
T('lug: needs a type and wire size',/ชนิด/.test(E(()=>c.saveLug('tok','T',{id:'1',k:'X',wire:'1'},'')))&&/ขนาดสาย/.test(E(()=>c.saveLug('tok','T',{id:'1',k:'PIN'},''))));
T('lug: renaming onto a CODE that exists is refused',/มีอยู่แล้ว/.test(E(()=>c.saveLug('tok','T',{id:'47180060',m:'x',k:'RING_N',wire:'2'},'47999999'))));
T('group: connection type, wire range checked',/ชนิดขั้วต่อ/.test(E(()=>c.saveLug('tok','F',{key:'X1',n:'x',cn:'zz'},'')))&&/ต่ำสุดมากกว่า/.test(E(()=>c.saveLug('tok','F',{key:'X1',n:'x',cn:'box',a:5,b:1},''))));
T('device: groups must exist',/ไม่พบกลุ่มขั้วต่อ NOPE/.test(E(()=>c.saveLug('tok','D',{it:'43401016',f:'PYF14A,NOPE'},'43401016'))));
g=c.saveLug('tok','D',{it:'43401016',f:'PYF14A,PYFZ',pl:true,g:'AUXILIARY RELAY'},'43401016');
T('device: socket list edited',g.D.find(x=>x.it==='43401016').f==='PYF14A,PYFZ'&&g.D.filter(x=>x.it==='43401016').length===1);
T('delete group still used by a device → refused',/ยังมีอุปกรณ์/.test(E(()=>c.deleteLug('tok','F','PYFZ'))));
g=c.deleteLug('tok','T','47999999');
T('delete lug',!g.T.find(t=>t.id==='47999999'));
c.__user={e:'u:s',n:'S.Somchai',a:0};c.isAdmin_=()=>false;
T('user cannot save or delete',/Admin only/.test(E(()=>c.saveLug('tok','T',{id:'1',k:'PIN',wire:'1'},'')))&&/Admin only/.test(E(()=>c.deleteLug('tok','T','47180060'))));
console.log(ok+' passed, '+bad+' failed');process.exit(bad?1:0);
