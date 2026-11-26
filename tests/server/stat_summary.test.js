// สถิติและต้นทุน (our unit only: DS13 · MF → CONTROL WIRING · By Project CODEs …PP____C…): monthly totals kept by the Shipment sync → one summary (usage, BOM vs issued, stock health, purchasing); money only for Admin
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
const c=mk();
const zlib2=require('zlib');
const now=new Date(),ago=(mo,day)=>{const d=new Date(now.getFullYear(),now.getMonth()-mo,day||5);return d};
const dmy=d=>('0'+d.getDate()).slice(-2)+'/'+('0'+(d.getMonth()+1)).slice(-2)+'/'+d.getFullYear(),ymd=d=>d.getFullYear()+('0'+(d.getMonth()+1)).slice(-2)+('0'+d.getDate()).slice(-2);
const H='FMTITEMNO,DESCRIPTION,DateShip,DOCNUM,CUSTNAME,CATEGORY,LOCATION,SHIPQTY,SHIPUNIT,UNITCOST,SHIPCOST,AUDTUSER';
const L=[
 `42046001,MCB C6,${dmy(ago(1))},MF1/PP1048B02-01,CONTROL WIRING,1048P2,WH1,100,PC,100,-10000,U1`,
 `42046001,MCB C6,${dmy(ago(1,9))},MF2,CONTROL WIRING,1048P2,WH1,2,PC,100,200,U1`,             // return 2
 `46700402,TERMINAL UK6N,${dmy(ago(2))},MF3,CONTROL WIRING,1048P2,WH1,300,PC,20,-6000,U1`,
 `47280030,LUG VF2-3M,${dmy(ago(2))},MF4,CONTROL WIRING,1048P2,WH1,500,PC,1,-500,U1`,            // not in the BOM
 `10190005,HEX BOLT,${dmy(ago(3))},T1,TANKING,1048P2,WH2,50,PC,2,-100,U3`,                 // another department: not ours
 `4556PP1048C00201,ANNUNCIATOR,${dmy(ago(2))},A1,ASSEMBLY,1048P2,WH1,1,EA,3000,-3000,U2`,   // By Project CODE: ours wherever issued
 `42046001,MCB C6,${dmy(ago(30))},OLD,CONTROL WIRING,P99901,WH1,99,PC,100,-9900,U1`];   // older than 24 months
FILE=Buffer.from(H+'\r\n'+L.join('\r\n')+'\r\n','utf8');props.SHIP_SRC_URL='https://example.sharepoint.com/:x:/g/personal/s_somchai/ABCDEF';
const imp=dmy(ago(4))+' 09:00:00';
c.bomProjects_=()=>[{project:'PP1048B02',title:'Local Control Cabinet (LCC)',status:'closed',imported:imp},{project:'PP1050B01',title:'RCC',status:'active',imported:imp}];
c.bomItems_=()=>({PP1048B02:[{code:'42046001',qty:4},{code:'46700402',qty:100},{code:'42040008',qty:2}]});
c.schedCached_=()=>[];
// Item stock: CODE DESC COMMENT QTY UNIT LOC PICK BY LASTCOST REQ | LASTSHIPDT ACTIVE (packed)
const srow=(code,desc,qty,cost,lu,act,by)=>[code,desc,'',qty,'PC','STORE1','',by||'DS13',cost,'',lu,act==null?'1':act].join('\t');
props.STOCK_SS_ID='STOCKSS';c.stockData_=()=>[srow('42046001','MCB C6',3,100,ymd(ago(1))),srow('46700402','TERMINAL UK6N',5000,20,ymd(ago(2))),srow('99990001','OLD RELAY',40,500,ymd(ago(40)),'0'),srow('99990002','NEVER USED',10,30,''),srow('88880001','OTHER UNIT PART',999,1000,ymd(ago(50)),'1','STORE/PL'),srow('4556PP1030C00201','LEFTOVER PANEL PART',2,5000,ymd(ago(20))),srow('77770001','NEW SPARE (ordered, not used yet)',5,100,'')];
c.saSettings_=()=>({'42046001':{min:1,max:2,lt:0}});c.rtMap_=()=>({'42046001':45});
const pend=[['42046001','PR1','PO1','','Open','MCB',dmy(ago(3)),'',dmy(ago(3,10)),dmy(ago(1,1)),'','','DS13','5','PC'],['42046001','PR2','PO2','','Received','MCB',dmy(ago(6)),'',dmy(ago(6,12)),dmy(ago(4,20)),'','','DS13','5','PC'],
  ['46700402','PR3','','','Open','TB',dmy(ago(5)),'','','','','','DS13','10','PC'],['88880001','PR9','PO9','','Open','X',dmy(ago(8)),'',dmy(ago(8,9)),dmy(ago(6)),'','','STORE/PL','1','PC'],['42046001','PR7','PO7','','On P/O','MCB OLD',dmy(ago(48)),'',dmy(ago(48,9)),dmy(ago(46)),'','','DS13','2','PC'],['77770001','PR6','PO6','','Received','SPARE',dmy(ago(20)),'',dmy(ago(20,9)),dmy(ago(18)),'','','DS13','5','PC']].map(r=>r.join('\t')).join('\n');
c.pendingBundle_=()=>({hash:'p',b64:zlib2.gzipSync(Buffer.from(pend)).toString('base64')});
let r=c.syncShipNow('');
T('Shipment sync ok and the summary is built with it',r.ok&&/สรุปสถิติใหม่แล้ว/.test(r.msg),r.msg);
const S=c.getStatSummary('');
T('usage: 12-month net use per CODE, the return kept apart (issued 100, returned 2)',S.topQty.find(x=>x.code==='42046001').q===100&&S.topQty.find(x=>x.code==='42046001').r===2&&S.$topCost[0].code==='42046001'&&!S.topQty.some(x=>x.code==='10190005'),S.topQty);
T('usage: cost per CODE is issued − returned (10000 − 200), the share of the top 20',S.$topCost.find(x=>x.code==='42046001').$c===9800&&S.$share20===100,S.$topCost);
T('usage: only CONTROL WIRING + By Project CODEs (TANKING left out), lines older than 24 months not in the chart',S.depts[0].name==='CONTROL WIRING'&&S.depts[0].$tot===16300&&S.depts[1].name==='ASSEMBLY'&&S.depts[1].$tot===3000&&S.depts.length===2,S.depts.map(d=>[d.name,d.$tot]));
T('usage: returns list',S.returns.length===1&&S.returns[0].code==='42046001'&&S.returns[0].rp===2,S.returns);
T('projects: CONTROL WIRING cost and returns per project, type from the BOM title',S.$projects.some(p=>p.cat==='1048P2'&&p.$c===16300&&p.$cr===200&&p.type==='LCC'),S.$projects);
const bp=S.bom.projects.find(p=>p.project==='PP1048B02');
T('BOM vs issued: 2 of 3 BOM CODEs issued, 1 extra CODE (the lug), accuracy 2/(3+1)',bp&&bp.issued===2&&bp.notIssued===1&&bp.extra===1&&bp.extraCodes[0]==='47280030'&&bp.acc===50,bp);
T('BOM vs issued: a project with nothing issued gets no score',!S.bom.projects.find(p=>p.project==='PP1050B01'));
T('stock: only the unit CODEs (ordered by DS13) + By Project leftovers; another unit part left out',S.stock.$value===3*100+5000*20+40*500+10*30+2*5000+5*100&&S.stock.buckets[5].wait&&S.stock.buckets[5].n===1&&S.stock.buckets[0].n===2&&S.stock.buckets[3].n===1&&S.stock.buckets[4].n===1&&S.stock.pcN===1&&S.stock.$pcV===10000&&S.stock.items===6,S.stock);
T('never issued but a PR/PO within 5 years → รอใช้งาน, not in the idle lists',S.waiting.some(x=>x.code==='77770001'&&x.pr)&&!S.$dead.some(x=>x.code==='77770001')&&!S.dead.some(x=>x.code==='77770001')&&S.dead.some(x=>x.code==='99990002'),[S.waiting,S.dead]);
T('stock: idle list by value (Admin) and by idle days (all); the By Project leftover flagged',S.$dead.some(x=>x.code==='4556PP1030C00201'&&x.pc)&&S.$dead.find(x=>!x.pc).code==='99990001'&&S.dead.some(x=>x.code==='99990002')&&!('$v' in S.dead[0]),[S.$dead,S.dead]);
T('stock: days of cover — MCB 3 left at 98/yr is below its 45-day lead time; terminal 5000 at 300/yr is over 2 years',S.low.some(x=>x.code==='42046001'&&x.lt===45)&&S.over.some(x=>x.code==='46700402'),[S.low,S.over]);
T('stock: turnover per CODE group',S.groups.find(g=>g.g==='4670').turn===statTurn(),S.groups);
function statTurn(){return Math.round(6000/(5000*20)*100)/100}
T('purchasing: open lines aged, the line past its due date listed',S.po.open===2&&S.po.ages.reduce((a,b)=>a+b[1],0)===2&&!S.po.late.some(x=>x.pr==='PR9')&&!S.po.late.some(x=>x.pr==='PR7')&&S.po.late.length===1&&S.po.late[0].pr==='PR1',S.po);
T('purchasing: an open line more than 3 years past its due date is listed apart (likely never closed), not open / aged / late',S.po.staleN===1&&S.po.stale[0].pr==='PR7'&&S.po.stale[0].years>=3.7&&S.po.stale[0].days>1095,S.po);
T('purchasing: median PR → PO days',S.po.prpo>=5&&S.po.prpo<=7,S.po.prpo);
const mm=S.minmax.find(x=>x.code==='42046001');
T('usage: the unit\'s share — By Project CODE and DS13-ordered CODEs flagged',S.topQty.find(x=>x.code==='4556PP1048C00201').pc&&S.topQty.find(x=>x.code==='42046001').mine&&S.$mineCost12>0,S.topQty);
T('Min/Max: from use × lead time (StockRT 45 d) and 90 days of use, shown beside the current 1/2',mm&&mm.lt===45&&mm.ltSrc==='rt'&&mm.min===Math.ceil(98/365*45*1.5)&&mm.max===Math.ceil(mm.min+98/365*90)&&mm.curMin===1,mm);
// a user: no money anywhere
const u=mk();Object.assign(u,{stockData_:c.stockData_});u.__user={e:'u@x',n:'S.Somchai',a:0};u.isAdmin_=()=>false;
const U=u.getStatSummary('');
T('user: same summary without any "$" (money) key, quantities kept',!/"\$/.test(JSON.stringify(U))&&U.topQty.length&&!U.admin&&U.stock.items===6,Object.keys(U));
T('user cannot rebuild',/Admin/.test(E(()=>u.statRebuild(''))));
T('page may call them (allow-list)',['getStatSummary','statRebuild'].every(f=>!!c.API_FNS[f]));
console.log(ok+' passed, '+bad+' failed');process.exit(bad?1:0);
