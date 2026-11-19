// Set Library: Admin saves / imports / deletes sets of parts used together; everyone signed in reads; catalog seeds
const vm=require('vm'),fs=require('fs'),path=require('path'),crypto=require('crypto');
const ROOT=path.resolve(__dirname,'../..');
const src=fs.readdirSync(ROOT).filter(f=>f.endsWith('.gs')).map(f=>fs.readFileSync(path.join(ROOT,f),'utf8')).join('\n');
const mkSheet=()=>{let rows=[];const sh={rows,getLastRow:()=>rows.length,appendRow:r=>{rows.push(r.map(String))},deleteRow:r=>{rows.splice(r-1,1)},setFrozenRows(){},
 getRange(r,c,nr,nc){nr=nr||1;nc=nc||1;const R={setValues(v){v.forEach((row,i)=>{while(rows.length<r+i)rows.push([]);row.forEach((x,j)=>rows[r+i-1][c+j-1]=String(x))});return R},setNumberFormat:()=>R,
  getValues:()=>Array.from({length:nr},(_,i)=>Array.from({length:nc},(_,j)=>{const v=(rows[r+i-1]||[])[c+j-1];return v==null?'':v}))};return R}};return sh};
const props={ADMIN_PASSCODE:'control'},cache={},sheets={};
const SS={getSheetByName:n=>sheets[n]||null,insertSheet:n=>sheets[n]=mkSheet()};
const mk=()=>{const ctx={console:{log(){},warn(){}},JSON,Math,Date,String,Number,Array,Object,Error,RegExp,
 PropertiesService:{getScriptProperties:()=>({getProperty:k=>k in props?props[k]:null,setProperty:(k,v)=>{props[k]=String(v)},deleteProperty:k=>{delete props[k]},getProperties:()=>({...props})})},
 Utilities:{formatDate:()=>'19/11/2026 09:00:00',getUuid:()=>crypto.randomUUID(),sleep(){}},SpreadsheetApp:{openById:()=>SS,flush(){}},LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){},tryLock:()=>true})},
 CacheService:{getScriptCache:()=>({get:k=>k in cache?cache[k]:null,put:(k,v)=>{cache[k]=String(v)},remove:k=>{delete cache[k]}})}};
 ctx.globalThis=ctx;vm.createContext(ctx);vm.runInContext(src,ctx);ctx.audit_=()=>{};ctx.__trusted=true;return ctx};
let ok=0,bad=0;const T=(n,c,x)=>{c?ok++:bad++;console.log((c?'PASS ':'FAIL ')+n+(c?'':' → '+JSON.stringify(x)))};
const E=f=>{try{f();return''}catch(e){return e.message}};
const c=mk();c.__user={e:'admin:a',n:'A.Admin',a:1};
let r=c.saveSetLib('',{name:' Relay MY4N DC110V + Socket ',cat:'Relay + Socket',note:'x',source:'https://example.com/a',parts:[
  {role:'หลัก',per:1,unit:'pcs',optional:true,opts:[{code:' 43401016 '}]},{role:'Socket',per:'1',unit:'PCS',opts:[{code:'46420024'},{kw:'PYF14A-E'},{}]},{role:'Clip',per:0,optional:1,opts:[{kw:'PYC-A1'}]},{role:'ว่าง',opts:[]}]});
const s0=r.sets[0];
T('save: cleaned (trim, upper-case CODE, unit; empty options and empty parts dropped; per ≥ 0 → 1)',r.sets.length===1&&s0.name==='Relay MY4N DC110V + Socket'&&s0.parts.length===3&&s0.parts[0].opts[0].code==='43401016'&&s0.parts[0].unit==='PCS'&&s0.parts[1].opts.length===2&&s0.parts[2].per===1,s0);
T('save: the main item is never optional, the others keep their flag',s0.parts[0].optional===false&&s0.parts[2].optional===true);
T('save: an id is given, admin flag returned',/^S/.test(s0.id)&&r.admin===true);
r=c.saveSetLib('',{...s0,name:'Relay MY4N DC110V + Socket (แก้)'});
T('save with the same id replaces the set',r.sets.length===1&&r.sets[0].name==='Relay MY4N DC110V + Socket (แก้)');
r=c.saveSetLib('',[{name:'EATON PKZM0 + Aux contact',cat:'Motor CB',parts:[{role:'หลัก',opts:[{code:'42050002'}]},{role:'Aux',opts:[{kw:'NHI11-PKZ0'}]}]},{name:'MCB PLS6',cat:'MCB',parts:[{role:'หลัก',opts:[{code:'42040008'}]}]}]);
T('import a list: added, sorted by category then name',r.sets.length===3&&r.sets.map(s=>s.cat).join()==='MCB,Motor CB,Relay + Socket',r.sets.map(s=>s.cat));
T('a set without any part is refused',/อย่างน้อย 1/.test(E(()=>c.saveSetLib('',{name:'x',parts:[{role:'a',opts:[]}]}))));
T('a set without a name is refused',/ชื่อชุด/.test(E(()=>c.saveSetLib('',{name:' ',parts:[{opts:[{code:'1'}]}]}))));
const u=mk();u.__user={e:'u@x',n:'U',a:0};
T('a signed-in user reads the sets (not admin)',(()=>{const x=u.getSetLib('');return x.sets.length===3&&x.admin===false})());
T('a signed-in user cannot save, import or delete',['saveSetLib','deleteSetLib'].every(f=>/Admin/.test(E(()=>u[f]('',{name:'x',parts:[{opts:[{code:'1'}]}]})))));
r=c.deleteSetLib('',r.sets[0].id);
T('delete removes the set',r.sets.length===2&&!r.sets.some(s=>s.cat==='MCB'));
T('delete of a missing set says so',/ไม่พบ/.test(E(()=>c.deleteSetLib('','NOPE'))));
const seeds=c.getSetSeeds();
T('catalog seeds: each has a name, category, source, and parts with model keywords (main part first)',seeds.length>=15&&seeds.every(s=>s.name&&s.cat&&s.src&&s.parts.length&&s.parts[0].role==='หลัก'&&s.parts.every(p=>p.kw&&p.kw.length)),seeds.filter(s=>!(s.name&&s.src)).map(s=>s.name));
T('catalog seeds: MY4N pairs with PYFZ-14 / PYF14A sockets and the PYC-A1 clip',(()=>{const m=seeds.find(s=>/MY4N/.test(s.name));return m&&m.parts[1].kw.join()==='PYFZ-14,PYF14A'&&m.parts[2].kw[0]==='PYC-A1'&&m.parts[2].optional})());
T('catalog seeds: magnetic contactors, PT30 / PT40 / UK 5 N / UK 6 N / USST terminals and test terminals are there',['FUJI SC-03','FUJI SC-0 ','Mitsubishi S-T','Schneider LC1D','TOGI PT30','TOGI PT40','Phoenix UK 5 N','Phoenix UK 6 N','USST','Phoenix URTK/S','Test terminal Type B','GE PK-2'].every(n=>seeds.some(x=>x.name.indexOf(n)===0)),seeds.map(x=>x.name));
T('catalog seeds: SC-0 asks for that model exactly (not SC-03 / SC-05)',seeds.find(x=>/SC-0 /.test(x.name)).parts[0].kw[0]==='=SC-0');
T('page may call them (allow-list)',['getSetLib','getSetSeeds','saveSetLib','deleteSetLib'].every(f=>!!c.API_FNS[f]));
console.log(ok+' passed, '+bad+' failed');process.exit(bad?1:0);
