// Settings store: movable keys live in the sheet "Config"; secrets and version stamps stay in Script Properties
const vm=require('vm'),fs=require('fs'),path=require('path');
const ROOT=path.resolve(__dirname,'../..');
const src=fs.readdirSync(ROOT).filter(f=>f.endsWith('.gs')).map(f=>fs.readFileSync(path.join(ROOT,f),'utf8')).join('\n');
const mkSheet=name=>{const cells={};const sh={name,cells,getLastRow:()=>Object.keys(cells).reduce((m,k)=>Math.max(m,+k.split(',')[0]),0),getMaxColumns:()=>26,
 getRange(r,c,nr,nc){nr=nr||1;nc=nc||1;const R={setValues(v){v.forEach((row,i)=>row.forEach((x,j)=>cells[(r+i)+','+(c+j)]=String(x)));return R},setValue(v){cells[r+','+c]=String(v);return R},getValues:()=>Array.from({length:nr},(_,i)=>Array.from({length:nc},(_,j)=>cells[(r+i)+','+(c+j)]||'')),getDisplayValues(){return R.getValues()},setFontWeight:()=>R,setNumberFormat:()=>R};return R},setFrozenRows(){}};return sh};
const world=()=>{const sheets={},props={ADMIN_PASSCODE:'secret',ORD_VER:'5',STOCK_HASH:'abc',BOM_CFG:'{"a":1}',PO_SRC_URL:'https://x/y',PO_SYNC_HOURS:'6',BOM_TPL_SHORT:'tpl1',BOM_TPL_SHORT_VER:'v3',SA_CFG:'{"s":1}'},cache={};let reads=0,breakSheet=false;
 const SS={getSheetByName:n=>{if(breakSheet)throw new Error('sheet down');if(n==='Config')reads++;return sheets[n]||null},insertSheet:n=>sheets[n]=mkSheet(n)};
 const run=()=>{const ctx={console:{log(){},warn(){}},JSON,Math,Date,String,Number,Array,Object,Error,RegExp,
  PropertiesService:{getScriptProperties:()=>({getProperty:k=>k in props?props[k]:null,setProperty:(k,v)=>{props[k]=String(v)},deleteProperty:k=>{delete props[k]},getProperties:()=>({...props}),setProperties:o=>Object.assign(props,o)})},
  Utilities:{formatDate:()=>'13/11/2026 10:00:00'},SpreadsheetApp:{openById:()=>SS,flush(){}},
  CacheService:{getScriptCache:()=>({get:k=>k in cache?cache[k]:null,put:(k,v)=>{cache[k]=v},putAll:o=>Object.assign(cache,o),getAll:ks=>{const o={};ks.forEach(k=>{if(k in cache)o[k]=cache[k]});return o},remove:k=>{delete cache[k]}})}};
  ctx.globalThis=ctx;vm.createContext(ctx);vm.runInContext(src,ctx);ctx.__trusted=true;ctx.audit_=()=>{};return ctx};
 return{run,props,sheets,cache,reads:()=>reads,down:v=>breakSheet=v}};
let ok=0,bad=0;const T=(n,c,x)=>{c?ok++:bad++;console.log((c?'PASS ':'FAIL ')+n+(c?'':' → '+JSON.stringify(x)))};
const W=world();let c=W.run();                                   // each run() = one new execution (fresh memo)
T('first read moves the settings',c.schedProp_('BOM_CFG')==='{"a":1}'&&c.schedProp_('PO_SRC_URL')==='https://x/y');
T('moved keys left Script Properties',!('BOM_CFG' in W.props)&&!('PO_SRC_URL' in W.props)&&!('PO_SYNC_HOURS' in W.props)&&!('BOM_TPL_SHORT' in W.props)&&!('BOM_TPL_SHORT_VER' in W.props),Object.keys(W.props));
T('secrets, versions, hashes and SA_CFG stay',W.props.ADMIN_PASSCODE==='secret'&&W.props.ORD_VER==='5'&&W.props.STOCK_HASH==='abc'&&W.props.SA_CFG==='{"s":1}'&&c.schedProp_('ORD_VER')==='5');
const rows=Object.keys(W.sheets.Config.cells).filter(k=>k.endsWith(',1')).map(k=>W.sheets.Config.cells[k]);
T('sheet has the keys + the done mark',['BOM_CFG','PO_SRC_URL','PO_SYNC_HOURS','BOM_TPL_SHORT','BOM_TPL_SHORT_VER','__MOVED_FROM_PROPERTIES'].every(k=>rows.includes(k)),rows);
T('a key never set reads null',c.schedProp_('FC_CFG')===null);
c.schedProp_('FC_CFG','{"n":2}');T('write then read in the same run',c.schedProp_('FC_CFG')==='{"n":2}'&&!('FC_CFG' in W.props));
c=W.run();const r0=W.reads();T('next run reads the new value',c.schedProp_('FC_CFG')==='{"n":2}'&&c.schedProp_('BOM_CFG')==='{"a":1}');
c=W.run();const r1=W.reads();c.schedProp_('FC_CFG');c.schedProp_('BOM_CFG');T('a later run is served from the cache (no sheet read)',W.reads()===r1,[r1,W.reads()]);
// another execution saves while this one already loaded: the old cached copy must not survive
const a=W.run();a.schedProp_('BOM_CFG');const b=W.run();b.schedProp_('BOM_CFG','{"a":2}');a.cfgCachePut_({v:'stale',m:{BOM_CFG:'{"a":1}'}});
c=W.run();T('stale cache is ignored after a write',c.schedProp_('BOM_CFG')==='{"a":2}',c.schedProp_('BOM_CFG'));
c.schedProp_('FC_CFG',null);c=W.run();T('delete',c.schedProp_('FC_CFG')===null);
c.schedProp_('ORD_VER','9');T('non-movable write goes to Script Properties',W.props.ORD_VER==='9');
let e='';try{c.schedProp_('BOM_CFG','x'.repeat(50000))}catch(x){e=x.message}T('too long is refused',/ยาวเกินไป/.test(e),e);
// the move runs once
c=W.run();W.props.RS_CFG='late';T('a property that appears later is still read',c.schedProp_('RS_CFG')==='late');
// sheet not reachable → Script Properties keep the app working
const X=world();X.down(true);c=X.run();T('sheet down: read falls back to properties',c.schedProp_('BOM_CFG')==='{"a":1}'&&('BOM_CFG' in X.props));
c.schedProp_('FC_CFG','{"f":1}');T('sheet down: write falls back to properties',X.props.FC_CFG==='{"f":1}');
X.down(false);c=X.run();T('sheet back: values are moved and readable',c.schedProp_('FC_CFG')==='{"f":1}'&&c.schedProp_('BOM_CFG')==='{"a":1}'&&!('FC_CFG' in X.props));
console.log(ok+' passed, '+bad+' failed');process.exit(bad?1:0);
