// Order form pages (more than 8 items) + shared helper sets
const vm=require('vm'),fs=require('fs'),path=require('path');
const ROOT=require('path').resolve(__dirname,'../..');
const src=fs.readdirSync(ROOT).filter(f=>f.endsWith('.gs')).map(f=>fs.readFileSync(path.join(ROOT,f),'utf8')).join('\n');
// a tiny in-memory spreadsheet: sheets of string cells
const mkSheet=name=>{const cells={};const sh={name,cells,getLastRow:()=>Object.keys(cells).reduce((m,k)=>Math.max(m,+k.split(',')[0]),0),getMaxColumns:()=>26,
 getRange(r,c,nr,nc){nr=nr||1;nc=nc||1;const R={setValues(v){v.forEach((row,i)=>row.forEach((x,j)=>cells[(r+i)+','+(c+j)]=String(x)));return R},getDisplayValues:()=>Array.from({length:nr},(_,i)=>Array.from({length:nc},(_,j)=>cells[(r+i)+','+(c+j)]||'')),setFontWeight:()=>R,setNumberFormat:()=>R,getValue:()=>cells[r+','+c]||''};return R},
 setFrozenRows(){},deleteRow(r){const keep={};Object.keys(cells).forEach(k=>{const[a,b]=k.split(',').map(Number);if(a<r)keep[k]=cells[k];else if(a>r)keep[(a-1)+','+b]=cells[k]});Object.keys(cells).forEach(k=>delete cells[k]);Object.assign(cells,keep)}};return sh};
const sheets={};const SS={getSheetByName:n=>sheets[n]||null,insertSheet:n=>sheets[n]=mkSheet(n)};
const cache={adm_TOK:'{"u":"a","n":"Admin"}'};
const ctx={console:{log(){},warn(){}},JSON,Math,Date,String,Number,Array,Object,Error,RegExp,
 PropertiesService:{getScriptProperties:()=>({getProperty:()=>null,setProperty(){}})},Utilities:{formatDate:()=>'13/11/2026 10:00:00'},
 SpreadsheetApp:{openById:()=>SS},LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){}})},CacheService:{getScriptCache:()=>({get:k=>cache[k]||null,put(){}})}};
ctx.globalThis=ctx;vm.createContext(ctx);vm.runInContext(src,ctx);ctx.persistFiles_=f=>f||[];ctx.audit_=()=>{};
let ok=0,bad=0;const T=(n,c,x)=>{c?ok++:bad++;console.log((c?'PASS ':'FAIL ')+n+(c?'':' → '+JSON.stringify(x)))};
const it=n=>Array.from({length:n},(_,i)=>({code:'3694',material:'ITEM '+(i+1),qty:'1',unit:'PCS'}));
const base={des:'C.Premwit',delivery:'30/11/2026',issueDate:'13/11/2026',otype:'project',prodNo:'PP1044D',name:'LCC'};
const E=o=>{try{ctx.ordClean_(o);return ''}catch(e){return e.message}};
T('9 items accepted',E({...base,items:it(9)})==='');
T('24 items accepted',E({...base,items:it(24)})==='');
T('25 items rejected',/24/.test(E({...base,items:it(25)})),E({...base,items:it(25)}));
T('form pages',ctx.ordFormPages_({items:it(1)})===1&&ctx.ordFormPages_({items:it(8)})===1&&ctx.ordFormPages_({items:it(9)})===2&&ctx.ordFormPages_({items:it(24)})===3&&ctx.ordFormPages_({items:[]})===1);
// ordFill_ on a recording sheet: which values land on which page
const rec=()=>{const set={};let page='';const cell=a=>{const R=new Proxy({},{get:(t,k)=>k==='setValue'?v=>{set[a]=v;return R}:k==='getFontSizes'?()=>[[10]]:()=>R});return R};
 const sh=new Proxy({},{get:(t,k)=>k==='getRange'?(a,b,c,d)=>{const key=typeof a==='string'?a:a+','+b;const R=cell(key);return new Proxy({},{get:(t2,k2)=>k2==='createTextFinder'?()=>({useRegularExpression(){return this},findNext:()=>({setValue:v=>{page=String(v)}})}):R[k2]})}:k==='getColumnWidth'?()=>60:k==='getRowHeight'?()=>20:()=>sh});
 return{sh,set,page:()=>page}};
const o=ctx.ordClean_({...base,items:it(11),note:'N1',kit:{cab:'LCC',pages:2,groups:[{kind:'card',w:60,h:15,text:'A'}]}});
const p1=rec(),p2=rec();ctx.ordFill_(p1.sh,o,{i:0,F:2,N:4});ctx.ordFill_(p2.sh,o,{i:1,F:2,N:4});
T('page 1: items 1-8',p1.set.B21==='1'&&p1.set.F21==='ITEM 1'&&p1.set.F35==='ITEM 8',[p1.set.B21,p1.set.F21,p1.set.F35]);
T('page 2: items 9-11 numbered on',p2.set.B21==='9'&&p2.set.F21==='ITEM 9'&&p2.set.B25==='11'&&p2.set.F27==='',[p2.set.B21,p2.set.F21,p2.set.B25,p2.set.F27]);
T('page 1 keeps the note, page 2 says continued',p1.set.D38==='N1'&&/ต่อจากหน้า 1/.test(p2.set.D38),[p1.set.D38,p2.set.D38]);
const d1=rec();ctx.ordFill_(d1.sh,o);T('default call = page 1 of 1 + drawing pages',d1.set.B21==='1'&&d1.set.F21==='ITEM 1');
// shared sets
const kit={cab:'LCC',groups:[{kind:'card',w:60,h:15,text:'AMMETER'}]};
let r=ctx.saveKitTemplate('LCC มาตรฐาน',kit,'C.Premwit',false);T('save set',r.ok&&r.list.length===1&&r.list[0].name==='LCC มาตรฐาน'&&r.list[0].by==='C.Premwit',r);
r=ctx.saveKitTemplate('LCC มาตรฐาน',{cab:'RCC',groups:kit.groups},'X',false);T('same name asks first',r.exists===true&&ctx.getKitTemplate('LCC มาตรฐาน').cab==='LCC',r);
r=ctx.saveKitTemplate('LCC มาตรฐาน',{cab:'RCC',groups:kit.groups},'X',true);T('replace on confirm',r.ok&&r.list.length===1&&ctx.getKitTemplate('LCC มาตรฐาน').cab==='RCC',r);
ctx.saveKitTemplate('B',kit,'Y',false);T('list has both',ctx.getKitTemplates().length===2);
let e='';try{ctx.saveKitTemplate('',kit)}catch(x){e=x.message}T('name required',/ตั้งชื่อ/.test(e),e);
e='';try{ctx.saveKitTemplate('E',{groups:[]})}catch(x){e=x.message}T('empty set rejected',/ยังไม่มีรายการ/.test(e),e);
e='';try{ctx.deleteKitTemplate('bad','B')}catch(x){e=x.message}T('delete needs Admin',/Admin/.test(e)&&ctx.getKitTemplates().length===2,e);
r=ctx.deleteKitTemplate('TOK','B');T('Admin deletes',r.ok&&r.list.length===1&&r.list[0].name==='LCC มาตรฐาน',r);
e='';try{ctx.getKitTemplate('B')}catch(x){e=x.message}T('missing set',/ไม่พบชุด/.test(e),e);
console.log(ok+' passed, '+bad+' failed');process.exit(bad?1:0);
