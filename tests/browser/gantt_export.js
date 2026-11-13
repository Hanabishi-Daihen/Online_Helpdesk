// gantt export — prints what it sees; exits 1 when the page throws an error
const OUT=__dirname+'/../out/';
const {chromium}=require('playwright');(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});const ctx=await b.newContext({viewport:{width:1400,height:900}});const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
const http=require('http'),fs=require('fs');const srv=http.createServer((q,r)=>{r.setHeader('content-type','text/html; charset=utf-8');r.end(fs.readFileSync(OUT+'order.html'))}).listen(8824);
await p.goto('http://localhost:8824/');await p.waitForTimeout(1500);
const popP=ctx.waitForEvent('page');
await p.evaluate(()=>{const d=x=>{const t=new Date();t.setDate(t.getDate()+x);return String(t.getDate()).padStart(2,'0')+'/'+String(t.getMonth()+1).padStart(2,'0')+'/'+t.getFullYear()};
const items=[{projectNo:'PP1048B01',customer:'Gulf-BL Ayutthaya',process:'LCC',start:d(-5),finish:d(3)},{projectNo:'PP1048B01',customer:'Gulf-BL Ayutthaya',process:'RCC',start:d(0),finish:d(10),newDelivery:d(20)},{projectNo:'PP1045A01',customer:'PEA Phattha',process:'LCC',start:d(2),finish:d(9)}];
window.zqExportPdf('gantt',items,{LCC:{code:'LCC',name:'Local control cabinet',color:'#3b82f6'},RCC:{code:'RCC',name:'Remote control cabinet',color:'#5fb49c'}},{});
document.querySelector('#zq-pdf-dlg [data-act="ok"]').click()});
const w=await popP;await w.waitForTimeout(1500);
const html=await w.content();console.log('black bars',(html.match(/background:#141413;height:5px/g)||[]).length,'| proc bars',(html.match(/class="bar[^"]*" style="left/g)||[]).length,'| footer',/แท่งสี = Process ที่เลือก/.test(html),'| diamond',(html.match(/class="dm"/g)||[]).length);
await w.screenshot({path:OUT+'gantt1.png'});console.log('errors:',errs);srv.close();await b.close();process.exit(errs.length?1:0)})();
