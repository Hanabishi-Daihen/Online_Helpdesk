// order pages shared sets — prints what it sees; exits 1 when the page throws an error
const OUT=__dirname+'/../out/';
const {chromium}=require('playwright');(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});const ctx=await b.newContext({viewport:{width:1440,height:1000}});
await ctx.route('https://cdnjs.cloudflare.com/ajax/libs/pdf-lib/**',r=>r.fulfill({contentType:'application/javascript',body:require('fs').readFileSync(require.resolve('pdf-lib/dist/pdf-lib.min.js'),'utf8')}));
const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));let ans='ชุดใหม่';p.on('dialog',d=>d.type()==='prompt'?d.accept(ans):d.accept());
const http=require('http'),fs=require('fs');const srv=http.createServer((q,r)=>{r.setHeader('content-type','text/html; charset=utf-8');r.end(fs.readFileSync(OUT+'order.html'))}).listen(8825);
await p.goto('http://localhost:8825/');await p.waitForTimeout(2000);
await p.locator('aside nav button',{hasText:'Order Item'}).first().click();await p.waitForTimeout(1200);
await p.locator('.zq-or-tabs button').first().click();await p.waitForTimeout(300);
await p.locator('.zq-or-type button',{hasText:'Order Project'}).click();await p.locator('input[aria-label="Trans. Production No. (โปรเจกต์)"]').fill('PP1004B');
const rows=async()=>(await p.locator('.zq-or-items tbody tr.it').evaluateAll(T=>T.map(t=>[...t.querySelectorAll('input,textarea')].slice(0,4).map(i=>i.value).join(' | '))));
// 3 rows typed by hand
for(let i=0;i<3;i++){if(i)await p.locator('button:has-text("＋ เพิ่มรายการ")').click();const tr=p.locator('.zq-or-items tbody tr.it').nth(i);const ins=tr.locator('input,textarea');await ins.nth(0).fill('4221');await ins.nth(1).fill('MANUAL ITEM '+(i+1));await tr.locator('input[inputmode=decimal],input[type=number]').first().fill('1').catch(()=>{})}
console.log('add btn',await p.locator('button:has-text("＋ เพิ่มรายการ")').innerText());
await p.locator('button:has-text("ตัวช่วยหน้าแบบ")').click();await p.waitForTimeout(600);
const K=p.locator('[role=dialog]');const G=i=>K.locator('section.zq-kit-g').nth(i);
console.log('selects',(await K.locator('.zq-kit-h select').evaluateAll(S=>S.map(s=>s.options[0].text+' ['+[...s.options].slice(1).map(o=>o.text).join(' ; ')+']'))).join(' || '));
// shared set
await K.locator('.zq-kit-h select[aria-label="ชุดที่บันทึกไว้"]').selectOption('s:LCC มาตรฐาน PEA');await p.waitForTimeout(500);
console.log('after shared set:',JSON.stringify(await G(0).locator('textarea').first().inputValue()));
// copy from old order
await K.locator('.zq-kit-h select[aria-label="คัดลอกจากใบสั่งเดิม"]').selectOption('o:ORD26-0050');await p.waitForTimeout(400);
console.log('after copy from order:',JSON.stringify(await G(0).locator('textarea').first().inputValue()),'| cab',await K.locator('.zq-kit-h input').first().inputValue());
// save shared
await K.locator('button:has-text("บันทึกเป็นชุด")').click();await p.waitForTimeout(500);
console.log('save msg:',await K.locator('.zq-kit-h small').first().innerText(),'| call',JSON.stringify(await p.evaluate(()=>{const c=window.__calls.filter(c=>c[0]==='saveKitTemplate').pop();return c&&[c[1][0],c[1][1].groups.length,c[1][3]]})));
// build > 8 rows: cam with 3 types + meter 3 rows + this card = 7 helper rows + 3 manual = 10
await K.locator('button:has-text("＋ Cam switch plate")').click();await G(1).locator('textarea').first().fill('43FLR | FAN SELECTOR SW. | LOCAL / REMOTE\n43FL | FAN CONTROL | MANUAL / AUTO | CHR10 A214 EG\n43HR | HEATER SW. | AUTO / OFF / ON | CA10 A210 EG');
await K.locator('button:has-text("＋ Ammeter / Voltmeter")').click();const M=G(2).locator('.zq-kit-mt');await M.nth(0).locator('input').first().fill('2300');await G(2).locator('button:has-text("＋ Ammeter")').click();await M.nth(1).locator('input').first().fill('750');await G(2).locator('button:has-text("＋ Voltmeter")').click();await M.nth(2).locator('input').first().fill('22');
await p.waitForTimeout(500);console.log('footer:',(await K.locator('.zq-kit-f small').innerText()).slice(0,70),'| preview heads',(await K.locator('.zq-kit-r small').allInnerTexts()).map(t=>t.slice(0,9)).join(','));
await K.locator('button:has-text("ใส่ลงใบสั่ง")').click();await p.waitForTimeout(600);
let R=await rows();console.log('rows',R.length);console.log(R.filter(r=>/P\.\d/.test(r)).join('\n'));
console.log('page label',await p.locator('.zq-or-paper').innerText().then(t=>(t.match(/Page 1\/\d+/)||[''])[0]),'| divider',await p.locator('.zq-or-items tbody tr.inf:has-text("ฟอร์มหน้า 2")').count());
// remove 2 manual rows → 8 rows → one form page → references move back
p.removeAllListeners('dialog');p.on('dialog',d=>d.accept());
for(let i=0;i<2;i++){await p.locator('.zq-or-items tbody tr.it').first().locator('button[title*="ลบ"],button:has-text("✕")').first().click();await p.waitForTimeout(300)}
R=await rows();console.log('rows after delete',R.length,'| refs',R.filter(r=>/P\.\d/.test(r)).map(r=>r.split(' | ').pop()).join(' ; '),'| page',await p.locator('.zq-or-paper').innerText().then(t=>(t.match(/Page 1\/\d+/)||[''])[0]));
console.log('errors:',errs);srv.close();await b.close();process.exit(errs.length?1:0)})();
