// helper strip direction — prints what it sees; exits 1 when the page throws an error
const OUT=__dirname+'/../out/';
const {chromium}=require('playwright');(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});const p=await b.newPage({viewport:{width:1440,height:1000},deviceScaleFactor:2});const errs=[];p.on('pageerror',e=>errs.push(e.message));p.on('dialog',d=>d.accept());
const http=require('http'),fs=require('fs');const srv=http.createServer((q,r)=>{r.setHeader('content-type','text/html; charset=utf-8');r.end(fs.readFileSync(OUT+'order.html'))}).listen(8822);
await p.goto('http://localhost:8822/');await p.waitForTimeout(2000);
await p.locator('aside nav button',{hasText:'Order Item'}).first().click();await p.waitForTimeout(1200);
await p.locator('.zq-or-tabs button').first().click();await p.waitForTimeout(300);
await p.locator('.zq-or-type button',{hasText:'Order Project'}).click();await p.locator('input[aria-label="Trans. Production No. (โปรเจกต์)"]').fill('PP1040D');
await p.locator('button:has-text("ตัวช่วยหน้าแบบ")').click();await p.waitForTimeout(400);
const K=p.locator('[role=dialog]');await K.locator('section.zq-kit-g').nth(0).locator('button[aria-label="ลบรายการ"]').click();
for(const d of ['VER','HOR']){await K.locator('button:has-text("＋ Strip marker")').click();const G=K.locator('section.zq-kit-g').last();await G.locator('select').first().selectOption(d);await G.locator('textarea').fill('1-10\n--\nM11-M13\n--\nL+, L-, PE*2\n--\n100-109');}
await p.waitForTimeout(600);const F=K.locator('figure.zq-kit-term');console.log('figs',await F.count());for(let i=0;i<2;i++)await F.nth(i).screenshot({path:OUT+'term_'+i+'.png'});const cs=K.locator('.zq-kit-r canvas');for(let i=0;i<2;i++)await cs.nth(i).screenshot({path:OUT+'strip_'+i+'.png'});
console.log('rows',(await K.locator('.zq-kit-f small').innerText()).slice(0,200));console.log('errors:',errs);srv.close();await b.close();process.exit(errs.length?1:0)})();
