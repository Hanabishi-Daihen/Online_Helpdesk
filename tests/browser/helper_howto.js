// helper howto — prints what it sees; exits 1 when the page throws an error
const OUT=__dirname+'/../out/';
const {chromium}=require('playwright');(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});const p=await b.newPage({viewport:{width:1440,height:1000},deviceScaleFactor:2});const errs=[];p.on('pageerror',e=>errs.push(e.message));p.on('dialog',d=>d.accept());
const http=require('http'),fs=require('fs');const srv=http.createServer((q,r)=>{r.setHeader('content-type','text/html; charset=utf-8');r.end(fs.readFileSync(OUT+'order.html'))}).listen(8820);
await p.goto('http://localhost:8820/');await p.waitForTimeout(2000);
await p.locator('aside nav button',{hasText:'Order Item'}).first().click();await p.waitForTimeout(1200);
await p.locator('.zq-or-tabs button').first().click();await p.waitForTimeout(300);
await p.locator('.zq-or-type button',{hasText:'Order Project'}).click();await p.locator('input[aria-label="Trans. Production No. (โปรเจกต์)"]').fill('PP1004B');
await p.locator('button:has-text("ตัวช่วยหน้าแบบ")').click();await p.waitForTimeout(400);
const K=p.locator('[role=dialog]');await K.locator('button:has-text("＋ Annunciator")').click();await p.waitForTimeout(200);
await K.locator('section.zq-kit-g').nth(1).locator('button:has-text("ใส่ตัวอย่าง")').click();await p.waitForTimeout(400);
console.log('help boxes',await K.locator('.zq-kit-help').count(),'| ann text',JSON.stringify((await K.locator('section.zq-kit-g').nth(1).locator('textarea').first().inputValue()).slice(0,40)),'| pages',await K.locator('.zq-kit-r canvas').count());
await K.locator('.zq-kit-l').screenshot({path:OUT+'help1.png'});
await K.locator('.zq-kit-help summary').first().click();await p.waitForTimeout(200);console.log('after hide: open boxes',await K.locator('.zq-kit-help[open]').count(),'| show btn',await K.locator('button:has-text("แสดงวิธีพิมพ์")').count());
console.log('errors:',errs);srv.close();await b.close();process.exit(errs.length?1:0)})();
