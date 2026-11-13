// helper cam annunciator — prints what it sees; exits 1 when the page throws an error
const OUT=__dirname+'/../out/';
const {chromium}=require('playwright');(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});const ctx=await b.newContext({viewport:{width:1440,height:1000}});
await ctx.route('https://cdnjs.cloudflare.com/ajax/libs/pdf-lib/**',r=>r.fulfill({contentType:'application/javascript',body:require('fs').readFileSync(require.resolve('pdf-lib/dist/pdf-lib.min.js'),'utf8')}));
await ctx.route('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/**',r=>r.abort());
const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));p.on('dialog',d=>d.accept());
const http=require('http'),fs=require('fs');const srv=http.createServer((q,r)=>{r.setHeader('content-type','text/html; charset=utf-8');r.end(fs.readFileSync(OUT+'order.html'))}).listen(8819);
await p.goto('http://localhost:8819/');await p.waitForTimeout(2000);
await p.locator('aside nav button',{hasText:'Order Item'}).first().click();await p.waitForTimeout(1200);
await p.locator('.zq-or-tabs button').first().click();await p.waitForTimeout(300);
await p.locator('.zq-or-type button',{hasText:'Order Project'}).click();await p.waitForTimeout(200);
await p.locator('input[aria-label="Trans. Production No. (โปรเจกต์)"]').fill('PP1004B');
await p.locator('button:has-text("ตัวช่วยหน้าแบบ")').click();await p.waitForTimeout(400);
const K=p.locator('[role=dialog][aria-label="ตัวช่วย Order Project"]');const G=i=>K.locator('section.zq-kit-g').nth(i);
await G(0).locator('button[aria-label="ลบรายการ"]').click();
await K.locator('button:has-text("＋ Cam switch plate")').click();
await G(0).locator('textarea').first().fill(`43FLR | FAN SELECTOR SW. | LOCAL / REMOTE
43FG | FAN FIRST GROUP | GROUP1 / GROUP2
43FL | FAN CONTROL | MANUAL / AUTO | CHR10 A214 EG
43FL1 | FAN GP.1 MANUAL | OFF / ON | CHR10 A214 EG
43FL2 | FAN GP.2 MANUAL | OFF / ON | CHR10 A214 EG
43HR | HEATER SW. | AUTO / OFF / ON | CA10 A210 EG
43RS | SELECTOR CONTROL | REMOTE / SCPS
43MA | OLTC CONTROL | MANUAL / AUTO | CHR10 A214 EG
43NF | OLTC CONTROL | ON / OFF | CHR10 A214 EG
43LR | OLTC MANUAL | LOWER / RAISE | CHR10 A214 EG
43IP | OLTC CONTROL | IND. / PAR.
43FR | FAN CONTROL | MANUAL / AUTO | CHR10 A214 EG
43FR1 | FAN GP.1 MANUAL | OFF / ON | CHR10 A214 EG
43FR2 | FAN GP.2 MANUAL | OFF / ON | CHR10 A214 EG`);
await K.locator('button:has-text("＋ Annunciator")').click();
const A=['TR. BUCH./TRIP','TR. OIL TEMP./TRIP','TR. WIND. TEMP./TRIP','TR. PRES./RELIEF/TRIP','DIV.SW. PRES./RELIEF/TRIP','DIV.SW.SURGE OIL/TRIP','(SPARE)','TR. BUCH./ALARM','TR. OIL TEMP./ALARM','TR. WIND. TEMP./ALARM','TR. OIL LEVEL/HIGH/ALARM','TR. OIL LEVEL/LOW/ALARM','DIV.SW. OIL LEVEL/HIGH/ALARM','DIV.SW. OIL LEVEL/LOW/ALARM','OLTC MOTOR/BKR.','TAP CHANGE/UPPER LIMIT','TAP CHANGE/LOWER LIMIT','TAP CHANGE/DELAY','AVR/FAILURE','PARALLEL/FAILURE','AVR/U<,U>,I>','OLTC/OVERCURRENT','VT./FAILURE','FAN GP.1 & GP.2/MOTOR/BKR.','FAN GP.1 & GP.2/THERM./RELAY','FAN STAND-BY/MOTOR/BKR.','FAN STAND-BY/THERM./RELAY','_','FAN MOTOR GP. 1/PROTECTION/OPERATED','FAN MOTOR GP. 2/PROTECTION/OPERATED','FAN STAND-BY/PROTECTION/OPERATED','GROUND/FAULT','AC. SUPPLY/FAILURE','AC. CONTROL/FAILURE','_'];
await G(1).locator('textarea').first().fill(A.join('\n'));
await K.locator('button:has-text("＋ Annunciator")').click();await G(2).locator('.zq-kit-sz button',{hasText:'สั่ง Engraving แยก'}).click();
await G(2).locator('textarea').first().fill('OLTC BKR./FAIL ALARM');
await p.waitForTimeout(800);
console.log('pages',await K.locator('.zq-kit-r canvas').count(),'|',(await K.locator('.zq-kit-f small').innerText()).slice(0,300));
const labels=await K.locator('.zq-kit-r small').allInnerTexts();console.log(labels.join(' | '));
const cs=K.locator('.zq-kit-r canvas');for(let i=0;i<await cs.count();i++)await cs.nth(i).screenshot({path:OUT+'k2_p'+(i+2)+'.png'});
await K.locator('button:has-text("ใส่ลงใบสั่ง")').click();await p.waitForTimeout(500);
const rows=await p.locator('.zq-or-items tbody tr.it').evaluateAll(T=>T.map(t=>[...t.querySelectorAll('input,textarea')].slice(0,7).map(i=>i.value).join(' | ')));console.log(rows.join('\n'));
console.log('note:',JSON.stringify(await p.locator('.zq-or-paper textarea').last().inputValue().catch(()=>'')));
console.log('errors:',errs);srv.close();await b.close();process.exit(errs.length?1:0)})();
