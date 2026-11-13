// ux: order bar hints, picture ↔ text link in the helper, dark mode, status colour legend — prints what it sees; exits 1 when the page throws an error
const OUT=__dirname+'/../out/';
const {chromium}=require('playwright');(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});const p=await b.newPage({viewport:{width:1440,height:1000}});const errs=[];p.on('pageerror',e=>errs.push(e.message));p.on('dialog',d=>d.accept());
const http=require('http'),fs=require('fs');const srv=http.createServer((q,r)=>{r.setHeader('content-type','text/html; charset=utf-8');r.end(fs.readFileSync(OUT+'order.html'))}).listen(8830);
await p.goto('http://localhost:8830/');await p.waitForTimeout(2000);
await p.locator('aside nav button',{hasText:'Order Item'}).first().click();await p.waitForTimeout(1200);
await p.locator('.zq-or-tabs button').first().click();await p.waitForTimeout(300);
await p.locator('.zq-or-type button',{hasText:'Order Project'}).click();await p.waitForTimeout(300);
console.log('missing chip:',await p.locator('.zq-or-miss').innerText().catch(()=>'-'),'| draft:',await p.locator('.zq-or-draft').innerText().catch(()=>'-'));
await p.locator('.zq-or-miss').click();await p.waitForTimeout(500);
console.log('after click: marked fields',await p.locator('.zq-or-paper .bad').count(),'| focused',await p.evaluate(()=>document.activeElement&&(document.activeElement.getAttribute('aria-label')||document.activeElement.placeholder||document.activeElement.tagName)));
await p.locator('input[aria-label="Trans. Production No. (โปรเจกต์)"]').fill('PP1004B');
await p.locator('button:has-text("ตัวช่วยหน้าแบบ")').click();await p.waitForTimeout(500);
const K=p.locator('[role=dialog][aria-label="ตัวช่วย Order Project"]'),ta=K.locator('section.zq-kit-g').first().locator('textarea').first();
await ta.fill('AMMETER (X1)\nAMMETER (X2)\nVOLTMETER ×2');await p.waitForTimeout(400);
const cv=K.locator('.zq-kit-r canvas').first(),bb=await cv.boundingBox();
// click the 2nd label in the picture (60x15 card: 2nd row sits 22 mm below the 1st, which starts at y = 52 mm)
await p.mouse.click(bb.x+bb.width*105/210,bb.y+bb.height*81.5/297);await p.waitForTimeout(400);
console.log('picture → text: selected',JSON.stringify(await ta.evaluate(t=>t.value.slice(t.selectionStart,t.selectionEnd))));
const px=async(xmm,ymm)=>cv.evaluate((c,[x,y])=>{const d=c.getContext('2d').getImageData(Math.round(x*3),Math.round(y*3),1,1).data;return[d[0],d[1],d[2]].join(',')},[xmm,ymm]);
console.log('highlight on 2nd label (not white):',await px(80,78),'| 1st label (white):',await px(80,56));
// caret on line 3 → both VOLTMETER labels light up
await ta.evaluate(t=>{const a=t.value.indexOf('VOLT');t.focus();t.setSelectionRange(a,a);t.dispatchEvent(new Event('select',{bubbles:true}))});await ta.press('End');await p.waitForTimeout(400);
console.log('text → picture: 3rd label',await px(80,100),'| 4th label',await px(80,122),'| 2nd back to white',await px(80,78));
await K.screenshot({path:OUT+'ux_helper_link.png'});
await K.locator('button[aria-label="ปิด"]').click();
// status colour legend + dark mode
await p.locator('.zq-look button[aria-label="ความหมายของสีสถานะ"]').click();await p.waitForTimeout(300);
console.log('legend rows',await p.locator('.zq-lg-r').count());await p.locator('.zq-lg').screenshot({path:OUT+'ux_legend.png'});await p.locator('.zq-lg button:has-text("ปิด")').click();
await p.locator('.zq-look button').first().click();await p.waitForTimeout(400);
console.log('dark class',await p.evaluate(()=>document.documentElement.classList.contains('zq-dark')),'| saved',await p.evaluate(()=>localStorage.getItem('zqDark')));
await p.screenshot({path:OUT+'ux_dark_order.png'});
await p.locator('aside nav button',{hasText:'Project 360'}).first().click().catch(()=>{});await p.waitForTimeout(1500);await p.screenshot({path:OUT+'ux_dark_projects.png'});
await p.reload();await p.waitForTimeout(1500);console.log('dark after reload',await p.evaluate(()=>document.documentElement.classList.contains('zq-dark')));
await p.locator('.zq-look button').first().click();console.log('light again',!(await p.evaluate(()=>document.documentElement.classList.contains('zq-dark'))));
console.log('errors:',errs);srv.close();await b.close();process.exit(errs.length?1:0)})();
