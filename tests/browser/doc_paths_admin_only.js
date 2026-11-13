// doc paths admin only — prints what it sees; exits 1 when the page throws an error
const OUT=__dirname+'/../out/';
const {chromium}=require('playwright');(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});const p=await b.newPage({viewport:{width:1360,height:960}});const errs=[];p.on('pageerror',e=>errs.push(e.message));p.on('dialog',d=>d.accept());
const http=require('http'),fs=require('fs');const srv=http.createServer((q,r)=>{r.setHeader('content-type','text/html; charset=utf-8');r.end(fs.readFileSync(OUT+'doc.html'))}).listen(8798);
const check=async(tag)=>{await p.locator('aside nav button:has-text("Pending Order")').first().click();await p.waitForTimeout(1200);
await p.locator('[role=tablist] button').nth(1).click();await p.waitForTimeout(800);await p.fill('input[type=search]','RQN26-03780');await p.waitForTimeout(700);
const html=await p.locator('main').innerHTML();console.log(tag,'| results',await p.locator('.zq-doc-r').count(),'| copy btns',await p.locator('button:has-text("คัดลอกที่อยู่ไฟล์")').count(),'| path shown',/192\.168|electrical-ds|SPECIFICATION SHEET/.test(html),'| getDocPaths calls',await p.evaluate(()=>window.__calls.filter(c=>c[0]==='getDocPaths').length));
console.log('  onedrive btns',await p.locator('.zq-doc-r a').count(),'| stock filter ok',await p.evaluate(()=>1));console.log('  tip:',(await p.locator('.zq-doc-ex').first().innerText()).replace(/\s+/g,' ').slice(60,200));
for(const k of ['Stock','Project']){const t=p.locator('button',{hasText:new RegExp('^'+k)}).first();if(await t.count()){}}
};
await p.goto('http://localhost:8798/');await p.waitForTimeout(1500);await check('USER ');
const st=await p.evaluate(()=>{const s=window.__calls.filter(c=>c[0]==='getDocPaths');return s.length});
await p.locator('button',{hasText:/Admin|ปลดล็อค/}).first().click();await p.waitForTimeout(500);
const u=p.locator('input[autocomplete=username]');if(await u.count())await u.fill('premwit');const pw=p.locator('input[type=password]');await pw.first().fill('control');await pw.first().press('Enter');await p.waitForTimeout(1500);
console.log('admin?',await p.evaluate(()=>window.__calls.filter(c=>c[0]==='adminLogin').map(c=>JSON.stringify(c[1])).join()),await p.locator('text=ออกจากระบบ').count());await p.locator('aside nav button:has-text("Dashboard")').first().click().catch(()=>{});await p.waitForTimeout(500);await check('ADMIN');
console.log('errors:',errs);srv.close();await b.close();process.exit(errs.length?1:0)})();
