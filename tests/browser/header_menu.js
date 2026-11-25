// Top bar: no Google Sheets refresh, no search button; tables dense by default; Admin arranges menu groups and order in Settings
const OUT = __dirname + '/../out/', fs = require('fs');
const { chromium } = require('playwright');
(async () => {
  const b = await chromium.launch({ executablePath: process.env.CHROMIUM_PATH || undefined });
  const p = await b.newPage({ viewport: { width: 1440, height: 960 } }), errs = [], fails = [];
  p.on('pageerror', e => errs.push(e.message)); p.on('dialog', d => d.accept());
  const check = (n, c, x) => { console.log((c ? 'PASS ' : 'FAIL ') + n + (c ? '' : ' → ' + JSON.stringify(x))); c || fails.push(n); };
  const srv = require('http').createServer((q, r) => { r.setHeader('content-type', 'text/html; charset=utf-8'); r.end(fs.readFileSync(OUT + 'doc.html')); }).listen(8852);
  await p.goto('http://localhost:8852/'); await p.waitForTimeout(2000);
  const head = await p.locator('header').first().innerText();
  check('top bar: no "Google Sheets" refresh and no ค้นหา / ⌘K button', !/Google Sheets/.test(head) && !/⌘K/.test(head) && !/ค้นหา/.test(head), head);
  check('tables dense by default (no choice saved yet)', await p.evaluate(() => document.documentElement.classList.contains('zq-dense')) && await p.locator('.zq-den-btn[aria-pressed=true]').count() >= 1);
  await p.locator('.zq-den-btn').first().click(); await p.waitForTimeout(200);
  await p.reload(); await p.waitForTimeout(2000);
  check('choosing โปร่ง is remembered', await p.evaluate(() => !document.documentElement.classList.contains('zq-dense') && localStorage.getItem('zqDense') === '0'));
  // Admin → Settings
  await p.locator('button', { hasText: /Admin|ปลดล็อค/ }).first().click(); await p.waitForTimeout(500);
  const u = p.locator('input[autocomplete=username]'); if (await u.count()) await u.fill('a.admin');
  const pw = p.locator('input[type=password]'); await pw.first().fill('control'); await pw.first().press('Enter'); await p.waitForTimeout(1500);
  await p.locator('aside nav button:has-text("ตั้งค่าระบบ")').first().click(); await p.waitForTimeout(1200);
  await p.locator('.zq-sett-tabs button', { hasText: 'เมนู' }).click(); await p.waitForTimeout(300);   // the layout card sits under ☰ เมนู
  const ed = p.locator('section[aria-label="จัดหมวดและลำดับเมนู"]');
  check('Settings: menu layout editor lists the groups and menus', await ed.count() === 1 && /สถิติและต้นทุน/.test(await ed.innerText()) && await ed.locator('input[aria-label^="ชื่อหมวด"]').count() >= 3, await ed.innerText().catch(() => ''));
  await ed.locator('input[aria-label="ชื่อหมวดใหม่"]').fill('หน่วยงาน DS13'); await ed.locator('button:has-text("+ เพิ่มหมวด")').click();
  await ed.locator('select[aria-label="หมวดของ สถิติและต้นทุน"]').selectOption('หน่วยงาน DS13');
  await ed.locator('select[aria-label="หมวดของ Booking BOM (จอง Item ตาม BOM)"]').selectOption('หน่วยงาน DS13').catch(async () => { const s = ed.locator('select[aria-label^="หมวดของ Booking BOM"]'); await s.selectOption('หน่วยงาน DS13'); });
  // the new group to the top
  for (let i = 0; i < 12; i++) { const up = ed.locator('.zq-ml-g', { has: p.locator('input[value="หน่วยงาน DS13"]') }).locator('button[aria-label="เลื่อนหมวดขึ้น"]'); if (await up.isDisabled()) break; await up.click(); }
  await ed.locator('button:has-text("บันทึกเมนู")').click(); await p.waitForTimeout(1200);
  const nav = (await p.locator('aside nav').first().innerText()).split('\n').map(x => x.trim()).filter(Boolean);
  const gi = nav.indexOf('หน่วยงาน DS13▾') >= 0 ? nav.indexOf('หน่วยงาน DS13▾') : nav.findIndex(x => /^หน่วยงาน DS13/.test(x));
  check('sidebar: the new group comes first after the ungrouped menus, holding the two moved menus', gi > 0 && gi < 6 && /สถิติและต้นทุน|Booking BOM/.test(nav.slice(gi + 1, gi + 6).join(' ')) && /Booking BOM/.test(nav.slice(gi + 1, gi + 8).join(' ')), nav.slice(0, 16));
  const settingsSaved = await p.evaluate(() => { try { return JSON.stringify(window.__zqNavAll && window.__zqNavAll.length) } catch (e) { return '' } });
  check('the layout keeps every menu (none lost)', Number(settingsSaved) > 20, settingsSaved);
  await p.screenshot({ path: OUT + 'header_menu.png' });
  await ed.locator('button:has-text("ค่าเริ่มต้น")').click(); await p.waitForTimeout(1000);
  check('↺ ค่าเริ่มต้น puts the default groups back', !/หน่วยงาน DS13/.test(await p.locator('aside nav').first().innerText()));
  check('no page errors', errs.length === 0, errs);
  srv.close(); await b.close(); process.exit(fails.length || errs.length ? 1 : 0);
})();
