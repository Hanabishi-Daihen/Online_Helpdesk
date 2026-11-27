// ตั้งค่าระบบ: one topic at a time (tabs) + search across every card
const OUT = __dirname + '/../out/', fs = require('fs');
const { chromium } = require('playwright');
(async () => {
  const b = await chromium.launch({ executablePath: process.env.CHROMIUM_PATH || undefined });
  const p = await b.newPage({ viewport: { width: 1440, height: 960 } }), errs = [], fails = [];
  p.on('pageerror', e => errs.push(e.message));
  const check = (n, c, x) => { console.log((c ? 'PASS ' : 'FAIL ') + n + (c ? '' : ' → ' + JSON.stringify(x))); c || fails.push(n); };
  const srv = require('http').createServer((q, r) => { r.setHeader('content-type', 'text/html; charset=utf-8'); r.end(fs.readFileSync(OUT + 'doc.html')); }).listen(8853);
  await p.goto('http://localhost:8853/'); await p.waitForTimeout(1500);
  await p.evaluate(() => {
    const G = window.google.script, d = Object.getOwnPropertyDescriptor(G, 'run'); window.__unit = [];
    let cfg = { by: 'DS13', mfDept: 'CONTROL WIRING', svPrefix: 'SV', projLetter: 'C', codeLen: 8 };
    Object.defineProperty(G, 'run', { configurable: true, get() {
      const inner = d.get.call(G); let ok = () => {};
      const w = { withSuccessHandler(f) { ok = f; inner.withSuccessHandler(f); return w; }, withFailureHandler(f) { inner.withFailureHandler(f); return w; },
        api(t, fn, a) {
          if (fn === 'getUnitCfg') { setTimeout(() => ok({ cfg, def: cfg, admin: true }), 20); return; }
          if (fn === 'saveUnitCfg') { window.__unit.push(a[1]); cfg = { ...a[1], codeLen: Number(a[1].codeLen) }; setTimeout(() => ok({ cfg, def: cfg, admin: true }), 20); return; }
          return inner.api(t, fn, a);
        } };
      return new Proxy(w, { get: (o, k) => k in o ? o[k] : inner[k] });
    } });
  });
  await p.locator('button', { hasText: /Admin|ปลดล็อค/ }).first().click(); await p.waitForTimeout(400);
  const u = p.locator('input[autocomplete=username]'); if (await u.count()) await u.fill('a.admin');
  const pw = p.locator('input[type=password]'); await pw.first().fill('control'); await pw.first().press('Enter'); await p.waitForTimeout(1500);
  await p.locator('aside nav button:has-text("ตั้งค่าระบบ")').first().click(); await p.waitForTimeout(1800);
  const vis = () => p.evaluate(() => Array.from(document.querySelectorAll('.zq-sett section')).filter(s => !s.parentElement.closest('section') && s.offsetParent).map(s => { const h = s.querySelector('h2,h3,h4'); return (h ? h.textContent : Array.from(s.childNodes).filter(x => x.nodeName !== 'STYLE').map(x => x.textContent).join(' ')).trim().slice(0, 40); }));
  const tabs = await p.locator('.zq-sett-tabs button').allInnerTexts();
  check('seven topics as tabs (หน่วยงาน first), each with its card count', tabs.length === 7 && /หน่วยงาน/.test(tabs[0]) && /ข้อมูลและการซิงค์/.test(tabs[1]) && /บัญชีและสิทธิ์/.test(tabs.join()) && /การแจ้งเตือน/.test(tabs.join()), tabs);
  let v = await vis();
  check('ข้อมูลและการซิงค์: Google Sheets + every OneDrive sync card, nothing else', v.length >= 7 && v.every(x => /Google Sheets|ซิงค์/.test(x)), v);
  const tab = async t => { await p.locator('.zq-sett-tabs button', { hasText: t }).click(); await p.waitForTimeout(300); return vis(); };
  v = await tab('หน่วยงาน');
  const un = p.locator('section[aria-label="หน่วยงาน"]');
  check('หน่วยงาน: one card with the orderer, MF department, skipped prefixes, By Project letter, CODE length', v.length === 1 && /หน่วยงาน/.test(v[0]) && (await un.locator('input').count()) === 5 && await un.locator('input').first().inputValue() === 'DS13', v);
  await un.locator('input').nth(2).fill('SV, SR'); await un.locator('button:has-text("บันทึก")').click(); await p.waitForTimeout(500);
  const su = await p.evaluate(() => window.__unit[0]);
  check('หน่วยงาน: saves all five values', su && su.by === 'DS13' && su.mfDept === 'CONTROL WIRING' && su.svPrefix === 'SV, SR' && su.projLetter === 'C' && Number(su.codeLen) === 8 && /บันทึกแล้ว/.test(await un.innerText()), su);
  v = await tab('บัญชีและสิทธิ์');
  check('บัญชีและสิทธิ์: Admin access, Admin accounts, user accounts, price code', v.length === 4 && v.some(x => /Admin Access/.test(x)) && v.some(x => /บัญชีผู้ใช้งาน/.test(x)) && v.some(x => /รหัสดูราคา/.test(x)), v);
  v = await tab('เมนู');
  const body = () => p.evaluate(() => Array.from(document.querySelectorAll('.zq-sett > :not(.zq-sett-hd):not(style)')).map(e => e.innerText).join('\n'));
  const menuTxt = await body();
  check('เมนู: who sees which menu (from the theme card, without the colours), menu layout, extra menus', v.some(x => /จัดหมวดและลำดับเมนู/.test(x)) && v.some(x => /เมนูเพิ่มเติม/.test(x)) && /สิทธิ์การแสดงผลฟังก์ชั่น/.test(menuTxt) && !/ชุดสี|ฟอนต์/.test(menuTxt), v);
  v = await tab('ธีมและหน้าตา');
  const lookTxt = await body();
  check('ธีมและหน้าตา: theme card without the menu visibility block', v.length === 1 && /Theme/.test(v[0]) && !/สิทธิ์การแสดงผลฟังก์ชั่น/.test(lookTxt), v);
  v = await tab('การแจ้งเตือน'); check('การแจ้งเตือน: e-mail card', v.length === 1 && /อีเมล/.test(v[0]), v);
  v = await tab('อื่น ๆ'); check('อื่น ๆ: Project 360 and connection speed', v.some(x => /Project 360/.test(x)) && v.some(x => /ความเร็ว/.test(x)), v);
  await p.screenshot({ path: OUT + 'settings_tabs.png' });
  await p.locator('input[aria-label="ค้นหาการตั้งค่า"]').fill('รหัส'); await p.waitForTimeout(500);
  v = await vis();
  check('search "รหัส": cards from several topics that mention it, tabs hidden', v.some(x => /รหัสดูราคา/.test(x)) && v.length >= 2 && await p.locator('.zq-sett-tabs').count() === 0 && /พบ \d+ หัวข้อ/.test(await p.locator('.zq-sett-hd').innerText()), v);
  await p.locator('.zq-sett-hd button:has-text("ล้างคำค้น")').click(); await p.waitForTimeout(300);
  check('clear search → back to the last tab', (await vis()).some(x => /Project 360/.test(x)));
  check('remembered tab after leaving and coming back', await (async () => { await p.locator('aside nav button:has-text("งานของฉัน")').first().click(); await p.waitForTimeout(800); await p.locator('aside nav button:has-text("ตั้งค่าระบบ")').first().click(); await p.waitForTimeout(1200); return /อื่น ๆ/.test(await p.locator('.zq-sett-tabs button[aria-selected=true]').innerText()); })());
  check('no page errors', errs.length === 0, errs);
  srv.close(); await b.close(); process.exit(fails.length || errs.length ? 1 : 0);
})();
