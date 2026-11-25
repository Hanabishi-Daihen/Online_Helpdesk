// สถิติและต้นทุน: four tabs from one summary; a user sees no money, Admin sees cost columns and can rebuild
const OUT = __dirname + '/../out/', fs = require('fs');
const { chromium } = require('playwright');
const M = Array.from({ length: 24 }, (_, i) => { const d = new Date(2024, 11 + i, 1); return d.getFullYear() + ('0' + (d.getMonth() + 1)).slice(-2); });
const S = { v: 1, built: '25/11/2026 03:20:00', notes: [], months: M,
  depts: [{ name: 'CONTROL WIRING', $m: M.map((m, i) => 100000 + i * 5000), n: M.map((m, i) => 40 + i), $tot: 3780000 }, { name: 'TANKING', $m: M.map(() => 2000), n: M.map(() => 3), $tot: 48000 }],
  $topCost: [{ code: '42046001', desc: 'MCB C6', unit: 'PC', q: 100, r: 2, $c: 9800, $cr: 200, rp: 2, depts: 1 }], topQty: [{ code: '46700402', desc: 'TERMINAL UK6N', unit: 'PC', q: 300, r: 0, $c: 6000, $cr: 0, rp: 0, depts: 1 }],
  returns: [{ code: '42046001', desc: 'MCB C6', unit: 'PC', q: 100, r: 2, $c: 9800, $cr: 200, rp: 2, depts: 1 }], $costAll12: 16300, $share20: 100, codes12: 3,
  $projects: [{ project: 'PP1048B02', cat: '1048P2', title: 'Local Control Cabinet (LCC)', type: 'LCC', n: 5, $c: 16400, $cr: 200, last: '09/10/2026' }], $types: [{ type: 'LCC', n: 1, $avg: 16400, $min: 16400, $max: 16400 }],
  bom: { projects: [{ project: 'PP1048B02', title: 'LCC', status: 'closed', imported: '25/07/2026', n: 3, issued: 2, notIssued: 1, under: 0, extra: 1, extraCodes: ['47280030'], acc: 50 }], avg: 50, over: [{ code: '46700402', desc: 'TERMINAL UK6N', n: 3, pct: 40 }] },
  stock: { $value: 120600, buckets: [{ name: 'ไม่เกิน 1 ปี', n: 2, $v: 100300 }, { name: '1–2 ปี', n: 0, $v: 0 }, { name: '2–3 ปี', n: 0, $v: 0 }, { name: 'เกิน 3 ปี', n: 1, $v: 20000 }, { name: 'ไม่เคยเบิก', n: 1, $v: 300 }], items: 4 },
  $dead: [{ code: '99990001', desc: 'OLD RELAY', unit: 'PC', on: 40, $v: 20000, last: '05/07/2023', idle: 1240, act: false }], dead: [{ code: '99990002', desc: 'NEVER USED', unit: 'PC', on: 10, last: '', idle: null, act: true }],
  groups: [{ g: '4670', n: 1, q: 300, $v: 100000, $use: 6000, turn: 0.06 }], low: [{ code: '42046001', desc: 'MCB C6', unit: 'PC', on: 3, use12: 98, cover: 11, lt: 45, $v: 300 }],
  over: [{ code: '46700402', desc: 'TERMINAL UK6N', unit: 'PC', on: 5000, use12: 300, cover: 6083, lt: 30, $v: 100000 }],
  po: { open: 2, ages: [['≤ 30 วัน', 0], ['31–60 วัน', 0], ['61–90 วัน', 1], ['เกิน 90 วัน', 1]], prpo: 6, podue: 55, def: 30, lateN: 1, late: [{ code: '42046001', name: 'MCB', pr: 'PR1', po: 'PO1', status: 'Open', due: '01/10/2026', days: 55, qty: 5, unit: 'PC', by: 'DS13' }], diff: [{ code: '42046001', desc: 'MCB C6', n: 2, pend: 80, prpo: 6, podue: 70, rt: 45 }] },
  minmax: [{ code: '42046001', desc: 'MCB C6', unit: 'PC', use12: 98, lt: 45, ltSrc: 'rt', min: 19, max: 44, curMin: 1, curMax: 2, on: 3, $c: 9800 }], secs: 4 };
const pub = JSON.parse(JSON.stringify(S, (k, v) => k && k[0] === '$' ? undefined : v));
(async () => {
  const b = await chromium.launch({ executablePath: process.env.CHROMIUM_PATH || undefined });
  const p = await b.newPage({ viewport: { width: 1440, height: 960 } }), errs = [], fails = [];
  p.on('pageerror', e => errs.push(e.message)); p.on('dialog', d => d.accept());
  const check = (n, c, x) => { console.log((c ? 'PASS ' : 'FAIL ') + n + (c ? '' : ' → ' + JSON.stringify(x))); c || fails.push(n); };
  const srv = require('http').createServer((q, r) => { r.setHeader('content-type', 'text/html; charset=utf-8'); r.end(fs.readFileSync(OUT + 'doc.html')); }).listen(8851);
  await p.goto('http://localhost:8851/'); await p.waitForTimeout(1500);
  await p.evaluate(([S, pub]) => {
    const G = window.google.script, d = Object.getOwnPropertyDescriptor(G, 'run'); window.__rb = 0;
    Object.defineProperty(G, 'run', { configurable: true, get() {
      const inner = d.get.call(G); let ok = () => {};
      const w = { withSuccessHandler(f) { ok = f; inner.withSuccessHandler(f); return w; }, withFailureHandler(f) { inner.withFailureHandler(f); return w; },
        api(t, fn, a) {
          if (fn === 'getStatSummary') { setTimeout(() => ok(window.__adm ? { ...S, admin: true } : { ...pub, admin: false }), 20); return; }
          if (fn === 'statRebuild') { window.__rb++; setTimeout(() => ok({ ...S, admin: true, built: '25/11/2026 09:00:00' }), 30); return; }
          if (fn === 'getCodeGroups') { setTimeout(() => ok({ '4670': 'Terminal block' }), 10); return; }
          return inner.api(t, fn, a);
        } };
      return new Proxy(w, { get: (o, k) => k in o ? o[k] : inner[k] });
    } });
  }, [S, pub]);
  const open = async () => { await p.locator('aside nav button:has-text("สถิติและต้นทุน")').first().click(); await p.waitForTimeout(1200); };
  const tab = async t => { await p.locator('.zq-sx-tabs button', { hasText: t }).click(); await p.waitForTimeout(300); return p.locator('.zq-sx').innerText(); };
  await open();
  let t = await tab('การเบิกใช้');
  check('user: menu item and page; no ฿ anywhere, no rebuild button', /สถิติและต้นทุน/.test(t) && !/฿/.test(t) && await p.locator('button:has-text("คำนวณใหม่")').count() === 0, t.slice(0, 300));
  check('user: monthly chart of lines per month (24 bars), top CODE by quantity', (await p.locator('.zq-sx-bar').count()) === 24 && /Top CODE ตามจำนวนเบิก/.test(t) && /46700402/.test(t), t.slice(0, 800));
  t = await tab('สุขภาพสต็อก');
  check('user stock: counts by idle time, idle list by days, no value', /จำนวน CODE ตามเวลาที่ไม่เคลื่อนไหว/.test(t) && /99990002/.test(t) && !/99990001/.test(t) && !/฿/.test(t), t.slice(0, 600));
  for (const x of ['BOM เทียบของจริง', 'การจัดซื้อ']) { t = await tab(x); check('user ' + x + ': no money', !/฿/.test(t)); }
  // Admin
  await p.evaluate(() => { window.__adm = true; });
  await p.locator('button', { hasText: /Admin|ปลดล็อค/ }).first().click(); await p.waitForTimeout(500);
  const u = p.locator('input[autocomplete=username]'); if (await u.count()) await u.fill('a.admin');
  const pw = p.locator('input[type=password]'); await pw.first().fill('control'); await pw.first().press('Enter'); await p.waitForTimeout(1500);
  await open();
  t = await tab('การเบิกใช้');
  check('Admin usage: 12-month cost, top-20 share, cost per CODE, project and type tables', /฿16,300/.test(t) && /100%/.test(t) && /฿9,800/.test(t) && /ต้นทุนเฉลี่ยตามชนิดงาน/.test(t) && /PP1048B02/.test(t), t.slice(0, 900));
  await p.locator('.zq-sx-bar').last().hover(); await p.waitForTimeout(200);
  check('chart: hovering a bar shows the month and value', /฿215,000/.test(await p.locator('.zq-sx-bar').last().innerText()), await p.locator('.zq-sx-bar').last().innerText());
  await p.locator('select[aria-label="แผนก"]').selectOption('TANKING'); await p.waitForTimeout(200);
  check('chart: department switch', /฿2,000/.test(await p.locator('.zq-sx-bar').last().innerText()));
  t = await tab('BOM เทียบของจริง');
  check('BOM tab: accuracy, issued / not issued / outside BOM, CODEs used over BOM', /50%/.test(t) && /PP1048B02/.test(t) && /CODE ที่มักเบิกเกิน BOM/.test(t) && /\+40%/.test(t), t.slice(0, 700));
  t = await tab('สุขภาพสต็อก');
  check('Admin stock: value, idle list by value, cover and turnover', /฿120,600/.test(t) && /99990001/.test(t) && /11 วัน/.test(t) && /16.7 ปี/.test(t) && /0.06×/.test(t) && /Terminal block/.test(t), t.slice(0, 900));
  t = await tab('การจัดซื้อ');
  check('purchasing: open lines, late list, lead time vs StockRT, suggested Min / Max', /PR1 \/ PO1/.test(t) && /55 วัน/.test(t) && /80 วัน/.test(t) && /19 \/ 44/.test(t) && /1 \/ 2/.test(t), t.slice(0, 900));
  await p.screenshot({ path: OUT + 'stat_summary.png' });
  await p.locator('button:has-text("คำนวณใหม่")').click(); await p.waitForTimeout(600);
  check('Admin: ↻ คำนวณใหม่ rebuilds and shows the new time', (await p.evaluate(() => window.__rb)) === 1 && /09:00:00/.test(await p.locator('.zq-sx').innerText()));
  check('no page errors', errs.length === 0, errs);
  srv.close(); await b.close(); process.exit(fails.length || errs.length ? 1 : 0);
})();
