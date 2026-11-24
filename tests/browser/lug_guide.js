// หางปลา × อุปกรณ์: by device (a relay lists its sockets), by lug (drawing + devices), the matrix, fit rules, order a lug, Admin edit
const OUT = __dirname + '/../out/', fs = require('fs'), vm = require('vm');
const { chromium } = require('playwright');
const ctx = {}; vm.runInNewContext(fs.readFileSync(__dirname + '/../../LugGuide.gs', 'utf8'), ctx);
const SEED = ctx.LUG_SEED;
(async () => {
  const b = await chromium.launch({ executablePath: process.env.CHROMIUM_PATH || undefined });
  const p = await b.newPage({ viewport: { width: 1440, height: 960 } }), errs = [], fails = [];
  p.on('pageerror', e => errs.push(e.message)); p.on('dialog', d => d.accept());
  const check = (n, c, x) => { console.log((c ? 'PASS ' : 'FAIL ') + n + (c ? '' : ' → ' + JSON.stringify(x))); c || fails.push(n); };
  const srv = require('http').createServer((q, r) => { r.setHeader('content-type', 'text/html; charset=utf-8'); r.end(fs.readFileSync(OUT + 'doc.html')); }).listen(8849);
  await p.goto('http://localhost:8849/'); await p.waitForTimeout(1500);
  await p.evaluate(seed => {
    const G = window.google.script, d = Object.getOwnPropertyDescriptor(G, 'run'); window.__saved = [];
    const stock = [['43401016', 'AUXILIARY RELAY', 'MY4N-GS-R DC100/110V OMRON', '48', 'PCS', 'STORE1'], ['47280021', 'TERMINAL', 'RF1.5-3 T-LUG', '6200', 'PCS', 'STORE1'], ['47280030', 'TERMINAL', 'VF2-3M KENION', '300', 'PCS', 'STORE1'],
      ['47180060', 'TERMINAL', 'RNBL2-5 KST', '1000', 'PCS', 'STORE1'], ['47280022', 'TERMINAL', 'YF1.5-3S T-LUG', '2000', 'PCS', 'STORE1'], ['47280036', 'Cable end sleeve', 'CE1.5-10 T-LUG', '1000', 'PCS', 'STORE1']];
    const tsv = stock.map(f => [f[0], f[1], f[2], f[3], f[4], f[5], '', 'DS13', '', '1'].join('\t')).join('\n');
    window.__bundle = (async () => { const cs = new CompressionStream('gzip'), w = cs.writable.getWriter(); w.write(new TextEncoder().encode(tsv)); w.close();
      const buf = new Uint8Array(await new Response(cs.readable).arrayBuffer()); let bin = ''; buf.forEach(c => bin += String.fromCharCode(c)); return btoa(bin); })();
    const data = () => ({ T: seed.T, F: seed.F, D: seed.D, TKN: seed.TKN, CONN: seed.CONN, CONNN: seed.CONNN, custom: false, admin: window.__adm || false });
    Object.defineProperty(G, 'run', { configurable: true, get() {
      const inner = d.get.call(G); let ok = () => {};
      const w = { withSuccessHandler(f) { ok = f; inner.withSuccessHandler(f); return w; }, withFailureHandler(f) { inner.withFailureHandler(f); return w; },
        api(t, fn, a) {
          if (fn === 'getStockMeta') { setTimeout(() => ok({ hash: 'lg12345678', rows: stock.length, items: 5, locations: ['STORE1'], bys: [], updated: '24/11/2026 05:00' }), 20); return; }
          if (fn === 'getStockBundle') { window.__bundle.then(b64 => ok({ hash: 'lg12345678', b64, v: 3 })); return; }
          if (fn === 'getCodeGroups' || fn === 'checkOrderDup') { setTimeout(() => ok({}), 10); return; }
          if (fn === 'getLugGuide') { setTimeout(() => ok(data()), 20); return; }
          if (fn === 'saveLug') { window.__saved.push(JSON.parse(JSON.stringify(a))); const r = data(); if (a[1] === 'F') r.F = { ...r.F, [a[2].key]: a[2] }; setTimeout(() => ok({ ...r, custom: true }), 30); return; }
          return inner.api(t, fn, a);
        } };
      return new Proxy(w, { get: (o, k) => k in o ? o[k] : inner[k] });
    } });
  }, SEED);
  // the fit rules on their own
  // the fit rules, taken from index.html and run here
  const html = fs.readFileSync(__dirname + '/../../index.html', 'utf8'), a0 = html.indexOf('zqLugKind='), a1 = html.indexOf('zqLugSvg=');
  const E = {}; vm.runInNewContext('var ' + html.slice(a0, a1).replace(/^\/\*.*\*\/$/gm, '').trim().replace(/,$/, ';') + 'R={zqLugFit};', E);
  const G = SEED, TT = id => G.T.find(x => x.id === id), FIT = E.R.zqLugFit;
  const fit = { ring_m3_pyf: FIT(G.F.PYF14A, TT('47280021'), G)[0], ring_m3_finger: FIT(G.F.PYF14A_N, TT('47280021'), G)[0], too_wide: FIT(G.F.PYF14A, TT('47280030'), G), spade_finger: FIT(G.F.PYF14A_N, TT('47280022'), G)[0],
    ring_m5_pyf: FIT(G.F.PYF14A, TT('47180060'), G), big_on_pt30: FIT(G.F.TB_PT30, TT('47280024'), G), cord_box: FIT(G.F.MCB, TT('47280036'), G)[0],
    ring_captive: FIT(G.F.SC03, TT('47180075'), G), ov: FIT({ ...G.F.PYF14A, ov: { '47280021': [0, 'ขั้วจริงแคบ'] } }, TT('47280021'), G) };
  check('fit: M3 ring fits the PYF14A socket (△: captive screw comes out); the finger-protection PYF14A-N/-E takes the fork, not the ring', fit.ring_m3_pyf === 2 && fit.ring_m3_finger === 0 && fit.spade_finger === 1, fit);
  check('fit: wrong screw / wire too big / lug too wide are refused with the reason', fit.too_wide[0] === 0 && /กว้าง 6.6 > 6/.test(fit.too_wide[1]) && fit.ring_m5_pyf[0] === 0 && /M5 ≠ สกรู M3/.test(fit.ring_m5_pyf[1]) && fit.big_on_pt30[0] === 0 && /ใหญ่กว่าขั้วรับ/.test(fit.big_on_pt30[1]), fit);
  check('fit: ferrule in a box clamp ✓ · ring on a captive screw △ · Admin exception wins', fit.cord_box === 1 && fit.ring_captive[0] === 2 && /ถอดสกรู/.test(fit.ring_captive[1]) && fit.ov[0] === 0 && /Admin/.test(fit.ov[1]), fit);
  await p.locator('aside nav button:has-text("หางปลา")').first().click(); await p.waitForTimeout(1500);
  await p.locator('input[aria-label="ค้นหา"]').fill('43401016'); await p.waitForTimeout(600);
  let t = await p.locator('.zq-st-det').innerText();
  check('by device: the MY4N relay shows its three sockets, plugs in (wired at the socket), name from Item stock', /MY4N-GS-R/.test(t) && /ขั้วต่อ 1 · Omron Socket PYF14A \(MY4N\)/.test(t) && /PYF14A-N/.test(t) && /PYFZ-14/.test(t) && /ต่อสายที่ Socket/.test(t), t.slice(0, 600));
  const cards = p.locator('.zq-lg-fam');
  const c1 = await cards.nth(0).innerText(), c2 = await cards.nth(1).innerText();
  check('by device: PYF14A lists the M3 ring RF1.5-3 with its Stock; the finger-protection socket does not; the 6.6 mm VF2-3M is not listed', /RF1.5-3/.test(c1) && /Stock 6,200/.test(c1) && !/RF1.5-3/.test(c2) && !/VF2-3M/.test(c1) && /YF1.5-3S/.test(c2), [c1.slice(0, 400), c2.slice(0, 400)]);
  check('by device: no Admin buttons for a user', await p.locator('button:has-text("แก้กลุ่มขั้วต่อ")').count() === 0);
  await cards.nth(0).locator('tr:has-text("RF1.5-3") button:has-text("+ ใบสั่ง")').first().click(); await p.waitForTimeout(1200);
  const dr = await p.evaluate(() => JSON.parse(localStorage.getItem('zqOrderDraft') || 'null'));
  check('order: the lug goes to an Order Item draft (100 pcs)', dr && dr.items.length === 1 && dr.items[0].code === '47280021' && String(dr.items[0].qty) === '100' && /หางปลา RF1.5-3/.test(dr.remark), dr && { items: dr.items, remark: dr.remark });
  await p.locator('aside nav button:has-text("หางปลา")').first().click(); await p.waitForTimeout(1000);
  await p.locator('button[role=tab]:has-text("ค้นจากหางปลา")').click(); await p.waitForTimeout(300);
  await p.locator('input[aria-label="ค้นหา"]').fill('47280021'); await p.waitForTimeout(500);
  t = await p.locator('.zq-st-det').innerText();
  check('by lug: drawing with dimensions, spec, devices that take it (relay via its socket)', await p.locator('.zq-lg-big svg').count() === 1 && /d2 \(mm\)\s*3.2/.test(t) && /อุปกรณ์ที่ใช้หางปลานี้ได้/.test(t) && /43401016/.test(t), t.slice(0, 700));
  await p.locator('button[role=tab]:has-text("ตารางรวม")').click(); await p.locator('input[aria-label="ค้นหา"]').fill(''); await p.waitForTimeout(300);
  const mx = await p.locator('.zq-lg-mx').innerText();
  check('matrix: one row per terminal group with ✓/△ counts per lug type', (await p.locator('.zq-lg-mx tbody tr').count()) === Object.keys(SEED.F).length && /✓ \d/.test(mx), mx.slice(0, 300));
  await p.screenshot({ path: OUT + 'lug_guide.png' });
  // Admin: edit a terminal group with an exception
  await p.evaluate(() => { window.__adm = true; });
  await p.locator('button', { hasText: /Admin|ปลดล็อค/ }).first().click(); await p.waitForTimeout(500);
  const u = p.locator('input[autocomplete=username]'); if (await u.count()) await u.fill('a.admin');
  const pw = p.locator('input[type=password]'); await pw.first().fill('control'); await pw.first().press('Enter'); await p.waitForTimeout(1500);
  await p.locator('aside nav button:has-text("หางปลา")').first().click(); await p.waitForTimeout(1200);
  await p.locator('button[role=tab]:has-text("ค้นจากอุปกรณ์")').click(); await p.locator('input[aria-label="ค้นหา"]').fill('43401016'); await p.waitForTimeout(500);
  await p.locator('button:has-text("แก้กลุ่มขั้วต่อ")').first().click(); await p.waitForTimeout(300);
  await p.locator('[role=dialog] textarea').fill('47280021=0 ขั้วจริงแคบ'); await p.locator('[role=dialog] button:has-text("บันทึก")').click(); await p.waitForTimeout(800);
  const sv = (await p.evaluate(() => window.__saved))[0] || [];
  check('Admin: group saved with its key and the exception', sv[1] === 'F' && sv[2].key === 'PYF14A' && sv[3] === 'PYF14A' && JSON.stringify(sv[2].ov) === JSON.stringify({ '47280021': [0, 'ขั้วจริงแคบ'] }), sv);
  t = await p.locator('.zq-lg-fam').nth(0).innerText();
  check('Admin: after saving, the relay no longer lists RF1.5-3 on PYF14A', !/RF1.5-3/.test(t), t.slice(0, 400));
  check('no page errors', errs.length === 0, errs);
  srv.close(); await b.close();
  process.exit(fails.length || errs.length ? 1 : 0);
})();
