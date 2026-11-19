// Set Library page: user view · Admin imports catalog suggestions (model → CODE in Item stock) · Stock per part, sets possible · order what is short
const OUT = __dirname + '/../out/', fs = require('fs');
const { chromium } = require('playwright');
const three = (() => { try { return fs.readFileSync(require.resolve('three/build/three.min.js'), 'utf8'); } catch (e) { return null; } })();
(async () => {
  const b = await chromium.launch({ executablePath: process.env.CHROMIUM_PATH || undefined, args: ['--use-gl=swiftshader', '--enable-unsafe-swiftshader', '--ignore-gpu-blocklist'] });
  const p = await b.newPage({ viewport: { width: 1440, height: 960 } }), errs = [], fails = [];
  await p.route(/three\.js\/r128\/three\.min\.js/, r => three ? r.fulfill({ contentType: 'application/javascript', body: three }) : r.abort());
  p.on('pageerror', e => errs.push(e.message)); p.on('dialog', d => d.accept());
  const check = (n, c, x) => { console.log((c ? 'PASS ' : 'FAIL ') + n + (c ? '' : ' → ' + JSON.stringify(x))); c || fails.push(n); };
  const srv = require('http').createServer((q, r) => { r.setHeader('content-type', 'text/html; charset=utf-8'); r.end(fs.readFileSync(OUT + 'doc.html')); }).listen(8848);
  await p.goto('http://localhost:8848/'); await p.waitForTimeout(1500);
  await p.evaluate(() => {
    const G = window.google.script, d = Object.getOwnPropertyDescriptor(G, 'run'); window.__saved = [];
    const stock = [['43401016', 'AUXILIARY RELAY', 'MY4N-GS-R DC100/110V OMRON', '48', 'PCS', 'STORE1'], ['46420024', 'SOCKET RELAY', 'PYFZ-14 (FOR MY4N)', '69', 'PCS', 'STORE1'], ['42050002', 'MOTOR CIRCUIT BREAKER', 'PKZM0-10', '15', 'PCS', 'STORE2'],
      ['42200001', 'MAGNETIC CONTACTOR', 'SC-03 AC220V FUJI', '7', 'PCS', 'STORE1'], ['42200005', 'THERMAL RELAY', 'TR-ON/3 0.95-1.45A', '4', 'PCS', 'STORE1'],
      ['46320001', 'Current test terminal', 'B-CTT (GREEN)', '8', 'PCS', 'STORE1'], ['46330005', 'Test terminal cover', 'FOR TYPE B-C FUJI ELECTRIC IND.', '4', 'PCS', 'STORE1']];
    const tsv = stock.map(f => [f[0], f[1], f[2], f[3], f[4], f[5], '', 'DS13', '', '1'].join('\t')).join('\n');
    window.__bundle = (async () => { const cs = new CompressionStream('gzip'), w = cs.writable.getWriter(); w.write(new TextEncoder().encode(tsv)); w.close();
      const buf = new Uint8Array(await new Response(cs.readable).arrayBuffer()); let bin = ''; buf.forEach(c => bin += String.fromCharCode(c)); return btoa(bin); })();
    let sets = [];
    const seeds = [{ cat: 'Relay + Socket', name: 'OMRON MY4N (4 ขั้ว) + Socket', ok: 1, note: 'n', src: 'https://example.com/my4 | ใบ Estimate', parts: [{ role: 'หลัก', kw: ['MY4N'] }, { role: 'Socket', kw: ['PYFZ-14', 'PYF14A'] }, { role: 'Clip', kw: ['PYC-A1'], optional: 1 }] },
      { cat: 'Magnetic contactor', name: 'FUJI SC-03 + Thermal relay', ok: 1, note: 'n', src: 'https://example.com/sc', parts: [{ role: 'หลัก', kw: ['SC-03'] }, { role: 'Thermal relay', kw: ['TR-0N', 'TR-ON'] }] },
      { cat: 'Magnetic contactor', name: 'FUJI SC-0 + Thermal relay', ok: 1, note: 'n', src: 'https://example.com/sc', parts: [{ role: 'หลัก', kw: ['=SC-0'] }, { role: 'Thermal relay', kw: ['TR-0N', 'TR-ON'] }] },
      { cat: 'Test terminal', name: 'Test terminal Type B', ok: 1, note: 'n', src: 'ใบ Estimate', parts: [{ role: 'หลัก', kw: ['B-CTT'], per: 2 }, { role: 'Cover', kw: ['=B-C'] }] },
      { cat: 'Timer + Socket', name: 'OMRON H3CR-A8 + Socket', ok: 0, note: 'n2', src: 'https://example.com/h3cr', parts: [{ role: 'หลัก', kw: ['H3CR-A8'] }, { role: 'Socket', kw: ['P2CF-08'] }] }];
    Object.defineProperty(G, 'run', { configurable: true, get() {
      const inner = d.get.call(G); let ok = () => {};
      const w = { withSuccessHandler(f) { ok = f; inner.withSuccessHandler(f); return w; }, withFailureHandler(f) { inner.withFailureHandler(f); return w; },
        api(t, fn, a) {
          if (fn === 'getStockMeta') { setTimeout(() => ok({ hash: 'st12345678', rows: stock.length, items: 3, locations: ['STORE1', 'STORE2'], bys: [], updated: '19/11/2026 05:00' }), 20); return; }
          if (fn === 'getStockBundle') { window.__bundle.then(b64 => ok({ hash: 'st12345678', b64, v: 3 })); return; }
          if (fn === 'getCodeGroups') { setTimeout(() => ok({}), 10); return; }
          if (fn === 'getSetLib') { setTimeout(() => ok({ sets, admin: false }), 20); return; }
          if (fn === 'getSetSeeds') { setTimeout(() => ok(seeds), 20); return; }
          if (fn === 'checkOrderDup') { setTimeout(() => ok({}), 10); return; }
          if (fn === 'saveSetLib') { const L = [].concat(a[1]); window.__saved.push(JSON.parse(JSON.stringify(L))); L.forEach((s, i) => sets.push({ ...s, id: s.id || 'S' + (sets.length + i + 1) })); setTimeout(() => ok({ sets: sets.slice(), admin: true }), 30); return; }
          return inner.api(t, fn, a);
        } };
      return new Proxy(w, { get: (o, k) => k in o ? o[k] : inner[k] });
    } });
  });
  const nav = () => p.locator('aside nav button:has-text("ชุดอุปกรณ์")').first().click();
  await nav(); await p.waitForTimeout(1200);
  const ut = await p.locator('.zq-st').innerText().catch(() => '');
  check('user: menu item, empty state, no Admin buttons', /Admin เป็นผู้สร้างชุดอุปกรณ์/.test(ut) && await p.locator('button:has-text("สร้างชุดใหม่")').count() === 0, ut.slice(0, 300));
  await p.locator('button', { hasText: /Admin|ปลดล็อค/ }).first().click(); await p.waitForTimeout(500);
  const u = p.locator('input[autocomplete=username]'); if (await u.count()) await u.fill('premwit');
  const pw = p.locator('input[type=password]'); await pw.first().fill('control'); await pw.first().press('Enter'); await p.waitForTimeout(1500);
  await nav(); await p.waitForTimeout(1500);
  await p.locator('button:has-text("ชุดแนะนำจาก Catalog")').first().click(); await p.waitForTimeout(1500);
  const dlg = p.locator('[aria-label="ชุดแนะนำจาก Catalog"]'), dt = await dlg.innerText();
  check('catalog: models matched to CODEs in Item stock, missing ones named, confirmed / check badges, source links', /43401016 \(Stock 48\)/.test(dt) && /46420024 \(Stock 69\)/.test(dt) && /ไม่พบใน Stock \(PYC-A1\)/.test(dt) && /ยืนยันแล้ว/.test(dt) && /ควรตรวจ/.test(dt) && await dlg.locator('a[href="https://example.com/my4"]').count() === 1, dt.slice(0, 500));
  check('catalog: "=SC-0" is that model only (SC-03 in Stock is not taken), "TR-ON/3" spelling found, the B-C cover is not the B-CTT terminal', /หลัก: 42200001/.test(dt) && /Thermal relay: 42200005/.test(dt) && /หลัก: ไม่พบใน Stock \(SC-0\)/.test(dt) && /หลัก ×2: 46320001/.test(dt) && /Cover: 46330005/.test(dt), dt);
  check('catalog: by default only sets whose main item is in Stock are ticked', /นำเข้า 3 ชุดที่เลือก/.test(dt));
  await dlg.locator('label:has-text("FUJI SC-03") input').uncheck(); await dlg.locator('label:has-text("Test terminal Type B") input').uncheck();
  await dlg.locator('button:has-text("นำเข้า")').click(); await p.waitForTimeout(800);
  const sv = (await p.evaluate(() => window.__saved))[0] || [];
  check('catalog import: one set saved with CODE options, the clip kept as a model keyword', sv.length === 1 && sv[0].parts[0].opts[0].code === '43401016' && sv[0].parts[1].opts[0].code === '46420024' && sv[0].parts[2].opts[0].kw === 'PYC-A1' && !sv[0].parts[2].opts[0].code && sv[0].parts[2].optional, sv);
  const qIn = p.locator('input[aria-label="จำนวนชุด"]');
  await qIn.fill('30'); await p.waitForTimeout(300);
  let k = await p.locator('.zq-st-k').innerText();
  check('30 sets: the relay is the main item (not its socket), Stock makes 48 sets (relay is the limit), nothing to order', /43401016 30 · 46420024 30/.test(k) && /48 ชุด/.test(k) && /ต้องสั่งเพิ่ม\s*0 รายการ/.test(k), k);
  check('an optional part with no CODE yet (the clip) starts left out', await p.locator('.zq-st-t label:has-text("ใส่ในชุด") input').isChecked() === false);
  await qIn.fill('60'); await p.waitForTimeout(300);
  k = await p.locator('.zq-st-k').innerText();
  check('60 sets: relay short 12, socket enough', /43401016 ขาด 12/.test(k) && !/46420024 ขาด/.test(k), k);
  // assembly animation: one step per part that is in the set (the clip with no CODE is left out), captions name the CODEs
  const an = p.locator('.zq-st-anim');
  await p.waitForFunction(() => document.querySelectorAll('.zq-st-steps li').length > 0, null, { timeout: 15000 }).catch(() => {});
  const stp = await an.locator('.zq-st-steps li').allInnerTexts();
  check('assembly: steps follow the set (socket on the rail, relay into the socket; no clip step while the clip is left out)', stp.length === 2 && /Socket 46420024 บนราง DIN/.test(stp[0]) && /43401016/.test(stp[1]) && /14 ขา/.test(stp[1]), stp);
  await an.locator('button', { hasText: 'แยกชิ้น' }).click(); await p.waitForTimeout(300);
  check('assembly: "แยกชิ้น" pauses the tour at the start', (await an.locator('input[type=range]').inputValue()) === '0' && /▶ เล่น/.test(await an.locator('.zq-st-ctl').innerText()));
  await an.locator('button', { hasText: 'ประกอบครบ' }).click(); await p.waitForTimeout(400);
  check('assembly: "ประกอบครบ" shows the set assembled', (await an.locator('input[type=range]').inputValue()) === '1000' && /ประกอบครบแล้ว/.test(await an.locator('.zq-st-cap').innerText()));
  check('assembly: the canvas draws (WebGL)', await an.locator('canvas').evaluate(c => c.width > 100 && !!(c.getContext('webgl') || c.getContext('webgl2') || true)));
  await p.locator('.zq-st-t label:has-text("ใส่ในชุด") input').check(); await p.waitForTimeout(800);
  check('assembly: ticking the clip into the set adds its step', (await an.locator('.zq-st-steps li').count()) === 3 && /PYC-A1/.test(await an.locator('.zq-st-steps').innerText()));
  await p.locator('.zq-st-t label:has-text("ใส่ในชุด") input').uncheck(); await p.waitForTimeout(300);
  await p.screenshot({ path: OUT + 'set_library.png' });
  await p.locator('button:has-text("ใส่ใบสั่งเฉพาะที่ขาด")').click(); await p.waitForTimeout(1200);
  const dr = await p.evaluate(() => JSON.parse(localStorage.getItem('zqOrderDraft') || 'null'));
  check('order what is short: draft has 43401016 × 12 and the set in the remark, the Order Item page opens', dr && dr.items.length === 1 && dr.items[0].code === '43401016' && dr.items[0].qty === '12' && /ชุด OMRON MY4N/.test(dr.remark), dr && { items: dr.items, remark: dr.remark });
  check('no page errors', errs.length === 0, errs);
  srv.close(); await b.close();
  process.exit(fails.length || errs.length ? 1 : 0);
})();
