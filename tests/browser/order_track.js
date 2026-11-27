// ติดตามใบสั่ง: the list must follow the menu badge — an old copy of the list (7) next to a newer badge (8) reloads by itself
const OUT = __dirname + '/../out/', fs = require('fs');
const { chromium } = require('playwright');
(async () => {
  const b = await chromium.launch({ executablePath: process.env.CHROMIUM_PATH || undefined });
  const p = await b.newPage({ viewport: { width: 1440, height: 900 } }), errs = [], fails = [];
  p.on('pageerror', e => errs.push(e.message)); p.on('dialog', d => d.accept());
  const check = (n, c, x) => { console.log((c ? 'PASS ' : 'FAIL ') + n + (c ? '' : ' → ' + JSON.stringify(x))); c || fails.push(n); };
  const srv = require('http').createServer((q, r) => { r.setHeader('content-type', 'text/html; charset=utf-8'); r.end(fs.readFileSync(OUT + 'doc.html')); }).listen(8851);
  await p.goto('http://localhost:8851/'); await p.waitForTimeout(1500);
  await p.evaluate(() => {
    const G = window.google.script, d = Object.getOwnPropertyDescriptor(G, 'run');
    const ord = i => ({ id: 'ORD26-00' + (10 + i), otype: 'stock', spec: 'DS1C26-' + (1400 + i), rqn: 'RQN26-0' + (4600 + i), des: 'S.Somchai', date: '2' + (i % 8) + '/11/2026', remark: '', state: 'exported', stateTh: 'ส่งออกแล้ว',
      stage: 'po', late: 0, noPr: 0, waitDays: null, fill: null, auto: null, refMissing: false,
      steps: [{ k: 'saved', label: 'บันทึก', at: '', done: true }, { k: 'sent', label: 'ส่ง', at: '', done: true }, { k: 'approved', label: 'อนุมัติ', at: '', done: true }, { k: 'pr', label: 'PR', at: '', done: true }, { k: 'po', label: 'PO', at: '', done: true }, { k: 'recv', label: 'รับ', at: '', done: false }],
      items: [{ code: '4204000' + i, material: 'MCB', qty: 1, unit: 'PC', st: 'po', pr: 'PR' + i, po: 'PO' + i, status: '', due: '', late: 0 }] });
    window.__trkCalls = 0; window.__trkN = 7; window.__zqTrkS = { open: 7, issue: 0, st: { po: 7 } };
    const list = n => ({ orders: Array.from({ length: n }, (_, i) => ord(i)), time: '27/11/2026 10:0' + n, po: '', ageCfg: { warn: 2, over: 3 }, admin: false });
    Object.defineProperty(G, 'run', { configurable: true, get() {
      const inner = d.get.call(G); let ok = () => {};
      const w = { withSuccessHandler(f) { ok = f; inner.withSuccessHandler(f); return w; }, withFailureHandler(f) { inner.withFailureHandler(f); return w; },
        api(t, fn, a) {
          // the first answer is the old copy (7, cached before the 8th order was committed); the server now has 8
          if (fn === 'getOrderTrack') { window.__trkCalls++; const n = window.__trkN; setTimeout(() => ok(list(n)), 30); return; }
          if (fn === 'getOrderTrackSummary') { const n = window.__trkN; setTimeout(() => ok({ open: n, issue: 0, st: { po: n } }), 30); return; }
          return inner.api(t, fn, a);
        } };
      return new Proxy(w, { get: (o, k) => k in o ? o[k] : inner[k] });
    } });
  });
  await p.locator('aside nav button:has-text("ติดตามใบสั่ง")').first().click(); await p.waitForTimeout(800);
  check('list shows the server copy (7)', /ยังไม่ครบ\s*7/.test(await p.locator('main').innerText()));
  const c0 = await p.evaluate(() => { window.__trkN = 8; return window.__trkCalls; });   // the 8th order is committed on the server
  // the menu badge refreshes (pulse "ord") with the server's newer count
  await p.evaluate(() => { window.__zqTrkS = { open: 8, issue: 0, st: { po: 8 } }; window.dispatchEvent(new Event('zq:trks')); });
  await p.waitForFunction(() => /ยังไม่ครบ\s*8/.test(document.querySelector('main') && document.querySelector('main').innerText || ''), null, { timeout: 8000 }).catch(() => {});
  const t = await p.locator('main').innerText();
  check('old list (7) next to the badge (8) → reloaded by itself, list shows 8', /ยังไม่ครบ\s*8/.test(t) && /ORD26-0017/.test(t), t.slice(0, 600));
  const calls = await p.evaluate(() => window.__trkCalls);
  check('reloaded once, no loop', calls === c0 + 1, [c0, calls]);
  await p.waitForTimeout(1500);
  check('still once after the badge event settles', await p.evaluate(() => window.__trkCalls) === c0 + 1, await p.evaluate(() => window.__trkCalls));
  console.log('errors:', errs); srv.close(); await b.close(); process.exit(errs.length || fails.length ? 1 : 0);
})();
