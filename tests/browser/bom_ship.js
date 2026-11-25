// Booking BOM × Shipment: "เบิกแล้ว ยังไม่ปิด" card filters, list pill, next step, the "เทียบ Shipment" tab, menu badge
const OUT = __dirname + '/../out/', fs = require('fs');
const { chromium } = require('playwright');
(async () => {
  const b = await chromium.launch({ executablePath: process.env.CHROMIUM_PATH || undefined });
  const p = await b.newPage({ viewport: { width: 1440, height: 960 } }), errs = [], fails = [];
  p.on('pageerror', e => errs.push(e.message)); p.on('dialog', d => d.accept());
  const check = (n, c, x) => { console.log((c ? 'PASS ' : 'FAIL ') + n + (c ? '' : ' → ' + JSON.stringify(x))); c || fails.push(n); };
  const srv = require('http').createServer((q, r) => { r.setHeader('content-type', 'text/html; charset=utf-8'); r.end(fs.readFileSync(OUT + 'doc.html')); }).listen(8850);
  await p.goto('http://localhost:8850/'); await p.waitForTimeout(1500);
  await p.evaluate(() => {
    const G = window.google.script, d = Object.getOwnPropertyDescriptor(G, 'run');
    const mod = v => { const P = (v.projects || []).filter(x => x.status === 'active'); window.__shipP = P[0] && P[0].project; window.__nAct = P.length;
      if (P[0]) { const its = P[0].items, by = {}; its.forEach((x, i) => { by[String(x.code).toUpperCase()] = i === 0 ? x.qty / 2 : i === its.length - 1 ? 0 : x.qty; });
        const full = its.length - 2; P[0].ship = { n: its.length, full, part: 1, pct: Math.round(full / its.length * 1000) / 1000, last: '21/11/2026', alert: '22/11/2026', closeOn: '06/12/2026', by, cat: '1099P1', shared: true, tagged: 1, lines: 9 };
        window.__ship = P[0].ship; window.__its = its.map(x => x.code); }
      return v; };
    Object.defineProperty(G, 'run', { configurable: true, get() {
      const inner = d.get.call(G); let ok = () => {};
      const w = { withSuccessHandler(f) { ok = f; inner.withSuccessHandler(f); return w; }, withFailureHandler(f) { inner.withFailureHandler(f); return w; },
        api(t, fn, a) {
          if (fn === 'getBom') { inner.withSuccessHandler(v => ok(mod(v))); return inner.api(t, fn, a); }
          if (fn === 'getBomSummary') { setTimeout(() => ok({ short: 2, late: 0, projects: 3, ship: 1 }), 20); return; }
          return inner.api(t, fn, a);
        } };
      return new Proxy(w, { get: (o, k) => k in o ? o[k] : inner[k] });
    } });
  });
  await p.locator('aside nav button:has-text("Booking BOM")').first().click(); await p.waitForTimeout(2500);
  const P = await p.evaluate(() => window.__shipP), S = await p.evaluate(() => window.__ship), nAct = await p.evaluate(() => window.__nAct);
  check('there is an open project to test with', !!P && nAct > 1, { P, nAct });
  await p.evaluate(() => { window.__zqBomS = { short: 2, late: 0, projects: 3, ship: 1 }; window.dispatchEvent(new Event('zq:boms')); }); await p.waitForTimeout(300);
  const nav = await p.locator('aside nav button:has-text("Booking BOM")').first().innerText();
  check('menu badge: 🚚1 next to the short count', /🚚1/.test(nav), nav);
  const card = p.locator('.zq-bx-kc', { hasText: 'เบิกแล้ว ยังไม่ปิด' });
  check('KPI card "เบิกแล้ว ยังไม่ปิด" counts 1', /^\s*1\s/.test(await card.innerText()), await card.innerText());
  await card.click(); await p.waitForTimeout(400);
  const items = await p.locator('.zq-bx-li').allInnerTexts();
  check('the card filters the list to that project, with its 🚚 pill', items.length === 1 && items[0].includes(P) && /🚚 เบิก \d+%/.test(items[0]), items);
  await p.locator('.zq-bx-li').first().click(); await p.waitForTimeout(400);
  const nx = await p.locator('.zq-bx-nx').innerText();
  check('next step: MF issued x/n CODE, last date, auto-close date', new RegExp('MF เบิกไปแล้ว ' + S.full + '/' + S.n + ' CODE').test(nx) && /21\/11\/2026/.test(nx) && /06\/12\/2026/.test(nx), nx);
  await p.locator('.zq-bx-nx button:has-text("ดูเทียบ Shipment")').click(); await p.waitForTimeout(400);
  const tb = await p.locator('.zq-bm-det').innerText();
  check('tab "เทียบ Shipment": CONTROL WIRING, units note, full / part / not issued per CODE', /เทียบ Shipment \(\d+%\)/.test(tb) && /CONTROL WIRING/.test(tb) && /หลาย Unit/.test(tb) && /✅ เบิกครบ/.test(tb) && /◐ ขาดอีก/.test(tb) && /— ยังไม่เบิก/.test(tb), tb.slice(0, 900));
  await p.locator('.zq-bm-det button:has-text("ยังไม่เบิก")').first().click(); await p.waitForTimeout(300);
  check('chip "ยังไม่เบิก" leaves the one CODE not issued', (await p.locator('.zq-bm-det tbody tr').count()) === 1);
  await p.screenshot({ path: OUT + 'bom_ship.png' });
  check('no page errors', errs.length === 0, errs);
  srv.close(); await b.close(); process.exit(fails.length || errs.length ? 1 : 0);
})();
