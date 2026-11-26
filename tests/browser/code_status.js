// ทะเบียน CODE เลิกใช้ in the pages: Order Item (reason box, swap, save refused for a user), Item stock, Stock Alert (reason, editor, filter), Booking BOM
const OUT = __dirname + '/../out/', fs = require('fs');
const { chromium } = require('playwright');
(async () => {
  const b = await chromium.launch({ executablePath: process.env.CHROMIUM_PATH || undefined });
  const p = await b.newPage({ viewport: { width: 1440, height: 960 } }), errs = [], fails = [];
  p.on('pageerror', e => errs.push(e.message));
  const check = (n, c, x) => { console.log((c ? 'PASS ' : 'FAIL ') + n + (c ? '' : ' → ' + JSON.stringify(x))); c || fails.push(n); };
  const srv = require('http').createServer((q, r) => { r.setHeader('content-type', 'text/html; charset=utf-8'); r.end(fs.readFileSync(OUT + 'doc.html')); }).listen(8855);
  await p.goto('http://localhost:8855/'); await p.waitForTimeout(1500);
  await p.evaluate(() => {
    const G = window.google.script, d = Object.getOwnPropertyDescriptor(G, 'run'); window.__cs = []; window.__saved = [];
    const stock = [['46001008', 'PILOT LAMP', 'APW-116DD-PW DC110V', '5', 'PC', 'STORE1', '', 'DS13', '', '0'], ['46001009', 'PILOT LAMP', 'APW-199DS DC110V', '120', 'PC', 'STORE1', '', 'DS13', '', '1'], ['42050002', 'MOTOR CIRCUIT BREAKER', 'PKZM0-10', '12', 'PCS', 'STORE1', '', 'DS13', '', '1']];
    const tsv = stock.map(f => [f[0], f[1], f[2], f[3], f[4], f[5], '', f[7], '', f[9]].join('\t')).join('\n');
    window.__bundle = (async () => { const cs = new CompressionStream('gzip'), w = cs.writable.getWriter(); w.write(new TextEncoder().encode(tsv)); w.close();
      const buf = new Uint8Array(await new Response(cs.readable).arrayBuffer()); let bin = ''; buf.forEach(c => bin += String.fromCharCode(c)); return btoa(bin); })();
    const R = [['eol', 'ผู้ผลิตเลิกผลิต'], ['model', 'เปลี่ยนรุ่น / มีรุ่นแทน'], ['quality', 'ปัญหาคุณภาพ / หน้างาน'], ['spec', 'ลูกค้า / สเปกไม่รับ'], ['price', 'ราคา / ผู้ขาย'], ['dup', 'CODE ซ้ำ'], ['other', 'อื่น ๆ'], ['official', 'ปิดในระบบ Official — ยังไม่ระบุสาเหตุ'], ['sa', 'ตั้งเลิกใช้ใน Stock Alert — ยังไม่ระบุสาเหตุ']];
    let map = { '46001008': { code: '46001008', level: 'off', reason: 'eol', detail: 'ผู้ผลิตเลิกผลิตรุ่น DD', replace: '46001009', since: '01/11/2026', by: 'A.Admin' } };
    Object.defineProperty(G, 'run', { configurable: true, get() {
      const inner = d.get.call(G); let ok = () => {};
      const w = { withSuccessHandler(f) { ok = f; inner.withSuccessHandler(f); return w; }, withFailureHandler(f) { inner.withFailureHandler(f); return w; },
        api(t, fn, a) {
          if (fn === 'getStockMeta') { setTimeout(() => ok({ hash: 'cs12345678', rows: stock.length, items: 3, locations: ['STORE1'], bys: [], updated: '26/11/2026 05:00' }), 20); return; }
          if (fn === 'getStockBundle') { window.__bundle.then(b64 => ok({ hash: 'cs12345678', b64, v: 3 })); return; }
          if (fn === 'getCodeStatus') { setTimeout(() => ok({ map, reasons: R, admin: !!window.__adm }), 20); return; }
          if (fn === 'saveCodeStatus') { window.__cs.push(a[1]); map = { ...map, [a[1].code]: { ...a[1], by: 'A.Admin' } }; setTimeout(() => ok({ map, reasons: R, admin: true }), 30); return; }
          return inner.api(t, fn, a);
        } };
      return new Proxy(w, { get: (o, k) => k in o ? o[k] : inner[k] });
    } });
  });
  // Order Item (user)
  await p.locator('aside nav button', { hasText: 'Order Item' }).first().click(); await p.waitForTimeout(1500);
  const row = p.locator('.zq-or-items tbody tr.it').first();
  await row.locator('input,textarea').first().fill('46001008'); await p.waitForTimeout(2500);
  const box = p.locator('.zq-or-items .zq-cs-box').first();
  const bt = await box.innerText().catch(() => '');
  check('Order Item: typing a retired CODE shows why, since when, by whom, and what to use instead', /CODE นี้เลิกใช้แล้ว/.test(bt) && /ผู้ผลิตเลิกผลิต/.test(bt) && /ผู้ผลิตเลิกผลิตรุ่น DD/.test(bt) && /01\/11\/2026/.test(bt) && /46001009/.test(bt) && /Stock 120/.test(bt), bt);
  await p.screenshot({ path: OUT + 'code_status_order.png' });
  await box.locator('button:has-text("เปลี่ยนเป็น 46001009")').click(); await p.waitForTimeout(600);
  check('Order Item: "เปลี่ยนเป็น 46001009" swaps the CODE on the line', (await row.locator('input,textarea').first().inputValue()) === '46001009' && await p.locator('.zq-or-items .zq-cs-box').count() === 0);
  // Item stock list + detail
  await p.locator('aside nav button', { hasText: 'Item stock' }).first().click(); await p.waitForTimeout(1500);
  const sq = p.locator('input[placeholder^="พิมพ์ CODE หรือรายละเอียดสินค้า"]').first(); await sq.fill('46001008'); await p.waitForTimeout(1500);
  const st = await p.locator('main').innerText();
  check('Item stock: the line carries ⛔ เลิกใช้ with the reason', /⛔ เลิกใช้/.test(st) && /46001009/.test(st), st.slice(0, 600));
  // Admin edits the reason from the Item stock detail
  await p.evaluate(() => { window.__adm = true; });
  await p.locator('button', { hasText: /Admin|ปลดล็อค/ }).first().click(); await p.waitForTimeout(400);
  const u = p.locator('input[autocomplete=username]'); if (await u.count()) await u.fill('a.admin');
  const pw = p.locator('input[type=password]'); await pw.first().fill('control'); await pw.first().press('Enter'); await p.waitForTimeout(1500);
  await p.evaluate(() => window.zqCsLoad && window.zqCsLoad(true)).catch(() => {});
  await p.locator('aside nav button', { hasText: 'Item stock' }).first().click(); await p.waitForTimeout(1200);
  await p.locator('input[placeholder^="พิมพ์ CODE หรือรายละเอียดสินค้า"]').first().fill('46001008'); await p.waitForTimeout(1500);
  await p.locator('.zq-sk-wrap .zq-sk-r').first().click(); await p.waitForTimeout(800);
  const det = p.locator('.zq-sk-det .zq-cs-box').first();
  check('Item stock detail: the reason box, with an edit button for Admin', /ผู้ผลิตเลิกผลิต/.test(await det.innerText().catch(() => '')) && await det.locator('button:has-text("แก้เหตุผล")').count() === 1, await det.innerText().catch(e => e.message));
  await det.locator('button:has-text("แก้เหตุผล")').click(); await p.waitForTimeout(400);
  const dlg = p.locator('[role=dialog][aria-label="เหตุผล CODE เลิกใช้"]');
  await dlg.locator('select').selectOption('quality'); await dlg.locator('textarea').fill('ไฟดับเร็วในงาน PP1030');
  await dlg.locator('button:has-text("บันทึก")').click(); await p.waitForTimeout(800);
  const sv = await p.evaluate(() => window.__cs[0]);
  check('editor: saves level, reason, detail and the replacement', sv && sv.code === '46001008' && sv.level === 'off' && sv.reason === 'quality' && /PP1030/.test(sv.detail) && sv.replace === '46001009', sv);
  check('the box shows the new reason at once', /ปัญหาคุณภาพ/.test(await p.locator('.zq-sk-det .zq-cs-box').first().innerText()));
  await p.screenshot({ path: OUT + 'code_status.png' });
  check('no page errors', errs.length === 0, errs);
  srv.close(); await b.close(); process.exit(fails.length || errs.length ? 1 : 0);
})();
