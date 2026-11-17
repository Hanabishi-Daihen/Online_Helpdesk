// Booking BOM: book from Item stock with a cart · edit a project's CODE list (Admin) · the item panel's "📋 จอง BOM"
// prints what it sees; exits 1 when the page throws an error or a check fails
const OUT = __dirname + '/../out/', fs = require('fs');
const { chromium } = require('playwright');
(async () => {
  const b = await chromium.launch({ executablePath: process.env.CHROMIUM_PATH || undefined });
  const p = await b.newPage({ viewport: { width: 1440, height: 960 } }), errs = [], fails = [];
  p.on('pageerror', e => errs.push(e.message)); p.on('dialog', d => d.accept());
  const check = (n, c, x) => { console.log((c ? 'PASS ' : 'FAIL ') + n + (c ? '' : ' → ' + JSON.stringify(x))); c || fails.push(n); };
  const srv = require('http').createServer((q, r) => { r.setHeader('content-type', 'text/html; charset=utf-8'); r.end(fs.readFileSync(OUT + 'doc.html')); }).listen(8846);
  await p.goto('http://localhost:8846/'); await p.waitForTimeout(1500);
  // fake server for the new calls; getBom passes through (and is kept to answer with an updated view)
  await p.evaluate(() => {
    const G = window.google.script, d = Object.getOwnPropertyDescriptor(G, 'run'); window.__bk = [];
    const stock = [['42050002', 'MOTOR CIRCUIT BREAKER', 'PKZM0-10', '12', 'PCS', 'STORE1'], ['42050002', 'MOTOR CIRCUIT BREAKER', 'PKZM0-10', '3', 'PCS', 'STORE2'], ['43401027', 'AUXILIARY RELAY', 'RH4B-U', '10', 'PCS', 'STORE1'], ['42310010', 'CAM SWITCH', 'CA10 A222 EG', '4', 'PCS', 'STORE1']];
    Object.defineProperty(G, 'run', { configurable: true, get() {
      const inner = d.get.call(G); let ok = () => {};
      const w = { withSuccessHandler(f) { ok = f; inner.withSuccessHandler(v => { if (window.__lastFn === 'getBom') window.__bom = v; f(v); }); return w; }, withFailureHandler(f) { inner.withFailureHandler(f); return w; },
        api(t, fn, a) {
          window.__lastFn = fn;
          if (fn === 'searchStock') { const q = String(a[1] || '').toLowerCase(); setTimeout(() => ok({ total: 2, page: 1, pages: 1, rows: stock.filter(f => (f[0] + ' ' + f[1]).toLowerCase().includes(q)), all: true }), 20); return; }
          if (fn === 'bomBook' || fn === 'bomSaveItems') {
            window.__bk.push([fn, JSON.parse(JSON.stringify(a))]);
            const v = JSON.parse(JSON.stringify(window.__bom || { projects: [], codes: [] }));
            if (fn === 'bomBook') { const r = a[1]; v.projects.unshift({ project: r.project, title: r.title, rev: '', status: 'active', needDate: r.needDate, daysLeft: 14, count: { ok: r.items.length, incoming: 0, late: 0, short: 0 }, total: r.items.length, items: r.items.map((x, i) => ({ item: '', code: x.code, qty: x.qty, material: 'MOTOR CIRCUIT BREAKER', st: 'ok', by: 'DS13', fromStock: x.qty })) });
              setTimeout(() => ok({ project: r.project, added: r.items.length, created: true, reopened: false, view: v }), 30); }
            else setTimeout(() => ok({ changes: ['42050002 1→5', '43401027 ⇄ X'], view: v }), 30);
            return;
          }
          return inner.api(t, fn, a);
        } };
      return new Proxy(w, { get: (o, k) => k in o ? o[k] : inner[k] });
    } });
  });
  const nav = () => p.locator('aside nav button:has-text("Booking BOM")').first().click();
  await nav(); await p.waitForTimeout(1500);
  check('user: no cart tab, no edit button', await p.locator('button:has-text("จองจาก Stock")').count() === 0 && await p.locator('button:has-text("แก้ไขรายการ CODE")').count() === 0);
  // Admin
  await p.locator('button', { hasText: /Admin|ปลดล็อค/ }).first().click(); await p.waitForTimeout(500);
  const u = p.locator('input[autocomplete=username]'); if (await u.count()) await u.fill('premwit');
  const pw = p.locator('input[type=password]'); await pw.first().fill('control'); await pw.first().press('Enter'); await p.waitForTimeout(1500);
  await nav(); await p.waitForTimeout(1200);
  // ── edit a project ──
  await p.locator('button:has-text("แก้ไขรายการ CODE")').first().click(); await p.waitForTimeout(500);
  const ed = p.locator('table.zq-bk-e'); check('edit: the editor replaces the table', await ed.count() === 1 && await ed.locator('tbody tr').count() >= 3, await ed.count());
  const r0 = ed.locator('tbody tr').nth(0), code0 = (await r0.locator('code').innerText()).trim();
  await r0.locator('input').first().fill('5');
  const r1 = ed.locator('tbody tr').nth(1), code1 = (await r1.locator('code').innerText()).trim();
  await r1.locator('button:has-text("ไม่ใช้")').click();
  const r2 = ed.locator('tbody tr').nth(2), code2 = (await r2.locator('code').innerText()).trim();
  await r2.locator('button:has-text("เปลี่ยน")').click(); await p.waitForTimeout(200);
  await p.locator('.zq-bk-s input[type=search]').fill('4205'); await p.waitForTimeout(700);
  await p.locator('.zq-bk-r button').first().click(); await p.waitForTimeout(200);
  check('edit: swapping to a CODE still in the list is refused with a clear message', /มีอยู่ในรายการแล้ว/.test(await p.locator('.zq-em-msg.err').innerText().catch(() => '')));
  await p.locator('.zq-bk-s input[type=search]').fill('cam'); await p.waitForTimeout(700);
  await p.locator('.zq-bk-r button:has-text("ใส่ตะกร้า")').first().click(); await p.waitForTimeout(200);
  check('edit: changed / not used / swapped rows are marked', await ed.locator('tr.ch').count() >= 2 && await ed.locator('tr.dl').count() === 1 && /แทน/.test(await ed.innerText()));
  await p.screenshot({ path: OUT + 'bom_edit.png' });
  await p.fill('input[aria-label="เหตุผลที่แก้"]', 'ลูกค้าเปลี่ยนรุ่น');
  await p.locator('button:has-text("บันทึกการแก้ไข")').click(); await p.waitForTimeout(600);
  const sv = (await p.evaluate(() => window.__bk)).find(x => x[0] === 'bomSaveItems'), A = sv && sv[1];
  const its = A ? A[2] : [];
  check('edit: saves the whole new list with the starting list\'s signature, the swap and the reason',
    !!A && its.find(x => x.code === code0).qty === 5 && !its.find(x => x.code === code1) && !its.find(x => x.code === code2) && !!its.find(x => x.code === '42310010') && A[4]['42310010'] === code2 && /=/.test(A[3]) && A[5] === 'ลูกค้าเปลี่ยนรุ่น', A);
  check('edit: back to the table with the result', await p.locator('table.zq-bk-e').count() === 0 && /บันทึกแล้ว/.test(await p.locator('.zq-em-msg').first().innerText().catch(() => '')));
  // ── the cart ──
  await p.locator('button:has-text("จองจาก Stock")').first().click(); await p.waitForTimeout(500);
  await p.locator('.zq-bk-s input[type=search]').fill('4205'); await p.waitForTimeout(700);
  const row = p.locator('.zq-bk-r').first();
  check('cart: one line per CODE, all locations added up, booked and free shown', /42050002/.test(await row.innerText()) && /15/.test(await row.locator('.zq-bk-n').innerText()) && /ว่าง/.test(await row.innerText()), await row.innerText());
  await row.locator('button').click(); await row.locator('button').click(); await p.waitForTimeout(200);
  check('cart: pressing again adds one', /✓ 2/.test(await row.locator('button').innerText()) && await p.locator('.zq-bk-l').count() === 1);
  await p.locator('.zq-bk-seg button:has-text("โปรเจกต์ใหม่")').click();
  await p.fill('input[aria-label="เลขโปรเจกต์ใหม่"]', 'pp1099a01'); await p.fill('input[aria-label="วันที่ต้องใช้"]', '01/12/2026'); await p.fill('input[aria-label="ชื่อโปรเจกต์"]', 'Local Control Cabinet (LCC.)');
  await p.screenshot({ path: OUT + 'bom_cart.png' });
  const go = p.locator('.zq-bk-c button.zq-or-btn');
  check('cart: the book button names the project and the count', /PP1099A01 \(1 รายการ\)/.test(await go.innerText()), await go.innerText());
  await go.click(); await p.waitForTimeout(800);
  const bk = (await p.evaluate(() => window.__bk)).find(x => x[0] === 'bomBook'), R = bk && bk[1][1];
  check('cart: books the project (upper-case), date, title, CODE × 2', !!R && R.project === 'PP1099A01' && R.mode === 'new' && R.needDate === '01/12/2026' && R.items.length === 1 && R.items[0].code === '42050002' && R.items[0].qty === 2, R);
  check('cart: emptied, and the new project opens', await p.evaluate(() => JSON.parse(localStorage.getItem('zqBomCart')).items.length) === 0 && /PP1099A01/.test(await p.locator('.zq-bm-det').innerText().catch(() => '')));
  // ── from the item panel ──
  await p.evaluate(() => window.dispatchEvent(new CustomEvent('zq:item', { detail: '46700501' }))); await p.waitForTimeout(1500);
  const ib = p.locator('button:has-text("จอง BOM")').first();
  if (await ib.count()) { await ib.click(); await p.waitForTimeout(200); }
  check('item panel: "📋 จอง BOM" puts the CODE in the cart', await p.evaluate(() => (JSON.parse(localStorage.getItem('zqBomCart') || '{"items":[]}').items || []).some(x => x.code === '46700501')));
  check('the cart tab counts it', /จองจาก Stock \(1\)/.test(await p.locator('button:has-text("จองจาก Stock")').first().innerText().catch(() => '')));
  console.log('errors:', errs); srv.close(); await b.close(); process.exit(errs.length || fails.length ? 1 : 0);
})();
