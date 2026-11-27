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
    const stock = [['42050002', 'MOTOR CIRCUIT BREAKER', 'PKZM0-10', '12', 'PCS', 'STORE1'], ['42050002', 'MOTOR CIRCUIT BREAKER', 'PKZM0-10', '3', 'PCS', 'STORE2'], ['43401027', 'AUXILIARY RELAY', 'RH4B-U', '10', 'PCS', 'STORE1'], ['42310010', 'CAM SWITCH', 'CA10 A222 EG', '4', 'PCS', 'STORE1'], ['10190005', 'HEX BOLT', 'M6x20', '50', 'PCS', 'STORE3'], ['00030001', 'OLD PART', '', '0', 'PCS', 'STORE9']];
    window.__srv = [];
    // every open project also holds a By Project CODE (stock left by project 1034) booked from Stock · onhand 1 more than all bookings
    window.__pj = v => { if (!v || !v.projects) return; const K = '4556PP1034C00701'; let n = 0;
      v.projects.forEach(pr => { if (pr.status !== 'active') return; n++; if (pr.items.some(i => i.code === K)) return; pr.items.push({ item: '', code: K, qty: 2, material: 'TERMINAL BLOCK UK5N', st: 'ok', by: 'DS13', fromStock: 2, unit: 'PCS', pj: 1 }); pr.total++; pr.count.ok++; });
      v.codes = (v.codes || []).filter(c => c.code !== K).concat([{ code: K, pj: 1, desc: 'TERMINAL BLOCK UK5N', unit: 'PCS', by: 'DS13', onhand: 2 * n + 1, incoming: 0, booked: 2 * n, balance: 1, max: null, over: false, short: false, known: true, projects: [], po: [] }]); v.pjL = 'C'; };
    const tsv = stock.map(f => [f[0], f[1], f[2], f[3], f[4], f[5], '', 'DS13', '', '1'].join('\t')).join('\n');
    window.__bundle = (async () => { const cs = new CompressionStream('gzip'), w = cs.writable.getWriter(); w.write(new TextEncoder().encode(tsv)); w.close();
      const buf = new Uint8Array(await new Response(cs.readable).arrayBuffer()); let bin = ''; buf.forEach(c => bin += String.fromCharCode(c)); return btoa(bin); })();
    Object.defineProperty(G, 'run', { configurable: true, get() {
      const inner = d.get.call(G); let ok = () => {};
      const w = { withSuccessHandler(f) { ok = f; inner.withSuccessHandler(v => { if (window.__lastFn === 'getBom') window.__bom = v; f(v); }); return w; }, withFailureHandler(f) { inner.withFailureHandler(f); return w; },
        api(t, fn, a) {
          window.__lastFn = fn;
          window.__srv.push(fn);
          if (fn === 'getStockMeta') { setTimeout(() => ok({ hash: 'bk12345678', rows: stock.length, items: 4, locations: ['STORE1', 'STORE2'], bys: [], updated: '17/11/2026 09:00' }), 20); return; }
          if (fn === 'getStockBundle') { window.__bundle.then(b64 => ok({ hash: 'bk12345678', b64, v: 3 })); return; }
          if (fn === 'getCodeGroups') { setTimeout(() => ok({ '4205': 'Motor circuit breaker', '4340': 'Relay', '4231': 'Cam switch' }), 20); return; }
          if (fn === 'searchStock') { const q = String(a[1] || '').toLowerCase(); setTimeout(() => ok({ total: 2, page: 1, pages: 1, rows: stock.filter(f => (f[0] + ' ' + f[1]).toLowerCase().includes(q)), all: true }), 20); return; }
          if (fn === 'bomBook' || fn === 'bomSaveItems') {
            window.__bk.push([fn, JSON.parse(JSON.stringify(a))]);
            const v = JSON.parse(JSON.stringify(window.__bom || { projects: [], codes: [] }));
            if (fn === 'bomBook') { const r = a[1]; v.projects.unshift({ project: r.project, title: r.title, rev: '', status: 'active', needDate: r.needDate, daysLeft: 14, count: { ok: r.items.length, incoming: 0, late: 0, short: 0 }, total: r.items.length, items: r.items.map((x, i) => ({ item: '', code: x.code, qty: x.qty, material: 'MOTOR CIRCUIT BREAKER', st: 'ok', by: 'DS13', fromStock: x.qty })) });
              setTimeout(() => ok({ project: r.project, added: r.items.length, created: true, reopened: false, view: v }), 30); }
            else setTimeout(() => ok({ changes: ['42050002 1→5', '43401027 ⇄ X'], view: v }), 30);
            return;
          }
          if (fn === 'getBom') inner.withSuccessHandler(v => { window.__pj(v); window.__bom = v; ok(v); });
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
  // ── new layout: KPI cards filter the list · views vs actions · next step · ⋯ menu · tabs ──
  const kc = p.locator('.zq-bx-kc');
  check('layout: KPI cards are buttons (6 incl. เบิกแล้ว ยังไม่ปิด + overstock), the view switch and the action buttons are apart', await kc.count() === 7 && await p.locator('.zq-bx-seg button').count() === 2 && await p.locator('.zq-bx-act button:has-text("จองจาก Stock")').count() === 1 && await p.locator('.zq-bx-act button:has-text("Import BOM")').count() === 1);
  await kc.nth(3).click(); await p.waitForTimeout(300);
  const rk = await p.locator('.zq-bx-li').count();
  await kc.nth(3).click(); await p.waitForTimeout(300);
  check('layout: a KPI card filters the project list, clicking it again clears it', /กรอง: พร้อมเบิก/.test(await p.locator('main').innerText()) === false && await p.locator('.zq-bx-li').count() >= rk, rk);
  check('layout: grouped compact list + next-step box + detail tabs', await p.locator('.zq-bx-gh').count() >= 1 && await p.locator('.zq-bx-nx').count() === 1 && await p.locator('.zq-bx-dt [role=tab]').count() === 3 && await p.locator('th:has-text("จัดสรร")').count() === 1 && await p.locator('th:has-text("จาก Stock")').count() === 0);
  check('layout: close / delete live in the ⋯ menu, not on the page', await p.locator('.zq-bm-det button:has-text("ปิดโปรเจกต์")').count() === 0);
  check('By Project CODE: not in the stock CODE list, a link to Order Project instead', !/4556PP1034C00701/.test(await p.locator('.zq-bm-det table.zq-bm-t').first().innerText()) && await p.locator('.zq-bm-det button:has-text("CODE By Project ที่จองไว้ 1 รายการ")').count() === 1 && /รายการ CODE \(\d+\)/.test(await p.locator('.zq-bx-dt [role=tab]').first().innerText()) && /จอง 1/.test(await p.locator('.zq-bx-dt [role=tab]:has-text("Order Project")').innerText()));
  await p.locator('.zq-bx-dt [role=tab]:has-text("Order Project")').click(); await p.waitForTimeout(400);
  const pjBox = await p.locator('.zq-bm-det .zq-bm-w').first().innerText();
  check('Order Project: booked By Project CODE listed on top with its allocation, Admin can edit / remove it', /จองจาก Stock — CODE By Project \(1\)/.test(pjBox) && /4556PP1034C00701/.test(pjBox) && /Stock 2/.test(pjBox) && await p.locator('.zq-bm-det button[aria-label="ลบ 4556PP1034C00701"]').count() === 1, pjBox);
  await p.locator('.zq-bm-det').screenshot({ path: OUT + 'bom_pj.png' });
  const poRows = await p.locator('.zq-bm-det .zq-bm-t').last().locator('tbody tr').allInnerTexts();
  check('Order Project: the list shows straight away, every line (received / cancelled too), open ones first, no show / hide button', poRows.length === 4 && /ออก PO|รอออก PO/.test(poRows[0]) && poRows.some(r => /รับครบ|รับแล้ว/.test(r)) && poRows.some(r => /ยกเลิก/.test(r)) && await p.locator('.zq-bm-det button:has-text("ดูรายการ")').count() === 0 && await p.locator('.zq-bm-det button:has-text("ซ่อนรายการ")').count() === 0, poRows);
  await p.locator('.zq-bx-dt [role=tab]:has-text("รายการ CODE")').click(); await p.waitForTimeout(200);
  await p.locator('button[aria-label="คำสั่งเพิ่มเติม"]').click(); await p.waitForTimeout(200);
  const mt = await p.locator('.zq-bx-m').innerText();
  check('⋯ menu: exports, edit, PDF, then close / delete after a divider', /By Stock/.test(mt) && /By Project/.test(mt) && /แก้ไขรายการ CODE/.test(mt) && /ปิดโปรเจกต์|เปิดโปรเจกต์/.test(mt) && /ลบโปรเจกต์/.test(mt) && await p.locator('.zq-bx-m hr').count() === 1, mt);
  await p.screenshot({ path: OUT + 'bom_ux.png' });
  await p.locator('.zq-bx-m button:has-text("แก้ไขรายการ CODE")').first().click(); await p.waitForTimeout(500);
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
  await p.locator('.zq-bk-r button').first().click(); await p.waitForTimeout(200);
  check('edit: changed / not used / swapped rows are marked', await ed.locator('tr.ch').count() >= 2 && await ed.locator('tr.dl').count() === 1 && /แทน/.test(await ed.innerText()));
  await p.screenshot({ path: OUT + 'bom_edit.png' });
  await p.fill('input[aria-label="เหตุผลที่แก้"]', 'ลูกค้าเปลี่ยนรุ่น');
  await p.locator('button:has-text("บันทึกการแก้ไข")').click(); await p.waitForTimeout(600);
  const sv = (await p.evaluate(() => window.__bk)).find(x => x[0] === 'bomSaveItems'), A = sv && sv[1];
  const its = A ? A[2] : [];
  check('edit: saves the whole new list with the starting list\'s signature, the swap and the reason',
    !!A && its.find(x => x.code === code0).qty === 5 && !its.find(x => x.code === code1) && !its.find(x => x.code === code2) && !!its.find(x => x.code === '42310010') && A[4]['42310010'] === code2 && /=/.test(A[3]) && A[5] === 'ลูกค้าเปลี่ยนรุ่น', A);
  check('edit: back to the table with the result', await p.locator('table.zq-bk-e').count() === 0 && /บันทึกแล้ว/.test(await p.locator('.zq-em-msg').first().innerText().catch(() => '')));
  // ── fix a wrong By Project CODE from the Order Project tab ──
  await p.locator('.zq-bx-dt [role=tab]:has-text("Order Project")').click(); await p.waitForTimeout(300);
  await p.locator('.zq-bm-det button[aria-label="แก้จำนวน 4556PP1034C00701"]').click(); await p.fill('input[aria-label="จำนวนใหม่ 4556PP1034C00701"]', '1');
  await p.locator('.zq-bm-det button:has-text("บันทึก")').first().click(); await p.waitForTimeout(500);
  let L2 = (await p.evaluate(() => window.__bk)).filter(x => x[0] === 'bomSaveItems').pop()[1];
  check('Order Project: edit qty saves the whole list with the new qty + signature + reason', L2[2].find(x => x.code === '4556PP1034C00701').qty === 1 && L2[2].length > 1 && /4556PP1034C00701=2/.test(L2[3]) && /แก้จำนวน/.test(L2[5]), L2);
  await p.locator('.zq-bx-dt [role=tab]:has-text("Order Project")').click().catch(() => {}); await p.waitForTimeout(300);
  await p.locator('.zq-bm-det button[aria-label="ลบ 4556PP1034C00701"]').click(); await p.waitForTimeout(500);
  L2 = (await p.evaluate(() => window.__bk)).filter(x => x[0] === 'bomSaveItems').pop()[1];
  check('Order Project: 🗑 removes the wrong CODE (rest kept)', !L2[2].some(x => x.code === '4556PP1034C00701') && L2[2].length > 0 && /ลบ CODE By Project/.test(L2[5]), L2);
  // ── the cart ──
  await p.locator('button:has-text("จองจาก Stock")').first().click(); await p.waitForTimeout(500);
  await p.waitForSelector('.zq-shop-c', { timeout: 8000 }).catch(() => {});
  const rail = p.locator('.zq-shop-cl');
  const rt = await rail.innerText();
  check('shop: named categories (Stock ช่วยใช้) by default, name first; unnamed prefixes in one "อื่น ๆ"; empty ones hidden', /Motor circuit breaker\s*4205/.test(rt) && /Relay\s*4340/.test(rt) && /อื่น ๆ \(ยังไม่ตั้งชื่อหมวด\)\s*1/.test(rt) && !/\b1019\b/.test(rt) && !/\b0003\b/.test(rt), rt);
  check('shop: the category list scrolls inside its box', await p.evaluate(() => { const n = document.querySelector('.zq-shop-cats'), l = document.querySelector('.zq-shop-cl'); return l.getBoundingClientRect().bottom <= n.getBoundingClientRect().bottom + 1; }));
  await rail.locator('button', { hasText: '4205' }).click(); await p.waitForTimeout(200);
  const cards = p.locator('.zq-shop-c'), card = cards.first();
  check('shop: a category shows only its CODEs; one card per CODE, locations added up', await cards.count() === 1 && /42050002/.test(await card.innerText()) && /15/.test(await card.locator('.zq-shop-oh').innerText()) && /STORE1, STORE2/.test(await card.innerText()), await card.innerText());
  check('shop: onhand is a bold green badge', await card.locator('.zq-shop-oh:not(.out)').count() === 1 && /15/.test(await card.locator('.zq-shop-oh').innerText()) && await card.locator('.zq-shop-oh').evaluate(e => getComputedStyle(e).backgroundColor) === 'rgb(209, 250, 229)');
  await card.locator('input').fill('2'); await card.locator('button.zq-bk-add').click(); await p.waitForTimeout(200);
  check('shop: the quantity on the card goes to the cart', await p.locator('.zq-bk-l').count() === 1 && await p.locator('.zq-bk-l input[aria-label^="จำนวน"]').inputValue() === '2' && /✓ 2/.test(await card.locator('button.zq-bk-add').innerText()));
  await rail.locator('button', { hasText: 'ทั้งหมด' }).first().click(); const before = (await p.evaluate(() => window.__srv)).length;
  await p.fill('input[aria-label="ค้นหาสินค้า"]', 'relay'); await p.waitForTimeout(150);
  check('shop: typing filters at once, with no server call', await cards.count() === 1 && /43401027/.test(await cards.first().innerText()) && (await p.evaluate(() => window.__srv)).slice(before).length === 0, (await p.evaluate(() => window.__srv)).slice(before));
  await p.fill('input[aria-label="ค้นหาสินค้า"]', ''); await rail.locator('button', { hasText: 'อื่น ๆ' }).click(); await p.waitForTimeout(150);
  check('shop: "อื่น ๆ" holds the unnamed prefixes', await cards.count() === 1 && /10190005/.test(await cards.first().innerText()));
  await rail.locator('button', { hasText: 'เพิ่มล่าสุด' }).click(); await p.waitForTimeout(150);
  check('shop: "เพิ่มล่าสุด" lists what was added', /42050002/.test(await p.locator('.zq-shop-g').innerText().catch(() => '')));
  await rail.locator('button', { hasText: 'ทั้งหมด' }).first().click();
  await p.locator('.zq-bk-seg button:has-text("โปรเจกต์ใหม่")').click();
  await p.fill('input[aria-label="เลขโปรเจกต์ใหม่"]', 'pp1099a01'); await p.fill('input[aria-label="วันที่ต้องใช้"]', '01/12/2026'); await p.fill('input[aria-label="ชื่อโปรเจกต์"]', 'Local Control Cabinet (LCC.)');
  await p.screenshot({ path: OUT + 'bom_cart.png' });
  const go = p.locator('.zq-bk-c button.zq-or-btn');
  check('cart: the book button names the project and the count', /PP1099A01 \(1 รายการ\)/.test(await go.innerText()), await go.innerText());
  await go.click(); await p.waitForTimeout(800);
  const bk = (await p.evaluate(() => window.__bk)).find(x => x[0] === 'bomBook'), R = bk && bk[1][1];
  check('cart: books the project (upper-case), date, title, CODE × 2', !!R && R.project === 'PP1099A01' && R.mode === 'new' && R.needDate === '01/12/2026' && R.items.length === 1 && R.items[0].code === '42050002' && R.items[0].qty === 2, R);
  check('cart: emptied, and the new project opens', await p.evaluate(() => JSON.parse(localStorage.getItem('zqBomCart')).items.length) === 0 && /PP1099A01/.test(await p.locator('.zq-bm-det').innerText().catch(() => '')));
  // ── over-booking guard in the cart ──
  await p.evaluate(() => { localStorage.setItem('zqBomCart', JSON.stringify({ mode: 'add', project: window.__bom.projects.find(x => x.status === 'active').project, items: [{ code: '4556PP1034C00701', qty: '2', desc: 'TERMINAL BLOCK UK5N', unit: 'PCS' }, { code: '42050002', qty: '99', desc: 'MOTOR CIRCUIT BREAKER', unit: 'PCS', onhand: 15 }] })); window.dispatchEvent(new Event('zq:bomcart')); });
  await p.locator('button:has-text("จองจาก Stock")').first().click(); await p.waitForTimeout(500);
  const ct = await p.locator('.zq-bk-c').innerText();
  await p.locator('.zq-bk-c').screenshot({ path: OUT + 'bom_pj_cart.png' });
  check('cart: By Project CODE over free → tagged, blocked with the free qty; book disabled', /By Project/.test(ct) && /จองได้ไม่เกินของว่าง 1/.test(ct) && await go.isDisabled(), ct);
  await p.fill('input[aria-label="จำนวน 4556PP1034C00701"]', '1'); await p.waitForTimeout(100);
  check('cart: stock CODE over free still needs "ยืนยันจองเกิน"', await go.isDisabled() && await p.locator('input[aria-label="ยืนยันจองเกิน 42050002"]').count() === 1);
  await p.locator('input[aria-label="ยืนยันจองเกิน 42050002"]').check(); await p.waitForTimeout(100);
  check('cart: within free + ticked → can book', !(await go.isDisabled()));
  await go.click(); await p.waitForTimeout(600);
  const bk2 = (await p.evaluate(() => window.__bk)).filter(x => x[0] === 'bomBook').pop()[1][1];
  check('cart: sends the ticked stock CODE as allowShort (never the By Project one)', JSON.stringify(bk2.allowShort) === '["42050002"]' && bk2.items.find(x => x.code === '4556PP1034C00701').qty === 1, bk2);
  // ── from the item panel ──
  await p.evaluate(() => window.dispatchEvent(new CustomEvent('zq:item', { detail: '46700501' }))); await p.waitForTimeout(1500);
  const ib = p.locator('button:has-text("จอง BOM")').first();
  if (await ib.count()) { await ib.click(); await p.waitForTimeout(200); }
  check('item panel: "📋 จอง BOM" puts the CODE in the cart', await p.evaluate(() => (JSON.parse(localStorage.getItem('zqBomCart') || '{"items":[]}').items || []).some(x => x.code === '46700501')));
  check('the cart tab counts it', /จองจาก Stock \(1\)/.test(await p.locator('button:has-text("จองจาก Stock")').first().innerText().catch(() => '')));
  console.log('errors:', errs); srv.close(); await b.close(); process.exit(errs.length || fails.length ? 1 : 0);
})();
