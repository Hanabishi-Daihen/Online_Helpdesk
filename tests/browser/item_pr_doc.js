// item details → each PO / order line gets "📄 เปิดไฟล์" when the FileList index has its PR (RQN) / Spec (DS1) file, else "ไม่พบใบ PR"
// prints what it sees; exits 1 when the page throws an error or a check fails
const OUT = __dirname + '/../out/', fs = require('fs');
const { chromium } = require('playwright');
(async () => {
  const b = await chromium.launch({ executablePath: process.env.CHROMIUM_PATH || undefined });
  const p = await b.newPage({ viewport: { width: 1360, height: 960 } }), errs = [], fails = [];
  p.on('pageerror', e => errs.push(e.message));
  const check = (n, c, x) => { console.log((c ? 'PASS ' : 'FAIL ') + n + (c ? '' : ' → ' + JSON.stringify(x))); c || fails.push(n); };
  const srv = require('http').createServer((q, r) => { r.setHeader('content-type', 'text/html; charset=utf-8'); r.end(fs.readFileSync(OUT + 'doc.html')); }).listen(8843);
  await p.goto('http://localhost:8843/'); await p.waitForTimeout(1500);
  // the item: one PO whose PR file is in the index, one received PO whose PR has no file, one order without PR / Spec
  await p.evaluate(() => {
    const G = window.google.script, d = Object.getOwnPropertyDescriptor(G, 'run');
    const item = { code: '3694PP1006C01401', key: '3694PP1006C01401', found: true, desc: 'Card Holder-LCC', unit: 'Pcs.', by: 'DS13', onhand: 10, locs: [], admin: false,
      incoming: [{ pr: 'RQN26-03780', spec: 'DS1C26-1161', po: 'PO26-00001', qty: 5, unit: 'Pcs.', due: '01/12/2026' }],
      received: [{ pr: 'RQN26-09999', spec: '', po: 'PO26-04098', qty: 10, unit: 'Pcs.', req: '22/09/2026', status: 'Received' }],
      orders: [{ id: 'ORD26-0001', qty: 2, unit: 'Pcs.', state: 'saved', rqn: '', spec: '' }] };
    Object.defineProperty(G, 'run', { configurable: true, get() {
      const inner = d.get.call(G); let ok = () => {};
      const w = { withSuccessHandler(f) { ok = f; inner.withSuccessHandler(f); return w; }, withFailureHandler(f) { inner.withFailureHandler(f); return w; },
        api(t, fn, a) { if (fn === 'getItem360') { setTimeout(() => ok(JSON.parse(JSON.stringify(item))), 20); return; } return inner.api(t, fn, a); } };
      return new Proxy(w, { get: (o, k) => k in o ? o[k] : inner[k] });
    } });
    window.dispatchEvent(new CustomEvent('zq:item', { detail: '3694PP1006C01401' }));
  });
  await p.waitForTimeout(2000);
  const box = p.locator('section', { hasText: 'ใบสั่งและ PO' }).last();
  const rows = box.locator('.zq-i360-l'), txt = await rows.allInnerTexts();
  console.log(txt.map(t => '  ' + t.replace(/\s+/g, ' ')).join('\n'));
  const open = box.locator('a.zq-po-of');
  check('PO with a PR file: open-file button', await open.count() === 1 && /เปิดไฟล์/.test(await open.innerText()), await open.count());
  const href = await open.getAttribute('href').catch(() => '');
  check('the button searches that PR in the OneDrive folder', /RQN26-03780/.test(decodeURIComponent(href || '')), href);
  check('received PO without a file: "ไม่พบใบ PR"', /PO26-04098[\s\S]*ไม่พบใบ PR/.test(txt.join('\n')), txt);
  check('order without a PR number: "ไม่พบใบ PR"', /ORD26-0001[\s\S]*ไม่พบใบ PR/.test(txt.join('\n')), txt);
  check('why it is missing is in the tooltip', /ไม่พบไฟล์ของ RQN26-09999/.test(await box.locator('.zq-pr-no').evaluateAll(e => e.map(x => x.title).join('|'))));
  await box.screenshot({ path: OUT + 'item_pr_doc.png' }).catch(() => {});
  console.log('errors:', errs); srv.close(); await b.close(); process.exit(errs.length || fails.length ? 1 : 0);
})();
