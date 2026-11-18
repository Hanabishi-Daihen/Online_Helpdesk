// Estimate Cost (Admin): department → project (typed or picked) → lines, merge by CODE, report settings, PDF
const OUT = __dirname + '/../out/', fs = require('fs');
const { chromium } = require('playwright');
(async () => {
  const b = await chromium.launch({ executablePath: process.env.CHROMIUM_PATH || undefined });
  const ctx = await b.newContext({ viewport: { width: 1440, height: 960 } }), p = await ctx.newPage(), errs = [], fails = [];
  p.on('pageerror', e => errs.push(e.message)); p.on('dialog', d => d.accept());
  const check = (n, c, x) => { console.log((c ? 'PASS ' : 'FAIL ') + n + (c ? '' : ' → ' + JSON.stringify(x))); c || fails.push(n); };
  const srv = require('http').createServer((q, r) => { r.setHeader('content-type', 'text/html; charset=utf-8'); r.end(fs.readFileSync(OUT + 'doc.html')); }).listen(8847);
  await p.goto('http://localhost:8847/'); await p.waitForTimeout(1500);
  await p.evaluate(() => {
    const M = window.__M;
    M.getEstMeta = { ok: true, ver: '1', range: '01/01/2023 – 05/11/2026', last: { ok: true, time: '18/11/2026 04:05:00', msg: 'ok' }, cfg: {}, hasUrl: true,
      sum: [['CONTROL WIRING', '1030P1', 'PP103001', 5, -43047.8, '05/11/2026', '20261105'], ['CONTROL WIRING', '1024P2', 'PP102402', 2, -900, '01/10/2026', '20261001'],
        ['TANKING', '1030P1', 'PP103001', 1, -75, '01/11/2026', '20261101'], ['CORE SHEARING', '1043P1', 'PP104301', 4, -10, '01/09/2026', '20260901']] };
    M.getEstCost = { cat: '1030P1', project: 'PP103001', cols: [], lines: [
      ['37000017', 'Winding Thermometer', '20261105', 'MF241A/PP1030D01-02', 'CONTROL WIRING', 'WH1', 2, 'EA', 33949.07, -67898.15, 'USR1'],
      ['4556PP1030C00201', 'Annunciator', '20261104', '154862/3098', 'CONTROL WIRING', 'WH1', 1, 'EA', 37800, 37800, 'USR1'],
      ['4556PP1030C00201', 'Annunciator', '20261103', '154860/3098', 'CONTROL WIRING', 'WH1', 1, 'EA', 37800, -37800, 'USR1'],
      ['46420024', 'SOCKET RELAY PYFZ-14 (FOR MY4N)', '20261102', '154850/3090', 'CONTROL WIRING', 'WH1', 67, 'PC', 190.575, -12768.5, 'USR2'],
      ['46420024', 'SOCKET RELAY PYFZ-14 (FOR MY4N)', '20261101', '154849/3090', 'CONTROL WIRING', 'WH1', 2, 'PC', 190.575, -381.15, 'USR2'],
      ['10190005', 'HEX BOLT M6x20', '20261101', '154000/1', 'TANKING', 'WH2', 50, 'PC', 1.5, -75, 'USR3']] };
    M.saveEstCfg = { 'CONTROL WIRING': { title: 'ESTIMATE COST OF CONTROL CABINET', sign: ['P.Approver', 'C.Checker', 'Y.Preparer'] } };
    M.getShipSource = { url: '', hours: 0, last: null, sheetUrl: '' };
  });
  check('user: no Estimate Cost in the menu', await p.locator('aside nav button:has-text("Estimate Cost")').count() === 0);
  await p.locator('button', { hasText: /Admin|ปลดล็อค/ }).first().click(); await p.waitForTimeout(500);
  const u = p.locator('input[autocomplete=username]'); if (await u.count()) await u.fill('premwit');
  const pw = p.locator('input[type=password]'); await pw.first().fill('control'); await pw.first().press('Enter'); await p.waitForTimeout(1500);
  await p.locator('aside nav button:has-text("Estimate Cost")').first().click(); await p.waitForTimeout(1000);
  const dept = p.locator('select[aria-label="แผนก"]');
  check('departments from the data, with project counts', /CONTROL WIRING \(2 โปรเจกต์\)/.test(await dept.innerText()) && /ทุกแผนก/.test(await dept.innerText()), await dept.innerText());
  await dept.selectOption('CONTROL WIRING'); await p.waitForTimeout(300);
  check('project list of the department, last ship first', (await p.locator('.zq-ec-pj').allInnerTexts()).map(t => t.split(/\s/)[0]).join() === 'PP103001,PP102402');
  await p.fill('input[aria-label="โปรเจกต์"]', '1024'); await p.waitForTimeout(200);
  check('typing filters the list', await p.locator('.zq-ec-pj').count() === 1);
  await p.fill('input[aria-label="โปรเจกต์"]', 'PP1030D01'); await p.press('input[aria-label="โปรเจกต์"]', 'Enter'); await p.waitForTimeout(800);
  const rowsN = () => p.locator('.zq-ec-tb tbody tr').count(), sumT = () => p.locator('.zq-ec-sum').innerText();
  check('PP1030D01 → 1030P1: the department\'s lines one by one (TANKING left out), newest first', await rowsN() === 5 && /^37000017/.test(await p.locator('.zq-ec-tb tbody tr').first().innerText()), await rowsN());
  check('TOTAL = sum of the lines, negative', /-81,047\.80/.test(await sumT()), await sumT());
  check('a reversal (+) row is marked', await p.locator('.zq-ec-tb tr.rv').count() === 1);
  check('other departments of the project are named', /TANKING/.test(await sumT()));
  await p.locator('label:has-text("รวมบรรทัดตาม CODE") input').check(); await p.waitForTimeout(300);
  check('merge by CODE: returns cancel out, quantities add up, same TOTAL', await rowsN() === 2 && /-81,047\.80/.test(await sumT()) && /69\.00/.test(await p.locator('.zq-ec-tb').innerText()), [await rowsN(), await sumT()]);
  await p.locator('label:has-text("รวมบรรทัดตาม CODE") input').uncheck();
  await p.locator('button:has-text("หัวรายงาน / ผู้ลงนาม")').click(); await p.waitForTimeout(200);
  check('report settings start with the department name as the title', (await p.locator('.zq-ws-card input[placeholder="ESTIMATE COST OF CONTROL WIRING"]').count()) === 1);
  await p.locator('button:has-text("บันทึก")').last().click(); await p.waitForTimeout(500);
  check('saved title / signers show on the page', /ESTIMATE COST OF CONTROL CABINET/.test(await p.locator('main').innerText()) && /P\.Approver \/ C\.Checker \/ Y\.Preparer/.test(await p.locator('main').innerText()));
  await p.screenshot({ path: OUT + 'estimate_cost.png' });
  const [pop] = await Promise.all([ctx.waitForEvent('page', { timeout: 5000 }).catch(() => null), p.locator('button:has-text("🖨 PDF")').click()]);
  if (pop) { await pop.waitForTimeout(800); const t = await pop.locator('body').innerText();
    check('PDF: title, PROJECT, TOTAL (THB/UNIT), 5 lines, signers', /ESTIMATE COST OF CONTROL CABINET/.test(t) && /PROJECT : PP103001/.test(t) && /TOTAL : \(THB\/UNIT\) -81,048/.test(t) && /APPROVED BY/.test(t) && /Y\.Preparer/.test(t) && await pop.locator('table tr').count() >= 7, t.slice(0, 400));
    await pop.close(); } else check('PDF window opened', false);
  check('no page errors', errs.length === 0, errs);
  srv.close(); await b.close();
  process.exit(fails.length || errs.length ? 1 : 0);
})();
