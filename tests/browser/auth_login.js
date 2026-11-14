// sign-in page: held calls, wrong / right password, register + e-mail code, forgot password, remembered session, 3D story
// prints what it sees; exits 1 when the page throws an error or a check fails
const OUT = __dirname + '/../out/', fs = require('fs'), path = require('path');
const { chromium } = require('playwright');
const lib = p => { try { return fs.readFileSync(require.resolve(p), 'utf8'); } catch (e) { return null; } };
(async () => {
  const b = await chromium.launch({ executablePath: process.env.CHROMIUM_PATH || undefined, args: ['--use-gl=swiftshader', '--enable-unsafe-swiftshader', '--ignore-gpu-blocklist'] });
  const ctx = await b.newContext({ viewport: { width: 1360, height: 860 } });
  // CDN → local copies (cdnjs / jsDelivr / unpkg are not reachable from the test machine)
  const MAP = [[/three\.js\/r128\/three\.min\.js/, 'three/build/three.min.js'], [/three@0\.128\.0\/(examples\/js\/.*\.js)/, m => 'three/' + m[1]], [/gsap\/3\.12\.5\/gsap\.min\.js/, 'gsap/dist/gsap.min.js'],
    [/gsap\/3\.12\.5\/ScrollTrigger\.min\.js/, 'gsap/dist/ScrollTrigger.min.js'], [/lenis@1\.1\.13\/dist\/lenis\.min\.js/, 'lenis/dist/lenis.min.js']];
  let libsServed = 0;
  await ctx.route(/cdnjs\.cloudflare\.com|cdn\.jsdelivr\.net|unpkg\.com/, r => { const u = r.request().url(); for (const [re, p] of MAP) { const m = re.exec(u); if (m) { const body = lib(typeof p === 'function' ? p(m) : p); if (body) { libsServed++; return r.fulfill({ contentType: 'application/javascript', body }); } } } return r.abort(); });
  await ctx.route(/fonts\.(googleapis|gstatic)\.com/, r => r.abort());
  const p = await ctx.newPage(), errs = [], fails = [];
  p.on('pageerror', e => errs.push(e.message)); p.on('dialog', d => d.accept());
  const check = (n, c, x) => { console.log((c ? 'PASS ' : 'FAIL ') + n + (c ? '' : ' → ' + JSON.stringify(x))); c || fails.push(n); };
  const http = require('http'); const srv = http.createServer((q, r) => { r.setHeader('content-type', 'text/html; charset=utf-8'); r.end(fs.readFileSync(OUT + 'auth.html')); }).listen(8832);
  await p.goto('http://localhost:8832/'); await p.waitForTimeout(1500);
  const A = p.locator('#zq-auth'), des0 = await p.evaluate(() => localStorage.getItem('zqOrDes') || '');
  check('sign-in page shows without a session', await A.isVisible());
  const held = await p.evaluate(() => (window.__apiLog || []).filter(x => !/^auth|getAppUrl/.test(x[0])).length);
  check('page calls wait for sign-in (none sent yet)', held === 0, held);
  await p.screenshot({ path: OUT + 'auth_hero.png' });
  // wrong then right password
  await p.fill('#za-id', 'somchai@example.co.th'); await p.fill('#za-pw', 'nope'); await p.click('.za-card .za-btn[type=submit]'); await p.waitForTimeout(300);
  check('wrong password message', /ไม่ถูกต้อง/.test(await p.locator('.za-msg').innerText()));
  // register: code view, wrong code, right code → pending
  await p.click('.za-tabs button:nth-child(2)'); await p.waitForTimeout(100);
  await p.fill('#za-name', 'New User'); await p.fill('#za-email', 'new@example.co.th'); await p.fill('#za-pw', 'Abcdef12'); await p.fill('#za-pw2', 'Abcdef12');
  check('password rules all green', (await p.locator('.za-rules span.ok').count()) === 3);
  await p.click('.za-card .za-btn[type=submit]'); await p.waitForTimeout(300);
  check('code boxes shown', (await p.locator('.za-otp input').count()) === 6);
  await p.locator('.za-otp input').first().fill('111111'); await p.waitForTimeout(80);
  await p.click('.za-card .za-btn[type=submit]'); await p.waitForTimeout(300);
  check('wrong code message', /ไม่ถูกต้อง/.test(await p.locator('.za-msg').innerText()));
  await p.locator('.za-otp input').first().fill('123456'); await p.click('.za-card .za-btn[type=submit]'); await p.waitForTimeout(300);
  check('waits for Admin approval', /รอ Admin อนุมัติ/.test(await p.locator('.za-card').innerText()));
  // forgot password → new password → signed in
  await p.click('.za-card .za-btn.ghost'); await p.click('text=ลืมรหัสผ่าน?'); await p.fill('#za-email', 'somchai@example.co.th'); await p.click('.za-card .za-btn[type=submit]'); await p.waitForTimeout(300);
  await p.locator('.za-otp input').first().fill('654321'); await p.fill('#za-pw', 'Newpass77'); await p.fill('#za-pw2', 'Newpass77');
  await p.click('.za-card .za-btn[type=submit]'); await p.waitForTimeout(1400);
  check('signed in: page closed', await A.count() === 0);
  const sent = await p.evaluate(() => (window.__apiLog || []).filter(x => !/^auth|getAppUrl/.test(x[0])));
  check('held calls went out with the session', sent.length > 0 && sent.every(x => x[1]), sent.slice(0, 5));
  check('user button shows initials', (await p.locator('#zq-me').innerText()) === 'SS');
  await p.click('#zq-me'); check('account panel', /S\.Somchai/.test(await p.locator('#zq-me-p').innerText()) && /ออกจากระบบ/.test(await p.locator('#zq-me-p').innerText()));
  const des1 = await p.evaluate(() => localStorage.getItem('zqOrDes'));
  check('DES name: filled when empty, kept when already set', des0 ? des1 === des0 : des1 === 'S.Somchai', [des0, des1]);
  // remembered session: reload → straight in
  await p.reload(); await p.waitForTimeout(1500);
  check('remembered session: no sign-in page after reload', await A.count() === 0);
  // 3D story (signed out again)
  await p.evaluate(() => { localStorage.removeItem('zqSes'); sessionStorage.removeItem('zqSes'); }); await p.reload(); await p.waitForTimeout(3500);
  check('3D libraries served', libsServed >= 10, libsServed);
  check('3D canvas on', await p.locator('.za-cv.on').count() === 1);
  await p.screenshot({ path: OUT + 'auth_3d_hero.png' });
  const H = await p.evaluate(() => { const e = document.getElementById('zq-auth'); return { h: e.scrollHeight, v: e.clientHeight }; });
  for (const [name, f] of [['part1', 0.12], ['part4', 0.45], ['part6', 0.7], ['assembled', 0.995]]) {
    await p.evaluate(y => { const e = document.getElementById('zq-auth'); e.scrollTop = y; e.dispatchEvent(new Event('scroll')); }, Math.round((H.h - H.v) * f));
    await p.waitForTimeout(2600); await p.screenshot({ path: OUT + 'auth_3d_' + name + '.png' });
    console.log('  ' + name + ': panel "' + ((await p.locator('.za-sec.cur h3, .za-end.cur h3').first().innerText().catch(() => '-'))) + '" · dot "' + ((await p.locator('.za-prog button.cur span').innerText().catch(() => '-'))) + '"');
  }
  const fps = await p.evaluate(() => new Promise(ok => { let n = 0, t0 = performance.now(); (function f() { n++; performance.now() - t0 < 1000 ? requestAnimationFrame(f) : ok(n); })(); }));
  console.log('  frames in 1 s (software GL, no GPU): ' + fps);
  // phone + reduced motion
  const m = await b.newContext({ viewport: { width: 390, height: 844 }, deviceScaleFactor: 1, reducedMotion: 'reduce', hasTouch: true, isMobile: true });
  await m.route(/cdnjs\.cloudflare\.com|cdn\.jsdelivr\.net|unpkg\.com/, r => { const u = r.request().url(); for (const [re, pp] of MAP) { const mm = re.exec(u); if (mm) { const body = lib(typeof pp === 'function' ? pp(mm) : pp); if (body) return r.fulfill({ contentType: 'application/javascript', body }); } } return r.abort(); });
  const pm = await m.newPage(); pm.on('pageerror', e => errs.push('phone: ' + e.message));
  await pm.goto('http://localhost:8832/'); await pm.waitForTimeout(3000);
  check('phone: no sideways scroll', !(await pm.evaluate(() => { const e = document.getElementById('zq-auth'); return e.scrollWidth > e.clientWidth; })));
  await pm.screenshot({ path: OUT + 'auth_phone.png' });
  // CDN blocked: sign-in still works
  const n = await b.newContext({ viewport: { width: 1200, height: 800 } }); await n.route(/cdnjs|jsdelivr|unpkg|fonts\./, r => r.abort());
  const pn = await n.newPage(); pn.on('pageerror', e => errs.push('nocdn: ' + e.message));
  await pn.goto('http://localhost:8832/'); await pn.waitForTimeout(1500);
  await pn.fill('#za-id', 'somchai@example.co.th'); await pn.fill('#za-pw', 'Passw0rd99'); await pn.click('.za-card .za-btn[type=submit]'); await pn.waitForTimeout(1200);
  check('CDN blocked: sign-in still works', await pn.locator('#zq-auth').count() === 0);
  console.log('errors:', errs); srv.close(); await b.close(); process.exit(errs.length || fails.length ? 1 : 0);
})();
