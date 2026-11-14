// sign-in page: first page plays by itself (wheel speeds it up) → "เข้าสู่ระบบ" → forms; held calls; register / forgot;
// remembered session; role from the account (no separate Admin unlock); phone; CDN blocked
// prints what it sees; exits 1 when the page throws an error or a check fails
const OUT = __dirname + '/../out/', fs = require('fs');
const { chromium } = require('playwright');
const lib = p => { try { return fs.readFileSync(require.resolve(p), 'utf8'); } catch (e) { return null; } };
const MAP = [[/three\.js\/r128\/three\.min\.js/, 'three/build/three.min.js'], [/three@0\.128\.0\/(examples\/js\/.*\.js)/, m => 'three/' + m[1]]];
let libsServed = 0;
const cdn = r => { const u = r.request().url(); for (const [re, p] of MAP) { const m = re.exec(u); if (m) { const body = lib(typeof p === 'function' ? p(m) : p); if (body) { libsServed++; return r.fulfill({ contentType: 'application/javascript', body }); } } } return r.abort(); };
(async () => {
  const b = await chromium.launch({ executablePath: process.env.CHROMIUM_PATH || undefined, args: ['--use-gl=swiftshader', '--enable-unsafe-swiftshader', '--ignore-gpu-blocklist'] });
  const ctx = await b.newContext({ viewport: { width: 1360, height: 860 } });
  await ctx.route(/cdnjs\.cloudflare\.com|cdn\.jsdelivr\.net|unpkg\.com/, cdn); await ctx.route(/fonts\.(googleapis|gstatic)\.com/, r => r.abort());
  const p = await ctx.newPage(), errs = [], fails = [];
  p.on('pageerror', e => errs.push(e.message)); p.on('dialog', d => d.accept());
  const check = (n, c, x) => { console.log((c ? 'PASS ' : 'FAIL ') + n + (c ? '' : ' → ' + JSON.stringify(x))); c || fails.push(n); };
  const srv = require('http').createServer((q, r) => { r.setHeader('content-type', 'text/html; charset=utf-8'); r.end(fs.readFileSync(OUT + 'auth.html')); }).listen(8832);
  await p.goto('http://localhost:8832/'); await p.waitForTimeout(3000);
  const A = p.locator('#zq-auth'), des0 = await p.evaluate(() => localStorage.getItem('zqOrDes') || '');
  check('first page: intro with the sign-in button', await A.getAttribute('data-mode') === 'intro' && await p.locator('#za-go').isVisible());
  check('page calls wait for sign-in', await p.evaluate(() => (window.__apiLog || []).filter(x => !/^auth|getAppUrl/.test(x[0])).length) === 0);
  check('3D libraries served + canvas on', libsServed >= 5 && await p.locator('.za-cv.on').count() === 1, libsServed);
  const capNow = () => p.locator('.za-mode').innerText().catch(() => '');
  const seen = async ms => { const S = new Set(); const t0 = Date.now(); while (Date.now() - t0 < ms) { S.add(await capNow()); await p.waitForTimeout(150); } return S; };
  await p.screenshot({ path: OUT + 'auth_intro.png' });
  check('labels on the model (overview)', await p.locator('.za-lb:visible').count() >= 5, await p.locator('.za-lb:visible').count());
  const pr0 = await p.evaluate(() => window.zqAuth.progress()); await p.waitForTimeout(2500); const pr1 = await p.evaluate(() => window.zqAuth.progress());
  console.log('  story progress ' + pr0.toFixed(2) + ' → ' + pr1.toFixed(2) + ' in 2.5 s without input');
  check('plays by itself', pr1 > pr0, [pr0, pr1]);
  const sp0 = await p.evaluate(() => window.zqAuth.speed()); await p.mouse.move(600, 500);
  for (let i = 0; i < 6; i++) { await p.mouse.wheel(0, 240); await p.waitForTimeout(40); }
  const sp1 = await p.evaluate(() => window.zqAuth.speed()); console.log('  speed ×' + sp0.toFixed(2) + ' → ×' + sp1.toFixed(2) + ' after 6 wheel notches');
  check('mouse wheel speeds it up (and the page does not scroll)', sp1 > 3 && sp0 < 1.5 && await p.evaluate(() => document.getElementById('zq-auth').scrollTop === 0), [sp0, sp1]);
  await p.screenshot({ path: OUT + 'auth_part.png' });
  // the console: live numbers from one model, parts you can isolate, drag to orbit, click a part, exploded view
  const v0 = await p.evaluate(() => window.zqAuth.view());
  const lvTxt = await p.locator('[data-v="lv"]').first().innerText(), hsTxt = await p.locator('[data-v="hs"]').first().innerText();
  check('panels read the model: 33 kV side held near 33 kV by the tap changer (it lags while the tour runs ×10), hot-spot shown', Math.abs(v0.lv - 33) < 1 && /^\d+\.\d{2} kV$/.test(lvTxt) && /°C/.test(hsTxt), [v0.lv, lvTxt, hsTxt]);
  await p.locator('.za-asm [data-k="6"]').click(); await p.waitForTimeout(1500);
  check('assembly button isolates the part', /แกนเหล็ก/.test(await capNow()) && (await p.evaluate(() => window.zqAuth.view())).k === 6, await capNow());
  check('a hand on the controls stops the tour', (await p.evaluate(() => window.zqAuth.view())).play === false);
  await p.locator('.za-disp [data-d="ex"]').click(); await p.waitForTimeout(800);
  check('exploded view', (await p.evaluate(() => window.zqAuth.view())).ex === 1 && /EXPLODED/.test(await capNow()));
  await p.locator('.za-asm [data-k="0"]').click(); await p.waitForTimeout(1500);
  const box = await p.locator('.za-cv').boundingBox();
  await p.mouse.move(box.x + box.width / 2, box.y + box.height / 2); await p.mouse.down(); await p.mouse.move(box.x + box.width / 2 + 160, box.y + box.height / 2 + 30, { steps: 6 }); await p.mouse.up();
  check('drag orbits (and counts as the user)', (await p.evaluate(() => window.zqAuth.view())).user === true);
  await p.mouse.click(box.x + box.width / 2, box.y + box.height * .56); await p.waitForTimeout(600);
  const kk = (await p.evaluate(() => window.zqAuth.view())).k; check('click on the model isolates the part under the pointer', kk >= 1 && kk <= 7, kk);
  await p.locator('.za-pp').click(); await p.waitForTimeout(300);
  check('play button hands the tour back', (await p.evaluate(() => window.zqAuth.view())).play === true);
  // "เข้าสู่ระบบ" → the form slides in
  await p.click('#za-go'); await p.waitForTimeout(1200);
  check('form page after the button', await A.getAttribute('data-mode') === 'form' && await p.locator('#za-id').isVisible());
  await p.screenshot({ path: OUT + 'auth_form.png' });
  await p.fill('#za-id', 'somchai@example.co.th'); await p.fill('#za-pw', 'nope'); await p.click('.za-card .za-btn[type=submit]'); await p.waitForTimeout(300);
  check('wrong password message', /ไม่ถูกต้อง/.test(await p.locator('.za-msg').innerText()));
  await p.click('.za-tabs button:nth-child(2)'); await p.fill('#za-name', 'New User'); await p.fill('#za-email', 'new@example.co.th'); await p.fill('#za-pw', 'Abcdef12'); await p.fill('#za-pw2', 'Abcdef12');
  check('password rules all green', (await p.locator('.za-rules span.ok').count()) === 3);
  await p.click('.za-card .za-btn[type=submit]'); await p.waitForTimeout(300);
  await p.locator('.za-otp input').first().fill('111111'); await p.click('.za-card .za-btn[type=submit]'); await p.waitForTimeout(300);
  check('wrong code message', /ไม่ถูกต้อง/.test(await p.locator('.za-msg').innerText()));
  await p.locator('.za-otp input').first().fill('123456'); await p.click('.za-card .za-btn[type=submit]'); await p.waitForTimeout(300);
  check('waits for Admin approval', /รอ Admin อนุมัติ/.test(await p.locator('.za-card').innerText()));
  await p.click('.za-card .za-btn.ghost'); await p.click('text=ลืมรหัสผ่าน?'); await p.fill('#za-email', 'somchai@example.co.th'); await p.click('.za-card .za-btn[type=submit]'); await p.waitForTimeout(300);
  await p.locator('.za-otp input').first().fill('654321'); await p.fill('#za-pw', 'Newpass77'); await p.fill('#za-pw2', 'Newpass77'); await p.click('.za-card .za-btn[type=submit]'); await p.waitForTimeout(1500);
  check('signed in: page closed', await A.count() === 0);
  const sent = await p.evaluate(() => (window.__apiLog || []).filter(x => !/^auth|getAppUrl/.test(x[0])));
  check('held calls went out with the session', sent.length > 0 && sent.every(x => x[1]), sent.slice(0, 4));
  // forgot-password sign-in carries no start-up data in this fake; the plain sign-in below does → its first calls are answered locally
  const side = await p.locator('aside').innerText();
  check('sidebar: user card, no Admin unlock', /S\.Somchai/.test(side) && /ผู้ใช้งาน/.test(side) && !/ปลดล็อค|Admin: Locked/.test(side), side.slice(-160));
  await p.click('#zq-me'); check('account panel', /ออกจากระบบ/.test(await p.locator('#zq-me-p').innerText()));
  const des1 = await p.evaluate(() => localStorage.getItem('zqOrDes'));
  check('DES name: filled when empty, kept when already set', des0 ? des1 === des0 : des1 === 'S.Somchai', [des0, des1]);
  await p.reload(); await p.waitForTimeout(1500);
  check('remembered session: no sign-in page after reload', await A.count() === 0);
  // Admin role from the account
  await p.evaluate(() => { localStorage.removeItem('zqSes'); sessionStorage.removeItem('zqSes'); }); await p.reload(); await p.waitForTimeout(1500);
  await p.click('#za-go'); await p.waitForTimeout(800); await p.fill('#za-id', 'admin1'); await p.fill('#za-pw', 'Admin2026x'); await p.click('.za-card .za-btn[type=submit]'); await p.waitForTimeout(1500);
  const side2 = await p.locator('aside').innerText();
  check('Admin account: Admin mode on, no unlock button', /A\.Admin/.test(side2) && /ADMIN/.test(side2) && !/ปลดล็อค/.test(side2), side2.slice(0, 120) + ' … ' + side2.slice(-120));
  await p.reload(); await p.waitForTimeout(2000);
  check('Admin mode comes back after reload', /ADMIN/.test(await p.locator('aside').innerText()));
  await p.screenshot({ path: OUT + 'auth_admin_app.png' });
  // sign out: Apps Script nests the page in iframes, so no location.reload() (white page); the app URL opens in the top window
  await p.evaluate(() => { window.zqApp = window.zqApp || {}; window.zqApp.url = location.origin + '/?signed-out'; });
  await p.click('#zq-me'); await p.click('#zq-me-p button:has-text("ออกจากระบบ")'); await p.waitForTimeout(2500);
  check('sign out: app URL opened in the top window, sign-in page shown', /signed-out/.test(p.url()) && await A.count() === 1 && await p.evaluate(() => !localStorage.getItem('zqSes')), p.url());
  // when the top window cannot be opened: the sign-in page comes up in place and Admin mode is dropped
  await p.goto('http://localhost:8832/'); await p.waitForTimeout(1200); await p.click('#za-go'); await p.waitForTimeout(500);
  await p.fill('#za-id', 'admin1'); await p.fill('#za-pw', 'Admin2026x'); await p.click('.za-card .za-btn[type=submit]'); await p.waitForTimeout(1500);
  await p.evaluate(() => { window.zqApp = window.zqApp || {}; window.zqApp.url = ''; if (window.__ZQ_BOOT__ && window.__ZQ_BOOT__.r) window.__ZQ_BOOT__.r.getAppUrl = ''; });
  await p.click('#zq-me'); await p.click('#zq-me-p button:has-text("ออกจากระบบ")'); await p.waitForTimeout(1500);
  check('sign out without top window: sign-in page in place, Admin mode off', await A.count() === 1 && await A.getAttribute('data-mode') === 'form' && !/ADMIN/.test(await p.locator('aside:not(.za-cap)').innerText()));
  await p.fill('#za-id', 'somchai@example.co.th'); await p.fill('#za-pw', 'Passw0rd99'); await p.click('.za-card .za-btn[type=submit]'); await p.waitForTimeout(1800);
  check('next sign-in after that works', await A.count() === 0 && /S\.Somchai/.test(await p.locator('aside').innerText()));
  // phone
  const m = await b.newContext({ viewport: { width: 390, height: 844 }, hasTouch: true, isMobile: true }); await m.route(/cdnjs\.cloudflare\.com|cdn\.jsdelivr\.net|unpkg\.com/, cdn);
  const pm = await m.newPage(); pm.on('pageerror', e => errs.push('phone: ' + e.message));
  await pm.goto('http://localhost:8832/'); await pm.waitForTimeout(3000); await pm.screenshot({ path: OUT + 'auth_phone.png' });
  await pm.click('#za-go'); await pm.waitForTimeout(1200); await pm.screenshot({ path: OUT + 'auth_phone_form.png' });
  check('phone: form reachable, nothing sideways', await pm.locator('#za-id').isVisible() && !(await pm.evaluate(() => document.documentElement.scrollWidth > innerWidth)));
  // reduced motion: still picture, buttons work
  const rm = await b.newContext({ viewport: { width: 1200, height: 800 }, reducedMotion: 'reduce' }); await rm.route(/cdnjs\.cloudflare\.com|cdn\.jsdelivr\.net|unpkg\.com/, cdn);
  const pr = await rm.newPage(); pr.on('pageerror', e => errs.push('reduced: ' + e.message));
  await pr.goto('http://localhost:8832/'); await pr.waitForTimeout(2500); const c1 = await pr.locator('.za-mode b').innerText(), t1 = await pr.evaluate(() => window.zqAuth.progress()); await pr.waitForTimeout(3000);
  check('reduced motion: no autoplay', c1 === await pr.locator('.za-mode b').innerText() && t1 === await pr.evaluate(() => window.zqAuth.progress()), c1);
  // CDN blocked
  const n = await b.newContext({ viewport: { width: 1200, height: 800 } }); await n.route(/cdnjs|jsdelivr|unpkg|fonts\./, r => r.abort());
  const pn = await n.newPage(); pn.on('pageerror', e => errs.push('nocdn: ' + e.message));
  await pn.goto('http://localhost:8832/'); await pn.waitForTimeout(1500); check('CDN blocked: the console still reads the model (no 3D)', /^\d+\.\d$/.test(await pn.locator('[data-v="mva"]').innerText()) && await pn.locator('.za-cv.on').count() === 0, await pn.locator('[data-v="mva"]').innerText()); await pn.click('#za-go'); await pn.waitForTimeout(800);
  await pn.fill('#za-id', 'somchai@example.co.th'); await pn.fill('#za-pw', 'Passw0rd99'); await pn.click('.za-card .za-btn[type=submit]'); await pn.waitForTimeout(1200);
  check('CDN blocked: sign-in still works', await pn.locator('#zq-auth').count() === 0);
  await pn.waitForTimeout(800); const sent2 = await pn.evaluate(() => (window.__apiLog || []).map(x => x[0]));
  check('start-up data came with the sign-in answer: no separate getSettings / getPublicTemplates / getPulse calls', !sent2.some(f => /^(getSettings|getPublicTemplates|getPulse)$/.test(f)) && sent2.length > 1, sent2);
  console.log('errors:', errs); srv.close(); await b.close(); process.exit(errs.length || fails.length ? 1 : 0);
})();
