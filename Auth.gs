/**
 * ═════════════════════════════════════════════════════════════════════
 *  SIGN-IN — every page call goes through api(token, fn, args)
 *  • Users (sheet "Users"): register with an e-mail code, sign in, forgot password (e-mail code → new password),
 *    change password, sign out (this device / every device). Admin approves new accounts unless their e-mail
 *    domain is on the auto-approve list. Admin accounts (sheet Settings, A13) can sign in with the same form.
 *  • Sessions (sheet "Sessions" + script cache): a random token kept by the browser; the sheet keeps only its hash.
 *  • api() is the only door: it checks the session, then runs one of the functions in API_FNS. Each of those
 *    functions starts with apiOnly_(), so calling it straight from the browser console (google.script.run.x())
 *    is refused — it only runs when api(), doGet or a trigger / editor entry point started the execution.
 *  • Settings (Config sheet key AUTH_CFG): { on, domains, approve, days } — Admin → ตั้งค่าระบบ → บัญชีผู้ใช้งาน.
 *  ต้องมี Code.gs (admUsers_, adminLogin, admHash_, withLock_, sheet_, schedProp_, audit_, nowStr_)
 * ═════════════════════════════════════════════════════════════════════
 */
var __trusted = false, __user = null;                       // set for this execution by api() / an entry point
function apiOnly_() { if (!__trusted) throw new Error('AUTH_REQUIRED: กรุณาเข้าสู่ระบบผ่านหน้าเว็บ'); }
function trustEntry_() { __trusted = true; }                // doGet, triggers, editor helpers, the e-mail approval page

// functions the web page may call through api() — tests/server/auth.test.js checks every call in index.html is listed
var API_FNS = (function () {
  var o = {};
  (
  'adminActOrderApproval adminLogin backupNow bomExportLookup bomExportPrep bomPreview bomProjectAction ' +
  'changeAdminPassword checkDataVers checkOrderDup claimRemain closeRemainClaim createIssue deleteIssues ' +
  'deleteKitTemplate deleteOrder deleteOrderDbRow diagnose docLookup exportBomForm exportBomShortReport exportOrder ' +
  'exportRemainReport getAdminUsers getAppUrl getBom getBomOrdersAll getBomProjOrders getBomShortMailCfg getBomSummary ' +
  'getDataVers getDocIndex getDocPaths getDocSource getEmailConfig getEmailRoster getExportPdf getFeedback getForecast ' +
  'getIssueFull getIssues getIssuesVer getItem360 getKitTemplate getKitTemplates getMany getMyDay getNotify ' +
  'getOrderAgeSummary getOrderApprovalPulse getOrderBoot getOrderDb getOrderTemplate getOrderTrack getOrderTrackSummary ' +
  'getOrders getP360Cfg getPendingBundle getPendingMeta getPendingSource getProject360 getProjects getPublicTemplates ' +
  'getPulse getRemainFlags getRemainStock getRtSource getScheduleSource getSchedules getSettings getStockAlert ' +
  'getStockAlertMailCfg getStockAlertSummary getStockAlertUndo getStockBundle getStockCostCode getStockCosts getStockMeta ' +
  'getStockSource getSystemHealth getUsageSummary importBom importScheduleRows installTriggers listBackups logUsage ' +
  'orderItemInfo previewBomShortMail previewEmail previewStockAlertEmail removeAdminUser resetAdminPassword ' +
  'resetOrderTemplate saveAdminUser saveAllIssues saveAllSchedules saveBomCfg saveCategories saveEmailRoster ' +
  'saveEmailTemplates saveFcCfg saveIssueTemplates saveKitTemplate saveOrder saveOrderAgeCfg saveOrderApprovers ' +
  'saveOrderDbRow saveOrderDbRows saveOrderDupCfg saveOrderKitPdf saveP360Cfg saveRemainCfg saveRemainGroups ' +
  'saveRemainTag saveReportTemplate saveSettings saveStockAlertBulk saveStockAlertItem searchStock sendBomShortMail ' +
  'sendStockAlertEmail sendTestEmail sendTestEmails setDocOneDrive setDocSource setFeedbackStatus setPendingSource ' +
  'setRtSource setScheduleSource setStockAlertConfig setStockCostCode setStockSource submitFeedback submitOrderApproval ' +
  'syncDocsNow syncPendingNow syncRtNow syncSchedulesNow syncStockNow trackFillRef trackScanNow undoStockAlertBulk ' +
  'uploadOrderTemplate upsertIssues verifyStockCostCode ' +
  '').split(' ').forEach(function (k) { if (k) o[k] = 1; });
  return o;
})();
// callable before signing in
var AUTH_OPEN = { authLogin: 1, authRegisterStart: 1, authRegisterVerify: 1, authForgotStart: 1, authForgotReset: 1, authMe: 1, authInfo: 1, getAppUrl: 1 };

function api(tok, fn, args) {
  fn = String(fn || '');
  var G = globalThis;
  if (!/^[A-Za-z][A-Za-z0-9]{1,60}$/.test(fn) || !(API_FNS[fn] || AUTH_OPEN[fn] || AUTH_FNS[fn]) || typeof G[fn] !== 'function') throw new Error('not allowed: ' + fn);
  __trusted = true;
  if (!AUTH_OPEN[fn] && authCfg_().on) {
    var u = authSession_(tok);
    if (!u) throw new Error('AUTH_REQUIRED: เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่');
    __user = u;
  } else if (tok) { try { __user = authSession_(tok); } catch (e) {} }
  return G[fn].apply(null, Array.isArray(args) ? args.slice(0, 20) : []);
}
// signed-in calls that are not page functions
var AUTH_FNS = { authLogout: 1, authChangePassword: 1, authProfile: 1, authUsers: 1, authUserAction: 1, authSaveCfg: 1 };

// ── settings ──
var AUTH_DEF = { on: true, domains: '', approve: true, days: 30 };
function authCfg_() {
  var c = {}; try { c = JSON.parse(schedProp_('AUTH_CFG') || '{}') || {}; } catch (e) {}
  return { on: c.on !== false, domains: String(c.domains == null ? AUTH_DEF.domains : c.domains), approve: c.approve !== false, days: Math.min(90, Math.max(1, Number(c.days) || AUTH_DEF.days)) };
}
function authDomains_() { return authCfg_().domains.toLowerCase().split(/[\s,;]+/).map(function (d) { return d.replace(/^@/, ''); }).filter(Boolean); }
function authDomainOk_(email) { var d = String(email).split('@')[1] || '', L = authDomains_(); return L.indexOf(d) >= 0; }

// ── users sheet ──
var USR_SHEET = 'Users', USR_HEAD = ['Email', 'Name', 'Dept', 'Status', 'Salt', 'Hash', 'Created', 'Last login', 'Note'];
function usrSheet_() { var s = sheet_(USR_SHEET, USR_HEAD); if (s.getLastRow() === 0) s.appendRow(USR_HEAD); return s; }
function usrAll_() {
  var s = usrSheet_(), n = s.getLastRow() - 1;
  return n > 0 ? s.getRange(2, 1, n, USR_HEAD.length).getDisplayValues().map(function (r, i) {
    return { row: i + 2, e: String(r[0]).trim().toLowerCase(), n: r[1], d: r[2], st: r[3] || 'active', s: r[4], h: r[5], c: r[6], t: r[7], note: r[8] };
  }).filter(function (u) { return u.e; }) : [];
}
function usrFind_(email) { email = authEmail_(email); return usrAll_().filter(function (u) { return u.e === email; })[0] || null; }
function usrWrite_(u) {
  var s = usrSheet_(), row = u.row || s.getLastRow() + 1, rg = s.getRange(row, 1, 1, USR_HEAD.length);
  fmtText_(rg); rg.setValues([[u.e, u.n || '', u.d || '', u.st || 'active', u.s || '', u.h || '', u.c || nowStr_(), u.t || '', u.note || '']]);
  u.row = row; return u;
}
function authEmail_(e) { e = String(e || '').trim().toLowerCase(); return /^[^@\s<>"']{1,64}@[a-z0-9.-]{3,190}\.[a-z]{2,24}$/.test(e) ? e : ''; }
function authName_(n) { return String(n || '').replace(/[<>"]/g, '').replace(/\s+/g, ' ').trim().substring(0, 40); }
function authPassCheck_(pw, email, name) {
  pw = String(pw == null ? '' : pw);
  if (pw.length < 8 || pw.length > 64) throw new Error('รหัสผ่านต้องยาว 8–64 ตัวอักษร');
  if (!/[A-Za-z]/.test(pw) || !/\d/.test(pw)) throw new Error('รหัสผ่านต้องมีทั้งตัวอักษรและตัวเลข');
  var low = pw.toLowerCase();
  if (low === String(email).split('@')[0] || (name && low === String(name).toLowerCase())) throw new Error('รหัสผ่านต้องไม่ซ้ำกับอีเมลหรือชื่อ');
  if (/^(password|12345678|qwerty12|daihen12|daihen123|abcd1234)/i.test(pw)) throw new Error('รหัสผ่านนี้เดาง่ายเกินไป');
  return pw;
}
function authPub_(u) { return { email: u.e, name: u.n, dept: u.d, status: u.st, admin: !!u.admin }; }

// ── sessions ──
var SES_SHEET = 'Sessions', SES_HEAD = ['Token hash', 'Email', 'Expires', 'Created', 'Device'];
function sesSheet_() { var s = sheet_(SES_SHEET, SES_HEAD); if (s.getLastRow() === 0) s.appendRow(SES_HEAD); return s; }
function sesHash_(tok) { return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'ses|' + String(tok), Utilities.Charset.UTF_8)).substring(0, 43); }
function sesNew_(u, remember, device) {
  var tok = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, ''), h = sesHash_(tok);
  var exp = Date.now() + (remember ? authCfg_().days * 864e5 : 12 * 3600e3), v = { e: u.e, n: u.n, d: u.d || '', a: u.admin ? 1 : 0, x: exp };
  withLock_(function () {
    var s = sesSheet_(), rg = s.getRange(s.getLastRow() + 1, 1, 1, SES_HEAD.length); fmtText_(rg);
    rg.setValues([[h, u.e, String(exp), nowStr_(), String(device || '').substring(0, 80)]]);
    if (s.getLastRow() > 3000) sesPrune_(s);
  });
  CacheService.getScriptCache().put('ses_' + h, JSON.stringify(v), 21600);
  return { token: tok, expires: exp };
}
function sesPrune_(s) {
  var n = s.getLastRow() - 1; if (n < 1) return;
  var now = Date.now(), keep = s.getRange(2, 1, n, SES_HEAD.length).getValues().filter(function (r) { return Number(r[2]) > now; });
  s.getRange(2, 1, n, SES_HEAD.length).clearContent();
  if (keep.length) { var rg = s.getRange(2, 1, keep.length, SES_HEAD.length); fmtText_(rg); rg.setValues(keep); }
}
// → { e, n, d, a } or null
function authSession_(tok) {
  tok = String(tok || ''); if (tok.length < 40 || tok.length > 80) return null;
  var h = sesHash_(tok), C = CacheService.getScriptCache(), v = null;
  try { v = JSON.parse(C.get('ses_' + h) || 'null'); } catch (e) {}
  if (!v) {
    var s = sesSheet_(), n = s.getLastRow() - 1; if (n < 1) return null;
    var hit = s.getRange(2, 1, n, 3).getValues().filter(function (r) { return r[0] === h; })[0];
    if (!hit || Number(hit[2]) < Date.now()) return null;
    var em = String(hit[1]), u = em.indexOf('admin:') === 0 ? authAdminUser_(em.substring(6)) : usrFind_(em);
    if (!u || u.st !== 'active') return null;
    v = { e: u.e, n: u.n, d: u.d || '', a: u.admin ? 1 : 0, x: Number(hit[2]) };
    C.put('ses_' + h, JSON.stringify(v), 21600);
  }
  if (!v || v.x < Date.now()) return null;
  return v;
}
function sesRevoke_(email, exceptTok) {                       // sign one account out everywhere
  var keepH = exceptTok ? sesHash_(exceptTok) : '';
  withLock_(function () {
    var s = sesSheet_(), n = s.getLastRow() - 1; if (n < 1) return;
    var rows = s.getRange(2, 1, n, SES_HEAD.length).getValues(), gone = [], keep = rows.filter(function (r) { var mine = String(r[1]) === email && r[0] !== keepH; mine && gone.push('ses_' + r[0]); return !mine; });
    if (!gone.length) return;
    s.getRange(2, 1, n, SES_HEAD.length).clearContent();
    if (keep.length) { var rg = s.getRange(2, 1, keep.length, SES_HEAD.length); fmtText_(rg); rg.setValues(keep); }
    CacheService.getScriptCache().removeAll(gone);
  });
}

// ── e-mail codes (register / forgot) — 6 digits, 15 minutes, 5 tries, rate-limited ──
var OTP_MIN = 15, OTP_TRIES = 5, OTP_PER_HOUR = 4, OTP_PER_DAY = 60;
function otpKey_(kind, email) { return 'otp_' + kind + '_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, email)).substring(0, 16); }
function otpHash_(code, email) { return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'otp|' + email + '|' + code, Utilities.Charset.UTF_8)); }
function otpSend_(kind, email, name, extra) {
  var C = CacheService.getScriptCache(), hk = 'otpn_' + otpKey_('', email), day = 'otpday_' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd');
  var nh = Number(C.get(hk) || 0), nd = Number(C.get(day) || 0), last = Number(C.get(hk + '_t') || 0);
  if (Date.now() - last < 60000) throw new Error('เพิ่งส่งรหัสไปแล้ว กรุณารอ 1 นาทีก่อนขอใหม่');
  if (nh >= OTP_PER_HOUR) throw new Error('ขอรหัสเกิน ' + OTP_PER_HOUR + ' ครั้งในชั่วโมงนี้ กรุณาลองใหม่ภายหลัง');
  if (nd >= OTP_PER_DAY) throw new Error('ระบบส่งอีเมลรหัสครบโควตาของวันนี้แล้ว กรุณาติดต่อ Admin');
  var code = String(Math.floor(100000 + Math.random() * 900000));
  C.put(otpKey_(kind, email), JSON.stringify({ h: otpHash_(code, email), x: Date.now() + OTP_MIN * 60000, n: 0, extra: extra || null }), OTP_MIN * 60);
  var what = kind === 'reg' ? 'ยืนยันอีเมลเพื่อสมัครใช้งาน' : 'ตั้งรหัสผ่านใหม่';
  MailApp.sendEmail({ to: email, subject: '[DAIHEN Helpdesk] รหัส ' + code + ' สำหรับ' + what, name: 'DAIHEN Online Helpdesk',
    body: 'รหัสสำหรับ' + what + ': ' + code + '\nรหัสใช้ได้ ' + OTP_MIN + ' นาที\nถ้าคุณไม่ได้ขอรหัสนี้ ไม่ต้องทำอะไร',
    htmlBody: '<div style="font-family:Tahoma,Arial,sans-serif;max-width:460px;margin:auto;color:#141413"><div style="background:#0093d8;color:#fff;padding:14px 18px;border-radius:10px 10px 0 0;font-weight:700;letter-spacing:.04em">DAIHEN ELECTRIC · Online Helpdesk</div>' +
      '<div style="border:1px solid #d9e2ea;border-top:0;border-radius:0 0 10px 10px;padding:18px">' + (name ? '<p>เรียน ' + esc_(name) + '</p>' : '') +
      '<p>รหัสสำหรับ' + what + '</p><p style="font-size:30px;letter-spacing:.35em;font-weight:700;margin:10px 0;color:#0b5cab">' + code + '</p>' +
      '<p style="color:#5b6876;font-size:13px">รหัสใช้ได้ ' + OTP_MIN + ' นาที · ถ้าคุณไม่ได้ขอรหัสนี้ ไม่ต้องทำอะไร และอย่าส่งรหัสนี้ให้ผู้อื่น</p></div></div>' });
  C.put(hk, String(nh + 1), 3600); C.put(hk + '_t', String(Date.now()), 3600); C.put(day, String(nd + 1), 86400);
}
function otpCheck_(kind, email, code) {
  var C = CacheService.getScriptCache(), k = otpKey_(kind, email), v = null;
  try { v = JSON.parse(C.get(k) || 'null'); } catch (e) {}
  if (!v || v.x < Date.now()) throw new Error('รหัสหมดอายุหรือยังไม่ได้ขอรหัส — กดขอรหัสใหม่');
  if (v.n >= OTP_TRIES) { C.remove(k); throw new Error('ใส่รหัสผิดเกิน ' + OTP_TRIES + ' ครั้ง — กดขอรหัสใหม่'); }
  if (otpHash_(String(code || '').replace(/\D/g, ''), email) !== v.h) { v.n++; C.put(k, JSON.stringify(v), Math.max(1, Math.round((v.x - Date.now()) / 1000))); Utilities.sleep(500); throw new Error('รหัสไม่ถูกต้อง (เหลือ ' + (OTP_TRIES - v.n) + ' ครั้ง)'); }
  C.remove(k);
  return v.extra;
}

// ── public: before signing in ──
function authInfo() { var c = authCfg_(); return { on: c.on, approve: c.approve, domains: authDomains_(), days: c.days }; }
function authAdminUser_(u) {
  var a = admUsers_().filter(function (x) { return x.u === String(u).toLowerCase(); })[0];
  return a ? { e: 'admin:' + a.u, n: a.n, d: 'Admin', st: 'active', admin: true } : null;
}
// id = e-mail, or an Admin user name (sheet Settings) → { token, expires, user, admin?: {token,user,name,names} } · { first: true } · { err }
function authLogin(id, pass, remember, device, newPass) {
  var raw = String(id || '').trim(), C = CacheService.getScriptCache();
  if (!raw || !pass) return { err: 'กรอกอีเมลและรหัสผ่าน' };
  if (raw.indexOf('@') < 0) {                                    // Admin account
    var r = adminLogin(raw, pass, newPass);
    if (!r || r.err || r.first) return r || { err: 'เข้าสู่ระบบไม่สำเร็จ' };
    var au = authAdminUser_(r.user), s = sesNew_(au, !!remember, device);
    audit_([['User', 'เข้าสู่ระบบ', '', au.n + ' (Admin)']]);
    return { token: s.token, expires: s.expires, user: authPub_(au), admin: r };
  }
  var email = authEmail_(raw); if (!email) return { err: 'รูปแบบอีเมลไม่ถูกต้อง' };
  var fk = 'usrfail_' + email;
  if (Number(C.get(fk) || 0) >= 5) return { err: 'ใส่รหัสผิดเกิน 5 ครั้ง — พักการเข้าสู่ระบบของอีเมลนี้ 15 นาที หรือกด "ลืมรหัสผ่าน"' };
  var u = usrFind_(email);
  var bad = function () { C.put(fk, String(Number(C.get(fk) || 0) + 1), 900); Utilities.sleep(700); return { err: 'อีเมลหรือรหัสผ่านไม่ถูกต้อง' }; };
  if (!u || !u.h || admHash_(u.s, pass) !== u.h) return bad();
  if (u.st === 'pending') return { err: 'บัญชีนี้รอ Admin อนุมัติ — แจ้ง Admin ของระบบเพื่อเปิดใช้งาน' };
  if (u.st !== 'active') return { err: 'บัญชีนี้ถูกระงับการใช้งาน — ติดต่อ Admin' };
  C.remove(fk);
  withLock_(function () { var x = usrFind_(email); if (x) { x.t = nowStr_(); usrWrite_(x); } });
  var ses = sesNew_(u, !!remember, device);
  return { token: ses.token, expires: ses.expires, user: authPub_(u) };
}
function authRegisterStart(email, name, dept, pass) {
  email = authEmail_(email); name = authName_(name); dept = String(dept || '').replace(/[^\w .-]/g, '').substring(0, 20);
  if (!email) return { err: 'รูปแบบอีเมลไม่ถูกต้อง' };
  if (!name) return { err: 'กรอกชื่อที่ใช้ในระบบ เช่น S.Somchai' };
  try { authPassCheck_(pass, email, name); } catch (e) { return { err: e.message }; }
  var cfg = authCfg_(), ex = usrFind_(email);
  if (ex && ex.st !== 'rejected') return { err: ex.st === 'pending' ? 'อีเมลนี้สมัครไว้แล้ว รอ Admin อนุมัติ' : 'อีเมลนี้มีบัญชีอยู่แล้ว — ใช้ "ลืมรหัสผ่าน" ถ้าจำรหัสไม่ได้' };
  if (cfg.domains && !authDomainOk_(email) && !cfg.approve) return { err: 'สมัครได้เฉพาะอีเมล @' + authDomains_().join(', @') };
  var salt = Utilities.getUuid().replace(/-/g, '');
  try { otpSend_('reg', email, name, { n: name, d: dept, s: salt, h: admHash_(salt, pass) }); } catch (e) { return { err: e.message }; }
  return { sent: true, email: email, minutes: OTP_MIN };
}
function authRegisterVerify(email, code, remember, device) {
  email = authEmail_(email); if (!email) return { err: 'รูปแบบอีเมลไม่ถูกต้อง' };
  var x; try { x = otpCheck_('reg', email, code); } catch (e) { return { err: e.message }; }
  var auto = !authCfg_().approve || authDomainOk_(email);
  var u = withLock_(function () {
    var ex = usrFind_(email); if (ex && ex.st !== 'rejected') return null;
    return usrWrite_({ row: ex ? ex.row : 0, e: email, n: x.n, d: x.d, st: auto ? 'active' : 'pending', s: x.s, h: x.h, c: nowStr_(), t: auto ? nowStr_() : '' });
  });
  if (!u) return { err: 'อีเมลนี้มีบัญชีอยู่แล้ว' };
  audit_([['User', 'สมัครใช้งาน', '', u.n + ' <' + u.e + '>' + (auto ? '' : ' — รอ Admin อนุมัติ')]]);
  if (!auto) return { pending: true, name: u.n };
  var s = sesNew_(u, !!remember, device);
  return { token: s.token, expires: s.expires, user: authPub_(u) };
}
// always answers the same way, so the form cannot be used to find out who has an account
function authForgotStart(email) {
  email = authEmail_(email); if (!email) return { err: 'รูปแบบอีเมลไม่ถูกต้อง' };
  var u = usrFind_(email);
  if (u && u.st === 'active') { try { otpSend_('pw', email, u.n); } catch (e) { return { err: e.message }; } }
  else Utilities.sleep(900);
  return { sent: true, email: email, minutes: OTP_MIN };
}
function authForgotReset(email, code, pass, remember, device) {
  email = authEmail_(email); if (!email) return { err: 'รูปแบบอีเมลไม่ถูกต้อง' };
  var u0 = usrFind_(email);
  try { authPassCheck_(pass, email, u0 && u0.n); } catch (e) { return { err: e.message }; }
  try { otpCheck_('pw', email, code); } catch (e) { return { err: e.message }; }
  var u = withLock_(function () { var x = usrFind_(email); if (!x || x.st !== 'active') return null; x.s = Utilities.getUuid().replace(/-/g, ''); x.h = admHash_(x.s, pass); x.t = nowStr_(); return usrWrite_(x); });
  if (!u) return { err: 'ไม่พบบัญชีที่ใช้งานได้ของอีเมลนี้' };
  sesRevoke_(email);
  CacheService.getScriptCache().remove('usrfail_' + email);
  audit_([['User', 'ตั้งรหัสผ่านใหม่ (ลืมรหัสผ่าน)', '', u.n + ' <' + u.e + '>']]);
  var s = sesNew_(u, !!remember, device);
  return { token: s.token, expires: s.expires, user: authPub_(u) };
}
// → { user, cfg } for a valid session · { none: true }
function authMe(tok) { var v = authSession_(tok); return v ? { user: { email: v.e, name: v.n, dept: v.d, admin: !!v.a }, on: authCfg_().on } : { none: true, on: authCfg_().on }; }

// ── signed in ──
function authLogout(tok, all) {
  var v = __user || authSession_(tok); if (!v) return { ok: true };
  if (all) sesRevoke_(v.e); else {
    var h = sesHash_(tok);
    withLock_(function () {
      var s = sesSheet_(), n = s.getLastRow() - 1; if (n < 1) return;
      var rows = s.getRange(2, 1, n, 1).getValues(); for (var i = 0; i < rows.length; i++) if (rows[i][0] === h) { s.deleteRow(i + 2); break; }
    });
    CacheService.getScriptCache().remove('ses_' + h);
  }
  return { ok: true };
}
function authChangePassword(tok, oldPass, newPass) {
  var v = __user || authSession_(tok); if (!v) throw new Error('AUTH_REQUIRED: กรุณาเข้าสู่ระบบ');
  if (v.a) return { err: 'บัญชี Admin เปลี่ยนรหัสที่ ตั้งค่าระบบ → บัญชี Admin' };
  var u = usrFind_(v.e); if (!u || admHash_(u.s, oldPass) !== u.h) { Utilities.sleep(700); return { err: 'รหัสผ่านเดิมไม่ถูกต้อง' }; }
  try { authPassCheck_(newPass, u.e, u.n); } catch (e) { return { err: e.message }; }
  withLock_(function () { var x = usrFind_(v.e); x.s = Utilities.getUuid().replace(/-/g, ''); x.h = admHash_(x.s, newPass); usrWrite_(x); });
  sesRevoke_(v.e, tok);
  audit_([['User', 'เปลี่ยนรหัสผ่าน', '', u.n + ' <' + u.e + '>']]);
  return { ok: true };
}
function authProfile(tok, name, dept) {
  var v = __user || authSession_(tok); if (!v || v.a) throw new Error('AUTH_REQUIRED: กรุณาเข้าสู่ระบบ');
  name = authName_(name); if (!name) return { err: 'กรอกชื่อ' };
  var u = withLock_(function () { var x = usrFind_(v.e); x.n = name; x.d = String(dept || '').replace(/[^\w .-]/g, '').substring(0, 20); return usrWrite_(x); });
  sesRevoke_(v.e, tok); CacheService.getScriptCache().remove('ses_' + sesHash_(tok));
  return { ok: true, user: authPub_(u) };
}

// ── Admin ──
function authUsers(token) {
  requireAdmin_(token);
  return { cfg: authCfg_(), users: usrAll_().map(function (u) { return { email: u.e, name: u.n, dept: u.d, status: u.st, created: u.c, last: u.t }; }) };
}
// act: approve | disable | enable | reject | delete | reset (→ temporary password) | signout
function authUserAction(token, email, act) {
  requireAdmin_(token); email = authEmail_(email); if (!email) throw new Error('อีเมลไม่ถูกต้อง');
  var temp = '';
  var u = withLock_(function () {
    var x = usrFind_(email); if (!x) return null;
    if (act === 'delete') { usrSheet_().deleteRow(x.row); return x; }
    if (act === 'approve' || act === 'enable') x.st = 'active';
    else if (act === 'disable') x.st = 'disabled';
    else if (act === 'reject') x.st = 'rejected';
    else if (act === 'reset') { temp = 'Dh' + Math.random().toString(36).slice(2, 7) + Math.floor(10 + Math.random() * 89); x.s = Utilities.getUuid().replace(/-/g, ''); x.h = admHash_(x.s, temp); }
    else if (act !== 'signout') throw new Error('คำสั่งไม่ถูกต้อง');
    return usrWrite_(x);
  });
  if (!u) throw new Error('ไม่พบผู้ใช้ ' + email);
  if (act !== 'approve' && act !== 'enable') sesRevoke_(email);
  audit_([['Admin', { approve: 'อนุมัติผู้ใช้', disable: 'ระงับผู้ใช้', enable: 'เปิดใช้ผู้ใช้', reject: 'ปฏิเสธผู้สมัคร', delete: 'ลบผู้ใช้', reset: 'รีเซ็ตรหัสผ่านผู้ใช้', signout: 'ออกจากระบบทุกเครื่อง (ผู้ใช้)' }[act], '', u.n + ' <' + u.e + '>']]);
  var out = authUsers(token); if (temp) out.temp = temp; return out;
}
function authSaveCfg(token, c) {
  requireAdmin_(token); c = c || {};
  var v = { on: c.on !== false, domains: String(c.domains || '').toLowerCase().replace(/[^a-z0-9.,; @-]/g, '').substring(0, 300), approve: c.approve !== false, days: Math.min(90, Math.max(1, Number(c.days) || 30)) };
  schedProp_('AUTH_CFG', JSON.stringify(v));
  audit_([['Admin', 'ตั้งค่าการเข้าสู่ระบบ', '', JSON.stringify(v)]]);
  return authUsers(token);
}
