/**
 * ═════════════════════════════════════════════════════════════════════
 *  INSIGHTS — สถิติการใช้งาน (ไม่ระบุตัวตน) และข้อเสนอแนะจากผู้ใช้
 *  • logUsage(device, rows)   เบราว์เซอร์รวมยอดทุก 1 นาที: เปิดหน้า (view), เวลาเรียก Server (call),
 *                             ผิดพลาด (err / jserr) → ชีต UsageLog ใน Helpdesk_ItemStock
 *                             device = รหัสสุ่มของเบราว์เซอร์ ไม่ใช่ชื่อคน
 *  • getUsageSummary(token, days)   Admin: หน้าที่ใช้บ่อย, อุปกรณ์ที่ใช้งาน, เวลาโหลด, ข้อผิดพลาด
 *  • submitFeedback(fb)       ปุ่ม 💬 ข้อเสนอแนะ (ทุกคน) → ชีต Feedback · getFeedback / setFeedbackStatus (Admin)
 *  ต้องมี Code.gs, OrderService.gs (helper ร่วมกัน)
 * ═════════════════════════════════════════════════════════════════════
 */
var USE_SHEET = 'UsageLog', USE_HEADER = ['DATE', 'TIME', 'DEVICE', 'KIND', 'NAME', 'COUNT', 'AVG_MS', 'MAX_MS'], USE_MAX = 60000, USE_KEEP = 45000;
var USE_KINDS = { view: 1, call: 1, err: 1, jserr: 1 };
var FB_SHEET = 'Feedback', FB_HEADER = ['ID', 'TIME', 'TYPE', 'MESSAGE', 'NAME', 'PAGE', 'DEVICE', 'STATUS', 'NOTE', 'BROWSER'];
var FB_TYPES = { bug: 'ปัญหา', idea: 'ไอเดีย', praise: 'ชม', other: 'อื่น ๆ' }, FB_STATUS = { 'new': 'ใหม่', ack: 'รับทราบ', done: 'ทำแล้ว', no: 'ไม่ทำ' };

function useDev_(d) { var s = String(d || '').replace(/[^A-Za-z0-9]/g, '').substring(0, 16); return s || 'unknown'; }

// ── usage (public, best effort: a busy lock drops this batch instead of waiting) ──
function logUsage(device, rows) {
  var dev = useDev_(device), now = new Date(), day = Utilities.formatDate(now, 'Asia/Bangkok', 'dd/MM/yyyy'), tm = Utilities.formatDate(now, 'Asia/Bangkok', 'HH:mm');
  var add = (Array.isArray(rows) ? rows : []).slice(0, 200).map(function (r) {
    r = r || {};
    var k = String(r.k || ''), n = String(r.n || '').replace(/[^\w\-:.ก-๙ ]/g, '').substring(0, 60), c = Math.max(1, Math.min(100000, Math.round(Number(r.c) || 1)));
    if (!USE_KINDS[k] || !n) return null;
    return [day, tm, dev, k, n, String(c), String(Math.max(0, Math.min(6e5, Math.round(Number(r.ms) || 0)))), String(Math.max(0, Math.min(6e5, Math.round(Number(r.mx) || 0))))];
  }).filter(Boolean);
  if (!add.length) return 0;
  var L = LockService.getScriptLock();
  if (!L.tryLock(3000)) return -1;
  try {
    var ss = stockSS_(), sh = ss.getSheetByName(USE_SHEET);
    if (!sh) { sh = ss.insertSheet(USE_SHEET); sh.getRange(1, 1, 1, USE_HEADER.length).setValues([USE_HEADER]).setFontWeight('bold'); sh.setFrozenRows(1); }
    var at = sh.getLastRow() + 1;
    if (sh.getMaxRows() < at + add.length) sh.insertRowsAfter(sh.getMaxRows(), at + add.length - sh.getMaxRows());
    var rg = sh.getRange(at, 1, add.length, USE_HEADER.length); rg.setNumberFormat('@'); rg.setValues(add);
    var n = sh.getLastRow() - 1; if (n > USE_MAX) sh.deleteRows(2, n - USE_KEEP);
  } finally { L.releaseLock(); }
  return add.length;
}
function getUsageSummary(token, days) {
  requireAdmin_(token);
  days = [7, 14, 30, 90].indexOf(Number(days)) >= 0 ? Number(days) : 30;
  var key = 'use_' + days + '_' + Math.floor(Date.now() / 6e5), hit = ordCacheGet_(key);
  if (hit) return hit;
  var out = { days: days, pages: [], calls: [], errors: [], daily: [], devices: 0, views: 0, rows: 0, since: '' };
  var sh = stockSS_().getSheetByName(USE_SHEET), n = sh ? sh.getLastRow() - 1 : 0;
  if (n > 0) {
    var today = wsToday_(), cut = today - (days - 1) * 864e5, dev = {}, pg = {}, cl = {}, er = {}, dl = {};
    var rows = sh.getRange(2, 1, n, USE_HEADER.length).getDisplayValues();
    out.since = rows[0][0];
    rows.forEach(function (r) {
      var t = bomDate_(r[0]); if (t == null || t < cut) return;
      out.rows++;
      var c = Number(r[5]) || 1, ms = Number(r[6]) || 0, mx = Number(r[7]) || 0;
      dev[r[2]] = 1; (dl[t] = dl[t] || {})[r[2]] = 1;
      if (r[3] === 'view') { var P = pg[r[4]] || (pg[r[4]] = { name: r[4], views: 0, dev: {} }); P.views += c; P.dev[r[2]] = 1; out.views += c; }
      else if (r[3] === 'call') { var C = cl[r[4]] || (cl[r[4]] = { name: r[4], count: 0, ms: 0, max: 0 }); C.count += c; C.ms += ms * c; C.max = Math.max(C.max, mx); }
      else { var E = er[r[3] + '|' + r[4]] || (er[r[3] + '|' + r[4]] = { kind: r[3], name: r[4], count: 0, last: '' }); E.count += c; E.last = r[0] + ' ' + r[1]; }
    });
    out.devices = Object.keys(dev).length;
    out.pages = Object.keys(pg).map(function (k) { return { name: k, views: pg[k].views, devices: Object.keys(pg[k].dev).length }; }).sort(function (a, b) { return b.views - a.views; });
    out.calls = Object.keys(cl).map(function (k) { var C = cl[k]; return { name: k, count: C.count, avg: C.count ? Math.round(C.ms / C.count) : 0, max: C.max }; })
      .sort(function (a, b) { return b.avg - a.avg; }).slice(0, 40);
    out.errors = Object.keys(er).map(function (k) { return er[k]; }).sort(function (a, b) { return b.count - a.count; }).slice(0, 30);
    for (var t = cut; t <= today; t += 864e5) out.daily.push([bomFmt_(t), dl[t] ? Object.keys(dl[t]).length : 0]);
  }
  ordCachePut_(key, out, 600);
  return out;
}

// ── feedback ────────────────────────────────────────────────────────
function fbSheet_() {
  var s = ss_().getSheetByName(FB_SHEET);
  if (!s) { s = ss_().insertSheet(FB_SHEET); s.getRange(1, 1, 1, FB_HEADER.length).setValues([FB_HEADER]).setFontWeight('bold'); s.setFrozenRows(1); }
  return s;
}
function submitFeedback(fb) {
  fb = fb || {};
  var type = FB_TYPES[fb.type] ? fb.type : 'other', msg = ordStr_(fb.message, 2000);
  if (msg.length < 3) throw new Error('กรุณาพิมพ์ข้อความอย่างน้อย 3 ตัวอักษร');
  var dev = useDev_(fb.device), c = CacheService.getScriptCache(), rk = 'fb_' + dev, cnt = Number(c.get(rk) || 0);
  if (cnt >= 10) throw new Error('ส่งข้อเสนอแนะถี่เกินไป กรุณารอสักครู่');
  c.put(rk, String(cnt + 1), 3600);
  var id = 'FB-' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyMMdd-HHmmss') + '-' + dev.substring(0, 3);
  withLock_(function () {
    var s = fbSheet_(), r = s.getLastRow() + 1;
    var rg = s.getRange(r, 1, 1, FB_HEADER.length); rg.setNumberFormat('@');
    rg.setValues([[id, nowStr_(), type, msg, ordStr_(fb.name, 60), ordStr_(fb.page, 60), dev, 'new', '', ordStr_(fb.ua, 200)]]);
  });
  dataBump_('FB_VER');
  return { id: id };
}
function fbAll_() {
  var s = ss_().getSheetByName(FB_SHEET), n = s ? s.getLastRow() - 1 : 0;
  if (n < 1) return [];
  return s.getRange(2, 1, n, FB_HEADER.length).getDisplayValues().map(function (r, i) {
    return { row: i + 2, id: r[0], time: r[1], type: r[2], message: r[3], name: r[4], page: r[5], device: r[6], status: r[7] || 'new', note: r[8] };
  }).filter(function (x) { return x.id; });
}
function getFeedback(token) {
  requireAdmin_(token);
  var L = fbAll_().reverse();
  return { items: L.slice(0, 300), counts: L.reduce(function (m, x) { m[x.status] = (m[x.status] || 0) + 1; return m; }, {}) };
}
function setFeedbackStatus(token, id, status, note) {
  requireAdmin_(token);
  if (!FB_STATUS[status]) throw new Error('สถานะไม่ถูกต้อง');
  withLock_(function () {
    var x = fbAll_().filter(function (f) { return f.id === id; })[0];
    if (!x) throw new Error('ไม่พบข้อเสนอแนะ ' + id);
    var rg = fbSheet_().getRange(x.row, 8, 1, 2); rg.setNumberFormat('@'); rg.setValues([[status, note != null ? ordStr_(note, 300) : x.note]]);
  });
  dataBump_('FB_VER');
  return getFeedback(token);
}
function fbNewCount_() { try { return fbAll_().filter(function (x) { return x.status === 'new'; }).length; } catch (e) { return 0; } }
