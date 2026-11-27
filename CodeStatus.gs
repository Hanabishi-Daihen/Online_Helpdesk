/**
 * ทะเบียน CODE เลิกใช้ — why a CODE is retired, what to use instead; read by every page that shows a CODE.
 *
 * Sheet CodeStatus (main spreadsheet): CODE | LEVEL | REASON | DETAIL | REPLACE | SINCE | DOC | UPDATED | BY
 *   LEVEL  off   = ⛔ เลิกใช้: no new orders, no bookings (Order Item refuses it unless Admin confirms with a reason)
 *          nonew = ⚠ ห้ามสั่งใหม่: use up what is left (a warning only)
 *   REASON one of CS_REASONS · REPLACE = CODE(s) to use instead, comma separated
 * Also retired without a row here (shown as "ยังไม่ระบุสาเหตุ" until Admin adds one):
 *   ACTIVE = 0 in the Official system (Item stock) · ticked เลิกใช้ in Stock Alert.
 * Everyone reads; only Admin writes.
 */
var CS_SHEET = 'CodeStatus', CS_HEAD = ['CODE', 'LEVEL', 'REASON', 'DETAIL', 'REPLACE', 'SINCE', 'DOC', 'UPDATED', 'BY'], CS_CK = 'cs_v1';
var CS_REASONS = [['eol', 'ผู้ผลิตเลิกผลิต'], ['model', 'เปลี่ยนรุ่น / มีรุ่นแทน'], ['quality', 'ปัญหาคุณภาพ / หน้างาน'], ['spec', 'ลูกค้า / สเปกไม่รับ'],
  ['price', 'ราคา / ผู้ขาย'], ['dup', 'CODE ซ้ำ'], ['other', 'อื่น ๆ'], ['official', 'ปิดในระบบ Official — ยังไม่ระบุสาเหตุ'], ['sa', 'ตั้งเลิกใช้ใน Stock Alert — ยังไม่ระบุสาเหตุ']];

function csSheet_() { return sheet_(CS_SHEET, CS_HEAD); }
function csStr_(v, n) { return String(v == null ? '' : v).replace(/[\r\n\t]+/g, ' ').trim().substring(0, n); }
// the rows Admin wrote: { CODE: rec }
function csRows_() {
  var C = CacheService.getScriptCache(), hit = C.get(CS_CK);
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  var sh = csSheet_(), n = sh.getLastRow() - 1, out = {};
  if (n > 0) sh.getRange(2, 1, n, CS_HEAD.length).getDisplayValues().forEach(function (r) {
    var k = ordCodeKey_(r[0]); if (!k) return;
    out[k] = { code: k, level: r[1] === 'nonew' ? 'nonew' : 'off', reason: r[2] || 'other', detail: r[3], replace: r[4], since: r[5], doc: r[6], updated: r[7], by: r[8] };
  });
  try { var s = JSON.stringify(out); if (s.length < 95000) C.put(CS_CK, s, 21600); } catch (e) {}
  return out;
}
// every retired CODE the system knows of, Admin's rows first; withStock = also scan Item stock for ACTIVE = 0
function csAll_(withStock, only) {
  var out = {}, R = csRows_();
  Object.keys(R).forEach(function (k) { if (!only || only[k]) out[k] = R[k]; });
  try {
    var set = saSettings_();
    Object.keys(set).forEach(function (k) { if (set[k].notUse && !out[k] && (!only || only[k])) out[k] = { code: k, level: 'off', reason: 'sa', detail: String(set[k].note || ''), auto: 'sa' }; });
  } catch (e) { errLog_(e, 'CodeStatus:37'); }
  if (withStock) {
    try {
      var on = {}, off = {};
      (only ? stockRowsFor_(only) : stockData_()).forEach(function (p) {   // a few CODEs: CODE index buckets only
        var t = p.indexOf('\t'); if (t < 0) return;
        var k = ordCodeKey_(p.substring(0, t)); if (only && !only[k]) return;
        var f = p.split('\t'); if (f[STOCK_ACT_COL] === '0') off[k] = 1; else if (f[STOCK_ACT_COL] === '1') on[k] = 1;
      });
      Object.keys(off).forEach(function (k) { if (!on[k] && !out[k]) out[k] = { code: k, level: 'off', reason: 'official', detail: '', auto: 'official' }; });
    } catch (e) { console.warn('csAll_ stock', e); }
  }
  return out;
}
// the page: Admin's rows + Stock Alert ticks (ACTIVE = 0 comes with the Item stock bundle in the browser)
function getCodeStatus(token) { apiOnly_(); return { map: csAll_(false), reasons: CS_REASONS, admin: isAdmin_(token) }; }

function saveCodeStatus(token, rec) { apiOnly_();
  requireAdmin_(token);
  rec = rec || {};
  var k = ordCodeKey_(rec.code); if (!k) throw new Error('ใส่ CODE');
  var reason = csStr_(rec.reason, 20); if (!CS_REASONS.some(function (r) { return r[0] === reason; }) || reason === 'official' || reason === 'sa') throw new Error('เลือกเหตุผลที่เลิกใช้');
  var detail = csStr_(rec.detail, 300); if (reason === 'other' && !detail) throw new Error('เหตุผล "อื่น ๆ" ต้องเขียนรายละเอียด');
  var rep = String(rec.replace || '').split(/[,\s]+/).map(ordCodeKey_).filter(function (c) { return c && c !== k; }).slice(0, 5).join(', ');
  var since = csStr_(rec.since, 10); if (since && !/^\d{2}\/\d{2}\/\d{4}$/.test(since)) throw new Error('วันที่มีผลต้องเป็น วว/ดด/ปปปป');
  var doc = csStr_(rec.doc, 400); if (doc && !/^https:\/\//i.test(doc)) throw new Error('ลิงก์เอกสารต้องขึ้นต้นด้วย https://');
  var who = __admName || (typeof __user !== 'undefined' && __user && __user.n) || 'Admin';
  var row = [k, rec.level === 'nonew' ? 'nonew' : 'off', reason, detail, rep, since || nowStr_().substring(0, 10), doc, nowStr_(), who];
  withLock_(function () {
    var sh = csSheet_(), n = sh.getLastRow() - 1, codes = n > 0 ? sh.getRange(2, 1, n, 1).getValues().map(function (r) { return ordCodeKey_(r[0]); }) : [];
    var at = codes.indexOf(k), rg = at >= 0 ? sh.getRange(at + 2, 1, 1, CS_HEAD.length) : sh.getRange(sh.getLastRow() + 1, 1, 1, CS_HEAD.length);
    fmtText_(rg); rg.setValues([row]);
    CacheService.getScriptCache().remove(CS_CK);
  });
  try { saBump_(); } catch (e) {}
  audit_([['Admin', 'บันทึกเหตุผล CODE เลิกใช้', k, (row[1] === 'off' ? 'เลิกใช้' : 'ห้ามสั่งใหม่') + ' · ' + csReasonName_(reason) + (rep ? ' · ใช้แทน ' + rep : '')]]);
  return getCodeStatus(token);
}
function deleteCodeStatus(token, code) { apiOnly_();
  requireAdmin_(token);
  var k = ordCodeKey_(code), gone = false;
  withLock_(function () {
    var sh = csSheet_(), n = sh.getLastRow() - 1, codes = n > 0 ? sh.getRange(2, 1, n, 1).getValues().map(function (r) { return ordCodeKey_(r[0]); }) : [];
    var at = codes.indexOf(k); if (at >= 0) { sh.deleteRow(at + 2); gone = true; }
    CacheService.getScriptCache().remove(CS_CK);
  });
  if (!gone) throw new Error('ไม่พบ ' + k + ' ในทะเบียน');
  try { saBump_(); } catch (e) {}
  audit_([['Admin', 'ลบเหตุผล CODE เลิกใช้', k, '']]);
  return getCodeStatus(token);
}
function csReasonName_(r) { for (var i = 0; i < CS_REASONS.length; i++) if (CS_REASONS[i][0] === r) return CS_REASONS[i][1]; return r; }

// Order Item: CODEs newly added to an order that are ⛔ เลิกใช้ → refused, unless Admin confirms each with a reason (ack)
function csOrderCheck_(items, before, ack, tok) {
  var had = {}; (before || []).forEach(function (i) { had[ordCodeKey_(i && i.code)] = 1; });
  var want = {}; (items || []).forEach(function (i) { var k = ordCodeKey_(i && i.code); if (k && !had[k]) want[k] = 1; });
  if (!Object.keys(want).length) return [];
  var cs = csAll_(true, want), bad = Object.keys(cs).filter(function (k) { return cs[k].level === 'off'; });
  if (!bad.length) return [];
  var admin = false; try { admin = isAdmin_(tok || ''); } catch (e) {}
  var miss = bad.filter(function (k) { return !(admin && ack && csStr_(ack[k], 200)); });
  if (miss.length) throw new Error('CODE เลิกใช้แล้ว: ' + miss.map(function (k) { var r = cs[k]; return k + ' (' + csReasonName_(r.reason) + (r.replace ? ' — ใช้ ' + r.replace + ' แทน' : '') + ')'; }).join(', ') +
    (admin ? ' — Admin ยืนยันพร้อมเหตุผลเพื่อสั่งต่อ' : ' — เปลี่ยนเป็น CODE ที่ใช้แทน หรือให้ Admin ยืนยัน'));
  return bad.map(function (k) { return k + ': ' + csStr_(ack[k], 200); });
}
