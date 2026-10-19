/**
 * ══════════════════════════════════════════════════════════════════════
 *  ORDER ITEM — ใบสั่งสินค้า (FormOrder) สำหรับ Stock
 *  • DatabaseItemOrder : ฐานข้อมูลสินค้าที่สั่งได้ (Admin เพิ่ม/แก้ไขได้)
 *  • OrderItem          : บันทึกใบสั่งทุกใบ (แท็บในสเปรดชีตหลัก)
 *  • Export             : คัดลอกแม่แบบ FormOrder (Google Sheet) → เติมข้อมูล → PDF / Excel
 *  • Stock / PendingOrder ใช้ข้อมูลที่ซิงค์ไว้แล้วในโมดูล Item stock / Pending Order
 *  ต้องมีไฟล์ OrderTemplate.gs (แม่แบบ .xlsx แบบ base64) อยู่ในโปรเจกต์เดียวกัน
 * ══════════════════════════════════════════════════════════════════════
 */
var ORD_DB_SHEET = 'DatabaseItemOrder';
var ORD_DB_DEFAULT_HEADER = ['CODE', 'MATERIAL', 'TYPE', 'SPEC', 'UNIT', 'MANUFACTURER'];
var ORD_SHEET = 'OrderItem';
var ORD_HEADER = ['Order ID', 'RQN No', 'Spec No', 'Issue Date', "DES'D", "CHK'D", "APP'D", 'Revisions',
  'Manufacturer', 'Supplier', 'Delivery Date', 'Place', 'Remark', 'Items', 'Note', 'Images',
  'Status', 'Created', 'Updated', 'PDF', 'XLSX', 'Item Codes', 'Approval'];
var ORD_NCOL = ORD_HEADER.length;   // 23 (Approval = JSON, see OrderApproval.gs)
var ORD_MAX_ITEMS = 8, ORD_MAX_REVS = 3, ORD_MAX_IMAGES = 6;
var ORD_PLACE_DEFAULT = "DELCO'S FACTORY";
// cell map of the template (sheet "Order")
var ORD_ITEM_ROW0 = 21, ORD_NOTE_ROW = 38, ORD_IMG_ROW0 = 41, ORD_IMG_ROW1 = 65, ORD_PRINT_ROWS = 67, ORD_PRINT_COLS = 25;

// ── helpers ─────────────────────────────────────────────────────────
function ordStr_(v, max) { return String(v == null ? '' : v).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '').trim().substring(0, max || 300); }
function ordYY_(dmy) { var m = /(\d{4})\s*$/.exec(String(dmy || '')); return String(m ? +m[1] : new Date().getFullYear()).slice(-2); }
function ordToday_() { return Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy'); }
function ordNorm_(h) { return String(h || '').toUpperCase().replace(/[^A-Z0-9฀-๿]/g, ''); }
function ordCodeKey_(c) { return String(c || '').toUpperCase().replace(/\s+/g, ''); }
function ordFolder_() {
  var id = schedProp_('ORDER_FOLDER_ID');
  if (id) { try { var f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {} }
  var f2 = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID).createFolder('FormOrder');
  schedProp_('ORDER_FOLDER_ID', f2.getId());
  return f2;
}

// ── DatabaseItemOrder ───────────────────────────────────────────────
// Finds the tab in the main spreadsheet (or Helpdesk_ItemStock); creates it with a default header if missing.
function ordDbSheet_() {
  var books = [ss_()]; try { if (schedProp_('STOCK_SS_ID')) books.push(stockSS_()); } catch (e) {}
  for (var b = 0; b < books.length; b++) {
    var sh = books[b].getSheetByName(ORD_DB_SHEET);
    if (!sh) sh = books[b].getSheets().filter(function (s) { return ordNorm_(s.getName()) === ordNorm_(ORD_DB_SHEET); })[0];
    if (sh) {
      if (sh.getLastRow() < 1 || !sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getDisplayValues()[0].join('').trim())
        sh.getRange(1, 1, 1, ORD_DB_DEFAULT_HEADER.length).setValues([ORD_DB_DEFAULT_HEADER]).setFontWeight('bold');
      return sh;
    }
  }
  var s = ss_().insertSheet(ORD_DB_SHEET);
  s.getRange(1, 1, 1, ORD_DB_DEFAULT_HEADER.length).setValues([ORD_DB_DEFAULT_HEADER]).setFontWeight('bold');
  s.setFrozenRows(1);
  return s;
}
// header name → field (works with CODE / Item Code / ITEMNO, MATERIAL / DETAILS / DESC, MANUFAC' / Maker …)
var ORD_ALIAS = {
  code: ['CODE', 'ITEMCODE', 'ITEMNO', 'ITEMID', 'PARTNO', 'รหัส', 'รหัสสินค้า'],
  material: ['MATERIAL', 'DETAILS', 'DETAIL', 'DESC', 'DESCRIPTION', 'ITEMNAME', 'NAME', 'รายละเอียด', 'ชื่อสินค้า'],
  type: ['TYPE', 'MODEL', 'ประเภท'],
  spec: ['SPEC', 'SPECS', 'SPECIFICATION', 'สเปค'],
  unit: ['UNIT', 'UNITS', 'STOCKUNIT', 'UOM', 'หน่วย'],
  manufac: ['MANUFAC', 'MANUFACTURER', 'MANUFACT', 'MAKER', 'BRAND', 'MFR', 'MFG', 'ผู้ผลิต', 'ยี่ห้อ'],
  supplier: ['SUPPLIER', 'VENDOR', 'ผู้จัดหา']
};
function ordDbMap_(header) {
  var map = {}, H = header.map(ordNorm_);
  Object.keys(ORD_ALIAS).forEach(function (k) {
    var al = ORD_ALIAS[k], i;
    for (var a = 0; a < al.length && map[k] == null; a++) { i = H.indexOf(al[a]); if (i >= 0) map[k] = i; }
    if (map[k] == null && k === 'manufac') { i = H.findIndex(function (h) { return h.indexOf('MANUFAC') === 0; }); if (i >= 0) map[k] = i; }
  });
  if (map.code == null) map.code = 0;
  return map;
}
function ordDbRead_() {
  var sh = ordDbSheet_(), lc = Math.max(1, sh.getLastColumn()), lr = sh.getLastRow();
  var header = sh.getRange(1, 1, 1, lc).getDisplayValues()[0].map(function (h) { return String(h).trim(); });
  while (header.length > 1 && !header[header.length - 1]) header.pop();
  var rows = lr > 1 ? sh.getRange(2, 1, lr - 1, header.length).getDisplayValues() : [];
  var out = [];
  rows.forEach(function (r, i) { if (r.join('').trim()) out.push({ row: i + 2, v: r.map(function (x) { return String(x).trim(); }) }); });
  return { sheet: sh, header: header, rows: out, map: ordDbMap_(header) };
}
// public: whole item database (small table — sent once)
// small gzip+chunked JSON cache (CacheService values are limited to ~100KB each)
function ordCachePut_(key, obj, ttl) {
  try {
    var b64 = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(JSON.stringify(obj), 'application/json')).getBytes());
    var put = {}, n = Math.ceil(b64.length / 90000);
    for (var i = 0; i < n; i++) put[key + '_' + i] = b64.substring(i * 90000, (i + 1) * 90000);
    put[key + '_n'] = String(n);
    CacheService.getScriptCache().putAll(put, ttl || 21600);
  } catch (e) {}
}
function ordCacheGet_(key) {
  try {
    var c = CacheService.getScriptCache(), n = Number(c.get(key + '_n') || 0); if (!n) return null;
    var keys = []; for (var i = 0; i < n; i++) keys.push(key + '_' + i);
    var got = c.getAll(keys), parts = [];
    for (var j = 0; j < n; j++) { if (got[keys[j]] == null) return null; parts.push(got[keys[j]]); }
    return JSON.parse(Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(parts.join('')), 'application/x-gzip')).getDataAsString());
  } catch (e) { return null; }
}
function ordDbKey_() { return 'orddb_' + (schedProp_('ORD_DB_VER') || '0'); }
function ordDbBump_() { schedProp_('ORD_DB_VER', String(Date.now())); }
// public: whole item database. Cached 10 min (admin edits refresh it at once; force = read the sheet now)
function getOrderDb(force) {
  var key = ordDbKey_();
  if (!force) { var hit = ordCacheGet_(key); if (hit) return hit; }
  var d = ordDbRead_(), sh = d.sheet;
  var out = { header: d.header, rows: d.rows, map: d.map, sheetName: sh.getName(), book: sh.getParent().getName(),
    url: sh.getParent().getUrl() + '#gid=' + sh.getSheetId(), ver: key + '_' + Date.now() };
  ordCachePut_(key, out, 600);
  return out;
}
// admin: add (row = 0) or edit one row. expectCode protects against editing a row that moved.
function saveOrderDbRow(token, row, values, expectCode) {
  requireAdmin_(token);
  var isNew = !Number(row);
  return withLock_(function () {
    var d = ordDbRead_(), sh = d.sheet, n = d.header.length, ci = d.map.code;
    var v = []; for (var c = 0; c < n; c++) v.push(ordStr_((values || [])[c], 500));
    if (!v[ci]) throw new Error('ต้องกรอก ' + (d.header[ci] || 'CODE'));
    var key = ordCodeKey_(v[ci]);
    var dup = d.rows.filter(function (x) { return ordCodeKey_(x.v[ci]) === key && x.row !== Number(row); })[0];
    if (dup) throw new Error('มี CODE "' + v[ci] + '" อยู่แล้วในแถวที่ ' + dup.row);
    row = Number(row) || 0;
    if (row) {
      var cur = d.rows.filter(function (x) { return x.row === row; })[0];
      if (!cur || ordCodeKey_(cur.v[ci]) !== ordCodeKey_(expectCode)) throw new Error('ข้อมูลในชีตเปลี่ยนไปแล้ว กรุณาโหลดใหม่แล้วลองอีกครั้ง');
    } else row = sh.getLastRow() + 1;
    writeText_(sh, row, 1, v);
    audit_([['Admin', isNew ? 'เพิ่มสินค้าใน DatabaseItemOrder' : 'แก้ไข DatabaseItemOrder', v[ci], v.join(' | ')]]);
    ordDbBump_();
    return getOrderDb(true);
  });
}
function deleteOrderDbRow(token, row, expectCode) {
  requireAdmin_(token);
  return withLock_(function () {
    var d = ordDbRead_(), ci = d.map.code;
    var cur = d.rows.filter(function (x) { return x.row === Number(row); })[0];
    if (!cur || ordCodeKey_(cur.v[ci]) !== ordCodeKey_(expectCode)) throw new Error('ข้อมูลในชีตเปลี่ยนไปแล้ว กรุณาโหลดใหม่แล้วลองอีกครั้ง');
    d.sheet.deleteRow(cur.row);
    audit_([['Admin', 'ลบสินค้าจาก DatabaseItemOrder', cur.v[ci], cur.v.join(' | ')]]);
    ordDbBump_();
    return getOrderDb(true);
  });
}

// ── Stock + PendingOrder lookup for the codes on the form ───────────
var ORD_PO_DONE = /receiv|complete|close|finish|cancel|รับแล้ว|รับครบ|ครบ|ปิด|ยกเลิก/i;
function orderItemInfo(codes) {
  var want = {}, out = {};
  (codes || []).slice(0, 20).forEach(function (c) { var k = ordCodeKey_(c); if (k) { want[k] = 1; out[k] = { stock: [], total: null, unit: '', pending: [], pendQty: 0 }; } });
  if (!Object.keys(want).length) return out;
  try {                                                     // Item stock (cached packed rows)
    var rows = stockData_();
    for (var i = 0; i < rows.length; i++) {
      var p = rows[i], t = p.indexOf('\t'); if (t < 0) continue;
      var k = ordCodeKey_(p.substring(0, t)); if (!want[k]) continue;
      var f = p.split('\t'), q = f[3] === '' ? null : Number(f[3]);
      out[k].stock.push({ loc: f[5] || '', qty: q, unit: f[4] || '' });
      if (q != null && !isNaN(q)) out[k].total = (out[k].total || 0) + q;
      if (!out[k].unit && f[4]) out[k].unit = f[4];
    }
  } catch (e) {}
  try {                                                     // Pending Order (PR / PO not yet received) — cached bundle
    var b = pendingBundle_(false), txt = b && b.b64 ? Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(b.b64), 'application/x-gzip')).getDataAsString() : '';
    var now = new Date(), t0 = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime(), YEAR = 365 * 864e5;
    var day = function (v) { var m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(v || ''); return m ? new Date(+m[3], +m[2] - 1, +m[1]).getTime() : null; };
    txt.split('\n').forEach(function (l) {
      var t2 = l.indexOf('\t'); if (t2 < 0) return;
      var k2 = ordCodeKey_(l.substring(0, t2)); if (!want[k2]) return;
      var f = l.split('\t'); if (ORD_PO_DONE.test(f[4] || '')) return;
      var d = day(f[9]) || day(f[7]) || day(f[6]);             // PO_Delivery → EXPARRIVAL → Request_Date
      if (d != null && Math.abs(d - t0) > YEAR) { out[k2].stale = (out[k2].stale || 0) + 1; return; }   // > 1 ปี = ข้อมูลค้างจากต้นทาง
      out[k2].pending.push({ pr: f[1], po: f[2], spec: f[3], status: f[4], name: f[5], req: f[6], due: f[9] || f[7], by: f[12], qty: f[13], unit: f[14], dist: d == null ? 9e15 : Math.abs(d - t0) });
      out[k2].pendQty += Number(f[13]) || 0;
    });
    Object.keys(out).forEach(function (k3) { out[k3].pending.sort(function (x, y) { return x.dist - y.dist; }); });
  } catch (e) {}
  return out;
}

// ── Duplicate-order check (Stock Alert → สั่ง / บันทึกใบสั่งใหม่) ─────────────
// For each CODE, look for an order that may already be on its way:
//   pending   : open PR / PO in Pending Order (official system — PH)
//   approving : Order Item waiting for CHK / APP
//   approved  : Order Item approved (or saved / exported without the approval step) but its RQN / Spec
//               is not found in Pending Order yet → not recorded in the official system → check with the requester
// excludeId = the order being edited. Returns only the codes that have a hit.
// Admin-adjustable (Script property ORD_DUP_CFG): days = Order Item older than this is ignored (0 = no limit),
// rejected = also warn about orders that were rejected.
var ORD_DUP_DEF = { days: 120, rejected: false };
function ordDupCfg_() {
  var c = {}; try { c = JSON.parse(schedProp_('ORD_DUP_CFG') || '{}') || {}; } catch (e) {}
  var d = c.days == null || c.days === '' ? NaN : Number(c.days);
  return { days: isFinite(d) && d >= 0 ? Math.min(Math.round(d), 3650) : ORD_DUP_DEF.days, rejected: c.rejected === true };
}
function getOrderDupCfg(token) { requireAdmin_(token); return ordDupCfg_(); }
function saveOrderDupCfg(token, cfg) {
  requireAdmin_(token);
  cfg = cfg || {};
  var d = Number(cfg.days);
  if (!isFinite(d) || d < 0 || d > 3650) throw new Error('จำนวนวันต้องอยู่ระหว่าง 0–3650 (0 = ไม่จำกัด)');
  var c = { days: Math.round(d), rejected: cfg.rejected === true };
  schedProp_('ORD_DUP_CFG', JSON.stringify(c));
  audit_([['Admin', 'ตั้งค่าการเช็คสั่งซ้ำ', '', 'ย้อนหลัง ' + (c.days || 'ไม่จำกัด') + ' วัน · รวมใบที่ไม่อนุมัติ: ' + (c.rejected ? 'ใช่' : 'ไม่')]]);
  return c;
}
function ordRefKey_(v) { return String(v || '').toUpperCase().replace(/\s+/g, ''); }
function ordPoRefs_() {                                   // every RQN / Spec already in Pending Order (any status)
  var refs = {};
  try {
    var b = pendingBundle_(false), txt = b && b.b64 ? Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(b.b64), 'application/x-gzip')).getDataAsString() : '';
    txt.split('\n').forEach(function (l) { var f = l.split('\t'); if (f[1]) refs[ordRefKey_(f[1])] = 1; if (f[3]) refs[ordRefKey_(f[3])] = 1; });
  } catch (e) {}
  return refs;
}
function ordAgeDays_(dmy) {
  var m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(String(dmy || ''));
  return m ? (Date.now() - new Date(+m[3], +m[2] - 1, +m[1]).getTime()) / 864e5 : 0;
}
function checkOrderDup(codes, excludeId) {
  var want = {}, out = {};
  (codes || []).slice(0, 30).forEach(function (c) { var k = ordCodeKey_(c); if (k) want[k] = String(c); });
  var keys = Object.keys(want); if (!keys.length) return out;
  var hit = function (k) { return out[k] || (out[k] = { code: want[k], pending: [], orders: [] }); };
  var info = orderItemInfo(keys);
  keys.forEach(function (k) {
    var p = info[k] && info[k].pending || [];
    if (p.length) hit(k).pending = p.slice(0, 5).map(function (x) {
      return { pr: x.pr, po: x.po, spec: x.spec, status: x.status, due: x.due, by: x.by, qty: x.qty, unit: x.unit };
    });
  });
  try {                                                   // Stock ช่วยใช้: old / substitutable stock for these CODEs
    var rs = rsForOrder_(keys);
    Object.keys(rs).forEach(function (k) { hit(k).remain = rs[k]; });
  } catch (e) { console.warn(e); }
  var refs = null, cfg = ordDupCfg_();
  ordAll_().forEach(function (o) {
    if (o.id === excludeId) return;
    var a = o.approval || {}, st = a.state || '';
    if (st === 'rejected' && !cfg.rejected) return;
    var mine = (o.items || []).filter(function (i) { return want[ordCodeKey_(i.code)]; });
    if (!mine.length) return;
    var kind = '';
    if (st === 'chk' || st === 'app') kind = 'approving';
    else {
      if (cfg.days && ordAgeDays_(o.issueDate) > cfg.days) return;
      if (refs === null) refs = ordPoRefs_();
      var rq = /\d$/.test(o.rqn) ? ordRefKey_(o.rqn) : '', sp = /\d$/.test(o.spec) ? ordRefKey_(o.spec) : '';
      if ((rq && refs[rq]) || (sp && refs[sp])) return;        // already recorded officially → covered by "pending"
      kind = st === 'approved' ? 'approved' : st === 'rejected' ? 'rejected' : 'saved';
    }
    mine.forEach(function (i) {
      hit(ordCodeKey_(i.code)).orders.push({ id: o.id, spec: o.spec, rqn: o.rqn, des: o.des, issueDate: o.issueDate, qty: i.qty, unit: i.unit,
        kind: kind, state: OA_STATE_TH[st] || (o.status === 'Exported' ? 'ส่งออกแล้ว' : 'บันทึกแล้ว'),
        at: st === 'approved' && a.app ? a.app.at : st ? a.sent : o.updated });
    });
  });
  return out;
}

// ── Orders (sheet OrderItem) ────────────────────────────────────────
function ordSheet_() {
  var s = ss_().getSheetByName(ORD_SHEET);
  if (!s) { s = ss_().insertSheet(ORD_SHEET); s.getRange(1, 1, 1, ORD_NCOL).setValues([ORD_HEADER]).setFontWeight('bold'); s.setFrozenRows(1); }
  if (s.getMaxColumns() < ORD_NCOL) s.insertColumnsAfter(s.getMaxColumns(), ORD_NCOL - s.getMaxColumns());
  if (!s.getRange(1, ORD_NCOL).getValue()) s.getRange(1, ORD_NCOL).setValue(ORD_HEADER[ORD_NCOL - 1]).setFontWeight('bold');
  return s;
}
function ordRowToObj_(r, row) {
  var arr = function (v, d) { try { var x = JSON.parse(v); return Array.isArray(x) ? x : d; } catch (e) { return d; } };
  return {
    row: row, id: r[0], rqn: r[1], spec: r[2], issueDate: r[3], des: r[4], chk: r[5], app: r[6], revs: arr(r[7], []),
    manufacturer: r[8], supplier: r[9], delivery: r[10], place: r[11], remark: r[12], items: arr(r[13], []), note: r[14],
    images: arr(r[15], []), status: r[16] || 'Saved', created: r[17], updated: r[18], pdf: r[19], xlsx: r[20],
    approval: (function () { try { var a = JSON.parse(r[22] || 'null'); return a && typeof a === 'object' ? a : null; } catch (e) { return null; } })()
  };
}
function ordObjToRow_(o) {
  return [o.id, o.rqn, o.spec, o.issueDate, o.des, o.chk, o.app, JSON.stringify(o.revs || []), o.manufacturer, o.supplier,
    o.delivery, o.place, o.remark, JSON.stringify(o.items || []), o.note, JSON.stringify(o.images || []), o.status || 'Saved',
    o.created || '', o.updated || '', o.pdf || '', o.xlsx || '', (o.items || []).map(function (i) { return i.code; }).join(', '),
    o.approval ? JSON.stringify(o.approval) : ''];
}
// browser-safe copy: the approval step key and approvers' e-mails never leave the server
function ordPub_(o) {
  if (!o) return o;
  var x = {}; Object.keys(o).forEach(function (k) { x[k] = o[k]; });
  x.approval = oaPublic_(o.approval);
  return x;
}
function ordPubAll_() { return ordAll_().reverse().map(ordPub_); }
// The OrderItem sheet is read at most once per server call: the rows are kept in __ordMemo,
// writes update the memo in place (ordMemoPut_), deleteRow resets it, and withLock_ resets it so
// everything inside a lock works on a fresh read.
var __ordMemo = null;
function ordMemoReset_() { __ordMemo = null; }
function ordMemoPut_(o) {
  if (!__ordMemo) return;
  for (var i = 0; i < __ordMemo.length; i++) if (__ordMemo[i].id === o.id) { __ordMemo[i] = o; return; }
  __ordMemo.push(o);
}
function ordAll_() {
  if (__ordMemo) return __ordMemo.slice();
  var s = ordSheet_(), lr = s.getLastRow();
  if (lr < 2) return (__ordMemo = []).slice();
  __ordMemo = s.getRange(2, 1, lr - 1, ORD_NCOL).getDisplayValues().map(function (r, i) { return ordRowToObj_(r, i + 2); })
    .filter(function (o) { return o.id; });
  return __ordMemo.slice();
}
function ordFind_(id) { return ordAll_().filter(function (o) { return o.id === id; })[0] || null; }

// suggested running numbers for the year (from saved orders + PendingOrder Requisition_No / Specification_No)
function ordNext_(yy) {
  yy = yy || ordYY_(ordToday_());
  var rq = { max: 0, w: 4 }, sp = { max: 0, w: 3 };
  var take = function (o, re, v) { var m = re.exec(String(v || '').toUpperCase().replace(/\s+/g, '')); if (m && m[1] === yy) { var n = +m[2]; if (n > o.max) { o.max = n; o.w = m[2].length; } } };
  var RQ = /^RQN(\d{2})-(\d+)/, SP = /^DS1C(\d{2})-(\d+)/;
  ordAll_().forEach(function (o) { take(rq, RQ, o.rqn); take(sp, SP, o.spec); });
  try {
    var sh = stockSS_().getSheetByName(PO_SHEET);
    if (sh && sh.getLastRow() > 1) sh.getRange(2, 2, sh.getLastRow() - 1, 3).getDisplayValues().forEach(function (r) { take(rq, RQ, r[0]); take(sp, SP, r[2]); });
  } catch (e) {}
  var pad = function (n, w) { n = String(n); while (n.length < w) n = '0' + n; return n; };
  return { yy: yy, rqn: pad(rq.max + 1, rq.w), spec: pad(sp.max + 1, 3).slice(-3) };
}

// public: everything the page needs in one call
function getOrderBoot(token) {
  var t0 = Date.now(), admin = isAdmin_(token), tpl = null;
  if (admin) { try { tpl = ordTplInfo_(); } catch (e) { tpl = { ok: false, msg: e.message }; } }
  var out = { db: getOrderDb(), orders: ordPubAll_(), next: null, today: ordToday_(), admin: admin, tpl: tpl,
    approvers: oaNames_(), approverCfg: admin ? oaCfg_() : null, dupCfg: admin ? ordDupCfg_() : null, apVer: oaVer_() };
  perf_('getOrderBoot', t0);
  return out;
}
function getOrders() { return ordPubAll_(); }
function getOrderNext(yy) { return ordNext_(ordStr_(yy, 2)); }

function ordClean_(o) {
  var d = /^\d{2}\/\d{2}\/\d{4}$/;
  var x = {
    id: ordStr_(o.id, 40), issueDate: d.test(o.issueDate) ? o.issueDate : ordToday_(),
    des: ordStr_(o.des, 60), chk: ordStr_(o.chk, 60), app: ordStr_(o.app, 60),
    manufacturer: ordStr_(o.manufacturer, 120) || '-', supplier: ordStr_(o.supplier, 120) || '-',
    delivery: d.test(o.delivery) ? o.delivery : '', place: ordStr_(o.place, 120) || ORD_PLACE_DEFAULT,
    remark: ordStr_(o.remark, 200), note: ordStr_(o.note, 1500)
  };
  var yy = ordYY_(x.issueDate);
  var rqn = ordStr_(o.rqnNo, 12).replace(/[^0-9A-Za-z]/g, ''), spec = ordStr_(o.specNo, 3).replace(/\D/g, '');
  x.rqn = 'RQN' + yy + '-' + rqn; x.spec = 'DS1C' + yy + '-' + spec;
  x.rqnNo = rqn; x.specNo = spec;
  x.revs = (o.revs || []).slice(0, ORD_MAX_REVS).map(function (r) { return { rev: ordStr_(r.rev, 6), date: ordStr_(r.date, 12), person: ordStr_(r.person, 40) }; })
    .filter(function (r) { return r.rev || r.date || r.person; });
  x.items = (o.items || []).map(function (i) {
    return { code: ordStr_(i.code, 40), material: ordStr_(i.material, 200), type: ordStr_(i.type, 120), spec: ordStr_(i.spec, 200),
      qty: ordStr_(i.qty, 12).replace(/[^0-9.,]/g, ''), unit: ordStr_(i.unit, 12).toUpperCase(), manufac: ordStr_(i.manufac, 80) };
  }).filter(function (i) { return i.code || i.material; });
  if (!x.items.length) throw new Error('ต้องมีรายการสินค้าอย่างน้อย 1 รายการ');
  if (x.items.length > ORD_MAX_ITEMS) throw new Error('สั่งได้สูงสุด ' + ORD_MAX_ITEMS + ' รายการต่อใบ');
  x.items.forEach(function (i, n) { if (!i.code) throw new Error('รายการที่ ' + (n + 1) + ': ต้องมี CODE'); if (!(Number(i.qty.replace(/,/g, '')) > 0)) throw new Error('รายการที่ ' + (n + 1) + ' (' + i.code + '): จำนวนต้องมากกว่า 0'); });
  if (!x.des) throw new Error("ต้องกรอกชื่อผู้ Order (DES'D)");
  if (spec && spec.length !== 3) throw new Error('เลข Spec ต้องเป็นตัวเลข 3 หลัก (DS1C' + yy + '-___) หรือเว้นว่าง');
  if (!x.delivery) throw new Error('ต้องระบุ Delivery Date (วว/ดด/ปปปป ค.ศ.)');
  var imgs = (o.images || []).slice(0, ORD_MAX_IMAGES);
  imgs.forEach(function (f) { if (f && f.url && f.url.indexOf('data:image/') !== 0 && !/^https:\/\/drive\.google\.com\//.test(f.url)) throw new Error('รูปแนบไม่ถูกต้อง'); });
  x.images = persistFiles_(imgs.map(function (f) { return { name: f.name, type: f.type, size: f.size, url: f.url, thumbnail: f.thumbnail || true }; }))
    .map(function (f, i) { return { name: f.name, url: f.url, thumbnail: f.thumbnail, w: Number(imgs[i].w) || 0, h: Number(imgs[i].h) || 0 }; });
  return x;
}
// public: create / update an order
function saveOrder(o) {
  var x = ordClean_(o || {});
  return withLock_(function () {
    var s = ordSheet_(), all = ordAll_(), now = nowStr_();
    var dup = all.filter(function (y) { return x.specNo && y.spec === x.spec && y.id !== x.id; })[0];
    if (dup) throw new Error('เลข Spec ' + x.spec + ' ถูกใช้แล้วในใบสั่ง ' + dup.id + ' — กรุณาเปลี่ยนเลข');
    var cur = x.id ? all.filter(function (y) { return y.id === x.id; })[0] : null;
    if (x.id && !cur) throw new Error('ไม่พบใบสั่ง ' + x.id + ' (อาจถูกลบไปแล้ว)');
    if (cur) {
      x.created = cur.created; x.status = 'Saved'; x.pdf = ''; x.xlsx = '';
      var ap = cur.approval;                        // any edit voids a pending / given approval → must be re-submitted
      if (ap && (ap.state === 'chk' || ap.state === 'app' || ap.state === 'approved')) {
        ap = JSON.parse(JSON.stringify(ap));
        oaHist_(ap, 'DES', x.des, 'แก้ไขใบสั่ง — ยกเลิกสถานะอนุมัติเดิม ต้องส่งขออนุมัติใหม่', '');
        ap.state = ''; ap.k = '';
        if (ap.chk) ap.chk.at = ''; if (ap.app) ap.app.at = '';
      }
      x.approval = ap || null;
      if (ap) oaBump_();
    }
    else {
      var yy = ordYY_(x.issueDate), mx = 0;
      all.forEach(function (y) { var m = /^ORD(\d{2})-(\d+)$/.exec(y.id); if (m && m[1] === yy) mx = Math.max(mx, +m[2]); });
      x.id = 'ORD' + yy + '-' + ('000' + (mx + 1)).slice(-4); x.created = now; x.status = 'Saved';
    }
    x.updated = now;
    var row = cur ? cur.row : s.getLastRow() + 1;
    writeText_(s, row, 1, ordObjToRow_(x));
    x.row = row; ordMemoPut_(x);
    audit_([['User', cur ? 'แก้ไขใบสั่ง (Order Item)' : 'สร้างใบสั่ง (Order Item)', x.id, x.spec + ' · ' + x.rqn + ' · ' + x.items.length + ' รายการ · ' + x.des]]);
    x.row = row;
    return { order: ordPub_(x), orders: ordPubAll_() };
  });
}
function deleteOrder(token, id) {
  requireAdmin_(token);
  return withLock_(function () {
    var o = ordFind_(id); if (!o) throw new Error('ไม่พบใบสั่ง ' + id);
    ordSheet_().deleteRow(o.row); ordMemoReset_();
    if (o.approval) oaBump_();
    audit_([['Admin', 'ลบใบสั่ง (Order Item)', id, o.spec + ' · ' + o.rqn]]);
    return ordPubAll_();
  });
}

// ── Template (Google Sheet converted from the FormOrder .xlsx) ───────
function ordConvertXlsx_(blob, name) {
  var folder = ordFolder_(), meta = { name: name, mimeType: MimeType.GOOGLE_SHEETS, parents: [folder.getId()] };
  try {                                                     // Advanced Drive service (enabled in appsscript.json)
    if (typeof Drive !== 'undefined' && Drive.Files) {
      if (Drive.Files.create) return Drive.Files.create(meta, blob).id;
      return Drive.Files.insert({ title: name, mimeType: MimeType.GOOGLE_SHEETS, parents: [{ id: folder.getId() }] }, blob, { convert: true }).id;
    }
  } catch (e) { console.warn('Drive advanced service: ' + e); }
  var bd = 'ord' + Utilities.getUuid().replace(/-/g, '');  // fallback: Drive REST multipart upload with conversion
  var head = Utilities.newBlob('--' + bd + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' + JSON.stringify(meta) +
    '\r\n--' + bd + '\r\nContent-Type: ' + MimeType.MICROSOFT_EXCEL + '\r\n\r\n').getBytes();
  var tail = Utilities.newBlob('\r\n--' + bd + '--').getBytes();
  var res = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true', {
    method: 'post', contentType: 'multipart/related; boundary=' + bd, payload: head.concat(blob.getBytes()).concat(tail),
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true
  });
  if (res.getResponseCode() >= 300) throw new Error('แปลงแม่แบบ .xlsx เป็น Google Sheet ไม่สำเร็จ (' + res.getResponseCode() + '): ' + res.getContentText().substring(0, 200) + ' — เปิด Services → Drive API ในโปรเจกต์ Apps Script');
  return JSON.parse(res.getContentText()).id;
}
function ordTplCreate_(blob, why, src) {
  var id = ordConvertXlsx_(blob, 'FormOrder_Template');
  schedProp_('ORDER_TPL_SRC', src || 'upload');
  var old = schedProp_('ORDER_TPL_ID');
  if (old && old !== id) { try { DriveApp.getFileById(old).setTrashed(true); } catch (e) {} }
  schedProp_('ORDER_TPL_ID', id);
  audit_([['Admin', 'ตั้งแม่แบบ FormOrder', '', why]]);
  return id;
}
function ordTplId_() {
  var id = schedProp_('ORDER_TPL_ID');
  var ver = 'default:' + (typeof ORDER_TEMPLATE_VER === 'number' ? ORDER_TEMPLATE_VER : 1), src = schedProp_('ORDER_TPL_SRC') || 'default:1';
  var stale = src.indexOf('default:') === 0 && src !== ver;     // embedded default template was updated → rebuild once
  if (id && !stale) { try { if (!DriveApp.getFileById(id).isTrashed()) return id; } catch (e) {} }
  var b64 = typeof ORDER_TEMPLATE_XLSX_B64 === 'string' ? ORDER_TEMPLATE_XLSX_B64 : '';
  if (!b64) throw new Error('ไม่พบแม่แบบ FormOrder — ตรวจว่ามีไฟล์ OrderTemplate.gs ในโปรเจกต์');
  return ordTplCreate_(Utilities.newBlob(Utilities.base64Decode(b64), MimeType.MICROSOFT_EXCEL, 'FormOrder_Template.xlsx'), 'สร้างจากแม่แบบเริ่มต้นของระบบ (' + ver + ')', ver);
}
function ordTplInfo_() {
  var id = schedProp_('ORDER_TPL_ID'), ok = false;
  if (id) { try { ok = !DriveApp.getFileById(id).isTrashed(); } catch (e) {} }
  return { ok: ok, url: ok ? 'https://docs.google.com/spreadsheets/d/' + id + '/edit' : '', folder: (function () { try { return ordFolder_().getUrl(); } catch (e) { return ''; } })() };
}
function getOrderTemplate(token) { requireAdmin_(token); ordTplId_(); return ordTplInfo_(); }
function resetOrderTemplate(token) {
  requireAdmin_(token); schedProp_('ORDER_TPL_SRC', 'default:reset');
  ordTplId_(); return ordTplInfo_();
}
// admin: replace the template with an uploaded .xlsx (same cell layout as FormOrder_Template.xlsx)
function uploadOrderTemplate(token, dataUrl) {
  requireAdmin_(token);
  var c = String(dataUrl || '').indexOf(',');
  if (c < 0) throw new Error('ไฟล์ไม่ถูกต้อง');
  ordTplCreate_(Utilities.newBlob(Utilities.base64Decode(String(dataUrl).substring(c + 1)), MimeType.MICROSOFT_EXCEL, 'upload.xlsx'), 'อัปโหลดแม่แบบใหม่ (.xlsx)');
  return ordTplInfo_();
}

// ── Export ──────────────────────────────────────────────────────────
function ordSet_(sh, a1, v, opt) {
  var r = sh.getRange(a1); r.setNumberFormat('@').setValue(v == null ? '' : String(v));
  if (opt) { if (opt.size) r.setFontSize(opt.size); if (opt.h) r.setHorizontalAlignment(opt.h); if (opt.wrap) r.setWrap(true); if (opt.v) r.setVerticalAlignment(opt.v); }
  return r;
}
function ordFit_(s, base, steps) {                           // smaller font for long text
  var n = String(s || '').length; for (var i = 0; i < steps.length; i++) if (n > steps[i][0]) return steps[i][1]; return base;
}
function ordFill_(sh, o) {
  var yy = ordYY_(o.issueDate), sp = String(o.specNo || o.spec.split('-').pop() || '').slice(-3);
  ordSet_(sh, 'O4', 'RQN' + yy + '-', { h: 'center' });
  ordSet_(sh, 'Q4', o.rqnNo != null ? o.rqnNo : String(o.rqn || '').split('-').slice(1).join('-'), { h: 'center' });
  ['D', 'S', '1', 'C', yy.charAt(0), yy.charAt(1), '-', sp.charAt(0), sp.charAt(1), sp.charAt(2)].forEach(function (ch, i) {
    ordSet_(sh, String.fromCharCode(79 + i) + '5', ch, { h: 'center' });   // O5 … X5
  });
  ordSet_(sh, 'F6', 'Name.   ' + (((o.items || [])[0] || {}).material || ''));               // Name. = MATERIAL ของรายการที่ 1
  ordSet_(sh, 'U10', o.issueDate, { h: 'center' });
  ordSet_(sh, 'U11', o.des, { h: 'center' }); ordSet_(sh, 'U12', o.chk, { h: 'center' }); ordSet_(sh, 'U13', o.app, { h: 'center' });
  for (var k = 0; k < ORD_MAX_REVS; k++) {
    var rv = (o.revs || [])[k] || {}, r = 11 + k;
    ordSet_(sh, 'F' + r, rv.rev || '', { h: 'center' }); ordSet_(sh, 'I' + r, rv.date || '', { h: 'center' }); ordSet_(sh, 'L' + r, rv.person || '', { h: 'center' });
  }
  if ((o.revs || []).length) ordSet_(sh, 'X7', o.revs[o.revs.length - 1].rev, { h: 'center' });
  ordSet_(sh, 'K15', ': ' + (o.manufacturer || '-') + ' / ' + (o.supplier || '-'));
  ordSet_(sh, 'K16', ': ' + (o.delivery || ''));
  ordSet_(sh, 'K17', ':  ' + (o.place || ORD_PLACE_DEFAULT));
  ordSet_(sh, 'K18', ':  ' + (o.remark || ''));
  for (var i = 0; i < ORD_MAX_ITEMS; i++) {
    var it = (o.items || [])[i], row = ORD_ITEM_ROW0 + i * 2;
    if (!it) { ['B', 'C', 'F', 'K', 'O', 'S', 'V'].forEach(function (c) { ordSet_(sh, c + row, ''); }); continue; }
    var qty = it.qty ? it.qty + (it.unit ? ' ' + it.unit : '') : '';
    var o2 = { h: 'center', v: 'middle', wrap: true };
    ordSet_(sh, 'B' + row, String(i + 1), { h: 'center', v: 'middle', size: 14 });
    ordSet_(sh, 'C' + row, it.code, { h: 'center', v: 'middle', wrap: true, size: ordFit_(it.code, 13, [[14, 10], [10, 12]]) });
    ordSet_(sh, 'F' + row, it.material, { h: 'center', v: 'middle', wrap: true, size: ordFit_(it.material, 13, [[100, 9], [70, 10], [40, 11], [26, 12]]) });
    ordSet_(sh, 'K' + row, it.type, { h: 'center', v: 'middle', wrap: true, size: ordFit_(it.type, 13, [[70, 9], [45, 10], [28, 11], [18, 12]]) });
    ordSet_(sh, 'O' + row, it.spec, { h: 'center', v: 'middle', wrap: true, size: ordFit_(it.spec, 13, [[110, 9], [75, 10], [45, 11], [28, 12]]) });
    ordSet_(sh, 'S' + row, qty, o2).setFontSize(ordFit_(qty, 13, [[14, 11]]));
    ordSet_(sh, 'V' + row, it.manufac, { h: 'center', v: 'middle', wrap: true, size: ordFit_(it.manufac, 12, [[30, 8], [20, 9], [14, 10], [9, 11]]) });
  }
  var hasImg = (o.images || []).length > 0;
  if (!hasImg) { try { sh.getRange('D38:X40').breakApart(); sh.getRange('D38:W' + ORD_IMG_ROW1).merge(); } catch (e) {} }
  ordSet_(sh, 'D38', o.note || '', { h: 'left', v: 'top', wrap: true, size: ordFit_(o.note, 12, [[260, 9], [160, 10]]) });
  if (hasImg) ordImages_(sh, o.images);
  try {                                                     // keep the frames closed whatever the merges did
    sh.getRange('B20:X' + (ORD_IMG_ROW1 + 1)).setBorder(null, true, true, true, null, null);
    sh.getRange('A1:Y' + ORD_PRINT_ROWS).setBorder(true, true, true, true, null, null);
  } catch (e) {}
}
// place pictures in a grid inside B41:X65, keeping aspect ratio
function ordImages_(sh, imgs) {
  var c0 = 2, c1 = 24, r0 = ORD_IMG_ROW0, r1 = ORD_IMG_ROW1, cw = [], rh = [], W = 0, H = 0, pad = 8;
  for (var c = c0; c <= c1; c++) { cw.push(sh.getColumnWidth(c)); W += cw[cw.length - 1]; }
  for (var r = r0; r <= r1; r++) { rh.push(sh.getRowHeight(r)); H += rh[rh.length - 1]; }
  var n = imgs.length, cols = n <= 1 ? 1 : n <= 4 ? 2 : 3, rows = Math.ceil(n / cols);
  var cellW = (W - pad * (cols + 1)) / cols, cellH = (H - pad * (rows + 1)) / rows;
  var at = function (arr, px) { var i = 0; while (i < arr.length - 1 && px >= arr[i]) { px -= arr[i]; i++; } return [i, Math.max(0, Math.floor(px))]; };
  imgs.forEach(function (f, i) {
    try {
      var m = /\/d\/([\w-]+)/.exec(f.url || ''); if (!m) return;
      var blob = DriveApp.getFileById(m[1]).getBlob();
      var iw = f.w || 4, ih = f.h || 3, sc = Math.min(cellW / iw, cellH / ih), w = Math.floor(iw * sc), h = Math.floor(ih * sc);
      var gx = i % cols, gy = Math.floor(i / cols);
      var x = pad + gx * (cellW + pad) + (cellW - w) / 2, y = pad + gy * (cellH + pad) + (cellH - h) / 2;
      var ax = at(cw, x), ay = at(rh, y);
      var img = sh.insertImage(blob, c0 + ax[0], r0 + ay[0], ax[1], ay[1]);
      img.setWidth(w).setHeight(h);
    } catch (e) { console.warn('image ' + i + ': ' + e); }
  });
}
function ordFetchExport_(ssId, gid, fmt) {
  var base = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export?';
  var q = fmt === 'pdf'
    ? 'format=pdf&gid=' + gid + '&size=A4&portrait=true&scale=4&gridlines=false&printtitle=false&sheetnames=false&pagenum=UNDEFINED' +
      '&fzr=false&horizontal_alignment=CENTER&vertical_alignment=TOP&top_margin=0.35&bottom_margin=0.35&left_margin=0.35&right_margin=0.35' +
      '&r1=0&c1=0&r2=' + ORD_PRINT_ROWS + '&c2=' + ORD_PRINT_COLS
    : 'format=xlsx';
  var opt = { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true };
  for (var t = 0; t < 3; t++) {
    var res = UrlFetchApp.fetch(base + q, opt);
    if (res.getResponseCode() === 200) return res.getBlob();
    if (res.getResponseCode() !== 429 && res.getResponseCode() < 500) throw new Error('ส่งออกไฟล์ไม่สำเร็จ (' + res.getResponseCode() + ')');
    Utilities.sleep(1500 * (t + 1));
  }
  throw new Error('Google ปฏิเสธการส่งออกชั่วคราว (ส่งออกถี่เกินไป) กรุณาลองใหม่ในอีกสักครู่');
}
// public: fmt = 'pdf' | 'xlsx' | 'both' → files saved in Drive (FormOrder folder) + base64 for direct download
function exportOrder(id, fmt) {
  var o = ordFind_(id); if (!o) throw new Error('ไม่พบใบสั่ง ' + id + ' — กรุณาบันทึกก่อน');
  fmt = { pdf: 'pdf', xlsx: 'xlsx' }[fmt] || 'both';
  var folder = ordFolder_(), base = ((/\d$/.test(o.spec) ? o.spec : o.id) + (/\d$/.test(o.rqn) ? '_' + o.rqn : '') + (o.revs.length ? '_Rev' + o.revs[o.revs.length - 1].rev : '')).replace(/[\\\/:*?"<>|]+/g, '_');
  var copy = DriveApp.getFileById(ordTplId_()).makeCopy('_tmp_' + base, folder), files = [];
  try {
    var ss = SpreadsheetApp.openById(copy.getId()), sh = ss.getSheetByName('Order') || ss.getSheets()[0];
    ss.getSheets().forEach(function (x) { if (x.getSheetId() !== sh.getSheetId()) ss.deleteSheet(x); });
    sh.setName('Order');
    ordFill_(sh, o);
    SpreadsheetApp.flush();
    (fmt === 'both' ? ['pdf', 'xlsx'] : [fmt]).forEach(function (f) {
      var blob = ordFetchExport_(ss.getId(), sh.getSheetId(), f).setName(base + (f === 'pdf' ? '.pdf' : '.xlsx'));
      var file = folder.createFile(blob);
      try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
      var b = blob.getBytes();
      files.push({ fmt: f, name: blob.getName(), mime: blob.getContentType(), url: file.getUrl(), b64: b.length < 9 * 1024 * 1024 ? Utilities.base64Encode(b) : '' });
    });
  } finally { try { copy.setTrashed(true); } catch (e) {} }
  withLock_(function () {
    var cur = ordFind_(id); if (!cur) return;
    files.forEach(function (f) { cur[f.fmt] = f.url; });
    cur.status = 'Exported'; cur.updated = nowStr_();
    writeText_(ordSheet_(), cur.row, 1, ordObjToRow_(cur)); ordMemoPut_(cur);
  });
  audit_([['User', 'ส่งออกใบสั่ง (' + fmt.toUpperCase() + ')', id, o.spec]]);
  return { files: files, orders: ordPubAll_() };
}

// Run ONCE from the editor (▶ Run) to authorize Drive/export and create the template sheet.
function testOrderSetup() {
  var id = ordTplId_();
  Logger.log('✓ แม่แบบ: https://docs.google.com/spreadsheets/d/' + id + '/edit');
  var d = getOrderDb();
  Logger.log('✓ ' + d.sheetName + ' (' + d.book + '): ' + d.rows.length + ' แถว · คอลัมน์ ' + d.header.join(', ') + ' · map ' + JSON.stringify(d.map));
  Logger.log('✓ เลขถัดไป: ' + JSON.stringify(ordNext_()));
}