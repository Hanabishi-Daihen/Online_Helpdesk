/**
 * ══════════════════════════════════════════════════════════════════════
 *  STOCK ALERT (DS13) — แจ้งเตือนสินค้าใกล้หมด / หมด
 *  • รายการ DS13  = Item stock ที่ "ผู้สั่งล่าสุด" (LastOfRequest_By) = DS13 (+ รายการที่ Admin ตั้งค่าไว้เอง) — เฉพาะ CODE 8 หลัก
 *  • คงเหลือ       = QTYONHAND รวมทุกคลังของ CODE นั้น (Item stock)
 *  • ค้างรับ        = Pending Order ที่ยังไม่ปิด และวันที่ไม่ห่างจากวันนี้เกิน 1 ปี (ใกล้วันนี้ที่สุดก่อน)
 *  • LT            = Readtime เฉลี่ย (วัน) จาก Q_StockRT (OneDrive CSV) — Admin กำหนดทับได้รายตัว
 *  • ตั้งค่า        = แท็บ StockAlertDS13 ในสเปรดชีตหลัก (CODE, MIN, MAX, NOT_USE, LT_DAYS, NOTE, UPDATED, BY)
 *  ระดับเตือน (เมื่อคงเหลือต่ำกว่า Min และค้างรับไม่พอ):  LT > 30 วัน = สูงสุด · LT ≤ 7 วัน = ต่ำ · นอกนั้น = กลาง
 *  ของหมด (คงเหลือ ≤ 0) จะถูกยกระดับขึ้น 1 ขั้น
 *  ต้องมี Code.gs และ OrderService.gs (ใช้ helper ร่วมกัน)
 * ══════════════════════════════════════════════════════════════════════
 */
var SA_SET_SHEET = 'StockAlertDS13';
var SA_SET_HEADER = ['CODE', 'MIN', 'MAX', 'NOT_USE', 'LT_DAYS', 'NOTE', 'UPDATED', 'BY'];
var RT_SHEET = 'StockRT';
var SA_DEF = { by: 'DS13', ltLow: 7, ltHigh: 30 };
var SA_CODE_LEN = 8;                                 // ใช้เฉพาะ CODE 8 หลัก (Stock) — ไม่เอา 16 หลัก (ของเฉพาะงาน)
function saCodeOk_(k) { k = String(k || ''); return k.length === SA_CODE_LEN && !/^SV/i.test(k); }   // SV… = service / non-stock codes → never alerted

function saCfg_() {
  var c = {}; try { c = JSON.parse(schedProp_('SA_CFG') || '{}') || {}; } catch (e) {}
  return { by: String(c.by || SA_DEF.by), ltLow: Number(c.ltLow) || SA_DEF.ltLow, ltHigh: Number(c.ltHigh) || SA_DEF.ltHigh };
}
function saByMatch_(cfgBy) {
  var keys = String(cfgBy || '').split(/[,;]+/).map(function (x) { return x.toUpperCase().replace(/[^A-Z0-9]/g, ''); }).filter(Boolean);
  return function (v) { var k = String(v || '').toUpperCase().replace(/[^A-Z0-9]/g, ''); return keys.indexOf(k) >= 0; };
}
function saSetSheet_() {
  var s = ss_().getSheetByName(SA_SET_SHEET);
  if (!s) { s = ss_().insertSheet(SA_SET_SHEET); s.getRange(1, 1, 1, SA_SET_HEADER.length).setValues([SA_SET_HEADER]).setFontWeight('bold'); s.setFrozenRows(1); }
  return s;
}
function saNum_(v) { var s = String(v == null ? '' : v).replace(/[^0-9.\-]/g, ''); if (s === '' || s === '-' || s === '.') return null; var n = Number(s); return isNaN(n) ? null : n; }
function saSettings_() {
  var s = saSetSheet_(), lr = s.getLastRow(), map = {};
  if (lr < 2) return map;
  s.getRange(2, 1, lr - 1, SA_SET_HEADER.length).getDisplayValues().forEach(function (r, i) {
    var k = ordCodeKey_(r[0]); if (!k) return;
    map[k] = { row: i + 2, code: String(r[0]).trim(), min: saNum_(r[1]), max: saNum_(r[2]), notUse: /^(1|y|yes|true|x|✓|เลิก)/i.test(String(r[3]).trim()),
      lt: saNum_(r[4]), note: r[5], updated: r[6], by: r[7] };
  });
  return map;
}
function saVer_() { return schedProp_('SA_VER') || '0'; }
function saBump_() { schedProp_('SA_VER', String(Date.now())); }

// ── Q_StockRT (lead time) sync from OneDrive ────────────────────────
function rtInfo_() { try { return JSON.parse(schedProp_('RT_SYNC_LAST') || 'null'); } catch (e) { return null; } }
function syncRtRun_(why, force) { return syncMutex_('RT_SYNC_LAST', why, function () { return syncRtInner_(why, force); }); }
function syncRtInner_(why, force) {
  var url = schedProp_('RT_SRC_URL'), info = { time: nowStr_(), why: why }, t0 = Date.now();
  if (!url) throw new Error('ยังไม่ได้ตั้งค่าลิงก์ OneDrive ของ Q_StockRT');
  try {
    var text = downloadShared_(url);
    var hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, text));
    var prev = rtInfo_();
    if (!force && prev && prev.ok && prev.hash === hash) {
      info.ok = true; info.unchanged = true; info.rows = prev.rows; info.hash = hash; info.msg = 'ไฟล์ไม่มีการเปลี่ยนแปลง';
      schedProp_('RT_SYNC_LAST', JSON.stringify(info)); return info;
    }
    var lines = text.split(/\r?\n/); text = null;
    var head = csvLine_(lines[0] || '').map(poKey_), ci = head.indexOf('itemid'), vi = head.findIndex(function (h) { return /readtime|leadtime|days?$/.test(h); });
    if (ci < 0 || vi < 0) throw new Error('หัวคอลัมน์ไม่ตรงรูปแบบ ต้องมี Item_Id และ Readtime_Display');
    var rows = [], bad = 0;
    for (var i = 1; i < lines.length; i++) {
      if (!lines[i]) continue;
      var f = csvLine_(lines[i]), code = String(f[ci] || '').trim(); if (!code || !saCodeOk_(ordCodeKey_(code))) continue;   // เก็บเฉพาะ 8 หลัก
      var m = /-?\d+(\.\d+)?/.exec(String(f[vi] || '')), d = m ? Number(m[0]) : null;
      if (d == null || d <= 0) { bad++; d = ''; }                 // ค่าติดลบ / 0 = ข้อมูลผิดจากต้นทาง → ไม่ใช้
      rows.push([code, d === '' ? '' : String(Math.round(d))]);
    }
    lines = null;
    if (!rows.length) throw new Error('ไฟล์ไม่มีข้อมูล');
    var ss = stockSS_(), sh = ss.getSheetByName(RT_SHEET) || ss.insertSheet(RT_SHEET), all = [['Item_Id', 'LT_Days']].concat(rows);
    sh.clearContents();
    if (sh.getMaxRows() < all.length) sh.insertRowsAfter(sh.getMaxRows(), all.length - sh.getMaxRows());
    var rg = sh.getRange(1, 1, all.length, 2); rg.setNumberFormat('@'); rg.setValues(all);
    sh.getRange(1, 1, 1, 2).setFontWeight('bold'); sh.setFrozenRows(1);
    schedProp_('RT_HASH', hash);
    info.ok = true; info.rows = rows.length; info.hash = hash; info.bad = bad; info.secs = Math.round((Date.now() - t0) / 1000);
    info.msg = 'อัปเดต ' + rows.length.toLocaleString() + ' รายการ (ข้ามค่าติดลบ/ศูนย์ ' + bad + ') ใช้เวลา ' + info.secs + ' วินาที';
    audit_([['System', 'ซิงค์ Q_StockRT จาก OneDrive', '', info.why + ' · ' + info.msg]]);
  } catch (e) {
    info.ok = false; info.msg = e.message || String(e);
    if (info.why !== 'อัตโนมัติ') audit_([['System', 'ซิงค์ Q_StockRT ไม่สำเร็จ', '', info.msg]]);
  }
  schedProp_('RT_SYNC_LAST', JSON.stringify(info));
  return info;
}
function rtAutoSync() { if (schedProp_('RT_SRC_URL')) syncRtRun_('อัตโนมัติ', false); }   // time trigger
function getRtSource(token) {
  requireAdmin_(token);
  var hrs = 0; ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'rtAutoSync') hrs = Number(schedProp_('RT_SYNC_HOURS') || 24); });
  var id = schedProp_('STOCK_SS_ID');
  return { url: schedProp_('RT_SRC_URL') || '', hours: hrs, last: rtInfo_(), sheetUrl: id ? 'https://docs.google.com/spreadsheets/d/' + id + '/edit' : '' };
}
function setRtSource(token, url, hours) {
  requireAdmin_(token);
  url = String(url || '').trim();
  if (url && !/^https:\/\/[^\s]+$/i.test(url)) throw new Error('ลิงก์ต้องขึ้นต้นด้วย https://');
  if (/_layouts\/15\/onedrive\.aspx|\/Forms\/AllItems\.aspx/i.test(url)) throw new Error('นี่คือลิงก์หน้าโฟลเดอร์ OneDrive ไม่ใช่ลิงก์แชร์ไฟล์ — คลิกขวาที่ไฟล์ .csv → Share → Copy link แล้ววางลิงก์นั้นแทน');
  schedProp_('RT_SRC_URL', url || null);
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'rtAutoSync') ScriptApp.deleteTrigger(t); });
  hours = Number(hours) || 0;
  if (url && hours > 0) {
    var tb = ScriptApp.newTrigger('rtAutoSync').timeBased();
    syncTrigger_(tb, hours, 6, 12);
    schedProp_('RT_SYNC_HOURS', String(hours));
  }
  audit_([['Admin', 'ตั้งค่าซิงค์ Q_StockRT', '', (url ? 'ตั้งลิงก์แล้ว' : 'ลบลิงก์') + ' · ' + syncLabel_(hours)]]);
  return getRtSource(token);
}
function syncRtNow(token) { requireAdmin_(token); return syncRtRun_('สั่งโดย Admin', true); }
function testRtSync() {
  if (!schedProp_('RT_SRC_URL')) { Logger.log('✗ ยังไม่ได้บันทึกลิงก์ในหน้า Stock Alert → การ์ด ☁ ซิงค์ Q_StockRT'); return; }
  var r = syncRtRun_('ทดสอบจาก Editor', true); Logger.log((r.ok ? '✓ ' : '✗ ') + r.msg);
}
function rtMap_(want) {
  var out = {};
  try {
    var sh = stockSS_().getSheetByName(RT_SHEET); if (!sh || sh.getLastRow() < 2) return out;
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
      var k = ordCodeKey_(r[0]); if (!want[k]) return;
      var d = Number(r[1]); if (d > 0) out[k] = d;
    });
  } catch (e) {}
  return out;
}

// ── Alert computation ───────────────────────────────────────────────
var SA_LV_ORDER = { high: 0, mid: 1, low: 2, wait: 3, nomin: 4, over: 5, ok: 6, off: 7 };
function saCompute_() {
  var cfg = saCfg_(), isBy = saByMatch_(cfg.by), set = saSettings_(), items = {}, want = {};
  var rows = []; try { rows = stockData_(); } catch (e) {}
  var codes = {};                                          // codes ordered by DS13 (any location)
  for (var i = 0; i < rows.length; i++) {
    var p = rows[i]; if (!p) continue;
    var f = p.split('\t'), kc = ordCodeKey_(f[0]); if (isBy(f[7]) && saCodeOk_(kc)) codes[kc] = 1;
  }
  Object.keys(set).forEach(function (k) { if (saCodeOk_(k)) codes[k] = 1; });   // + codes added in the settings sheet
  for (var j = 0; j < rows.length; j++) {
    var p2 = rows[j], t = p2.indexOf('\t'); if (t < 0) continue;
    var k = ordCodeKey_(p2.substring(0, t)); if (!codes[k]) continue;
    var g = p2.split('\t'), q = saNum_(g[3]);
    var it = items[k] || (items[k] = { code: g[0], desc: g[1], comment: g[2], unit: g[4], qty: 0, locs: [], by: g[7] });
    it.locs.push([g[5] || '-', q == null ? '' : q]); if (q != null) it.qty += q;
    if (!it.unit && g[4]) it.unit = g[4];
    if (isBy(g[7])) it.by = g[7];
  }
  Object.keys(codes).forEach(function (k) { if (!items[k]) items[k] = { code: (set[k] && set[k].code) || k, desc: '', comment: '', unit: '', qty: 0, locs: [], by: '', missing: true }; want[k] = 1; });
  var rt = rtMap_(want);
  // pending (valid ≤ 1 year, nearest first)
  var pend = {};
  try {
    var b = pendingBundle_(false), txt = b && b.b64 ? Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(b.b64), 'application/x-gzip')).getDataAsString() : '';
    var now = new Date(), t0 = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime(), YEAR = 365 * 864e5;
    var day = function (v) { var m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(v || ''); return m ? new Date(+m[3], +m[2] - 1, +m[1]).getTime() : null; };
    txt.split('\n').forEach(function (l) {
      var t2 = l.indexOf('\t'); if (t2 < 0) return;
      var k2 = ordCodeKey_(l.substring(0, t2)); if (!want[k2]) return;
      var f2 = l.split('\t'); if (ORD_PO_DONE.test(f2[4] || '')) return;
      var d = day(f2[9]) || day(f2[7]) || day(f2[6]);
      var P = pend[k2] || (pend[k2] = { qty: 0, list: [], stale: 0 });
      if (d != null && Math.abs(d - t0) > YEAR) { P.stale++; return; }
      P.qty += Number(f2[13]) || 0;
      P.list.push({ pr: f2[1], po: f2[2], due: f2[9] || f2[7], qty: f2[13], unit: f2[14], dist: d == null ? 9e15 : Math.abs(d - t0) });
    });
    Object.keys(pend).forEach(function (k3) { pend[k3].list.sort(function (x, y) { return x.dist - y.dist; }); pend[k3].list = pend[k3].list.slice(0, 5); });
  } catch (e) {}
  var bk = {}; try { bk = bomBooked_(); } catch (e) {}              // Booking BOM: stock already reserved for projects
  var out = [], cnt = { high: 0, mid: 0, low: 0, wait: 0, nomin: 0, over: 0, ok: 0, off: 0, out: 0 };
  Object.keys(items).forEach(function (k) {
    var it = items[k], s = set[k] || {}, P = pend[k] || { qty: 0, list: [], stale: 0 };
    it.qty = Math.round(it.qty * 1000) / 1000;
    it.min = s.min != null ? s.min : null; it.max = s.max != null ? s.max : null; it.notUse = !!s.notUse; it.note = s.note || '';
    it.lt = s.lt != null && s.lt > 0 ? s.lt : (rt[k] || null); it.ltSrc = s.lt != null && s.lt > 0 ? 'set' : rt[k] ? 'rt' : '';
    it.ltRt = rt[k] || null;
    it.pend = P.qty; it.pendList = P.list; it.pendStale = P.stale;
    var B = bk[k]; it.booked = B ? B.qty : 0; it.bookedBy = B ? B.projects : [];
    it.isOut = it.qty <= 0;
    var eff = it.qty - it.booked, avail = eff + it.pend, lv;          // eff = onhand left after BOM bookings
    if (it.notUse) lv = 'off';
    else if (it.min == null) lv = 'nomin';                    // ยังไม่ตั้ง Min → ไม่แจ้งเตือน (แสดงป้าย "หมด" ถ้าคงเหลือ ≤ 0)
    else if (eff < it.min) {
      if (avail >= it.min) lv = 'wait';
      else {
        lv = it.lt == null ? 'mid' : it.lt > cfg.ltHigh ? 'high' : it.lt <= cfg.ltLow ? 'low' : 'mid';
        if (eff <= 0) lv = lv === 'low' ? 'mid' : 'high';
      }
    }
    else if (it.max != null && it.qty > it.max) lv = 'over';
    else lv = 'ok';
    it.level = lv;
    var target = it.max != null ? it.max : it.min;
    it.suggest = target != null && (lv === 'high' || lv === 'mid' || lv === 'low') ? Math.max(0, Math.ceil(target - avail)) : 0;
    cnt[lv]++; if (it.isOut && !it.notUse) cnt.out++;
    out.push(it);
  });
  out.sort(function (a, b2) { return SA_LV_ORDER[a.level] - SA_LV_ORDER[b2.level] || (b2.lt || 0) - (a.lt || 0) || a.qty - b2.qty || String(a.code).localeCompare(String(b2.code)); });
  var si = stockInfo_ ? stockInfo_() : null, pi = poInfo_(), ri = rtInfo_();
  return { items: out, count: cnt, cfg: cfg, time: nowStr_(),
    src: { stock: si && si.time || '', pending: pi && pi.ok ? pi.time : '', rt: ri && ri.ok ? ri.time : '', rtRows: ri && ri.rows || 0 } };
}
function saKey_() {
  var raw = ['nosv', schedProp_('STOCK_HASH') || '', schedProp_('PO_HASH') || '', schedProp_('RT_HASH') || '', saVer_(), schedProp_('SA_CFG') || '', ordToday_()].join('|');
  return 'sa_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, raw)).replace(/=+$/, '');
}
function saData_(force) {
  var key = saKey_();
  if (!force) { var hit = ordCacheGet_(key); if (hit) return hit; }
  var d = saCompute_(); ordCachePut_(key, d, 21600); return d;
}
// public: full list for the Stock Alert page
function getStockAlert(token, force) {
  var d = saData_(!!force && isAdmin_(token)); d.admin = isAdmin_(token);
  try { d.groups = rsGroups_(); } catch (e) { d.groups = {}; }      // category names for the CODE 4-digit filter (RemainStock.gs)
  return d;
}
// public: small summary for the menu badge + dashboard card
function getStockAlertSummary() {
  var d = saData_(false);
  var top = d.items.filter(function (x) { return x.level === 'high' || x.level === 'mid' || x.level === 'low'; }).slice(0, 6)
    .map(function (x) { return { code: x.code, desc: x.desc, qty: x.qty, unit: x.unit, min: x.min, lt: x.lt, level: x.level, isOut: x.isOut }; });
  return { count: d.count, top: top, total: d.items.length, time: d.time, by: d.cfg.by };
}

// ── Admin: per-item settings + global config ────────────────────────
function saveStockAlertItem(token, code, v) {
  requireAdmin_(token);
  code = ordStr_(code, 40); if (!code) throw new Error('ไม่มี CODE');
  v = v || {};
  var num = function (x, nm) { if (x === '' || x == null) return ''; var n = Number(x); if (isNaN(n) || n < 0) throw new Error(nm + ' ต้องเป็นตัวเลขไม่ติดลบ'); return String(n); };
  var mn = num(v.min, 'Min'), mx = num(v.max, 'Max'), lt = num(v.lt, 'LT');
  if (mn !== '' && mx !== '' && Number(mx) < Number(mn)) throw new Error('Max ต้องไม่น้อยกว่า Min');
  return withLock_(function () {
    var s = saSetSheet_(), set = saSettings_(), k = ordCodeKey_(code), cur = set[k];
    var row = [code, mn, mx, v.notUse ? '1' : '', lt, ordStr_(v.note, 200), nowStr_(), 'Admin'];
    writeText_(s, cur ? cur.row : s.getLastRow() + 1, 1, row);
    saBump_();
    audit_([['Admin', 'ตั้งค่า Stock Alert', code, 'Min ' + (mn || '-') + ' · Max ' + (mx || '-') + ' · LT ' + (lt || 'auto') + (v.notUse ? ' · เลิกใช้' : '')]]);
    return getStockAlert(token, true);
  });
}
function setStockAlertConfig(token, c) {
  requireAdmin_(token);
  c = c || {};
  var by = ordStr_(c.by, 60) || SA_DEF.by, lo = Number(c.ltLow), hi = Number(c.ltHigh);
  if (!(lo > 0) || !(hi > lo)) throw new Error('เกณฑ์ LT ไม่ถูกต้อง (ต้องเป็น 0 < ระดับต่ำ < ระดับสูงสุด)');
  schedProp_('SA_CFG', JSON.stringify({ by: by, ltLow: lo, ltHigh: hi }));
  audit_([['Admin', 'ตั้งค่าเกณฑ์ Stock Alert', '', 'ผู้สั่ง ' + by + ' · LT ต่ำ ≤' + lo + ' · สูงสุด >' + hi]]);
  return getStockAlert(token, true);
}

// ── Admin: bulk edit (one call, all-or-nothing, with undo) ───────────
// list = [{ code, min?, max?, lt?, notUse?, note? }] — a field left out (undefined) keeps the current value, '' clears it
var SA_BULK_MAX = 3000;
function saveStockAlertBulk(token, list) {
  requireAdmin_(token);
  list = (list || []).slice(0, SA_BULK_MAX + 1);
  if (!list.length) throw new Error('ไม่มีรายการที่แก้ไข');
  if (list.length > SA_BULK_MAX) throw new Error('บันทึกได้ครั้งละไม่เกิน ' + SA_BULK_MAX + ' รายการ');
  return withLock_(function () {
    var s = saSetSheet_(), lr = s.getLastRow(), W = SA_SET_HEADER.length;
    var old = lr > 1 ? s.getRange(2, 1, lr - 1, W).getDisplayValues() : [];
    var tbl = old.map(function (r) { return r.slice(); }), idx = {};
    tbl.forEach(function (r, i) { var k = ordCodeKey_(r[0]); if (k) idx[k] = i; });
    var errs = [], now = nowStr_(), changed = 0, added = 0;
    var num = function (v, nm, code) {
      if (v === undefined) return undefined; if (v === '' || v == null) return '';
      var n = Number(String(v).replace(/,/g, '')); if (isNaN(n) || n < 0) { errs.push(code + ': ' + nm + ' "' + v + '" ไม่ใช่ตัวเลข'); return ''; }
      return String(n);
    };
    list.forEach(function (it) {
      var code = ordStr_(it && it.code, 40), k = ordCodeKey_(code);
      if (!k) { errs.push('มีแถวที่ไม่มี CODE'); return; }
      if (!saCodeOk_(k)) { errs.push(code + ': CODE ต้องยาว ' + SA_CODE_LEN + ' หลัก และไม่ขึ้นต้นด้วย SV'); return; }
      var i = idx[k], cur = i != null ? tbl[i] : [code, '', '', '', '', '', '', ''];
      var mn = num(it.min, 'Min', code), mx = num(it.max, 'Max', code), lt = num(it.lt, 'LT', code);
      var row = cur.slice();
      if (mn !== undefined) row[1] = mn; if (mx !== undefined) row[2] = mx;
      if (it.notUse !== undefined) row[3] = it.notUse ? '1' : '';
      if (lt !== undefined) row[4] = lt;
      if (it.note !== undefined) row[5] = ordStr_(it.note, 200);
      if (row[1] !== '' && row[2] !== '' && Number(row[2]) < Number(row[1])) errs.push(code + ': Max (' + row[2] + ') น้อยกว่า Min (' + row[1] + ')');
      if (row.slice(1, 6).join('|') === cur.slice(1, 6).join('|') && i != null) return;   // no real change
      row[6] = now; row[7] = 'Admin (bulk)';
      if (i != null) { tbl[i] = row; } else { idx[k] = tbl.length; tbl.push(row); added++; }
      changed++;
    });
    if (errs.length) throw new Error('ยังไม่ได้บันทึก — พบข้อผิดพลาด ' + errs.length + ' จุด: ' + errs.slice(0, 8).join(' · ') + (errs.length > 8 ? ' …' : ''));
    if (!changed) return getStockAlert(token, true);
    ordCachePut_('sa_undo', { rows: old, time: now, n: changed }, 21600);    // undo snapshot (6 h)
    if (tbl.length) { var rg = s.getRange(2, 1, tbl.length, W); rg.setNumberFormat('@'); rg.setValues(tbl); }
    saBump_();
    audit_([['Admin', 'แก้ไข Stock Alert หลายรายการ', '', 'แก้ไข ' + (changed - added) + ' · เพิ่มใหม่ ' + added + ' รายการ']]);
    var r = getStockAlert(token, true); r.saved = { changed: changed, added: added }; r.undo = { time: now, n: changed };
    return r;
  });
}
function undoStockAlertBulk(token) {
  requireAdmin_(token);
  return withLock_(function () {
    var u = ordCachePut_ && ordCacheGet_('sa_undo'); if (!u || !u.rows) throw new Error('ไม่มีการบันทึกที่ย้อนกลับได้ (เก็บไว้ 6 ชั่วโมง)');
    var s = saSetSheet_(), W = SA_SET_HEADER.length, lr = s.getLastRow();
    if (lr > 1) s.getRange(2, 1, lr - 1, W).clearContent();
    if (u.rows.length) { var rg = s.getRange(2, 1, u.rows.length, W); rg.setNumberFormat('@'); rg.setValues(u.rows); }
    CacheService.getScriptCache().remove('sa_undo_n');
    saBump_();
    audit_([['Admin', 'ย้อนกลับการแก้ไข Stock Alert', '', 'คืนค่าก่อน ' + u.time + ' (' + u.n + ' รายการ)']]);
    return getStockAlert(token, true);
  });
}
function getStockAlertUndo(token) { requireAdmin_(token); var u = ordCacheGet_('sa_undo'); return u ? { time: u.time, n: u.n } : null; }

// ── E-mail report to the team (preview first, then send) ────────────
// opt = { scope:'alert'|'selected', codes:[...], wait:bool, subject:'', note:'', due:'dd/MM/yyyy', to:'a@x,b@y', cc:'', csv:bool, sender:'' }
var SA_MAIL_MAX_ROWS = 150;
function saMailItems_(d, opt) {
  var want = null;
  if (opt.scope === 'selected') { want = {}; (opt.codes || []).forEach(function (c) { want[ordCodeKey_(c)] = 1; }); }
  return d.items.filter(function (x) {
    if (want) return !!want[ordCodeKey_(x.code)];
    return x.level === 'high' || x.level === 'mid' || x.level === 'low' || (opt.wait && x.level === 'wait');
  });
}
var SA_MAIL_LV = { high: ['สูงสุด', '#be123c', '#ffe4e6'], mid: ['กลาง', '#c2410c', '#ffedd5'], low: ['ต่ำ', '#854d0e', '#fef9c3'], wait: ['รอของเข้า', '#1d4ed8', '#dbeafe'],
  nomin: ['ยังไม่ตั้ง Min', '#57534e', '#f5f5f4'], over: ['เกิน Max', '#6d28d9', '#ede9fe'], ok: ['ปกติ', '#047857', '#d1fae5'], off: ['เลิกใช้', '#78716c', '#f5f5f4'] };
function saFmt_(n) { if (n === '' || n == null) return '-'; var x = Number(n); return isNaN(x) ? String(n) : x.toLocaleString('en-US', { maximumFractionDigits: 3 }); }
function saMailBuild_(opt) {
  opt = opt || {};
  var d = saData_(false), cfg = d.cfg, list = saMailItems_(d, opt), C = EMAIL_CFG;
  applyBrand_();
  var cnt = { high: 0, mid: 0, low: 0, wait: 0 }; list.forEach(function (x) { if (cnt[x.level] != null) cnt[x.level]++; });
  var by = cfg.by || 'DS13', today = ordToday_();
  var subject = ordStr_(opt.subject, 200) || ('[Stock Alert ' + by + '] สินค้าที่ควรพิจารณาสั่ง ' + list.length + ' รายการ' +
    (cnt.high ? ' · ด่วนสูงสุด ' + cnt.high : '') + ' — ' + today);
  var due = /^\d{2}\/\d{2}\/\d{4}$/.test(opt.due || '') ? opt.due : '';
  var intro = 'เรียน ทีม ' + by + '\n\nระบบ Stock Alert ตรวจพบสินค้าของแผนก ' + by + ' ที่ควรพิจารณาสั่งเพิ่มจำนวน ' + list.length + ' รายการ (ข้อมูล ณ ' + d.time + ') ' +
    'รบกวนทุกท่านช่วยกันพิจารณาว่ารายการใดควรเริ่มออก Order' + (due ? ' และแจ้งความเห็นภายในวันที่ ' + due : '') + ' ขอบคุณครับ';
  var pill = function (lv, n) { var L = SA_MAIL_LV[lv]; return '<td style="padding:4px 6px 4px 0;"><span style="display:inline-block;padding:6px 12px;border-radius:10px;background:' + L[2] + ';color:' + L[1] + ';font-size:12pt;font-weight:700;">' + esc_(L[0]) + ' ' + n + '</span></td>'; };
  var sum = '<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="margin:16px 0 4px;"><tr>' + pill('high', cnt.high) + pill('mid', cnt.mid) + pill('low', cnt.low) + (opt.wait || cnt.wait ? pill('wait', cnt.wait) : '') + '</tr></table>' +
    '<div style="font-size:10.5pt;color:' + C.muted + ';line-height:1.6;">เกณฑ์: คงเหลือต่ำกว่า Min · LT ≤ ' + cfg.ltLow + ' วัน = ต่ำ · ' + (cfg.ltLow + 1) + '–' + cfg.ltHigh + ' วัน = กลาง · มากกว่า ' + cfg.ltHigh + ' วัน = สูงสุด · ของหมดยกระดับขึ้น 1 ขั้น · แนะนำสั่ง = Max − (คงเหลือ + ค้างรับ)</div>';
  var th = function (t, al) { return '<th style="padding:8px 6px;font-size:10pt;color:' + C.muted + ';text-align:' + (al || 'left') + ';border-bottom:2px solid ' + C.line + ';white-space:nowrap;">' + t + '</th>'; };
  var rows = list.slice(0, SA_MAIL_MAX_ROWS).map(function (x, i) {
    var L = SA_MAIL_LV[x.level] || SA_MAIL_LV.ok, locs = (x.locs || []).length > 1 ? 'รวมจาก ' + x.locs.map(function (a) { return a[0] + ' ' + saFmt_(a[1]); }).join(' + ') : '';
    var p0 = (x.pendList || [])[0], pend = x.pend ? saFmt_(x.pend) + (p0 && p0.due ? '<br><span style="font-size:9pt;color:' + C.muted + ';">เข้า ' + esc_(p0.due) + '</span>' : '') : '-';
    var td = 'padding:8px 6px;border-bottom:1px solid ' + C.line + ';font-size:11pt;vertical-align:top;color:' + C.ink + ';';
    return '<tr style="background:' + (i % 2 ? '#ffffff' : '#fcfbf8') + ';">' +
      '<td style="' + td + 'border-left:4px solid ' + L[1] + ';"><span style="display:inline-block;padding:2px 8px;border-radius:999px;background:' + L[2] + ';color:' + L[1] + ';font-size:9.5pt;font-weight:700;white-space:nowrap;">' + esc_(L[0]) + '</span></td>' +
      '<td style="' + td + 'font-family:Consolas,monospace;font-weight:700;white-space:nowrap;">' + esc_(x.code) + '</td>' +
      '<td style="' + td + 'min-width:150px;">' + esc_(x.desc || '-') + (x.comment ? '<br><span style="font-size:9.5pt;color:' + C.muted + ';">' + esc_(x.comment) + '</span>' : '') + (x.note ? '<br><span style="font-size:9.5pt;color:#a9583e;">📝 ' + esc_(x.note) + '</span>' : '') + '</td>' +
      '<td style="' + td + 'text-align:right;white-space:nowrap;"><b style="color:' + (x.isOut ? '#be123c' : C.ink) + ';">' + (x.isOut ? 'หมด' : saFmt_(x.qty)) + '</b> ' + esc_(x.unit || '') + (locs ? '<br><span style="font-size:9pt;color:#a9583e;white-space:normal;">' + esc_(locs) + '</span>' : '') + '</td>' +
      '<td style="' + td + 'text-align:right;white-space:nowrap;">' + saFmt_(x.min) + ' / ' + saFmt_(x.max) + '</td>' +
      '<td style="' + td + 'text-align:right;white-space:nowrap;color:#1d4ed8;">' + pend + '</td>' +
      '<td style="' + td + 'text-align:right;white-space:nowrap;">' + (x.lt ? x.lt + ' วัน' : '-') + '</td>' +
      '<td style="' + td + 'text-align:right;white-space:nowrap;font-weight:700;">' + (x.suggest ? saFmt_(x.suggest) : '-') + '</td>' +
      '<td style="' + td + 'white-space:nowrap;font-size:10pt;color:' + C.muted + ';">☐ สั่ง<br>☐ ยังไม่สั่ง</td></tr>';
  }).join('');
  var table = list.length ? '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;margin:16px 0 8px;border:1px solid ' + C.line + ';">' +
    '<tr style="background:#faf9f5;">' + th('ระดับ') + th('CODE') + th('รายละเอียด') + th('คงเหลือ', 'right') + th('Min / Max', 'right') + th('ค้างรับ', 'right') + th('LT', 'right') + th('แนะนำสั่ง', 'right') + th('ความเห็น') + '</tr>' + rows + '</table>' +
    (list.length > SA_MAIL_MAX_ROWS ? '<div style="font-size:10.5pt;color:' + C.muted + ';">… และอีก ' + (list.length - SA_MAIL_MAX_ROWS) + ' รายการ (ดูทั้งหมดในไฟล์แนบ / ในระบบ)</div>' : '')
    : '<div style="padding:16px;border:1px dashed ' + C.line + ';border-radius:12px;font-size:12pt;color:' + C.muted + ';margin-top:14px;">ไม่มีรายการตามเงื่อนไขที่เลือก</div>';
  var how = '<div style="margin:14px 0 6px;padding:12px 14px;background:#f5f0e8;border-radius:12px;font-size:11pt;line-height:1.7;color:#3d3d3a;"><b>วิธีแจ้งความเห็น</b><br>' +
    'กด <b>Reply All</b> แล้วระบุ CODE พร้อม "สั่ง" (จำนวน) หรือ "ยังไม่สั่ง" (เหตุผล) — เช่น <span style="font-family:Consolas,monospace;">43401027 สั่ง 200</span><br>' +
    'ข้อมูลอ้างอิง: Item stock ' + esc_(d.src.stock || '-') + ' · Pending ' + esc_(d.src.pending || '-') + ' · Lead time จาก Q_StockRT</div>';
  var sender = ordStr_(opt.sender, 60);
  var o = { title: subject, preheader: list.length + ' รายการที่ควรพิจารณาสั่ง · สูงสุด ' + cnt.high + ' · กลาง ' + cnt.mid + ' · ต่ำ ' + cnt.low,
    heading: 'รายงานสินค้าที่ควรพิจารณาสั่งซื้อ (' + by + ')', intro: intro,
    message: ordStr_(opt.note, 1500), messageLabel: 'ข้อความจาก' + (sender ? ' ' + sender : 'ผู้ส่ง'),
    extraHtml: sum + table + how, buttons: [{ label: 'เปิด Stock Alert ในระบบ', url: baseUrl_() }],
    footer: 'อีเมลนี้ส่งจากระบบ Stock Alert (' + by + ') เพื่อใช้ประกอบการตัดสินใจภายในแผนก' + (sender ? ' · ผู้ส่ง: ' + sender : '') + ' · ตอบกลับอีเมลนี้เพื่อแจ้งความเห็นได้' };
  o.intro = o.intro.replace(/\n+/g, ' ');
  var html = emailHtml_(o).replace(esc_(o.intro), esc_(intro).replace(/\n/g, '<br>'));
  var text = [o.heading, '', intro, '', o.message ? o.messageLabel + ': ' + o.message + '\n' : '']
    .concat(list.slice(0, SA_MAIL_MAX_ROWS).map(function (x) { return '[' + (SA_MAIL_LV[x.level] || ['-'])[0] + '] ' + x.code + ' ' + (x.desc || '') + ' | คงเหลือ ' + saFmt_(x.qty) + ' ' + (x.unit || '') + ' | Min/Max ' + saFmt_(x.min) + '/' + saFmt_(x.max) + ' | ค้างรับ ' + saFmt_(x.pend) + ' | LT ' + (x.lt || '-') + ' | แนะนำสั่ง ' + (x.suggest || '-'); }))
    .concat(['', 'เปิดระบบ: ' + baseUrl_()]).join('\n');
  return { subject: subject.replace(/[\r\n]+/g, ' '), html: html, text: text, list: list, count: cnt };
}
function saMailCsv_(list) {
  var e = function (v) { v = String(v == null ? '' : v); return /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; };
  var H = ['ระดับ', 'CODE', 'DESC', 'COMMENT', 'คงเหลือ', 'หน่วย', 'รวมจากคลัง', 'ค้างรับ', 'MIN', 'MAX', 'LT (วัน)', 'แนะนำสั่ง', 'หมายเหตุ', 'ความเห็น (สั่ง/ยังไม่สั่ง)'];
  var rows = list.map(function (x) {
    return [(SA_MAIL_LV[x.level] || ['-'])[0], x.code, x.desc, x.comment, x.qty, x.unit, (x.locs || []).map(function (a) { return a[0] + ' ' + a[1]; }).join(' + '), x.pend, x.min == null ? '' : x.min, x.max == null ? '' : x.max, x.lt || '', x.suggest || '', x.note || '', ''];
  });
  return Utilities.newBlob('﻿' + [H].concat(rows).map(function (r) { return r.map(e).join(','); }).join('\r\n'), 'text/csv', 'StockAlert_' + ordToday_().replace(/\//g, '-') + '.csv');
}
function getStockAlertMailCfg(token) {
  requireAdmin_(token);
  var last = {}; try { last = JSON.parse(schedProp_('SA_MAIL_LAST') || '{}') || {}; } catch (e) {}
  return { to: last.to || '', cc: last.cc || '', sender: last.sender || '', quota: MailApp.getRemainingDailyQuota() };
}
function previewStockAlertEmail(token, opt) {
  requireAdmin_(token);
  var m = saMailBuild_(opt);
  return { subject: m.subject, html: m.html, n: m.list.length, count: m.count };
}
function sendStockAlertEmail(token, opt) {
  requireAdmin_(token);
  opt = opt || {};
  var to = uniq_(emailList_(opt.to)), cc = uniq_(emailList_(opt.cc)).filter(function (x) { return to.map(function (y) { return y.toLowerCase(); }).indexOf(x.toLowerCase()) < 0; });
  if (!to.length) throw new Error('กรุณาระบุผู้รับ (To) อย่างน้อย 1 อีเมล');
  if (to.length + cc.length > 50) throw new Error('ผู้รับรวมต้องไม่เกิน 50 อีเมล');
  var m = saMailBuild_(opt);
  if (!m.list.length) throw new Error('ไม่มีรายการให้ส่ง');
  if (MailApp.getRemainingDailyQuota() < to.length + cc.length) throw new Error('โควตาอีเมลวันนี้ไม่พอ (เหลือ ' + MailApp.getRemainingDailyQuota() + ')');
  var mo = { to: to.join(','), cc: cc.join(','), subject: m.subject, htmlBody: m.html, body: m.text, name: 'Stock Alert ' + (saCfg_().by || 'DS13') };
  var rp = emailList_(opt.replyTo)[0] || to[0]; if (rp) mo.replyTo = rp;
  if (opt.csv) mo.attachments = [saMailCsv_(m.list)];
  MailApp.sendEmail(mo);
  schedProp_('SA_MAIL_LAST', JSON.stringify({ to: to.join(', '), cc: cc.join(', '), sender: ordStr_(opt.sender, 60) }));
  logMail_('stock_alert', '', to.join(','), m.subject, 'ส่งแล้ว (' + m.list.length + ' รายการ)', cc.join(','));
  audit_([['Admin', 'ส่งอีเมล Stock Alert', '', m.list.length + ' รายการ → ' + to.concat(cc).length + ' ผู้รับ']]);
  return { ok: true, n: m.list.length, to: to, cc: cc, subject: m.subject, quota: MailApp.getRemainingDailyQuota() };
}