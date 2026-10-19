/**
 * ═════════════════════════════════════════════════════════════════════
 *  REMAIN STOCK — "Stock ช่วยใช้"  ของค้างนาน / RQN เก่า / ล็อตสุดท้าย → ช่วยกันนำไปใช้ ลดการ Scrap
 *  • ขอบเขต      = Item stock ที่ "ผู้สั่งล่าสุด" (LastOfRequest_By) ตรงกับค่าที่ตั้ง (เริ่มต้น DS13)
 *                  + CODE ที่ Admin ติดป้ายไว้เอง
 *  • ค้างนาน     = จำนวนคงเหลือรวมทุกคลังไม่เปลี่ยนมา ≥ N วัน (จากประวัติ StockHistory ที่ระบบจดทุกครั้งที่ซิงค์ Stock)
 *  • RQN เก่า     = เลข RQN ล่าสุดเก่ากว่า N ปี (RQNyy-xxxx)
 *  • ป้าย (Admin) = last (ล็อตสุดท้าย/เปลี่ยนรุ่น) · sub (ใช้แทนกันได้ → CODE ใหม่) · skip (ไม่ต้องแจ้ง)
 *  • ราคา / มูลค่า (QTY × LASTCOST) เห็นเฉพาะ Admin
 *  ต้องมี Code.gs, OrderService.gs และ StockAlert.gs (ใช้ helper ร่วมกัน)
 * ═════════════════════════════════════════════════════════════════════
 */
var RS_TAG_SHEET = 'RemainStockTags';
var RS_TAG_HEADER = ['CODE', 'TAG', 'REPLACE_BY', 'USE_BEFORE', 'NOTE', 'UPDATED', 'BY'];
var RS_HIST_SHEET = 'StockHistory';                    // in the Helpdesk_ItemStock workbook
var RS_HIST_HEADER = ['CODE', 'QTY', 'SINCE', 'LAST_SEEN'];
var RS_DEF = { days: 180, rqnYears: 2, by: 'DS13' };
var RS_TAGS = { last: 'ล็อตสุดท้าย / เปลี่ยนรุ่น', sub: 'ใช้แทนกันได้', skip: 'ไม่ต้องแจ้ง' };

function rsCfg_() {
  var c = {}; try { c = JSON.parse(schedProp_('RS_CFG') || '{}') || {}; } catch (e) {}
  var n = function (v, d, lo, hi) { v = Number(v); return isFinite(v) && v >= lo && v <= hi ? Math.round(v) : d; };
  return { days: n(c.days, RS_DEF.days, 1, 3650), rqnYears: n(c.rqnYears, RS_DEF.rqnYears, 0, 20),
    by: String(c.by || RS_DEF.by).substring(0, 200) };
}
function saveRemainCfg(token, cfg) {
  requireAdmin_(token);
  cfg = cfg || {};
  var d = Number(cfg.days), y = Number(cfg.rqnYears), by = String(cfg.by || '').replace(/[^A-Za-z0-9,; ]/g, '').trim();
  if (!(d >= 1 && d <= 3650)) throw new Error('จำนวนวันค้างนานต้องอยู่ระหว่าง 1–3650');
  if (!(y >= 0 && y <= 20)) throw new Error('อายุ RQN ต้องอยู่ระหว่าง 0–20 ปี (0 = ไม่ใช้เกณฑ์นี้)');
  if (!by) throw new Error('ต้องระบุผู้สั่งอย่างน้อย 1 ราย เช่น DS13');
  var c = { days: Math.round(d), rqnYears: Math.round(y), by: by };
  schedProp_('RS_CFG', JSON.stringify(c));
  audit_([['Admin', 'ตั้งค่า Stock ช่วยใช้', '', 'ค้าง ≥ ' + c.days + ' วัน · RQN ≥ ' + c.rqnYears + ' ปี · ผู้สั่ง ' + c.by]]);
  return c;
}

// ── tags (main spreadsheet) ─────────────────────────────────────────
function rsTagSheet_() {
  var s = ss_().getSheetByName(RS_TAG_SHEET);
  if (!s) { s = ss_().insertSheet(RS_TAG_SHEET); s.getRange(1, 1, 1, RS_TAG_HEADER.length).setValues([RS_TAG_HEADER]).setFontWeight('bold'); s.setFrozenRows(1); }
  return s;
}
function rsTags_() {
  var s = ss_().getSheetByName(RS_TAG_SHEET), out = {};
  if (!s || s.getLastRow() < 2) return out;
  s.getRange(2, 1, s.getLastRow() - 1, RS_TAG_HEADER.length).getDisplayValues().forEach(function (r, i) {
    var k = ordCodeKey_(r[0]); if (!k || !RS_TAGS[r[1]]) return;
    out[k] = { row: i + 2, tag: r[1], replaceBy: r[2], useBefore: r[3], note: r[4], updated: r[5] };
  });
  return out;
}
// tag = 'last' | 'sub' | 'skip' | '' (remove)
function saveRemainTag(token, code, t) {
  requireAdmin_(token);
  t = t || {};
  var k = ordCodeKey_(code), tag = String(t.tag || '');
  if (!k || k.length > 40) throw new Error('CODE ไม่ถูกต้อง');
  if (tag && !RS_TAGS[tag]) throw new Error('ป้ายไม่ถูกต้อง');
  var useBefore = String(t.useBefore || '').trim();
  if (useBefore && !/^\d{2}\/\d{2}\/\d{4}$/.test(useBefore)) throw new Error('วันที่ควรใช้ให้หมดก่อน ต้องเป็นรูปแบบ วว/ดด/ปปปป');
  return withLock_(function () {
    var s = rsTagSheet_(), cur = rsTags_()[k];
    var row = [k, tag, ordStr_(t.replaceBy, 40).toUpperCase().replace(/\s+/g, ''), useBefore, ordStr_(t.note, 300), nowStr_(), 'Admin'];
    if (!tag) { if (cur) s.deleteRow(cur.row); }
    else if (cur) writeText_(s, cur.row, 1, row);
    else writeText_(s, s.getLastRow() + 1, 1, row);
    audit_([['Admin', tag ? 'ติดป้าย Stock ช่วยใช้: ' + RS_TAGS[tag] : 'ลบป้าย Stock ช่วยใช้', k, row[2] ? '→ ' + row[2] : '']]);
    CacheService.getScriptCache().remove('rs_flags');
    return true;
  });
}

// ── history: when did each CODE's total quantity last change? (Helpdesk_ItemStock → StockHistory) ──
// Called after every successful Stock sync. One row per CODE, so it stays small.
function rsTrack_(rows) {
  var tot = {};
  rows.forEach(function (r) { var k = ordCodeKey_(r[0]); if (!k) return; var q = Number(r[3]); tot[k] = Math.round(((tot[k] || 0) + (isFinite(q) ? q : 0)) * 1000) / 1000; });
  var ss = stockSS_(), sh = ss.getSheetByName(RS_HIST_SHEET) || ss.insertSheet(RS_HIST_SHEET), today = ordToday_(), old = {};
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 4).getDisplayValues().forEach(function (r) { if (r[0]) old[r[0]] = r; });
  var out = Object.keys(tot).sort().map(function (k) {
    var o = old[k], q = String(tot[k]);
    return [k, q, o && o[1] === q && o[2] ? o[2] : today, today];
  });
  var all = [RS_HIST_HEADER].concat(out);
  sh.clearContents();
  if (sh.getMaxRows() < all.length) sh.insertRowsAfter(sh.getMaxRows(), all.length - sh.getMaxRows());
  var rg = sh.getRange(1, 1, all.length, 4); rg.setNumberFormat('@'); rg.setValues(all);
  sh.setFrozenRows(1);
  if (!schedProp_('RS_HIST_START')) schedProp_('RS_HIST_START', today);
  try { rsMeasure_(ss, tot, rows); } catch (e) { console.warn('rsMeasure_', e); }
}
// Compare the flagged list remembered at the previous sync with today's totals:
// every decrease of a flagged CODE = stock put to use instead of scrap → RemainStockSaved.
// Then remember today's flagged list (sheet RemainStockPrev in Helpdesk_ItemStock).
var RS_PREV_SHEET = 'RemainStockPrev', RS_SAVED_SHEET = 'RemainStockSaved';
var RS_SAVED_HEADER = ['DATE', 'CODE', 'DESC', 'QTY_USED', 'UNIT', 'VALUE', 'WHY'];
function rsMeasure_(ss, tot, rows) {
  var prevSh = ss.getSheetByName(RS_PREV_SHEET) || ss.insertSheet(RS_PREV_SHEET), today = ordToday_(), used = [];
  if (prevSh.getLastRow() > 1) prevSh.getRange(2, 1, prevSh.getLastRow() - 1, 6).getDisplayValues().forEach(function (r) {
    var k = r[0], q0 = Number(r[1]), q1 = tot[k] || 0, c = Number(r[2]) || 0;
    if (k && q1 < q0) used.push([today, k, r[3], String(Math.round((q0 - q1) * 1000) / 1000), r[4], String(Math.round((q0 - q1) * c * 100) / 100), r[5]]);
  });
  if (used.length) {
    var sv = rsSheet_(RS_SAVED_SHEET, RS_SAVED_HEADER);
    var rg = sv.getRange(sv.getLastRow() + 1, 1, used.length, RS_SAVED_HEADER.length); rg.setNumberFormat('@'); rg.setValues(used);
    rsCloseClaims_(used.map(function (u) { return u[1]; }));
  }
  var cur = rsCompute_(rows.map(function (r) { return r[STOCK_NCOL]; }), true).items.filter(function (x) { return !x.skip; })
    .map(function (x) { return [x.key, String(x.qty), String(x.cost || 0), x.desc || '', x.unit || '', (x.why || []).join(',')]; });
  var all = [['CODE', 'QTY', 'COST', 'DESC', 'UNIT', 'WHY']].concat(cur);
  prevSh.clearContents();
  var rg2 = prevSh.getRange(1, 1, all.length, 6); rg2.setNumberFormat('@'); rg2.setValues(all);
}
function rsSheet_(name, header) {
  var s = ss_().getSheetByName(name);
  if (!s) { s = ss_().insertSheet(name); s.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold'); s.setFrozenRows(1); }
  return s;
}
function rsSavedSummary_(admin) {
  var s = ss_().getSheetByName(RS_SAVED_SHEET), now = new Date(), mm = ('0' + (now.getMonth() + 1)).slice(-2) + '/' + now.getFullYear();
  var o = { month: { items: 0, value: 0 }, all: { items: 0, value: 0 }, recent: [] };
  if (s && s.getLastRow() > 1) {
    var rows = s.getRange(2, 1, s.getLastRow() - 1, RS_SAVED_HEADER.length).getDisplayValues();
    rows.forEach(function (r) {
      var v = Number(r[5]) || 0; o.all.items++; o.all.value += v;
      if (String(r[0]).slice(3) === mm) { o.month.items++; o.month.value += v; }
    });
    o.recent = rows.slice(-10).reverse().map(function (r) { return { date: r[0], code: r[1], desc: r[2], qty: r[3], unit: r[4], value: admin ? Number(r[5]) || 0 : null }; });
  }
  o.month.value = admin ? Math.round(o.month.value) : null; o.all.value = admin ? Math.round(o.all.value) : null;
  return o;
}

// ── "ฉันจะนำไปใช้" claims ────────────────────────────────────────────
var RS_CLAIM_SHEET = 'RemainStockClaims';
var RS_CLAIM_HEADER = ['ID', 'CODE', 'QTY', 'NAME', 'PURPOSE', 'CREATED', 'STATUS', 'CLOSED'];
function rsClaims_(openOnly) {
  var s = ss_().getSheetByName(RS_CLAIM_SHEET), out = {};
  if (!s || s.getLastRow() < 2) return out;
  s.getRange(2, 1, s.getLastRow() - 1, RS_CLAIM_HEADER.length).getDisplayValues().forEach(function (r, i) {
    if (!r[0] || (openOnly && r[6] !== 'open')) return;
    (out[r[1]] = out[r[1]] || []).push({ id: r[0], row: i + 2, qty: r[2], name: r[3], purpose: r[4], created: r[5], status: r[6] });
  });
  return out;
}
function claimRemain(code, c) {
  c = c || {};
  var k = ordCodeKey_(code), name = ordStr_(c.name, 60).trim(), purpose = ordStr_(c.purpose, 200).trim(), qty = ordStr_(c.qty, 12).replace(/[^0-9.]/g, '');
  if (!k) throw new Error('CODE ไม่ถูกต้อง');
  if (!name) throw new Error('กรุณากรอกชื่อผู้จะนำไปใช้');
  if (!purpose) throw new Error('กรุณาระบุงาน / โปรเจกต์ที่จะนำไปใช้');
  var rc = CacheService.getScriptCache(), rk = 'rs_cl_' + name.toLowerCase(), n = Number(rc.get(rk) || 0);
  if (n >= 20) throw new Error('จองถี่เกินไป กรุณารอสักครู่');
  return withLock_(function () {
    var s = rsSheet_(RS_CLAIM_SHEET, RS_CLAIM_HEADER), id = 'RS' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
    writeText_(s, s.getLastRow() + 1, 1, [id, k, qty, name, purpose, nowStr_(), 'open', '']);
    rc.put(rk, String(n + 1), 3600);
    audit_([['User: ' + name, 'จองใช้ Stock ช่วยใช้', k, (qty ? qty + ' · ' : '') + purpose]]);
    return rsClaims_(true);
  });
}
// status: 'done' | 'cancel'  (Admin)
function closeRemainClaim(token, id, status) {
  requireAdmin_(token);
  if (status !== 'done' && status !== 'cancel') throw new Error('สถานะไม่ถูกต้อง');
  return withLock_(function () {
    var all = rsClaims_(true), hit = null;
    Object.keys(all).forEach(function (k) { all[k].forEach(function (x) { if (x.id === id) hit = x; }); });
    if (!hit) throw new Error('ไม่พบรายการจอง ' + id);
    ss_().getSheetByName(RS_CLAIM_SHEET).getRange(hit.row, 7, 1, 2).setValues([[status, nowStr_()]]);
    audit_([['Admin', status === 'done' ? 'ปิดรายการจอง (นำไปใช้แล้ว)' : 'ยกเลิกรายการจอง', id, '']]);
    return rsClaims_(true);
  });
}
function rsCloseClaims_(codes) {                      // stock went down → open claims on that CODE are done
  var want = {}; codes.forEach(function (k) { want[k] = 1; });
  var s = ss_().getSheetByName(RS_CLAIM_SHEET); if (!s) return;
  var all = rsClaims_(true), t = nowStr_();
  Object.keys(all).forEach(function (k) { if (want[k]) all[k].forEach(function (x) { s.getRange(x.row, 7, 1, 2).setValues([['done', t]]); }); });
}

// ── for the duplicate-order popup: substitutable / leftover stock for the CODEs being ordered ──
function rsForOrder_(keys) {
  var want = {}, out = {}; keys.forEach(function (k) { want[k] = 1; });
  var rows = []; try { rows = stockData_(); } catch (e) { return out; }
  rsCompute_(rows, false).items.forEach(function (x) {
    if (x.skip) return;
    var target = x.replaceBy && want[x.replaceBy] ? x.replaceBy : want[x.key] ? x.key : '';
    if (!target) return;
    (out[target] = out[target] || []).push({ code: x.code, desc: x.desc, qty: x.qty, unit: x.unit, why: x.why, self: target === x.key,
      useBefore: x.tag && x.tag.useBefore || '' });
  });
  return out;
}
function rsHist_(want) {
  var out = {};
  try {
    var sh = stockSS_().getSheetByName(RS_HIST_SHEET);
    if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 3).getDisplayValues().forEach(function (r) { if (want[r[0]]) out[r[0]] = r[2]; });
  } catch (e) {}
  return out;
}
function rsDays_(dmy) {
  var m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(String(dmy || ''));
  if (!m) return null;
  var now = new Date(), t0 = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  return Math.max(0, Math.round((t0 - new Date(+m[3], +m[2] - 1, +m[1]).getTime()) / 864e5));
}
function rsRqnYear_(rqn) { var m = /RQN\s*(\d{2})\s*-/i.exec(String(rqn || '')); return m ? 2000 + Number(m[1]) : null; }

// ── the list ────────────────────────────────────────────────────────
function getRemainStock(token) {
  var t0 = Date.now(), admin = isAdmin_(token), rows = [];
  try { rows = stockData_(); } catch (e) {}
  var out = rsCompute_(rows, admin);
  out.claims = rsClaims_(true);
  out.saved = rsSavedSummary_(admin);
  perf_('getRemainStock', t0);
  return out;
}
// rows = packed Item-stock rows ("ITEMNO\tDESC\t…" strings)
function rsCompute_(rows, admin) {
  var cfg = rsCfg_(), isBy = saByMatch_(cfg.by), tags = rsTags_(), items = {};
  for (var i = 0; i < rows.length; i++) {
    var p = rows[i]; if (!p) continue;
    var f = p.split('\t'), k = ordCodeKey_(f[0]);
    if (!k || !(isBy(f[7]) || tags[k])) continue;
    var q = saNum_(f[3]), c = saNum_(f[8]);
    var it = items[k] || (items[k] = { code: f[0], desc: f[1], comment: f[2], unit: f[4], qty: 0, locs: [], cost: 0, rqn: '', by: f[7] });
    if (q != null) it.qty += q;
    it.locs.push([f[5] || '-', q == null ? '' : q]);
    if (c != null && c > it.cost) it.cost = c;
    if (f[9] && String(f[9]) > String(it.rqn)) it.rqn = f[9];
    if (!it.unit && f[4]) it.unit = f[4];
    if (isBy(f[7])) it.by = f[7];
  }
  var want = {}; Object.keys(items).forEach(function (k) { want[k] = 1; });
  var hist = rsHist_(want), year = new Date().getFullYear(), out = [], total = 0, start = schedProp_('RS_HIST_START') || '';
  Object.keys(items).forEach(function (k) {
    var it = items[k], tg = tags[k] || null;
    it.qty = Math.round(it.qty * 1000) / 1000;
    if (!(it.qty > 0)) return;
    it.since = hist[k] || ''; it.still = rsDays_(it.since);
    var ry = rsRqnYear_(it.rqn); it.rqnAge = ry ? year - ry : null;
    var why = [];
    if (it.still != null && it.still >= cfg.days) why.push('still');
    if (cfg.rqnYears > 0 && it.rqnAge != null && it.rqnAge >= cfg.rqnYears) why.push('rqn');
    if (tg && (tg.tag === 'last' || tg.tag === 'sub')) why.push(tg.tag);
    it.skip = !!(tg && tg.tag === 'skip');
    if (!why.length && !it.skip) return;
    it.why = why; it.key = k; it.len = k.length;
    if (tg) it.tag = { tag: tg.tag, replaceBy: tg.replaceBy, useBefore: tg.useBefore, note: tg.note };
    if (tg && tg.replaceBy) it.replaceBy = tg.replaceBy;
    it.value = Math.round(it.qty * it.cost * 100) / 100;
    if (!it.skip) total += it.value;
    if (!admin) { delete it.cost; delete it.value; delete it.rqn; }
    out.push(it);
  });
  out.sort(function (a, b) { return admin ? (b.value - a.value) : ((b.still || 0) - (a.still || 0) || (b.rqnAge || 0) - (a.rqnAge || 0)); });
  return { items: out, cfg: cfg, admin: admin, histStart: start, total: admin ? Math.round(total) : null,
    stockTime: (stockInfo_() || {}).time || '' };
}

// small public map for badges in Item stock: { CODE: 'last' | 'sub' | 'old' }
function getRemainFlags() {
  var c = CacheService.getScriptCache(), hit = c.get('rs_flags');
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  var m = {};
  try { getRemainStock('').items.forEach(function (x) { if (!x.skip) m[x.key] = x.why.indexOf('last') >= 0 ? 'last' : x.why.indexOf('sub') >= 0 ? 'sub' : 'old'; }); } catch (e) {}
  try { c.put('rs_flags', JSON.stringify(m), 3600); } catch (e) {}
  return m;
}

// ── PDF report ──────────────────────────────────────────────────────
// opt = { group: 'prefix4' | 'length' | 'none', scope: 'all' | '8' | '16', why: 'all' | 'still' | 'rqn' | 'last' | 'sub',
//         claims: true/false (show only items without a claim) }
// Everyone can export; cost / value columns appear only for Admin (same data the page gets).
function exportRemainReport(token, opt) {
  var t0 = Date.now(); opt = opt || {};
  var d = getRemainStock(token), admin = d.admin, claims = d.claims || {};
  var group = { prefix4: 1, length: 1, none: 1 }[opt.group] ? opt.group : 'prefix4';
  var list = d.items.filter(function (x) {
    if (x.skip) return false;
    if (opt.scope === '8' && x.len !== 8) return false;
    if (opt.scope === '16' && x.len !== 16) return false;
    if (opt.why && opt.why !== 'all' && (x.why || []).indexOf(opt.why) < 0) return false;
    if (opt.unclaimed && claims[x.key]) return false;
    return true;
  });
  var keyOf = function (x) { return group === 'prefix4' ? x.key.substring(0, 4) : group === 'length' ? (x.len === 8 ? 'CODE 8 หลัก' : x.len === 16 ? 'CODE 16 หลัก' : 'CODE ' + x.len + ' หลัก') : 'ทุกรายการ'; };
  var groups = {}, order = [];
  list.forEach(function (x) { var k = keyOf(x); if (!groups[k]) { groups[k] = { name: k, items: [], value: 0, qtyItems: 0, oldest: 0, claimed: 0 }; order.push(k); }
    var g = groups[k]; g.items.push(x); g.value += x.value || 0; g.oldest = Math.max(g.oldest, x.still || 0); if (claims[x.key]) g.claimed++; });
  order.sort(function (a, b) { return admin ? groups[b].value - groups[a].value : String(a).localeCompare(String(b)); });
  var WHY = { still: 'ค้างนาน', rqn: 'RQN เก่า', last: 'ล็อตสุดท้าย', sub: 'ใช้แทนได้' };
  var fmt = function (n) { return n == null || n === '' ? '-' : Number(n).toLocaleString('en-US', { maximumFractionDigits: 2 }); };
  var money = function (n) { return '฿' + Number(Math.round(n || 0)).toLocaleString('en-US'); };
  var cnt = function (w) { return list.filter(function (x) { return (x.why || []).indexOf(w) >= 0; }).length; };
  var total = list.reduce(function (s, x) { return s + (x.value || 0); }, 0);
  var claimedN = list.filter(function (x) { return claims[x.key]; }).length;
  var scopeTxt = { all: 'ทุกความยาว CODE', '8': 'CODE 8 หลัก', '16': 'CODE 16 หลัก' }[opt.scope] || 'ทุกความยาว CODE';
  var groupTxt = { prefix4: 'แบ่งหมวดตาม CODE 4 หลักแรก', length: 'แบ่งตามความยาว CODE (8 / 16 หลัก)', none: 'ไม่แบ่งหมวด' }[group];
  var C = '#cc785c', L = '#e6dfd8', M = '#6c6a64';
  var th = function (t, al) { return '<th style="background:#f5f0e8;border:1px solid ' + L + ';padding:4px 5px;font-size:8.5pt;text-align:' + (al || 'left') + ';">' + esc_(t) + '</th>'; };
  var td = function (t, al, st) { return '<td style="border:1px solid ' + L + ';padding:4px 5px;font-size:8.5pt;vertical-align:top;text-align:' + (al || 'left') + ';' + (st || '') + '">' + t + '</td>'; };
  var kpi = function (v, l, col) { return '<td style="border:1px solid ' + L + ';padding:8px;text-align:center;width:20%;"><div style="font-size:15pt;font-weight:bold;color:' + (col || '#141413') + ';">' + v + '</div><div style="font-size:8pt;color:' + M + ';">' + esc_(l) + '</div></td>'; };
  var top = list.slice().sort(function (a, b) { return admin ? b.value - a.value : (b.still || 0) - (a.still || 0); }).slice(0, 10);
  var h = [];
  h.push('<html><head><meta charset="utf-8"><style>body{font-family:Sarabun,Tahoma,Arial,sans-serif;color:#141413;font-size:9pt;}h1{font-size:17pt;margin:0;}h2{font-size:12pt;margin:16px 0 6px;border-bottom:2px solid ' + C + ';padding-bottom:3px;}table{border-collapse:collapse;width:100%;}</style></head><body>');
  h.push('<table><tr><td style="background:' + C + ';color:#fff;padding:10px 12px;"><h1>♻ รายงาน Stock ช่วยใช้ (Remain stock)</h1><div style="font-size:9pt;">DAIHEN ELECTRIC CO., LTD. · ช่วยกันนำของค้างไปใช้ เพื่อลดการ Scrap</div></td></tr></table>');
  h.push('<p style="font-size:8.5pt;color:' + M + ';margin:6px 0;">ออกรายงาน ' + esc_(nowStr_()) + ' · ข้อมูล Stock ' + esc_(d.stockTime || '-') + ' · ผู้สั่ง ' + esc_(d.cfg.by) + ' · เกณฑ์: ค้างนาน ≥ ' + d.cfg.days + ' วัน' + (d.cfg.rqnYears ? ', RQN เก่า ≥ ' + d.cfg.rqnYears + ' ปี' : '') + ' · ' + esc_(scopeTxt) + ' · ' + esc_(groupTxt) + (opt.why && opt.why !== 'all' ? ' · เฉพาะ "' + esc_(WHY[opt.why] || '') + '"' : '') + (opt.unclaimed ? ' · เฉพาะที่ยังไม่มีผู้จอง' : '') + (d.histStart ? ' · เริ่มเก็บประวัติ ' + esc_(d.histStart) : '') + '</p>');
  // 1) summary
  h.push('<h2>1. สรุปภาพรวม</h2><table><tr>' + kpi(fmt(list.length), 'รายการช่วยใช้') +
    (admin ? kpi(money(total), 'มูลค่าเสี่ยง Scrap', '#9f1239') : kpi(fmt(order.length), 'จำนวนหมวด')) +
    kpi(fmt(claimedN), 'มีผู้จองนำไปใช้แล้ว', '#5b21b6') +
    kpi(d.saved ? (admin ? money(d.saved.month.value) : fmt(d.saved.month.items)) : '-', admin ? 'ช่วยลด Scrap เดือนนี้' : 'รายการที่ถูกใช้เดือนนี้', '#065f46') +
    kpi(d.saved ? (admin ? money(d.saved.all.value) : fmt(d.saved.all.items)) : '-', admin ? 'ช่วยลด Scrap สะสม' : 'ถูกใช้สะสม (ครั้ง)', '#065f46') + '</tr></table>');
  h.push('<p style="font-size:8.5pt;margin:6px 0;">แยกตามเหตุผล: ค้างนาน ' + cnt('still') + ' · RQN เก่า ' + cnt('rqn') + ' · ล็อตสุดท้าย ' + cnt('last') + ' · ใช้แทนได้ ' + cnt('sub') + ' รายการ (1 รายการอาจมีหลายเหตุผล)</p>');
  // 2) top 10
  h.push('<h2>2. ' + (admin ? 'Top 10 มูลค่าสูงสุด — ควรเร่งนำไปใช้ก่อน' : 'Top 10 ค้างนานที่สุด') + '</h2><table><tr>' + th('#', 'center') + th('CODE') + th('รายการ') + th('คงเหลือ', 'right') + th('ไม่เคลื่อนไหว', 'right') + (admin ? th('มูลค่า', 'right') : '') + '</tr>');
  top.forEach(function (x, i) { h.push('<tr>' + td(String(i + 1), 'center') + td('<b>' + esc_(x.code) + '</b>') + td(esc_(x.desc || '-')) + td(fmt(x.qty) + ' ' + esc_(x.unit || ''), 'right') + td(x.still != null ? x.still + ' วัน' : '-', 'right') + (admin ? td(money(x.value), 'right', 'color:#9f1239;font-weight:bold;') : '') + '</tr>'); });
  h.push('</table>');
  // 3) group summary
  if (group !== 'none') {
    h.push('<h2>3. สรุปตามหมวด (' + esc_(groupTxt) + ')</h2><table><tr>' + th(group === 'prefix4' ? 'หมวด (4 หลักแรก)' : 'หมวด') + th('จำนวนรายการ', 'right') + th('มีผู้จอง', 'right') + th('ค้างนานสุด', 'right') + (admin ? th('มูลค่ารวม', 'right') + th('สัดส่วน', 'right') : '') + '</tr>');
    order.forEach(function (k) { var g = groups[k]; h.push('<tr>' + td('<b>' + esc_(g.name) + '</b>') + td(fmt(g.items.length), 'right') + td(fmt(g.claimed), 'right') + td(g.oldest ? g.oldest + ' วัน' : '-', 'right') + (admin ? td(money(g.value), 'right') + td(total ? (g.value / total * 100).toFixed(1) + '%' : '-', 'right') : '') + '</tr>'); });
    h.push('</table>');
  }
  // 4) detail
  h.push('<h2>' + (group !== 'none' ? '4' : '3') + '. รายละเอียด</h2>');
  order.forEach(function (k) {
    var g = groups[k];
    if (group !== 'none') h.push('<div style="margin:10px 0 4px;font-weight:bold;font-size:10pt;color:' + C + ';">หมวด ' + esc_(g.name) + ' · ' + g.items.length + ' รายการ' + (admin ? ' · ' + money(g.value) : '') + '</div>');
    h.push('<table><tr>' + th('CODE') + th('รายการ / Spec') + th('คงเหลือ', 'right') + th('คลัง') + th('ไม่เคลื่อนไหว', 'right') + th('เหตุผล / ป้าย') + th('ผู้จอง') + (admin ? th('มูลค่า', 'right') : '') + '</tr>');
    g.items.forEach(function (x) {
      var tg = x.tag ? [x.replaceBy ? 'ใช้แทน → ' + x.replaceBy : '', x.tag.useBefore ? 'ใช้ให้หมดก่อน ' + x.tag.useBefore : '', x.tag.note || ''].filter(String).map(esc_).join('<br>') : '';
      var cl = (claims[x.key] || []).map(function (c) { return esc_(c.name) + (c.qty ? ' (' + esc_(c.qty) + ')' : '') + ': ' + esc_(c.purpose); }).join('<br>');
      h.push('<tr>' + td('<b>' + esc_(x.code) + '</b>') + td(esc_(x.desc || '-') + (x.comment ? '<br><span style="color:' + M + ';">' + esc_(x.comment) + '</span>' : '')) +
        td(fmt(x.qty) + ' ' + esc_(x.unit || ''), 'right') + td(esc_((x.locs || []).map(function (l) { return l[0] + ' ' + fmt(l[1]); }).join(', '))) +
        td((x.still != null ? x.still + ' วัน' : '-') + (x.rqnAge != null ? '<br><span style="color:' + M + ';">RQN ' + x.rqnAge + ' ปี</span>' : ''), 'right') +
        td(esc_((x.why || []).map(function (w) { return WHY[w] || w; }).join(', ')) + (tg ? '<br><span style="color:' + M + ';">' + tg + '</span>' : '')) +
        td(cl || '-') + (admin ? td(money(x.value), 'right') : '') + '</tr>');
    });
    h.push('</table>');
  });
  if (!list.length) h.push('<p>ไม่มีรายการตามเงื่อนไขที่เลือก</p>');
  h.push('<p style="font-size:8pt;color:' + M + ';margin-top:14px;">หมายเหตุ: "ไม่เคลื่อนไหว" = จำนวนคงเหลือรวมทุกคลังไม่เปลี่ยนแปลง นับจากประวัติที่ระบบบันทึกทุกครั้งที่ซิงค์ Stock · "ถูกใช้" นับจากจำนวนที่ลดลงของรายการช่วยใช้' + (admin ? ' · มูลค่า = คงเหลือ × LASTCOST (เอกสารนี้มีข้อมูลราคา — ใช้ภายในเท่านั้น)' : '') + '</p></body></html>');
  var name = 'RemainStock_' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd_HHmm') + '_' + group + (opt.scope && opt.scope !== 'all' ? '_' + opt.scope : '') + '.pdf';
  var pdf = Utilities.newBlob(h.join(''), 'text/html', 'r.html').getAs('application/pdf').setName(name);
  audit_([[admin ? 'Admin' : 'User', 'ออกรายงาน PDF Stock ช่วยใช้', '', groupTxt + ' · ' + scopeTxt + ' · ' + list.length + ' รายการ']]);
  perf_('exportRemainReport', t0);
  return { name: name, mime: 'application/pdf', b64: Utilities.base64Encode(pdf.getBytes()), count: list.length };
}
