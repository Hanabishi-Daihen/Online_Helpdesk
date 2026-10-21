/**
 * ═════════════════════════════════════════════════════════════════════
 *  WORKSPACE — มุมมองรวมข้อมูลข้ามฟังก์ชัน (อ่านอย่างเดียว ใช้ snapshot/cache ที่มีอยู่)
 *  • getItem360(token, code)   แผงข้อมูล CODE: คลัง/PK, กำลังเข้า, จอง BOM, ใบสั่ง, Stock Alert,
 *                              Stock ช่วยใช้, ประวัติคงเหลือ (StockChangeLog) และคาดวันหมด
 *  • getMyDay(token)           หน้า "งานของฉัน" ทุกบทบาทในครั้งเดียว (จัดซื้อ / คลัง / ผู้อนุมัติ / Admin)
 *  • getProjects(token)        รายการโปรเจกต์ (ตารางงาน + Booking BOM) · getProject360(token, base) รายละเอียด
 *  • wsLogChanges_()           จดเฉพาะ CODE ที่จำนวนรวมเปลี่ยน ทุกครั้งที่ซิงค์ Stock (เรียกจาก rsTrack_)
 *  ต้องมี Code.gs, OrderService.gs, OrderApproval.gs, StockAlert.gs, RemainStock.gs, BomBooking.gs, BomExport.gs
 * ═════════════════════════════════════════════════════════════════════
 */
var WS_LOG_SHEET = 'StockChangeLog', WS_LOG_MAX = 300000, WS_LOG_KEEP = 240000;

// ── stock change log (Helpdesk_ItemStock → StockChangeLog: DATE, CODE, QTY) ──
// old = { CODE: [code, qty, since, lastSeen] } from StockHistory before this sync, tot = { CODE: qty } now
function wsLogChanges_(ss, old, tot) {
  var today = ordToday_(), add = [], first = !Object.keys(old).length;
  Object.keys(tot).forEach(function (k) {
    var q = String(tot[k]), o = old[k];
    if (first || !o || o[1] !== q) add.push([today, k, q]);
  });
  Object.keys(old).forEach(function (k) { if (!(k in tot) && old[k][1] !== '0') add.push([today, k, '0']); });
  if (!add.length) return 0;
  var sh = ss.getSheetByName(WS_LOG_SHEET);
  if (!sh) { sh = ss.insertSheet(WS_LOG_SHEET); sh.getRange(1, 1, 1, 3).setValues([['DATE', 'CODE', 'QTY']]); sh.setFrozenRows(1); }
  var at = sh.getLastRow() + 1;
  if (sh.getMaxRows() < at + add.length) sh.insertRowsAfter(sh.getMaxRows(), at + add.length - sh.getMaxRows());
  var rg = sh.getRange(at, 1, add.length, 3); rg.setNumberFormat('@'); rg.setValues(add);
  var n = sh.getLastRow() - 1;
  if (n > WS_LOG_MAX) sh.deleteRows(2, n - WS_LOG_KEEP);               // keep the newest rows
  return add.length;
}
function wsHistory_(k) {
  var out = [];
  try {
    var sh = stockSS_().getSheetByName(WS_LOG_SHEET);
    if (!sh || sh.getLastRow() < 2) return out;
    var hits = sh.getRange(2, 2, sh.getLastRow() - 1, 1).createTextFinder(k).matchEntireCell(true).findAll();
    hits.slice(-120).forEach(function (r) {
      var v = sh.getRange(r.getRow(), 1, 1, 3).getDisplayValues()[0];
      var t = bomDate_(v[0]); if (t != null) out.push([t, Number(v[2]) || 0]);
    });
  } catch (e) { console.warn('wsHistory_', e); }
  return out.sort(function (a, b) { return a[0] - b[0]; });
}
function wsToday_() { var d = new Date(); return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime(); }
function wsPendingRows_(match) {
  var out = [];
  try {
    var b = pendingBundle_(false), txt = b && b.b64 ? Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(b.b64), 'application/x-gzip')).getDataAsString() : '';
    txt.split('\n').forEach(function (l) {
      var t = l.indexOf('\t'); if (t < 0) return;
      var f = l.split('\t'); if (match(ordCodeKey_(l.substring(0, t)), f)) out.push(f);
    });
  } catch (e) {}
  return out;
}
function wsOrderState_(o) {
  var st = (o.approval || {}).state || '';
  return { key: st || (o.status === 'Exported' ? 'exported' : 'saved'), label: OA_STATE_TH[st] && st ? OA_STATE_TH[st] : (o.status === 'Exported' ? 'ส่งออกแล้ว' : 'บันทึกแล้ว') };
}

// ── Item 360 ─────────────────────────────────────────────────────────
function getItem360(token, code) {
  var t0 = Date.now(), k = ordCodeKey_(code), admin = isAdmin_(token);
  if (!k) throw new Error('CODE ไม่ถูกต้อง');
  var out = { code: String(code).trim(), key: k, found: false, desc: '', comment: '', unit: '', by: '', rqn: '', onhand: 0, locs: [], admin: admin };
  try {
    var rows = stockData_();
    for (var i = 0; i < rows.length; i++) {
      var p = rows[i], t = p.indexOf('\t'); if (t < 0 || ordCodeKey_(p.substring(0, t)) !== k) continue;
      var f = p.split('\t'), q = saNum_(f[3]) || 0;
      out.found = true; out.code = String(f[0]).trim();
      out.desc = out.desc || f[1] || ''; out.comment = out.comment || f[2] || ''; out.unit = out.unit || f[4] || '';
      out.by = out.by || String(f[7] || '').trim(); out.rqn = out.rqn || f[9] || '';
      if (admin && !out.cost && saNum_(f[8])) out.cost = saNum_(f[8]);
      out.onhand = Math.round((out.onhand + q) * 1000) / 1000;
      out.locs.push({ loc: String(f[5] || '').trim(), pk: String(f[6] || '').trim(), qty: q });
    }
    out.locs.sort(function (a, b) { return b.qty - a.qty; });
  } catch (e) {}
  // Pending Order: open (≤ 1 year) and recent closed, plus the planned lead time of closed POs
  var today = wsToday_(), YEAR = 365 * 864e5, open = [], done = [], lts = [];
  wsPendingRows_(function (kk) { return kk === k; }).forEach(function (g) {
    var due = bomDate_(g[9]) || bomDate_(g[7]), req = bomDate_(g[6]);
    var rec = { pr: g[1], po: g[2], status: g[4], due: g[9] || g[7], req: g[6], qty: Number(g[13]) || 0, unit: g[14], by: g[12] };
    if (ORD_PO_DONE.test(g[4] || '')) { done.push(rec); if (req != null && due != null && due > req) lts.push((due - req) / 864e5); return; }
    if (due != null && Math.abs(due - today) > YEAR) return;
    rec.t = due; open.push(rec);
  });
  open.sort(function (a, b) { return (a.t == null ? 9e15 : a.t) - (b.t == null ? 9e15 : b.t); });
  done.sort(function (a, b) { return (bomDate_(b.req) || 0) - (bomDate_(a.req) || 0); });
  out.incoming = open.slice(0, 8); out.incomingQty = open.reduce(function (s, x) { return s + x.qty; }, 0);
  out.received = done.slice(0, 5);
  lts = lts.slice(0, 6); out.ltPlan = lts.length ? Math.round(lts.reduce(function (s, x) { return s + x; }, 0) / lts.length) : null; out.ltN = lts.length;
  // Booking BOM
  try {
    var snap = bomSnapGet_() || {}, need = {}, c = null;
    (snap.projects || []).forEach(function (pr) {
      if (pr.status !== 'active') return;
      pr.items.forEach(function (it) { if (ordCodeKey_(it.code) === k) need[pr.project] = { project: pr.project, needDate: pr.needDate, daysLeft: pr.daysLeft, qty: it.qty, st: it.st, short: it.short || 0 }; });
    });
    (snap.codes || []).forEach(function (x) { if (x.code === k) c = x; });
    var list = Object.keys(need).map(function (x) { return need[x]; }).sort(function (a, b) { return (a.daysLeft == null ? 9e9 : a.daysLeft) - (b.daysLeft == null ? 9e9 : b.daysLeft); });
    out.bom = c || list.length ? { booked: c ? c.booked : 0, balance: c ? c.balance : null, short: list.reduce(function (s, x) { return s + x.short; }, 0), projects: list } : null;
  } catch (e) { out.bom = null; }
  // Order Item (newest first)
  try {
    var ords = [];
    ordAll_().forEach(function (o) {
      var st = wsOrderState_(o); if (st.key === 'rejected') return;
      (o.items || []).forEach(function (it) {
        if (ordCodeKey_(it.code) === k) ords.push({ id: o.id, spec: o.spec, rqn: o.rqn, date: o.issueDate, qty: it.qty, unit: it.unit, state: st.key, stateTh: st.label, remark: o.remark });
      });
    });
    out.orders = ords.reverse().slice(0, 6);
  } catch (e) { out.orders = []; }
  // Stock Alert
  try {
    var sa = saData_(false), hit = null;
    sa.items.forEach(function (x) { if (ordCodeKey_(x.code) === k) hit = x; });
    out.sa = hit ? { level: hit.level, min: hit.min, max: hit.max, lt: hit.lt, ltSrc: hit.ltSrc, suggest: hit.suggest, notUse: hit.notUse, note: hit.note } : null;
    out.saBy = sa.cfg.by;
  } catch (e) { out.sa = null; }
  // Stock ช่วยใช้ (this CODE or a substitute for it)
  try { out.remain = rsForOrder_([k])[k] || []; } catch (e) { out.remain = []; }
  // history + forecast
  var hist = wsHistory_(k), since = '';
  try { since = rsHist_((function () { var w = {}; w[k] = 1; return w; })())[k] || ''; } catch (e) {}
  out.since = since; out.histStart = schedProp_('RS_HIST_START') || '';
  var cut = today - 90 * 864e5, pts = hist.filter(function (x) { return x[0] >= cut; }), before = hist.filter(function (x) { return x[0] < cut; }).pop();
  if (before) pts.unshift([cut, before[1]]);
  if (!pts.length || pts[pts.length - 1][0] < today) pts.push([today, out.onhand]);
  out.series = pts.map(function (x) { return [bomFmt_(x[0]), x[1]]; });
  var used = 0; for (var j = 1; j < pts.length; j++) if (pts[j][1] < pts[j - 1][1]) used += pts[j - 1][1] - pts[j][1];
  var span = pts.length > 1 ? (pts[pts.length - 1][0] - pts[0][0]) / 864e5 : 0;
  out.usePerDay = span >= 14 && used > 0 ? Math.round(used / span * 100) / 100 : null;
  out.spanDays = Math.round(span);
  var free = out.onhand - (out.bom ? out.bom.booked || 0 : 0);
  out.daysLeft = out.usePerDay ? Math.max(0, Math.floor(free / out.usePerDay)) : null;
  out.suggestOrder = Math.max(out.bom ? Math.ceil(out.bom.short || 0) : 0, out.sa ? out.sa.suggest || 0 : 0);
  if (admin && out.cost) out.value = Math.round(out.onhand * out.cost);
  perf_('getItem360', t0);
  return out;
}

// ── My day (all roles in one call; the browser shows the chosen role) ──
function getMyDay(token) {
  var t0 = Date.now(), admin = isAdmin_(token), today = wsToday_(), out = { time: nowStr_(), admin: admin };
  var cfg = bomCfg_(), ob = String(cfg.orderBy || 'DS13').toUpperCase();
  out.orderBy = cfg.orderBy || 'DS13';
  // Stock Alert
  try {
    var sa = saData_(false);
    out.sa = { count: sa.count, by: sa.cfg.by, top: sa.items.filter(function (x) { return x.level === 'high' || x.level === 'mid'; }).slice(0, 8)
      .map(function (x) { return { code: x.code, desc: x.desc, qty: x.qty, unit: x.unit, min: x.min, lt: x.lt, level: x.level, suggest: x.suggest, pend: x.pend }; }) };
  } catch (e) { out.sa = null; }
  // Booking BOM
  var snap = null; try { snap = bomSnapGet_(); } catch (e) {}
  var projs = (snap && snap.projects || []).filter(function (p) { return p.status === 'active'; });
  out.bomShort = [];
  projs.forEach(function (p) { p.items.forEach(function (x) {
    if (x.st === 'short') out.bomShort.push({ project: p.project, needDate: p.needDate, daysLeft: p.daysLeft, code: x.code, material: x.material, short: x.short, unit: x.unit, by: x.by || '', mine: String(x.by || '').toUpperCase() === ob });
  }); });
  out.bomShort.sort(function (a, b) { return (a.daysLeft == null ? 9e9 : a.daysLeft) - (b.daysLeft == null ? 9e9 : b.daysLeft); });
  var exp = {}; try { exp = bomExpLast_(); } catch (e) {}
  out.projects = projs.map(function (p) { return { project: p.project, title: p.title, needDate: p.needDate, daysLeft: p.daysLeft, count: p.count, total: p.total, last: (exp[p.project] || [])[0] || null }; })
    .sort(function (a, b) { return (a.daysLeft == null ? 9e9 : a.daysLeft) - (b.daysLeft == null ? 9e9 : b.daysLeft); });
  // orders (last 120 days) — the browser picks "mine" by the DES'D name it remembers
  out.orders = [];
  try {
    ordAll_().forEach(function (o) {
      if (ordAgeDays_(o.issueDate) > 120) return;
      var st = wsOrderState_(o), a = o.approval || {};
      out.orders.push({ id: o.id, spec: o.spec, rqn: o.rqn, des: o.des, chk: o.chk, app: o.app, date: o.issueDate, items: (o.items || []).length, state: st.key, stateTh: st.label,
        sent: a.sent || '', remark: o.remark || '' });
    });
    out.orders.reverse();
  } catch (e) {}
  // incoming this week (Pending Order due in the next 7 days, not closed)
  var wk = [];
  wsPendingRows_(function (k, f) { return !ORD_PO_DONE.test(f[4] || ''); }).forEach(function (f) {
    var d = bomDate_(f[9]) || bomDate_(f[7]); if (d == null || d < today - 864e5 || d > today + 7 * 864e5) return;
    wk.push({ code: f[0], name: f[5], po: f[2] || f[1], due: f[9] || f[7], qty: Number(f[13]) || 0, unit: f[14], by: f[12], t: d });
  });
  wk.sort(function (a, b) { return a.t - b.t; });
  out.incoming = wk.slice(0, 30); out.incomingN = wk.length;
  out.incomingMine = wk.filter(function (x) { return String(x.by || '').toUpperCase() === ob; }).length;
  // Stock ช่วยใช้ claims waiting to be handed out
  try { var cl = rsClaims_(true), n = 0; Object.keys(cl).forEach(function (k) { n += cl[k].length; }); out.claims = n; } catch (e) { out.claims = 0; }
  // Admin: cases + system
  if (admin) {
    try {
      var s = ss_().getSheetByName(CONFIG.ISSUES), open = 0, old = 0;
      if (s && s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 14).getDisplayValues().forEach(function (r) {
        if (!r[0] || String(r[13]).trim() === '1' || /^(Resolved|Closed)$/.test(r[8])) return;
        open++; var d = dateOf_(r[7]); if (d && (today - d.getTime()) / 864e5 > 7) old++;
      });
      out.cases = { open: open, old: old };
    } catch (e) {}
    out.sync = { stock: stockInfo_(), po: poInfo_(), sched: schedSyncInfo_(), rt: rtInfo_() };
    try { out.backup = JSON.parse(schedProp_('LAST_BACKUP') || 'null'); } catch (e) {}
    try { out.mailQuota = MailApp.getRemainingDailyQuota(); } catch (e) {}
  }
  perf_('getMyDay', t0);
  return out;
}

// ── Project 360 ─────────────────────────────────────────────────────
// base = schedule ProjectNo (PP1024); BOM projects are matched by prefix (PP1024D01, PP1024D02)
function wsBase_(p) { var k = bomKey_(p), m = /^([A-Z]+\d+)[A-Z]/.exec(k); return m ? m[1] : k; }
function getProjects(token) {
  var t0 = Date.now(), sched = [], by = {}, today = wsToday_();
  try { sched = getSchedules(); } catch (e) {}
  var get = function (b) { return by[b] || (by[b] = { base: b, customer: '', mva: '', next: null, nextLabel: '', delivery: '', bom: [], total: 0, ready: 0, short: 0, late: 0 }); };
  sched.forEach(function (r) {
    var b = bomKey_(r[0]); if (!b) return;
    var P = get(b); P.customer = P.customer || r[4]; P.mva = P.mva || r[1]; P.delivery = P.delivery || r[7] || r[6] || r[5];
    var d = bomDate_(r[9]);
    if (d != null && d >= today && (P.next == null || d < P.next)) { P.next = d; P.nextLabel = String(r[8]).trim() + ' Start'; }
  });
  var snap = null; try { snap = bomSnapGet_(); } catch (e) {}
  (snap && snap.projects || []).forEach(function (p) {
    var b = wsBase_(p.project), hit = null;
    Object.keys(by).forEach(function (x) { if (bomKey_(p.project).indexOf(x) === 0 && (!hit || x.length > hit.length)) hit = x; });
    var P = get(hit || b);
    P.bom.push({ project: p.project, status: p.status });
    if (p.status !== 'active') return;
    var c = p.count || {}; P.total += p.total || 0; P.ready += (c.ok || 0) + (c.incoming || 0); P.short += c.short || 0; P.late += c.late || 0;
  });
  var list = Object.keys(by).map(function (b) { var P = by[b]; P.pct = P.total ? Math.round(P.ready / P.total * 100) : null; P.nextDate = bomFmt_(P.next);
    P.daysLeft = P.next != null ? Math.round((P.next - today) / 864e5) : null; return P; })
    .filter(function (P) { return P.bom.length || P.next != null; })
    .sort(function (a, b) { return (a.bom.some(function (x) { return x.status === 'active'; }) ? 0 : 1) - (b.bom.some(function (x) { return x.status === 'active'; }) ? 0 : 1) ||
      (a.next == null ? 9e15 : a.next) - (b.next == null ? 9e15 : b.next); });
  perf_('getProjects', t0);
  return { projects: list.slice(0, 200), built: snap && snap.built || '' };
}
function getProject360(token, base) {
  var t0 = Date.now(), B = bomKey_(base), today = wsToday_(), admin = isAdmin_(token);
  if (!B) throw new Error('ไม่ระบุโปรเจกต์');
  var out = { base: B, admin: admin, sched: [], bom: [], short: [], orders: [], exports: [], issues: [], events: [] };
  var sched = []; try { sched = getSchedules(); } catch (e) {}
  sched.forEach(function (r) {
    if (bomKey_(r[0]) !== B) return;
    out.sched.push({ process: String(r[8]).trim(), start: r[9], finish: r[10], mva: r[1], volt: r[2], customer: r[4], purchaser: r[3], delivery: r[7] || r[6] || r[5] });
  });
  var info = out.sched[0] || {}; out.customer = info.customer || ''; out.mva = info.mva || ''; out.volt = info.volt || ''; out.delivery = info.delivery || '';
  out.sched.forEach(function (s) { var d = bomDate_(s.start); if (d != null) out.events.push({ t: d, label: s.process + ' Start', kind: 'sched' }); });
  if (bomDate_(out.delivery) != null) out.events.push({ t: bomDate_(out.delivery), label: 'Delivery', kind: 'delivery' });
  var snap = null; try { snap = bomSnapGet_(); } catch (e) {}
  var exp = {}; try { exp = bomExpLast_(); } catch (e) {}
  var cfg = bomCfg_(), ob = String(cfg.orderBy || 'DS13').toUpperCase(), codes = {};
  (snap && snap.projects || []).forEach(function (p) {
    if (bomKey_(p.project).indexOf(B) !== 0) return;
    out.bom.push({ project: p.project, rev: p.rev, title: p.title, status: p.status, needDate: p.needDate, daysLeft: p.daysLeft, count: p.count, total: p.total, imported: p.imported, closed: p.closed });
    var imp = bomDate_(p.imported); if (imp != null) out.events.push({ t: imp, label: 'Import BOM ' + p.project.substring(B.length), kind: 'bom' });
    (exp[p.project] || []).forEach(function (e) {
      out.exports.push({ project: p.project, t: e.t, kind: e.k, n: e.n, pdf: e.pdf, xlsx: e.xlsx });
      var d = bomDate_(e.t); if (d != null) out.events.push({ t: d, label: 'Export ใบเบิก ' + p.project.substring(B.length), kind: 'export' });
    });
    if (p.status !== 'active') return;
    p.items.forEach(function (x) {
      if (x.st !== 'short' && x.st !== 'late') return;
      var k = ordCodeKey_(x.code), S = codes[k];
      if (!S) { S = codes[k] = { code: x.code, material: x.material, unit: x.unit, by: x.by || '', mine: String(x.by || '').toUpperCase() === ob, short: 0, late: 0, eta: '', sets: [], orders: [] }; out.short.push(S); }
      S.short += x.short || 0; S.late += x.late || 0; if (x.eta) S.eta = x.eta; S.sets.push(p.project.substring(B.length) || p.project);
    });
  });
  var tot = 0, ready = 0; out.bom.forEach(function (b) { if (b.status !== 'active') return; var c = b.count || {}; tot += b.total || 0; ready += (c.ok || 0) + (c.incoming || 0); });
  out.pct = tot ? Math.round(ready / tot * 100) : null;
  out.short.sort(function (a, b) { return (b.short ? 1 : 0) - (a.short ? 1 : 0) || b.short - a.short; });
  // orders for this project (remark FOR PROJECT …) or holding its short CODEs
  try {
    ordAll_().forEach(function (o) {
      var st = wsOrderState_(o); if (st.key === 'rejected') return;
      var forMe = bomKey_(o.remark).indexOf('FORPROJECT' + B) >= 0, mine = [];
      (o.items || []).forEach(function (it) { var S = codes[ordCodeKey_(it.code)]; if (S) { mine.push(it.code); S.orders.push({ id: o.id, stateTh: st.label, qty: it.qty }); } });
      if (!forMe && !mine.length) return;
      out.orders.push({ id: o.id, spec: o.spec, date: o.issueDate, items: (o.items || []).length, state: st.key, stateTh: st.label, codes: mine });
      var d = bomDate_(o.issueDate); if (d != null) out.events.push({ t: d, label: 'ใบสั่ง ' + o.id, kind: 'order' });
    });
    out.orders.reverse();
  } catch (e) {}
  // cases that mention the project
  try {
    var s = ss_().getSheetByName(CONFIG.ISSUES);
    if (s && s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 14).getDisplayValues().forEach(function (r) {
      if (!r[0] || String(r[13]).trim() === '1') return;
      if (bomKey_(r[1] + ' ' + r[2]).indexOf(B) < 0) return;
      out.issues.push({ id: r[0], project: r[1], desc: String(r[2]).substring(0, 120), status: r[8] || 'New', date: r[7] });
    });
    out.issues = out.issues.reverse().slice(0, 8);
  } catch (e) {}
  out.events.push({ t: today, label: 'วันนี้', kind: 'today' });
  out.events.sort(function (a, b) { return a.t - b.t; });
  out.events = out.events.map(function (e) { return { date: bomFmt_(e.t), label: e.label, kind: e.kind, days: Math.round((e.t - today) / 864e5) }; });
  perf_('getProject360', t0);
  return out;
}
