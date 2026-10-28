/**
 * ═════════════════════════════════════════════════════════════════════
 *  TRACKING — พยากรณ์ของขาด · ติดตามใบสั่งครบวงจร · แจ้งเตือนของเข้า
 *  • F1 พยากรณ์ (fcCode_)   ต่อ CODE ที่มีการจอง BOM: เดินยอดตามวันที่ — Onhand + PO ที่จะเข้า (ตามกำหนดส่ง)
 *                            − ของที่โปรเจกต์ต้องใช้ (ตามวันที่ต้องใช้) → วันที่ยอดติดลบ, โปรเจกต์ที่กระทบ,
 *                            ควรสั่งภายใน (= วันที่ขาด − Lead time) · คำนวณพร้อม snapshot ของ Booking BOM
 *                            Lead time: ตั้งใน Stock Alert → Readtime → ค่าเริ่มต้น (Admin ตั้งได้)
 *  • F4 ติดตามใบสั่ง          ใบสั่ง (Order Item) → อนุมัติ → พบใน Pending Order (RQN / Spec ตรงกัน) → ออก PO
 *                            → ของเข้าบางส่วน → ครบ · สถานะราย CODE จาก Status ของ Pending Order
 *  • F5 แจ้งเตือนในระบบ       ทุกครั้งที่ซิงค์ Stock / Pending Order เทียบกับรอบก่อน แล้วจดลงชีต Notify
 *                            (ออก PO แล้ว · ของเข้าแล้ว · Stock เพิ่มสำหรับ BOM / ใบสั่ง / Stock Alert) — ไม่ส่งอีเมล
 *  ต้องมี Code.gs, OrderService.gs, OrderApproval.gs, StockAlert.gs, BomBooking.gs, Workspace.gs
 * ═════════════════════════════════════════════════════════════════════
 */
var FC_DEF = { lt: 30 };
var TRK_STATE_SHEET = 'TrackState', TRK_CHUNK = 40000;
var NTF_SHEET = 'Notify', NTF_HEADER = ['TS', 'AT', 'TYPE', 'CODE', 'TITLE', 'DETAIL', 'REF', 'WHO', 'PROJECT'], NTF_MAX = 3000, NTF_KEEP = 2000;
var TRK_CANCEL = /cancel|ยกเลิก/i;

// ── F1: forecast ────────────────────────────────────────────────────
function fcCfg_() {
  var c = {}; try { c = JSON.parse(schedProp_('FC_CFG') || '{}') || {}; } catch (e) {}
  var lt = Number(c.lt);
  return { lt: lt >= 1 && lt <= 365 ? Math.round(lt) : FC_DEF.lt };
}
function saveFcCfg(token, cfg) {
  requireAdmin_(token);
  var lt = Number((cfg || {}).lt);
  if (!(lt >= 1 && lt <= 365)) throw new Error('Lead time เริ่มต้นต้องอยู่ระหว่าง 1–365 วัน');
  schedProp_('FC_CFG', JSON.stringify({ lt: Math.round(lt) }));
  audit_([['Admin', 'ตั้งค่าพยากรณ์ของขาด', '', 'Lead time เริ่มต้น ' + Math.round(lt) + ' วัน']]);
  bomSnapBuild_();
  return fcCfg_();
}
// onhand: now · ins: [{qty, due}] open POs · dl: [{t, p:{project}, it:{qty}}] bookings · min: Stock Alert Min (or null)
// → null when the CODE never runs short / below Min, else { k: 'short'|'late'|'min', … }
function fcCode_(onhand, ins, dl, min, lt, ltSrc, today) {
  var ev = [], noDate = 0, poNoDate = 0;
  ins.forEach(function (x) {
    if (x.due == null) { poNoDate += x.qty; return; }
    ev.push({ t: Math.max(x.due, today), q: x.qty, o: 0, label: 'PO ' + (x.po || x.pr || '-') + (x.due < today ? ' (เลยกำหนดส่ง ' + bomFmt_(x.due) + ')' : '') });
  });
  dl.forEach(function (d) {
    if (d.t == null) { noDate += d.it.qty; return; }
    ev.push({ t: Math.max(d.t, today), q: -d.it.qty, o: 1, label: d.p.project, p: d.p.project });
  });
  ev.sort(function (a, b) { return a.t - b.t || a.o - b.o; });           // same day: arrivals first (as the booking does)
  var bal = Math.max(0, onhand), low = bal, short = null, rec = null, below = min != null && bal < min ? today : null, hit = {}, out = [];
  ev.forEach(function (e) {
    bal = Math.round((bal + e.q) * 1000) / 1000;
    if (bal < low) low = bal;
    if (bal < 0 && !short) short = { t: e.t, q: -bal, p: e.p || '' };
    if (bal < 0 && e.p) hit[e.p] = 1;
    if (short && rec == null && bal >= 0 && e.t > short.t) rec = e.t;
    if (min != null && bal < min && below == null) below = e.t;
    out.push([bomFmt_(e.t), e.q, bal, e.label]);
  });
  var perm = bal < 0 ? -bal : 0, k = perm > 0 ? 'short' : short ? 'late' : below != null ? 'min' : '';
  if (!k) return null;
  var r = { k: k, lt: lt, ltSrc: ltSrc, end: bal, low: low, perm: Math.ceil(perm * 1000) / 1000, noDate: noDate, poNoDate: poNoDate, ev: out.slice(0, 16) };
  if (short) { r.t = bomFmt_(short.t); r.q = short.q; r.p = short.p; r.ps = Object.keys(hit); r.by = bomFmt_(short.t - lt * 864e5); }
  if (rec != null) r.rec = bomFmt_(rec);
  if (below != null) { r.below = bomFmt_(below); if (!short) r.by = bomFmt_(below - lt * 864e5); }
  if (min != null) r.min = min;
  return r;
}
// public: every CODE that will run short (or below Min) — from the Booking BOM snapshot
function getForecast(token) {
  var t0 = Date.now(), out = fcCached_();
  out.admin = isAdmin_(token);
  perf_('getForecast', t0);
  return out;
}
function fcCached_() { return wsCached_('fc', ['bom', 'ord', 'oa', 'sa'], schedProp_('FC_CFG') || '', fcList_); }
function fcList_() {
  var snap = null; try { snap = bomSnapGet_(); } catch (e) {}
  var cfg = bomCfg_(), ob = String(cfg.orderBy || 'DS13').toUpperCase(), list = [], want = {};
  var need = {};                                                      // project → need date
  (snap && snap.projects || []).forEach(function (p) { need[p.project] = p.needDate; });
  (snap && snap.codes || []).forEach(function (c) {
    if (!c.fc) return;
    want[c.code] = list.length;
    list.push({ code: c.code, desc: c.desc, unit: c.unit, by: c.by || '', mine: String(c.by || '').toUpperCase() === ob, onhand: c.onhand, incoming: c.incoming,
      booked: c.booked, fc: c.fc, need: (c.fc.ps || []).map(function (p) { return [p, need[p] || '']; }), po: c.po || [], orders: [] });
  });
  try {                                                               // Order Items already holding these CODEs (not rejected, ≤ 120 days)
    ordAll_().forEach(function (o) {
      var st = wsOrderState_(o); if (st.key === 'rejected' || ordAgeDays_(o.issueDate) > 120) return;
      (o.items || []).forEach(function (it) { var i = want[ordCodeKey_(it.code)]; if (i != null && list[i].orders.length < 3) list[i].orders.push({ id: o.id, stateTh: st.label, state: st.key, qty: it.qty }); });
    });
  } catch (e) {}
  return { items: list, built: snap && snap.built || '', cfg: fcCfg_(), orderBy: cfg.orderBy || 'DS13', time: nowStr_() };
}

// ── F4: order tracking ──────────────────────────────────────────────
// Pending Order rows indexed by RQN / Spec (the numbers the requester writes on the order)
function trkPendIdx_() {
  var idx = {}, byCode = {}, n = 0;
  try {
    var b = pendingBundle_(false), txt = b && b.b64 ? Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(b.b64), 'application/x-gzip')).getDataAsString() : '';
    txt.split('\n').forEach(function (l) {
      var f = l.split('\t'); if (f.length < 14) return;
      n++;
      var r = { id: n, code: ordCodeKey_(f[0]), pr: f[1], po: f[2], spec: f[3], status: f[4], name: f[5], req: f[6], poDate: f[8], due: f[9] || f[7], by: f[12], qty: Number(f[13]) || 0, unit: f[14] };
      (byCode[r.code] = byCode[r.code] || []).push(r);
      [f[1], f[3]].forEach(function (v) { var k = ordRefKey_(v); if (k) (idx[k] = idx[k] || []).push(r); });
    });
  } catch (e) { console.warn('trkPendIdx_', e); }
  return { idx: idx, byCode: byCode, n: n };
}
function trkRefOk_(v) { return /\d$/.test(String(v || '').trim()); }   // "RQN26-" / "DS1C26-" / "-" = not filled in yet
function trkRows_(o, idx) {
  var seen = {}, rows = [];
  [o.rqn, o.spec].forEach(function (v) {
    if (!trkRefOk_(v)) return;                                        // placeholders never match
    (idx[ordRefKey_(v)] || []).forEach(function (r) { if (!seen[r.id]) { seen[r.id] = 1; rows.push(r); } });
  });
  return rows;
}
// one CODE of an order: none → pr (in Pending) → po (PO number) → part (some lines received) → recv · cancel · gone (left Pending after being seen)
function trkItemSt_(rs) {
  if (!rs.length) return 'none';
  var live = rs.filter(function (r) { return !TRK_CANCEL.test(r.status || ''); });
  if (!live.length) return 'cancel';
  var done = live.filter(function (r) { return ORD_PO_DONE.test(r.status || ''); }).length;
  if (done === live.length) return 'recv';
  if (done) return 'part';
  return live.some(function (r) { return String(r.po || '').trim(); }) ? 'po' : 'pr';
}
var TRK_RANK = { none: 0, pr: 1, po: 2, part: 3, recv: 4, gone: 4, cancel: 5 };
function trkOrder_(o, idx, prev, today, opt, auto) {
  var a = o.approval || {}, st = wsOrderState_(o), rows = auto || trkRows_(o, idx), P = prev || {};
  opt = opt || {};
  var items = (o.items || []).filter(function (it) { return it.code || it.material; }).map(function (it) {
    var k = ordCodeKey_(it.code), rs = rows.filter(function (r) { return r.code === k; }), s = trkItemSt_(rs), p = P[k];
    if (s === 'none' && p && p[0] !== 'none') s = /^(pr|po|part)$/.test(p[0]) ? (opt.gone ? 'gone' : p[0]) : p[0];   // was in Pending, no longer listed → closed at the source
    var r0 = rs.filter(function (r) { return !TRK_CANCEL.test(r.status || ''); })[0] || rs[0] || {};
    var due = null; rs.forEach(function (r) { if (ORD_PO_DONE.test(r.status || '')) return; var d = bomDate_(r.due); if (d != null && (due == null || d > due)) due = d; });
    return { code: it.code, material: it.material, qty: it.qty, unit: it.unit, st: s, pr: r0.pr || '', po: r0.po || '', status: r0.status || '', due: bomFmt_(due),
      late: due != null && due < today && /^(pr|po|part)$/.test(s) ? Math.round((today - due) / 864e5) : 0, req: r0.req || '', poDate: r0.poDate || '',
      at: p && p[0] === s ? p[1] : '' };
  });
  var live = items.filter(function (i) { return i.st !== 'cancel'; }), lo = 9, hi = 0;
  live.forEach(function (i) { var r = TRK_RANK[i.st]; if (r < lo) lo = r; if (r > hi) hi = r; });
  if (!live.length) lo = hi = 0;
  var appr = a.state === 'approved' || (!a.state && o.status === 'Exported');
  var stage = st.key === 'rejected' ? 'rejected' : items.length && !live.length ? 'cancel' : !live.length ? (appr ? 'approved' : 'saved')
    : lo === 4 ? 'done' : hi >= 3 ? 'part' : lo >= 2 ? 'po' : hi >= 1 ? 'pr' : st.key === 'chk' || st.key === 'app' ? 'wait' : appr ? 'approved' : 'saved';
  var first = function (f) { var b = null; rows.forEach(function (r) { var d = bomDate_(r[f]); if (d != null && (b == null || d < b)) b = d; }); return bomFmt_(b); };
  var recvAt = ''; if (stage === 'done') live.forEach(function (i) { if (i.at && (!recvAt || bomDate_(i.at) > bomDate_(recvAt))) recvAt = i.at; });
  var apAt = a.state === 'approved' ? (a.app && a.app.at) || '' : !a.state && o.status === 'Exported' ? String(o.updated || '').split(' ')[0] : '';
  var steps = [
    { k: 'saved', label: 'บันทึก', at: String(o.created || o.issueDate || '').split(' ')[0], done: true },
    { k: 'sent', label: a.state || !appr ? 'ส่งขออนุมัติ' : 'ส่งออกใบสั่ง', at: String(a.sent || '').split(' ')[0], done: !!a.sent || appr },
    { k: 'approved', label: 'อนุมัติ', at: String(apAt).split(' ')[0], done: appr },
    { k: 'pr', label: 'พบใน Pending (PR)', at: first('req'), done: hi >= 1 },
    { k: 'po', label: 'ออก PO', at: first('poDate'), done: live.length > 0 && lo >= 2 },
    { k: 'recv', label: 'ของเข้าครบ', at: recvAt, done: stage === 'done' }
  ];
  var apDays = appr && hi === 0 && apAt ? Math.round((today - (bomDate_(apAt) || today)) / 864e5) : 0;
  return { id: o.id, spec: o.spec, rqn: o.rqn, des: o.des, chk: o.chk, app: o.app, date: o.issueDate, remark: o.remark || '', state: st.key, stateTh: st.label,
    auto: auto ? auto.map(function (r) { return r.pr; }).filter(function (v, i, A) { return v && A.indexOf(v) === i; }) : null, refMissing: !trkRefOk_(o.rqn) && !trkRefOk_(o.spec),
    fill: auto ? trkFill_(o, auto) : null,
    stage: stage, steps: steps, items: items, late: items.reduce(function (m, i) { return Math.max(m, i.late); }, 0), noPr: apDays, waitDays: st.key === 'chk' || st.key === 'app' ? oaWorkDays_(oaSince_(a)) : null };
}
// a Pending Order file much smaller than last time (partial export) must not turn every listed order into "received"
// Orders whose RQN / Spec is not filled in ("RQN26-") are matched by CODE + request date instead:
// a Pending Order line of the same CODE requested 3 days before … 90 days after the order date, not used by any
// other order (exact matches first, then older orders first), nearest date first, same quantity preferred.
function trkAuto_(orders, P) {
  var used = {}, out = {};
  orders.forEach(function (o) { trkRows_(o, P.idx).forEach(function (r) { used[r.id] = 1; }); });
  orders.forEach(function (o) {
    if (trkRefOk_(o.rqn) || trkRefOk_(o.spec) || wsOrderState_(o).key === 'rejected') return;
    var t0 = bomDate_(o.issueDate), rs = []; if (t0 == null) return;
    (o.items || []).forEach(function (it) {
      var k = ordCodeKey_(it.code), best = null, bs = 9e15; if (!k) return;
      (P.byCode[k] || []).forEach(function (r) {
        var d = bomDate_(r.req); if (used[r.id] || d == null || d < t0 - 3 * 864e5 || d > t0 + 90 * 864e5) return;
        var sc = Math.abs(d - t0) + (Number(it.qty) === r.qty ? 0 : 30 * 864e5); if (sc < bs) { bs = sc; best = r; }
      });
      if (best) { used[best.id] = 1; rs.push(best); }
    });
    if (rs.length) out[o.id] = rs;
  });
  return out;
}
// numbers to write back into an order matched by CODE + date (only the parts that are still a bare prefix)
// sure = one PR of the order's year for every CODE, same quantities → filled automatically at the next scan;
// otherwise the tracking card offers a button (trackFillRef)
function trkFill_(o, rows) {
  var uq = function (a) { return a.filter(function (v, i) { return v && a.indexOf(v) === i; }); };
  var yy = ordYY_(o.issueDate), its = (o.items || []).filter(function (i) { return ordCodeKey_(i.code); });
  var prs = uq(rows.map(function (r) { return String(r.pr || '').trim(); })), sps = uq(rows.map(function (r) { return String(r.spec || '').trim(); }));
  var f = { rqn: '', spec: '', sure: false };
  if (!trkRefOk_(o.rqn) && prs.length === 1 && new RegExp('^RQN' + yy + '-\\d+$', 'i').test(prs[0])) f.rqn = prs[0].toUpperCase();
  var sm = sps.length === 1 ? new RegExp('^DS1C' + yy + '-0?(\\d{3})$', 'i').exec(sps[0]) : null;   // FormOrder prints 3 Spec digits: "DS1C26-0326" → 326
  if (!trkRefOk_(o.spec) && sm) f.spec = 'DS1C' + yy + '-' + sm[1];
  if (!f.rqn && !f.spec) return null;
  f.sure = !!f.rqn && rows.length === its.length && its.every(function (i) { var r = rows.filter(function (x) { return x.code === ordCodeKey_(i.code); })[0]; return r && Number(i.qty) === r.qty; });
  return f;
}
// write the numbers into the order (caller holds the script lock)
function trkFillApply_(id, f, why) {
  ordMemoReset_();
  var o = ordFind_(id); if (!o || !f) return null;
  var old = [o.rqn, o.spec], ch = [];
  if (f.rqn && !trkRefOk_(o.rqn)) { o.rqn = f.rqn; ch.push('RQN ' + f.rqn); }
  if (f.spec && !trkRefOk_(o.spec)) { o.spec = f.spec; ch.push('Spec ' + f.spec); }
  if (!ch.length) return null;
  o.updated = nowStr_();
  writeText_(ordSheet_(), o.row, 1, ordObjToRow_(o)); ordMemoPut_(o); dataBump_('ORD_VER');
  audit_([['System', 'เติมเลข RQN / Spec จาก Pending Order (' + why + ')', id, (old[0] || '-') + ' / ' + (old[1] || '-') + ' → ' + ch.join(' · ')]]);
  return { o: o, text: ch.join(' · ') };
}
// public: "✓ ใช้เลขนี้ในใบสั่ง" on the tracking card
function trackFillRef(id) {
  return withLock_(function () {
    var x = trkAll_(trkStateGet_()).filter(function (y) { return y.id === id; })[0];
    if (!x || !x.fill) throw new Error('ใบสั่งนี้มีเลข RQN / Spec ครบแล้ว หรือยังไม่พบ PR ที่จับคู่ได้');
    var r = trkFillApply_(id, x.fill, 'ยืนยันจากหน้าติดตามใบสั่ง');
    if (!r) throw new Error('ไม่มีเลขที่ต้องเติม');
    return { id: id, rqn: r.o.rqn, spec: r.o.spec, text: r.text };
  });
}
function trkAll_(prevState) {
  var P = trkPendIdx_(), today = wsToday_(), prev = (prevState && prevState.o) || {}, out = [];
  var opt = { gone: !!(P.n && (!prevState || !prevState.pn || P.n >= prevState.pn * 0.5)) };
  var orders = ordAll_().filter(function (o) { return ordAgeDays_(o.issueDate) <= 365; }), auto = trkAuto_(orders, P);
  orders.forEach(function (o) { out.push(trkOrder_(o, P.idx, prev[o.id], today, opt, auto[o.id])); });
  out = out.reverse(); out.pn = P.n; out.gone = opt.gone;
  return out;
}
function trkCached_() {
  return wsCached_('otrk', ['ord', 'oa', 'po', 'ntf'], '', function () {
    var list = trkAll_(trkStateGet_()).filter(function (x) { return x.stage !== 'done' && x.stage !== 'cancel' || ordAgeDays_(x.date) <= 180; });
    return { orders: list.slice(0, 400), time: nowStr_(), po: (poInfo_() || {}).time || '', ageCfg: oaAgeCfg_() };
  });
}
// menu badge: orders not fully received yet (+ how many need chasing)
function getOrderTrackSummary() {
  var d = trkCached_(), ac = oaAgeCfg_(), s = { open: 0, issue: 0, st: {} };
  (d.orders || []).forEach(function (o) {
    if (/^(done|rejected|cancel)$/.test(o.stage)) return;
    s.open++; s.st[o.stage] = (s.st[o.stage] || 0) + 1;
    if (o.late > 0 || o.noPr >= 7 || (o.waitDays != null && o.waitDays >= ac.over)) s.issue++;
  });
  return s;
}
function getOrderTrack(token) {
  var t0 = Date.now(), out = trkCached_();
  out.admin = isAdmin_(token);
  perf_('getOrderTrack', t0);
  return out;
}

// ── state between syncs (gzip JSON in sheet TrackState of Helpdesk_ItemStock, cached 6 h) ──
function trkStateGet_() {
  var b64 = '';
  try { var c = CacheService.getScriptCache(), n = Number(c.get('trk_n') || 0); if (n) { var ks = []; for (var i = 0; i < n; i++) ks.push('trk_' + i); var g = c.getAll(ks); if (ks.every(function (k) { return g[k] != null; })) b64 = ks.map(function (k) { return g[k]; }).join(''); } } catch (e) {}
  if (!b64) try { var sh = stockSS_().getSheetByName(TRK_STATE_SHEET); if (sh && sh.getLastRow()) b64 = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(function (r) { return r[0]; }).join(''); } catch (e) {}
  if (!b64) return null;
  try { return JSON.parse(Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(b64), 'application/x-gzip')).getDataAsString()); } catch (e) { return null; }
}
function trkStatePut_(st) {
  var b64 = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(JSON.stringify(st), 'application/json')).getBytes()), parts = [];
  for (var i = 0; i < b64.length; i += TRK_CHUNK) parts.push(b64.substring(i, i + TRK_CHUNK));
  try { var ss = stockSS_(), sh = ss.getSheetByName(TRK_STATE_SHEET) || ss.insertSheet(TRK_STATE_SHEET); sh.clearContents(); sh.getRange(1, 1, parts.length, 1).setValues(parts.map(function (p) { return [p]; })); } catch (e) { console.warn(e); }
  try { var put = { trk_n: String(parts.length) }; parts.forEach(function (p, j) { put['trk_' + j] = p; }); CacheService.getScriptCache().putAll(put, 21600); } catch (e) {}
}

// ── F5: notifications (called after every Stock / Pending Order sync) ──
function trkScan_(why) {
  var t0 = Date.now();
  return withLock_(function () {
    var prev = trkStateGet_(), list = trkAll_(prev), today = ordToday_(), now = Date.now(), ev = [];
    list.forEach(function (x) {                                         // unambiguous CODE + date matches → write RQN / Spec into the order
      if (!x.fill || !x.fill.sure) return;
      var r = null; try { r = trkFillApply_(x.id, x.fill, 'อัตโนมัติ'); } catch (e) { console.warn('trkFillApply_', e); }
      if (r) { ev.push(['ref', '', 'เติมเลขในใบสั่งแล้ว · ' + x.id, r.text + ' (จาก Pending Order)', x.id, '|' + String(x.des || '').trim().toLowerCase() + '|', '']); x.rqn = r.o.rqn; x.spec = r.o.spec; x.fill = null; }
    });
    var st = { v: 1, at: nowStr_(), pn: list.gone ? list.pn : (prev && prev.pn) || list.pn || 0, o: {} };   // keep the full-file size while a partial file is in
    var P = prev && prev.o || {};
    list.forEach(function (x) {
      var m = st.o[x.id] = {}, p = P[x.id], po = [], recv = [], part = [];
      x.items.forEach(function (i) {
        var k = ordCodeKey_(i.code); if (!k) return;
        var old = p && p[k];
        m[k] = [i.st, old && old[0] === i.st ? old[1] : today];
        if (!p || !old) return;                                         // first time this order / CODE is seen → remember only
        var a = TRK_RANK[old[0]], b = TRK_RANK[i.st];
        if (i.st === old[0] || i.st === 'cancel' || b < a) return;
        if (i.st === 'po' && a < 2) po.push(i);
        else if (i.st === 'part') part.push(i);
        else if (i.st === 'recv' || i.st === 'gone') recv.push(i);
      });
      var who = '|' + String(x.des || '').trim().toLowerCase() + '|', codes = function (a) { return a.map(function (i) { return i.code; }).join(', '); };
      if (recv.length) ev.push(['recv', recv.length === 1 ? recv[0].code : '', 'ของเข้าแล้ว · ใบสั่ง ' + x.id, codes(recv) + (x.spec ? ' · Spec ' + x.spec : '') + (recv.some(function (i) { return i.st === 'gone'; }) ? ' (ปิดจาก Pending Order แล้ว)' : ''), x.id, who, '']);
      if (part.length) ev.push(['part', part.length === 1 ? part[0].code : '', 'ของเข้าบางส่วน · ใบสั่ง ' + x.id, codes(part), x.id, who, '']);
      if (po.length) ev.push(['po', po.length === 1 ? po[0].code : '', 'ออก PO แล้ว · ใบสั่ง ' + x.id, po.map(function (i) { return i.code + (i.po ? ' (PO ' + i.po + (i.due ? ' กำหนด ' + i.due : '') + ')' : ''); }).join(', '), x.id, who, '']);
    });
    // Stock went up (this run's Item stock sync) for a CODE someone is waiting for
    var inc = __wsInc || {}, ks = Object.keys(inc); __wsInc = null;          // consumed once per run
    if (ks.length && prev) {
      var bk = {}; try { bk = bomBooked_(); } catch (e) {}
      var snap = null; try { snap = bomSnapGet_(); } catch (e) {}
      var shortBy = {}; (snap && snap.codes || []).forEach(function (c) { shortBy[c.code] = c; });
      var sa = {}; try { saData_(false).items.forEach(function (x) { if (/^(high|mid|low|wait)$/.test(x.level)) sa[ordCodeKey_(x.code)] = x; }); } catch (e) {}
      var waitBy = {};                                                   // CODE → open orders not received yet
      list.forEach(function (x) { if (/^(rejected|done|cancel)$/.test(x.stage)) return; x.items.forEach(function (i) { var k = ordCodeKey_(i.code); if (/^(none|pr|po|part)$/.test(i.st)) (waitBy[k] = waitBy[k] || []).push(x); }); });
      var done = {}; ev.forEach(function (e) { if (e[1]) done[ordCodeKey_(e[1])] = 1; });
      ks.slice(0, 400).forEach(function (k) {
        if (done[k]) return;
        var B = bk[k], O = waitBy[k] || [], S = sa[k], C = shortBy[k];
        if (!B && !O.length && !S) return;
        var d = inc[k], add = Math.round((d[1] - d[0]) * 1000) / 1000, why = [];
        if (B) why.push('จอง BOM ' + B.projects.map(function (p) { return p[0]; }).slice(0, 3).join(', ') + (C && C.balance < 0 ? ' · ยังขาด ' + Math.round(-C.balance * 1000) / 1000 : ' · พอแล้ว'));
        if (O.length) why.push('ใบสั่ง ' + O.map(function (x) { return x.id; }).slice(0, 3).join(', '));
        if (S) why.push('Stock Alert');
        var who = '|' + O.map(function (x) { return String(x.des || '').trim().toLowerCase(); }).filter(String).join('|') + '|';
        ev.push(['stock', k, 'Stock เพิ่ม +' + add + ' · ' + k, (C && C.desc || S && S.desc || '') + ' · คงเหลือ ' + d[1] + ' · ' + why.join(' · '), O.length ? O[0].id : '', who, B ? B.projects.map(function (p) { return p[0]; }).slice(0, 5).join(', ') : '']);
      });
    }
    trkStatePut_(st);
    if (ev.length) ntfAdd_(ev.slice(0, 300), now);
    dataBump_('NTF_VER');
    perf_('trkScan_ (' + why + ') ' + ev.length + ' แจ้งเตือน', t0);
    return ev.length;
  });
}
function trkScanSafe_(why) {
  try { trkScan_(why); } catch (e) { console.warn('trkScan_', e); }
  if (typeof wsWarmSafe_ === 'function') wsWarmSafe_(why);              // Workspace.gs: pre-build the shared page caches
}
function ntfSheet_() {
  var ss = stockSS_(), sh = ss.getSheetByName(NTF_SHEET);
  if (!sh) { sh = ss.insertSheet(NTF_SHEET); sh.getRange(1, 1, 1, NTF_HEADER.length).setValues([NTF_HEADER]).setFontWeight('bold'); sh.setFrozenRows(1); }
  return sh;
}
function ntfAdd_(rows, ts) {
  var sh = ntfSheet_(), at = nowStr_(), add = rows.map(function (r, i) { return [String(ts + i), at].concat(r.map(function (v) { return ordStr_(v, 500); })); });
  var r0 = sh.getLastRow() + 1;
  if (sh.getMaxRows() < r0 + add.length) sh.insertRowsAfter(sh.getMaxRows(), r0 + add.length - sh.getMaxRows());
  var rg = sh.getRange(r0, 1, add.length, NTF_HEADER.length); rg.setNumberFormat('@'); rg.setValues(add);
  var n = sh.getLastRow() - 1; if (n > NTF_MAX) sh.deleteRows(2, n - NTF_KEEP);
}
// public: newest 300 notifications (everyone; the browser filters "mine" by the DES'D name it remembers)
function getNotify() {
  var key = 'ntf_' + (schedProp_('NTF_VER') || '0'), hit = ordCacheGet_(key);
  if (hit) return hit;
  var out = { ver: schedProp_('NTF_VER') || '', items: [] };
  try {
    var sh = stockSS_().getSheetByName(NTF_SHEET), n = sh ? sh.getLastRow() - 1 : 0;
    if (n > 0) {
      var m = Math.min(300, n);
      out.items = sh.getRange(n - m + 2, 1, m, NTF_HEADER.length).getDisplayValues().reverse().map(function (r) {
        return { ts: Number(r[0]) || 0, at: r[1], type: r[2], code: r[3], title: r[4], detail: r[5], ref: r[6], who: r[7], project: r[8] };
      });
    }
  } catch (e) { console.warn('getNotify', e); }
  ordCachePut_(key, out, 3600);
  return out;
}
// Admin: rebuild the order-tracking state now (no notifications for this run's first sight)
function trackScanNow(token) { requireAdmin_(token); return { events: trkScan_('สั่งโดย Admin') }; }
