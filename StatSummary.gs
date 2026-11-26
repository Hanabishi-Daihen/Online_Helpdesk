/**
 * สถิติและต้นทุน — one summary built after the daily syncs, read by the page in one call.
 *
 *   1 การเบิกใช้      Shipment (project categories): per month and department, top CODEs, project cost, returns
 *   2 BOM vs ของจริง  Booking BOM projects against what CONTROL WIRING issued (BomShip.gs rules)
 *   3 สุขภาพสต็อก     Item stock value, idle stock by LASTSHIPDT, turnover per CODE group, days of cover
 *   4 การจัดซื้อ      Pending Order lead time (PR → PO → due) against StockRT, open-line ageing, suggested Min/Max
 *
 * Nothing heavy runs when the page opens: the Shipment sync keeps monthly totals per CODE × department
 * (sheet ShipAgg, gzip), and statBuild_ combines them with Item stock, Pending Order, Stock Alert and Booking BOM
 * into one gzip blob (sheet Stat in Helpdesk_Shipment + cache). Built after the Shipment, Item stock and
 * Pending Order syncs, or by Admin's "คำนวณใหม่".
 * Money: every key that starts with "$" is a cost / value — removed before the summary reaches a non-Admin.
 */
var STAT_SHEET = 'Stat', STAT_AGG = 'ShipAgg', STAT_TOP = 60;

// ── monthly totals kept by the Shipment sync ──────────────────────────
// agg: { 'CODE\tDEPT\tyyyymm': [qtyIssued, qtyReturned, costIssued, costReturned] } · desc: { CODE: [desc, unit] }
function statAggAdd_(A, code, desc, unit, dept, d8, qty, cost) {
  code = ordCodeKey_(code); if (!code || !d8) return;
  var k = code + '\t' + dept + '\t' + d8.substring(0, 6), a = A.agg[k] || (A.agg[k] = [0, 0, 0, 0]), ret = cost > 0;
  qty = Math.abs(qty); cost = Math.abs(cost);
  if (ret) { a[1] += qty; a[3] += cost; } else { a[0] += qty; a[2] += cost; }
  if (!A.desc[code]) A.desc[code] = [String(desc || '').substring(0, 80), String(unit || '').substring(0, 10)];
}
function statBlobPut_(ss, name, obj) {
  var b = shipB64_(JSON.stringify(obj)), cells = [];
  for (var i = 0; i < b.length; i += SHIP_CELL) cells.push([b.substring(i, i + SHIP_CELL)]);
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clearContents(); ensureRows_(sh, cells.length + 1);
  var rg = sh.getRange(1, 1, cells.length, 1); fmtText_(rg); rg.setValues(cells);
  return b;
}
function statBlobGet_(ss, name) {
  var sh = ss && ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 1) return null;
  return sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(function (r) { return r[0]; }).join('');
}
function statAggSave_(A) {
  var rows = Object.keys(A.agg).map(function (k) { var a = A.agg[k]; return k + '\t' + a.map(function (x) { return Math.round(x * 100) / 100; }).join('\t'); });
  statBlobPut_(shipSS_(true), STAT_AGG, { ver: String(Date.now()), rows: rows.join('\n'), desc: A.desc });
}
function statAgg_() {
  var b = statBlobGet_(shipSS_(false), STAT_AGG); if (!b) return null;
  var o = JSON.parse(shipUnB64_(b)), list = [];
  (o.rows ? o.rows.split('\n') : []).forEach(function (l) { var f = l.split('\t'); if (f.length >= 7) list.push([f[0], f[1], f[2], +f[3], +f[4], +f[5], +f[6]]); });
  return { list: list, desc: o.desc || {} };
}

// ── helpers ───────────────────────────────────────────────────────────
function statYm_(t) { var d = new Date(t); return String(d.getFullYear()) + ('0' + (d.getMonth() + 1)).slice(-2); }
function statMonths_(n, now) { var out = [], d = new Date(now); d.setDate(1); for (var i = 0; i < n; i++) { out.unshift(statYm_(d.getTime())); d.setMonth(d.getMonth() - 1); } return out; }
function statR_(v, n) { var p = Math.pow(10, n == null ? 2 : n); return Math.round((Number(v) || 0) * p) / p; }
function statMedian_(a) { if (!a.length) return null; a = a.slice().sort(function (x, y) { return x - y; }); var m = Math.floor(a.length / 2); return a.length % 2 ? a[m] : (a[m - 1] + a[m]) / 2; }
function statD8t_(d8) { return d8 && d8.length === 8 ? new Date(+d8.substring(0, 4), +d8.substring(4, 6) - 1, +d8.substring(6, 8)).getTime() : null; }
function statProjC_(code) { return /^\d{4}PP\d{4}C\d{3,}$/i.test(String(code || '')); }   // By Project CODE: 4556PP1030C00201
function statGroupOf_(code) { return String(code).substring(0, 4); }

// ── build ─────────────────────────────────────────────────────────────
function statBuild_() {
  var t0 = Date.now(), now = t0, today = new Date(); today = new Date(today.getFullYear(), today.getMonth(), today.getDate()).getTime();
  // our unit only: Item stock / Pending Order lines ordered by the unit (DS13), MF issues to CONTROL WIRING,
  // and By Project CODEs (16 digits with the project number + C: 4556PP1030C00201)
  var by = String((bomCfg_() || {}).orderBy || 'DS13').trim().toUpperCase();
  var S = { v: 2, built: nowStr_(), notes: [], scope: { by: by, dept: 'CONTROL WIRING' } }, M24 = statMonths_(24, now), M12 = {}; statMonths_(12, now).forEach(function (m) { M12[m] = 1; });
  // Item stock: onhand / LASTCOST / LASTSHIPDT per CODE (all locations)
  var stock = {};
  var hasStock = !!schedProp_('STOCK_SS_ID');                         // never create the Item stock workbook from here
  if (hasStock) try {
    stockData_().forEach(function (row) {
      var f = row.split('\t'); var k = ordCodeKey_(f[0]); if (!k) return;
      var s = stock[k] || (stock[k] = { desc: f[1] || '', spec: f[2] || '', unit: f[4] || '', on: 0, cost: 0, lu: '', act: '', mine: false });
      if (String(f[7] || '').trim().toUpperCase() === by) s.mine = true;
      var q = saNum_(f[3]), c = saNum_(f[STOCK_COST_COL]); if (q != null) s.on += q; if (c && c > s.cost) s.cost = c;
      var lu = f[STOCK_LU_COL] || ''; if (lu > s.lu) s.lu = lu; if (f[STOCK_ACT_COL] === '0') s.act = '0';
    });
  } catch (e) { S.notes.push('อ่าน Item stock ไม่สำเร็จ: ' + (e.message || e)); }
  var ours = function (k) { return statProjC_(k) || !!(stock[k] && stock[k].mine); };
  // 1 · usage from the monthly totals: MF → CONTROL WIRING, plus By Project CODEs wherever they were issued
  var A = null; try { A = statAgg_(); } catch (e) { S.notes.push('อ่านยอดเบิกรายเดือนไม่สำเร็จ: ' + (e.message || e)); }
  var use = {}, dept = {}, months = {};
  if (A) {
    A.list.forEach(function (r) {
      var code = r[0], dp = r[1], ym = r[2];
      if (!BOM_SHIP_DEPT.test(dp) && !statProjC_(code)) return;
      if (M24.indexOf(ym) >= 0) { var dm = dept[dp] || (dept[dp] = {}), x = dm[ym] || (dm[ym] = [0, 0]); x[0] += r[5] - r[6]; x[1] += 1; months[ym] = 1; }
      if (!M12[ym]) return;
      var u = use[code] || (use[code] = { q: 0, r: 0, c: 0, cr: 0, d: {} });
      u.q += r[3]; u.r += r[4]; u.c += r[5]; u.cr += r[6]; u.d[dp] = 1;
    });
  } else S.notes.push('ยังไม่มียอดเบิกรายเดือน — ซิงค์ Shipment ใหม่หนึ่งครั้ง (ตั้งค่าระบบ → ☁ แหล่งข้อมูล OneDrive)');
  var nameOf = function (k) { var s = stock[k], d = A && A.desc[k]; return { desc: s ? (s.desc + (s.spec ? ' ' + s.spec : '')) : d ? d[0] : '', unit: s && s.unit || d && d[1] || '' }; };
  S.months = M24;
  S.depts = Object.keys(dept).map(function (dp) {
    var tot = 0; M24.forEach(function (m) { tot += (dept[dp][m] || [0])[0]; });
    return { name: dp, $m: M24.map(function (m) { return statR_((dept[dp][m] || [0])[0], 0); }), n: M24.map(function (m) { return (dept[dp][m] || [0, 0])[1]; }), $tot: statR_(tot, 0) };
  }).sort(function (a, b) { return b.$tot - a.$tot; });
  var codes = Object.keys(use).map(function (k) { var u = use[k], n = nameOf(k); return { code: k, desc: n.desc, unit: n.unit, q: statR_(u.q), r: statR_(u.r), $c: statR_(u.c - u.cr, 0), $cr: statR_(u.cr, 0), rp: u.q ? statR_(u.r / u.q * 100, 1) : 0, depts: Object.keys(u.d).length, mine: !!(stock[k] && stock[k].mine), pc: statProjC_(k) }; });
  S.$topCost = codes.slice().sort(function (a, b) { return b.$c - a.$c; }).slice(0, STAT_TOP);
  S.topQty = codes.slice().sort(function (a, b) { return b.q - a.q; }).slice(0, STAT_TOP);
  S.returns = codes.filter(function (c) { return c.r > 0; }).sort(function (a, b) { return b.r - a.r; }).slice(0, STAT_TOP);
  S.$costAll12 = statR_(codes.reduce(function (s, c) { return s + c.$c; }, 0), 0);
  S.$share20 = S.$costAll12 ? statR_(S.$topCost.slice(0, 20).reduce(function (s, c) { return s + c.$c; }, 0) / S.$costAll12 * 100, 1) : 0;
  S.codes12 = codes.length;
  var mineC = codes.filter(function (c) { return c.mine || c.pc; });
  S.$mineCost12 = statR_(mineC.reduce(function (s, c) { return s + c.$c; }, 0), 0); S.mine12 = mineC.length;
  // projects: cost per project and type (title from Booking BOM / schedule)
  try {
    var m = shipMeta_(), names = m ? shipNames_(m) : {}, P = {};
    (m ? m.sum : []).forEach(function (s) {
      if (!BOM_SHIP_DEPT.test(s[0])) return;                             // MF → CONTROL WIRING only
      var p = P[s[1]] || (P[s[1]] = { cat: s[1], n: 0, $c: 0, $cr: 0, last: '' });
      p.n += s[2]; p.$c += -s[3]; if (s[5]) p.$cr += s[5]; if (s[4] > p.last) p.last = s[4];
    });
    var typeOf = function (t) { t = String(t || '').toUpperCase(); var x = /\b(LCC|RCC|MCC|CP|RP|PCP|CRP|AC\/DC|DC|MARSHALLING|MK|TB)\b/.exec(t); return x ? x[1] : t ? 'อื่น ๆ' : 'ไม่ทราบชนิด'; };
    var pl = Object.keys(P).map(function (k) { var p = P[k], nm = names[k] || {}; return { project: nm.no || shipProj_(k), cat: k, title: nm.title || nm.cust || '', type: typeOf(nm.title || ''), n: p.n, $c: statR_(p.$c, 0), $cr: statR_(p.$cr, 0), last: shipShow_(p.last), d8: p.last }; })
      .sort(function (a, b) { return a.d8 < b.d8 ? 1 : a.d8 > b.d8 ? -1 : 0; });
    var ty = {};
    pl.forEach(function (p) { if (p.d8 < statMonths_(24, now)[0] + '01') return; var t = ty[p.type] || (ty[p.type] = { type: p.type, n: 0, $sum: 0, $min: null, $max: null }); t.n++; t.$sum += p.$c; t.$min = t.$min == null ? p.$c : Math.min(t.$min, p.$c); t.$max = t.$max == null ? p.$c : Math.max(t.$max, p.$c); });
    S.$projects = pl.slice(0, 120).map(function (p) { delete p.d8; return p; });
    S.$types = Object.keys(ty).map(function (k) { var t = ty[k]; return { type: t.type, n: t.n, $avg: statR_(t.$sum / t.n, 0), $min: statR_(t.$min, 0), $max: statR_(t.$max, 0) }; }).sort(function (a, b) { return b.n - a.n; });
  } catch (e) { S.notes.push('สรุปต้นทุนรายโปรเจกต์ไม่สำเร็จ: ' + (e.message || e)); }
  // 2 · BOM against what was issued (open + closed in the last 18 months, newest 40)
  try { S.bom = statBom_(now); } catch (e) { S.notes.push('เทียบ BOM ไม่สำเร็จ: ' + (e.message || e)); }
  // 3 · stock health
  var sa = {}; try { sa = saSettings_(); } catch (e) {}
  var y = function (n) { return statYm_(now - n * 365 * 864e5) + '01'; };
  // never issued but a PR / PO in the last 5 years = still planned for use (not idle stock)
  var pend = []; try { pend = statPend_(); } catch (e) { S.notes.push('อ่าน Pending Order ไม่สำเร็จ: ' + (e.message || e)); }
  var lastReq = {}; pend.forEach(function (g) { var k = ordCodeKey_(g[0]), r = bomDate_(g[6]) || bomDate_(g[8]); if (k && r != null && !(lastReq[k] >= r)) lastReq[k] = r; });
  var B = [['ไม่เกิน 1 ปี', y(1)], ['1–2 ปี', y(2)], ['2–3 ปี', y(3)], ['เกิน 3 ปี', '0'], ['ไม่เคยเบิก', null], ['ไม่เคยเบิก · มี PR/PO ใน 5 ปี (รอใช้งาน)', null]].map(function (b, i) { return { name: b[0], from: b[1], n: 0, $v: 0, wait: i === 5 }; });
  var waitL = [];
  var V = 0, grp = {}, dead = [], low = [], over = [], lt0 = 30;
  try { lt0 = fcCfg_().lt; } catch (e) {}
  var rt = {}; if (hasStock) try { rt = rtMap_(stock); } catch (e) {}
  var pcN = 0, pcV = 0;
  Object.keys(stock).forEach(function (k) {
    if (!ours(k)) return;
    var pc = statProjC_(k);
    var s = stock[k], val = s.on > 0 ? s.on * (s.cost || 0) : 0, u = use[k], lu = s.lu;
    if (s.on > 0) {
      V += val; if (pc) { pcN++; pcV += val; }
      var lr = lastReq[k], recent = lr != null && today - lr < 5 * 365 * 864e5;
      var bi = !lu ? (recent ? 5 : 4) : lu >= B[0].from ? 0 : lu >= B[1].from ? 1 : lu >= B[2].from ? 2 : 3; B[bi].n++; B[bi].$v += val;
      var idle = lu ? Math.round((today - statD8t_(lu)) / 864e5) : null;
      if (bi === 5) waitL.push({ code: k, desc: s.desc + (s.spec ? ' ' + s.spec : ''), unit: s.unit, on: statR_(s.on), $v: statR_(val, 0), pr: bomFmt_(lr), pc: pc });
      else if (bi > 0) dead.push({ pr: lr != null ? bomFmt_(lr) : '', code: k, desc: s.desc + (s.spec ? ' ' + s.spec : ''), unit: s.unit, on: statR_(s.on), $v: statR_(val, 0), last: lu ? shipShow_(lu) : '', idle: idle, act: s.act !== '0', pc: pc });
    }
    var g = grp[statGroupOf_(k)] || (grp[statGroupOf_(k)] = { g: statGroupOf_(k), n: 0, $v: 0, $use: 0, q: 0 });
    if (s.on > 0) { g.n++; g.$v += val; } if (u) { g.$use += u.c - u.cr; g.q += u.q; }
    if (!pc && u && u.q > u.r && s.on >= 0) {                        // By Project parts are one-off: no cover / Min-Max
      var daily = (u.q - u.r) / 365, cover = daily > 0 ? Math.round(s.on / daily) : null, lt = sa[k] && sa[k].lt > 0 ? sa[k].lt : rt[k] || lt0;
      var row = { code: k, desc: s.desc + (s.spec ? ' ' + s.spec : ''), unit: s.unit, on: statR_(s.on), use12: statR_(u.q - u.r), cover: cover, lt: lt, $v: statR_(val, 0) };
      if (cover != null && cover < lt) low.push(row); else if (cover != null && cover > 730) over.push(row);
    }
  });
  S.stock = { $value: statR_(V, 0), buckets: B.map(function (b) { return { name: b.name, n: b.n, $v: statR_(b.$v, 0), wait: b.wait }; }), items: Object.keys(stock).filter(function (k) { return ours(k) && stock[k].on > 0; }).length, pcN: pcN, $pcV: statR_(pcV, 0) };
  S.waiting = waitL.sort(function (a, b) { return b.$v - a.$v; }).slice(0, STAT_TOP);
  S.$dead = dead.slice().sort(function (a, b) { return b.$v - a.$v; }).slice(0, STAT_TOP);
  S.dead = dead.slice().sort(function (a, b) { return (b.idle == null ? 9e9 : b.idle) - (a.idle == null ? 9e9 : a.idle) || b.on - a.on; }).slice(0, STAT_TOP).map(function (x) { var o = {}; Object.keys(x).forEach(function (k) { if (k !== '$v') o[k] = x[k]; }); return o; });
  S.groups = Object.keys(grp).map(function (k) { var g = grp[k]; return { g: g.g, n: g.n, q: statR_(g.q), $v: statR_(g.$v, 0), $use: statR_(g.$use, 0), turn: g.$v > 0 ? statR_(g.$use / g.$v, 2) : null }; })
    .filter(function (g) { return g.n || g.q; }).sort(function (a, b) { return b.$v - a.$v || b.q - a.q; }).slice(0, 80);
  S.low = low.sort(function (a, b) { return a.cover / a.lt - b.cover / b.lt; }).slice(0, STAT_TOP);
  S.over = over.sort(function (a, b) { return b.$v - a.$v || b.cover - a.cover; }).slice(0, STAT_TOP);
  // 4 · purchasing
  try { S.po = statPo_(today, stock, rt, lt0, by, pend); } catch (e) { S.notes.push('สรุป Pending Order ไม่สำเร็จ: ' + (e.message || e)); }
  // suggested Min / Max from 12-month use: Min = use per day × lead time × 1.5 · Max = Min + 90 days of use
  var mm = [];
  Object.keys(use).forEach(function (k) {
    var u = use[k], net = u.q - u.r; if (!(net > 0) || statProjC_(k) || !(stock[k] && stock[k].mine)) return;   // items the unit orders for stock
    var daily = net / 365, lt = sa[k] && sa[k].lt > 0 ? sa[k].lt : rt[k] || lt0, mn = Math.ceil(daily * lt * 1.5), mx = Math.ceil(mn + daily * 90), cur = sa[k] || {};
    if (cur.notUse) return;
    var off = cur.min == null || cur.max == null || Math.abs(cur.min - mn) > Math.max(2, mn * 0.5) || Math.abs(cur.max - mx) > Math.max(2, mx * 0.5);
    if (!off) return;
    var n = nameOf(k);
    mm.push({ code: k, desc: n.desc, unit: n.unit, use12: statR_(net), lt: lt, ltSrc: sa[k] && sa[k].lt > 0 ? 'set' : rt[k] ? 'rt' : 'def', min: mn, max: mx, curMin: cur.min == null ? null : cur.min, curMax: cur.max == null ? null : cur.max, on: stock[k] ? statR_(stock[k].on) : 0, $c: statR_(u.c - u.cr, 0) });
  });
  S.minmax = mm.sort(function (a, b) { return b.$c - a.$c || b.use12 - a.use12; }).slice(0, 150);
  S.secs = Math.round((Date.now() - t0) / 1000);
  // store
  var ss = shipSS_(true), b = statBlobPut_(ss, STAT_SHEET, S);
  schedProp_('STAT_VER', String(Date.now()));
  try { var C = CacheService.getScriptCache(), put = {}, k2 = 0; for (var i = 0; i < b.length; i += 90000) put['stat_' + (k2++)] = b.substring(i, i + 90000); put.stat_n = String(k2); C.putAll(put, 21600); } catch (e) {}
  return S;
}
function statRebuild_(why) { try { return statBuild_(); } catch (e) { console.warn('statBuild_ ' + why, e); return null; } }

// Booking BOM projects against CONTROL WIRING lines (after the BOM import) → accuracy and CODEs used over the BOM
function statBom_(now) {
  var m = shipMeta_(); if (!m) return null;
  var projs = bomProjects_(), itemsBy = bomItems_(), from = statYm_(now - 548 * 864e5) + '01';
  projs = projs.filter(function (p) { var d = bomShipD8_(p.imported); return d && d >= from && m.idx[shipCat_(p.project)]; })
    .sort(function (a, b) { return bomShipD8_(b.imported) < bomShipD8_(a.imported) ? -1 : 1; }).slice(0, 40);
  var lines = {}, over = {}, out = [];
  projs.forEach(function (p) {
    var cat = shipCat_(p.project), d8 = bomShipD8_(p.imported), need = {}, got = {}, key = bomKey_(p.project);
    (itemsBy[p.project] || []).forEach(function (it) { var c = ordCodeKey_(it.code); if (c) need[c] = (need[c] || 0) + (Number(it.qty) || 0); });
    if (!lines[cat]) { try { lines[cat] = shipLines_(m, cat); } catch (e) { lines[cat] = []; } }
    lines[cat].forEach(function (f) {
      if (!BOM_SHIP_DEPT.test(f[4] || '') || String(f[2]) < d8) return;
      var x = /PP\d{4}[A-Z]{1,2}\d{2}/.exec(String(f[3] || '').toUpperCase()); if (x && bomKey_(x[0]) !== key) return;   // another unit's line
      var c = ordCodeKey_(f[0]), q = Number(f[6]) || 0; if (!c) return;
      got[c] = (got[c] || 0) + ((Number(f[9]) || 0) > 0 ? -Math.abs(q) : Math.abs(q));
    });
    var nc = Object.keys(need), issued = 0, under = 0, extra = Object.keys(got).filter(function (c) { return !need[c] && got[c] > 0; });
    nc.forEach(function (c) {
      var g = got[c] || 0; if (g > 0) issued++; if (g > 0 && g < need[c]) under++;
      if (g > need[c] && need[c] > 0) { var o = over[c] || (over[c] = { code: c, n: 0, pct: [] }); o.n++; o.pct.push((g - need[c]) / need[c] * 100); }
    });
    if (!Object.keys(got).length) return;                                // nothing issued yet: no score
    out.push({ project: p.project, title: p.title, status: p.status, imported: String(p.imported || '').substring(0, 10), n: nc.length, issued: issued, notIssued: nc.length - issued, under: under, extra: extra.length,
      extraCodes: extra.slice(0, 12), acc: statR_(issued / (nc.length + extra.length) * 100, 1) });
  });
  var desc = {}; try { var A = statAgg_(); desc = A ? A.desc : {}; } catch (e) {}
  return { projects: out, avg: out.length ? statR_(out.reduce(function (s, p) { return s + p.acc; }, 0) / out.length, 1) : null,
    over: Object.keys(over).map(function (c) { var o = over[c]; return { code: c, desc: (desc[c] || [''])[0], n: o.n, pct: statR_(statMedian_(o.pct), 0) }; })
      .filter(function (o) { return o.n >= 2; }).sort(function (a, b) { return b.n - a.n || b.pct - a.pct; }).slice(0, 40) };
}

// Pending Order: lead time per CODE (Request → PO → due), open-line ageing, lines past their due date
// Pending Order lines (split) — read once per build
function statPend_() {
  var b = pendingBundle_(false), txt = b && b.b64 ? shipUnB64_(b.b64) : '';
  return txt ? txt.split('\n').map(function (l) { return l.split('\t'); }).filter(function (g) { return g.length >= 10; }) : [];
}
// open lines older than 3 years are most likely received / used but never closed in the system (PH): listed apart,
// not counted as open, aged or late
var STAT_STALE_D = 3 * 365;
function statPo_(today, stock, rt, lt0, by, rows) {
  var from = today - 730 * 864e5;
  var lt = {}, ages = [0, 0, 0, 0], late = [], prpo = [], podue = [], openN = 0, stale = [];
  (rows || statPend_()).forEach(function (g) {
    var k = ordCodeKey_(g[0]), req = bomDate_(g[6]), iss = bomDate_(g[8]), due = bomDate_(g[9]) || bomDate_(g[7]), done = ORD_PO_DONE.test(g[4] || ''), cancel = /cancel|ยกเลิก/i.test(g[4] || '');
    if (cancel || !k) return;
    if (by && String(g[12] || '').trim().toUpperCase() !== by && !statProjC_(k)) return;   // the unit's lines only
    if (req != null && req >= from && due != null && due >= req) {
      var x = lt[k] || (lt[k] = { d: [], a: [], b: [] }); x.d.push((due - req) / 864e5);
      if (iss != null && iss >= req) { x.a.push((iss - req) / 864e5); prpo.push((iss - req) / 864e5); if (due >= iss) { x.b.push((due - iss) / 864e5); podue.push((due - iss) / 864e5); } }
    }
    if (!done) {
      var age = req != null ? (today - req) / 864e5 : due != null ? (today - due) / 864e5 : null;
      if (age != null && age > STAT_STALE_D) { stale.push({ code: k, name: g[5] || (stock[k] || {}).desc || '', pr: g[1], po: g[2], status: g[4], req: bomFmt_(req), due: bomFmt_(due), years: Math.round(age / 365 * 10) / 10, qty: Number(g[13]) || 0, unit: g[14] || '', by: g[12] || '' }); return; }
      openN++;
      if (age != null) ages[age <= 30 ? 0 : age <= 60 ? 1 : age <= 90 ? 2 : 3]++;   // 91 days – 3 years
      if (due != null && due < today) late.push({ code: k, name: g[5] || (stock[k] || {}).desc || '', pr: g[1], po: g[2], status: g[4], due: bomFmt_(due), days: Math.round((today - due) / 864e5), qty: Number(g[13]) || 0, unit: g[14] || '', by: g[12] || '' });
    }
  });
  var diff = [];
  Object.keys(lt).forEach(function (k) {
    var d = statMedian_(lt[k].d); if (d == null || lt[k].d.length < 2) return;
    var r = rt[k] || null; if (r && Math.abs(d - r) <= Math.max(7, r * 0.5)) return;
    diff.push({ code: k, desc: stock[k] ? stock[k].desc + (stock[k].spec ? ' ' + stock[k].spec : '') : '', n: lt[k].d.length, pend: Math.round(d), prpo: lt[k].a.length ? Math.round(statMedian_(lt[k].a)) : null, podue: lt[k].b.length ? Math.round(statMedian_(lt[k].b)) : null, rt: r });
  });
  return { open: openN, ages: [['≤ 30 วัน', ages[0]], ['31–60 วัน', ages[1]], ['61–90 วัน', ages[2]], ['91 วัน – 3 ปี', ages[3]]], staleN: stale.length,
    stale: stale.sort(function (a, b) { return b.years - a.years; }).slice(0, STAT_TOP),
    prpo: prpo.length ? Math.round(statMedian_(prpo)) : null, podue: podue.length ? Math.round(statMedian_(podue)) : null, def: lt0,
    late: late.sort(function (a, b) { return b.days - a.days; }).slice(0, STAT_TOP), lateN: late.length,
    diff: diff.sort(function (a, b) { return b.n - a.n; }).slice(0, STAT_TOP) };
}

// ── read ──────────────────────────────────────────────────────────────
function statGet_() {
  var C = CacheService.getScriptCache(), b = null;
  try { var n = Number(C.get('stat_n') || 0); if (n) { var keys = []; for (var i = 0; i < n; i++) keys.push('stat_' + i); var o = C.getAll(keys); if (keys.every(function (k) { return o[k]; })) b = keys.map(function (k) { return o[k]; }).join(''); } } catch (e) {}
  if (!b) b = statBlobGet_(shipSS_(false), STAT_SHEET);
  return b ? JSON.parse(shipUnB64_(b)) : null;
}
// "$…" = money: Admin only
function statPublic_(S) { return JSON.parse(JSON.stringify(S, function (k, v) { return k && k.charAt(0) === '$' ? undefined : v; })); }
function getStatSummary(token) { apiOnly_();
  var admin = isAdmin_(token), S = null;
  try { S = statGet_(); } catch (e) { console.warn('statGet_', e); }
  if (!S) return { none: true, admin: admin, hasShip: !!schedProp_('SHIP_SRC_URL') };
  var out = admin ? S : statPublic_(S); out.admin = admin; return out;
}
function statRebuild(token) { apiOnly_(); requireAdmin_(token); var S = statBuild_(); audit_([['Admin', 'คำนวณสถิติและต้นทุนใหม่', '', S.secs + ' วินาที']]); return getStatSummary(token); }
