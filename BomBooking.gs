/**
 * ═════════════════════════════════════════════════════════════════════
 *  BOOKING BOM — จอง Item ตาม BOM ของโปรเจกต์ที่จะผลิต เพื่อเช็คว่า Onhand + กำลังเข้า พอหรือไม่
 *  • Import (Admin)  : ไฟล์ BOM .xlsx อ่านในเบราว์เซอร์ (คอลัมน์ B = PROJECT, C = CODE, H = จำนวน)
 *                      Import โปรเจกต์เดิมซ้ำ = แทนที่ BOM ทั้งชุด (Revision ใหม่)
 *  • วันที่ต้องใช้      : จากตารางงานการผลิต (Q_SchedulesForUpdate) — ProjectNo ที่เป็นส่วนต้นของชื่อโปรเจกต์
 *                      (PP1024 ↔ PP1024D01) · Start-date ที่ใกล้ที่สุดของ LCC หรือ RCC (จองรวมกัน)
 *                      Admin กำหนดเองทับได้รายโปรเจกต์
 *  • Onhand          : Item stock รวมทุกคลัง · กำลังเข้า = Pending Order ที่ยังไม่ปิด (ไม่เกิน 1 ปี)
 *  • จัดสรร           : ต่อ CODE เรียงตามวันที่ต้องใช้ — ใช้ Onhand ก่อน แล้วของที่จะเข้าทันวัน แล้วของที่เข้าไม่ทัน
 *  • ปิดโปรเจกต์       : Admin ปิดเอง หรืออัตโนมัติเมื่อเลยวันที่ต้องใช้ N วัน (ตั้งค่าได้, 0 = ปิดเองเท่านั้น)
 *  • ความเร็ว          : ผลคำนวณเก็บเป็น snapshot (cache + ชีต BomSnap) คำนวณใหม่เมื่อ Import / ปิด / ตั้งค่า
 *                      และเมื่อซิงค์ Stock / Pending Order / ตารางงาน
 *  ต้องมี Code.gs, OrderService.gs, StockAlert.gs (ใช้ helper ร่วมกัน)
 * ═════════════════════════════════════════════════════════════════════
 */
var BOM_PROJ_SHEET = 'BomProjects';
var BOM_PROJ_HEADER = ['PROJECT', 'TITLE', 'REV', 'FILE', 'ITEMS', 'NEED_DATE', 'STATUS', 'IMPORTED', 'CLOSED', 'NOTE'];
var BOM_ITEM_SHEET = 'BomItems';
var BOM_ITEM_HEADER = ['PROJECT', 'CODE', 'QTY', 'MATERIAL', 'TYPE', 'RATED', 'MAKER', 'REMARK', 'ITEM'];
var BOM_SNAP_SHEET = 'BomSnap', BOM_SNAP_CHUNK = 40000;
var BOM_DEF = { process: 'auto', autoClose: 0 };

function bomCfg_() {
  var c = {}; try { c = JSON.parse(schedProp_('BOM_CFG') || '{}') || {}; } catch (e) {}
  var a = Number(c.autoClose);
  return { process: { auto: 1, LCC: 1, RCC: 1 }[c.process] ? c.process : BOM_DEF.process, autoClose: a >= 0 && a <= 365 ? Math.round(a) : 0 };
}
function saveBomCfg(token, cfg) {
  requireAdmin_(token);
  cfg = cfg || {};
  var a = Number(cfg.autoClose);
  if (!(a >= 0 && a <= 365)) throw new Error('จำนวนวันปิดอัตโนมัติต้องอยู่ระหว่าง 0–365 (0 = ปิดเองเท่านั้น)');
  var c = { process: { auto: 1, LCC: 1, RCC: 1 }[cfg.process] ? cfg.process : 'auto', autoClose: Math.round(a) };
  schedProp_('BOM_CFG', JSON.stringify(c));
  audit_([['Admin', 'ตั้งค่า Booking BOM', '', 'Process ' + c.process + ' · ปิดอัตโนมัติ ' + (c.autoClose ? c.autoClose + ' วันหลังวันที่ต้องใช้' : 'ปิด')]]);
  bomSnapBuild_();
  return c;
}

// ── sheets ──────────────────────────────────────────────────────────
function bomSheet_(name, header) {
  var s = ss_().getSheetByName(name);
  if (!s) { s = ss_().insertSheet(name); s.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold'); s.setFrozenRows(1); }
  return s;
}
function bomProjects_() {
  var s = ss_().getSheetByName(BOM_PROJ_SHEET), out = [];
  if (!s || s.getLastRow() < 2) return out;
  s.getRange(2, 1, s.getLastRow() - 1, BOM_PROJ_HEADER.length).getDisplayValues().forEach(function (r, i) {
    if (!r[0]) return;
    out.push({ row: i + 2, project: r[0], title: r[1], rev: r[2], file: r[3], items: Number(r[4]) || 0, needManual: r[5], status: r[6] || 'active',
      imported: r[7], closed: r[8], note: r[9] });
  });
  return out;
}
function bomItems_() {
  var s = ss_().getSheetByName(BOM_ITEM_SHEET), out = {};
  if (!s || s.getLastRow() < 2) return out;
  s.getRange(2, 1, s.getLastRow() - 1, BOM_ITEM_HEADER.length).getDisplayValues().forEach(function (r) {
    if (!r[0] || !r[1]) return;
    (out[r[0]] = out[r[0]] || []).push({ code: r[1], qty: Number(r[2]) || 0, material: r[3], type: r[4], rated: r[5], maker: r[6], remark: r[7], item: r[8] });
  });
  return out;
}
function bomKey_(p) { return String(p || '').toUpperCase().replace(/\s+/g, ''); }

// ── Admin: preview (stock name / onhand for the codes in the files) ──
function bomPreview(token, codes) {
  requireAdmin_(token);
  var want = {}, out = {};
  (codes || []).slice(0, 3000).forEach(function (c) { var k = ordCodeKey_(c); if (k) want[k] = 1; });
  var rows = []; try { rows = stockData_(); } catch (e) {}
  for (var i = 0; i < rows.length; i++) {
    var p = rows[i], t = p.indexOf('\t'); if (t < 0) continue;
    var k = ordCodeKey_(p.substring(0, t)); if (!want[k]) continue;
    var f = p.split('\t'), q = saNum_(f[3]), o = out[k] || (out[k] = { desc: f[1], unit: f[4], onhand: 0 });
    if (q != null) o.onhand = Math.round((o.onhand + q) * 1000) / 1000;
  }
  var exist = bomProjects_().reduce(function (m, p) { m[bomKey_(p.project)] = { rev: p.rev, status: p.status, items: p.items }; return m; }, {});
  return { stock: out, existing: exist };
}

// ── Admin: import ───────────────────────────────────────────────────
// list = [{ project, title, rev, file, items: [{ code, qty, material, type, rated, maker, remark, item }] }]
function importBom(token, list) {
  requireAdmin_(token);
  list = (list || []).slice(0, 50).map(function (p) {
    var items = {}, order = [];
    (p.items || []).slice(0, 2000).forEach(function (it) {
      var k = ordCodeKey_(it.code), q = Number(String(it.qty).replace(/,/g, ''));
      if (!k || !(q > 0)) return;
      if (!items[k]) { items[k] = { code: k, qty: 0, material: ordStr_(it.material, 120), type: ordStr_(it.type, 80), rated: ordStr_(it.rated, 80),
        maker: ordStr_(it.maker, 60), remark: ordStr_(it.remark, 120), item: ordStr_(it.item, 10) }; order.push(k); }
      else if (it.remark) items[k].remark = ordStr_((items[k].remark ? items[k].remark + ', ' : '') + it.remark, 120);
      items[k].qty = Math.round((items[k].qty + q) * 1000) / 1000;
    });
    return { project: bomKey_(ordStr_(p.project, 40)), title: ordStr_(p.title, 120), rev: ordStr_(p.rev, 10), file: ordStr_(p.file, 120),
      items: order.map(function (k) { return items[k]; }) };
  }).filter(function (p) { return /^[A-Z0-9][A-Z0-9_\-\/.]{2,39}$/.test(p.project) && p.items.length; });
  if (!list.length) throw new Error('ไม่พบข้อมูลที่ Import ได้ (ต้องมี PROJECT, CODE และจำนวน > 0)');
  var res = withLock_(function () {
    var ps = bomSheet_(BOM_PROJ_SHEET, BOM_PROJ_HEADER), is = bomSheet_(BOM_ITEM_SHEET, BOM_ITEM_HEADER);
    var names = {}; list.forEach(function (p) { names[p.project] = 1; });
    // items: keep other projects, replace these
    var keep = is.getLastRow() > 1 ? is.getRange(2, 1, is.getLastRow() - 1, BOM_ITEM_HEADER.length).getDisplayValues().filter(function (r) { return r[0] && !names[r[0]]; }) : [];
    var add = []; list.forEach(function (p) { p.items.forEach(function (it) { add.push([p.project, it.code, String(it.qty), it.material, it.type, it.rated, it.maker, it.remark, it.item]); }); });
    var all = keep.concat(add);
    if (is.getLastRow() > 1) is.getRange(2, 1, is.getLastRow() - 1, BOM_ITEM_HEADER.length).clearContent();
    if (all.length) {
      if (is.getMaxRows() < all.length + 1) is.insertRowsAfter(is.getMaxRows(), all.length + 1 - is.getMaxRows());
      var rg = is.getRange(2, 1, all.length, BOM_ITEM_HEADER.length); rg.setNumberFormat('@'); rg.setValues(all);
    }
    var cur = {}; bomProjects_().forEach(function (p) { cur[p.project] = p; });
    var now = nowStr_(), audits = [];
    list.forEach(function (p) {
      var c = cur[p.project];
      var row = [p.project, p.title || (c && c.title) || '', p.rev || '', p.file, String(p.items.length), c ? c.needManual : '', 'active', now, '',
        c ? ('Rev ' + (c.rev || '-') + ' → ' + (p.rev || '-') + ' (' + now.split(' ')[0] + ')') : ''];
      writeText_(ps, c ? c.row : ps.getLastRow() + 1, 1, row);
      audits.push(['Admin', c ? 'Import BOM ทับ (Revision ใหม่)' : 'Import BOM', p.project, p.items.length + ' รายการ · ' + p.file]);
    });
    audit_(audits);
    return list.map(function (p) { return { project: p.project, items: p.items.length, replaced: !!cur[p.project] }; });
  });
  bomSnapBuild_();
  return res;
}

// ── Admin: project actions ──────────────────────────────────────────
// action: 'close' | 'reopen' | 'delete' | 'needDate' (value = dd/MM/yyyy or '' to use the schedule)
function bomProjectAction(token, project, action, value) {
  requireAdmin_(token);
  var key = bomKey_(project);
  if (action === 'needDate' && value && !/^\d{2}\/\d{2}\/\d{4}$/.test(value)) throw new Error('วันที่ต้องเป็นรูปแบบ วว/ดด/ปปปป');
  withLock_(function () {
    var p = bomProjects_().filter(function (x) { return x.project === key; })[0];
    if (!p) throw new Error('ไม่พบโปรเจกต์ ' + key);
    var s = ss_().getSheetByName(BOM_PROJ_SHEET);
    if (action === 'close') s.getRange(p.row, 7, 1, 3).setValues([['closed', p.imported, nowStr_() + ' (Admin)']]);
    else if (action === 'reopen') s.getRange(p.row, 7, 1, 3).setValues([['active', p.imported, '']]);
    else if (action === 'needDate') { var r = s.getRange(p.row, 6); r.setNumberFormat('@'); r.setValue(value || ''); }
    else if (action === 'delete') {
      s.deleteRow(p.row);
      var is = ss_().getSheetByName(BOM_ITEM_SHEET);
      if (is && is.getLastRow() > 1) {
        var keep = is.getRange(2, 1, is.getLastRow() - 1, BOM_ITEM_HEADER.length).getDisplayValues().filter(function (r) { return r[0] && r[0] !== key; });
        is.getRange(2, 1, is.getLastRow() - 1, BOM_ITEM_HEADER.length).clearContent();
        if (keep.length) { var rg = is.getRange(2, 1, keep.length, BOM_ITEM_HEADER.length); rg.setNumberFormat('@'); rg.setValues(keep); }
      }
    } else throw new Error('คำสั่งไม่ถูกต้อง');
    audit_([['Admin', { close: 'ปิดโปรเจกต์ BOM', reopen: 'เปิดโปรเจกต์ BOM อีกครั้ง', needDate: 'กำหนดวันที่ต้องใช้ BOM', 'delete': 'ลบโปรเจกต์ BOM' }[action], key, value || '']]);
  });
  return bomView_(bomSnapBuild_());
}

// ── need dates from the production schedule ─────────────────────────
function bomDate_(s) { var m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(String(s || '')); return m ? new Date(+m[3], +m[2] - 1, +m[1]).getTime() : null; }
function bomFmt_(t) { return t == null ? '' : Utilities.formatDate(new Date(t), 'Asia/Bangkok', 'dd/MM/yyyy'); }
function bomNeed_(p, sched, cfg) {
  if (bomDate_(p.needManual) != null) return { t: bomDate_(p.needManual), src: 'manual', label: 'กำหนดเอง' };
  var key = bomKey_(p.project), best = null;
  sched.forEach(function (r) { var pn = bomKey_(r[0]); if (pn && key.indexOf(pn) === 0 && (!best || pn.length > best.length)) best = pn; });
  if (!best) return { t: null, src: 'none', label: 'ไม่พบในตารางงาน' };
  // LCC and RCC are booked together → need date = the earliest Start-date of either process
  var rows = sched.filter(function (r) { return bomKey_(r[0]) === best; });
  var first = function (rs) {
    var b = null;
    rs.forEach(function (r) { var d = bomDate_(r[9]) || bomDate_(r[10]); if (d != null && (!b || d < b.t)) b = { t: d, proc: String(r[8]).trim() }; });
    return b;
  };
  var cc = first(rows.filter(function (r) { return /^(LCC|RCC)$/.test(String(r[8]).toUpperCase().replace(/\s+/g, '')); }));
  if (cc) return { t: cc.t, src: 'sched', label: 'ตารางงาน ' + best + ' · ' + cc.proc + ' (Start ใกล้ที่สุดของ LCC/RCC)' };
  var t;
  var any = first(rows);
  if (any) return { t: any.t, src: 'sched', label: 'ตารางงาน ' + best + ' · ไม่มี LCC/RCC → ' + any.proc + ' (Start)' };
  t = rows.map(function (r) { return bomDate_(r[6]) || bomDate_(r[5]); }).filter(function (x) { return x != null; })[0];
  return t != null ? { t: t, src: 'sched', label: 'ตารางงาน ' + best + ' · Delivery' } : { t: null, src: 'none', label: 'ตารางงาน ' + best + ' ไม่มีวันที่' };
}

// ── compute ─────────────────────────────────────────────────────────
function bomCompute_() {
  var cfg = bomCfg_(), projs = bomProjects_(), itemsBy = bomItems_(), sched = [];
  try { sched = getSchedules(); } catch (e) {}
  var today = new Date(); today = new Date(today.getFullYear(), today.getMonth(), today.getDate()).getTime();
  // auto-close
  if (cfg.autoClose > 0) {
    var s = ss_().getSheetByName(BOM_PROJ_SHEET), closed = [];
    projs.forEach(function (p) {
      if (p.status !== 'active') return;
      var n = bomNeed_(p, sched, cfg);
      if (n.t != null && (today - n.t) / 864e5 > cfg.autoClose) {
        s.getRange(p.row, 7, 1, 3).setValues([['closed', p.imported, nowStr_() + ' (อัตโนมัติ)']]); p.status = 'closed'; p.closed = nowStr_() + ' (อัตโนมัติ)'; closed.push(p.project);
      }
    });
    if (closed.length) audit_([['System', 'ปิดโปรเจกต์ BOM อัตโนมัติ', '', closed.join(', ')]]);
  }
  var active = projs.filter(function (p) { return p.status === 'active'; });
  var want = {};
  active.forEach(function (p) { p.need = bomNeed_(p, sched, cfg); (itemsBy[p.project] || []).forEach(function (it) { want[ordCodeKey_(it.code)] = 1; }); });
  // onhand (all locations) + descriptions
  var stock = {};
  try {
    var rows = stockData_();
    for (var i = 0; i < rows.length; i++) {
      var pr = rows[i], tb = pr.indexOf('\t'); if (tb < 0) continue;
      var k = ordCodeKey_(pr.substring(0, tb)); if (!want[k]) continue;
      var f = pr.split('\t'), q = saNum_(f[3]), o = stock[k] || (stock[k] = { desc: f[1], unit: f[4], onhand: 0 });
      if (q != null) o.onhand = Math.round((o.onhand + q) * 1000) / 1000;
    }
  } catch (e) {}
  // incoming = Pending Order not closed (≤ 1 year)
  var inc = {};
  try {
    var b = pendingBundle_(false), txt = b && b.b64 ? Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(b.b64), 'application/x-gzip')).getDataAsString() : '';
    txt.split('\n').forEach(function (l) {
      var t2 = l.indexOf('\t'); if (t2 < 0) return;
      var k2 = ordCodeKey_(l.substring(0, t2)); if (!want[k2]) return;
      var g = l.split('\t'); if (ORD_PO_DONE.test(g[4] || '')) return;
      var due = bomDate_(g[9]) || bomDate_(g[7]) || bomDate_(g[6]);
      if (due != null && Math.abs(due - today) > 365 * 864e5) return;
      var qty = Number(g[13]) || 0; if (!(qty > 0)) return;
      (inc[k2] = inc[k2] || []).push({ qty: qty, due: due, pr: g[1], po: g[2], status: g[4] });
    });
  } catch (e) {}
  var sa = {}; try { sa = saSettings_(); } catch (e) {}
  // allocate per code by need date
  var demand = {};
  active.forEach(function (p) { (itemsBy[p.project] || []).forEach(function (it) {
    var k = ordCodeKey_(it.code); (demand[k] = demand[k] || []).push({ p: p, it: it, t: p.need.t });
  }); });
  var res = {}, codes = [];
  Object.keys(demand).forEach(function (k) {
    var st = stock[k] || { desc: '', unit: '', onhand: 0 }, left = Math.max(0, st.onhand);
    var ins = (inc[k] || []).map(function (x) { return { qty: x.qty, due: x.due, pr: x.pr, po: x.po, status: x.status, left: x.qty }; })
      .sort(function (a, b) { return (a.due == null ? 9e15 : a.due) - (b.due == null ? 9e15 : b.due); });
    var dl = demand[k].sort(function (a, b) { return (a.t == null ? 9e15 : a.t) - (b.t == null ? 9e15 : b.t) || String(a.p.project).localeCompare(String(b.p.project)); });
    var booked = 0;
    dl.forEach(function (d) {
      var need = d.it.qty, fromStock = Math.min(left, need); left -= fromStock; need -= fromStock; booked += d.it.qty;
      var onTime = 0, late = 0, eta = null;
      ins.forEach(function (x) {
        if (need <= 0 || x.left <= 0) return;
        if (d.t != null && x.due != null && x.due > d.t) return;      // first pass: in time only
        var u = Math.min(x.left, need); x.left -= u; need -= u; onTime += u; if (x.due != null) eta = Math.max(eta || 0, x.due);
      });
      ins.forEach(function (x) {
        if (need <= 0 || x.left <= 0) return;
        var u = Math.min(x.left, need); x.left -= u; need -= u; late += u; if (x.due != null) eta = Math.max(eta || 0, x.due);
      });
      var stt = need > 0 ? 'short' : late > 0 ? 'late' : onTime > 0 ? 'incoming' : 'ok';
      (res[d.p.project] = res[d.p.project] || []).push({ code: d.it.code, qty: d.it.qty, material: d.it.material || st.desc, type: d.it.type, rated: d.it.rated,
        maker: d.it.maker, remark: d.it.remark, item: d.it.item, unit: st.unit, st: stt, fromStock: fromStock, onTime: onTime, late: late, short: Math.max(0, need),
        eta: bomFmt_(eta), known: !!stock[k] });
    });
    var incoming = ins.reduce(function (s2, x) { return s2 + x.qty; }, 0), bal = Math.round((st.onhand + incoming - booked) * 1000) / 1000;
    var mx = sa[k] && sa[k].max != null ? sa[k].max : null;
    codes.push({ code: k, desc: st.desc || (dl[0].it.material || ''), unit: st.unit, onhand: st.onhand, incoming: Math.round(incoming * 1000) / 1000,
      booked: Math.round(booked * 1000) / 1000, balance: bal, max: mx, over: mx != null && bal > mx, short: bal < 0, known: !!stock[k],
      projects: dl.map(function (d) { return [d.p.project, d.it.qty]; }), po: (inc[k] || []).slice(0, 5).map(function (x) { return { qty: x.qty, due: bomFmt_(x.due), pr: x.pr, po: x.po, status: x.status }; }) });
  });
  var projects = projs.map(function (p) {
    var its = p.status === 'active' ? (res[p.project] || []) : (itemsBy[p.project] || []).map(function (it) { return { code: it.code, qty: it.qty, material: it.material, type: it.type, rated: it.rated, maker: it.maker, remark: it.remark, item: it.item, st: 'closed' }; });
    var c = { ok: 0, incoming: 0, late: 0, short: 0 }; its.forEach(function (x) { if (c[x.st] != null) c[x.st]++; });
    var n = p.need || (p.status === 'active' ? bomNeed_(p, sched, cfg) : { t: null, label: '' });
    return { project: p.project, title: p.title, rev: p.rev, file: p.file, status: p.status, imported: p.imported, closed: p.closed, note: p.note,
      needDate: bomFmt_(n.t), needSrc: n.src || '', needLabel: n.label || '', needManual: p.needManual, daysLeft: n.t != null ? Math.round((n.t - today) / 864e5) : null,
      count: c, total: its.length, items: its };
  });
  projects.sort(function (a, b) { return (a.status === 'active' ? 0 : 1) - (b.status === 'active' ? 0 : 1) || (a.daysLeft == null ? 9e9 : a.daysLeft) - (b.daysLeft == null ? 9e9 : b.daysLeft); });
  codes.sort(function (a, b) { return (a.balance - b.balance) || String(a.code).localeCompare(String(b.code)); });
  return { projects: projects, codes: codes, cfg: cfg, built: nowStr_(), stockTime: (stockInfo_() || {}).time || '', poTime: (poInfo_() || {}).time || '',
    schedTime: (schedSyncInfo_() || {}).time || '' };
}

// ── snapshot (CacheService + sheet BomSnap in Helpdesk_ItemStock) ────
function bomSnapPut_(snap) {
  var b64 = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(JSON.stringify(snap), 'application/json')).getBytes()), parts = [];
  for (var i = 0; i < b64.length; i += BOM_SNAP_CHUNK) parts.push(b64.substring(i, i + BOM_SNAP_CHUNK));
  try { var put = { bom_n: String(parts.length) }; parts.forEach(function (p, i) { put['bom_' + i] = p; }); CacheService.getScriptCache().putAll(put, 21600); } catch (e) { console.warn(e); }
  try {
    var ss = stockSS_(), sh = ss.getSheetByName(BOM_SNAP_SHEET) || ss.insertSheet(BOM_SNAP_SHEET);
    sh.clearContents(); sh.getRange(1, 1, parts.length, 1).setValues(parts.map(function (p) { return [p]; }));
  } catch (e) { console.warn(e); }
  schedProp_('BOM_VER', String(Date.now()));
  __bomBooked = null;
  try { saBump_(); } catch (e) {}                                   // Stock Alert recomputes with the new bookings
  try { CacheService.getScriptCache().remove('rs_flags'); } catch (e) {}
}
// { CODE: { qty: booked by active projects, balance, projects: [[project, qty], …] } } — read from the snapshot (cheap)
var __bomBooked = null;
function bomBooked_() {
  if (__bomBooked) return __bomBooked;
  var m = {}, s = null;
  try { s = bomSnapGet_(); } catch (e) {}
  (s && s.codes || []).forEach(function (c) { if (c.booked > 0) m[c.code] = { qty: c.booked, balance: c.balance, onhand: c.onhand, incoming: c.incoming, projects: c.projects }; });
  return (__bomBooked = m);
}
function bomSnapGet_() {
  var b64 = '';
  try {
    var c = CacheService.getScriptCache(), n = Number(c.get('bom_n') || 0);
    if (n) { var keys = []; for (var i = 0; i < n; i++) keys.push('bom_' + i); var got = c.getAll(keys);
      if (keys.every(function (k) { return got[k] != null; })) b64 = keys.map(function (k) { return got[k]; }).join(''); }
  } catch (e) {}
  if (!b64) {
    try {
      var sh = stockSS_().getSheetByName(BOM_SNAP_SHEET);
      if (sh && sh.getLastRow()) {
        b64 = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(function (r) { return r[0]; }).join('');
        var parts = {}, k = 0; for (var j = 0; j < b64.length; j += BOM_SNAP_CHUNK) parts['bom_' + (k++)] = b64.substring(j, j + BOM_SNAP_CHUNK);
        parts.bom_n = String(k); try { CacheService.getScriptCache().putAll(parts, 21600); } catch (e) {}
      }
    } catch (e) {}
  }
  if (!b64) return null;
  try { return JSON.parse(Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(b64), 'application/x-gzip')).getDataAsString()); } catch (e) { return null; }
}
function bomSnapBuild_() { var s = bomCompute_(); bomSnapPut_(s); return s; }
// called after Stock / Pending Order / Schedule syncs (only when there are BOM projects)
function bomRefresh_() { if (ss_().getSheetByName(BOM_PROJ_SHEET)) bomSnapBuild_(); }
function bomView_(snap) {
  snap = snap || { projects: [], codes: [] };
  var today = new Date(); today = new Date(today.getFullYear(), today.getMonth(), today.getDate()).getTime();
  (snap.projects || []).forEach(function (p) { var t = bomDate_(p.needDate); p.daysLeft = t != null ? Math.round((t - today) / 864e5) : null; });
  return snap;
}

// ── public: everything the page needs (everyone can view; Admin edits) ──
function getBom(token) {
  var t0 = Date.now(), snap = bomSnapGet_();
  if (!snap) snap = ss_().getSheetByName(BOM_PROJ_SHEET) ? bomSnapBuild_() : { projects: [], codes: [], cfg: bomCfg_() };
  var out = bomView_(snap); out.admin = isAdmin_(token); out.cfg = bomCfg_();
  try { out.exports = bomExpLast_(); } catch (e) { out.exports = {}; }
  perf_('getBom', t0);
  return out;
}
