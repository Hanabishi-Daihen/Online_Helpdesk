/**
 * ═════════════════════════════════════════════════════════════════════
 *  BOOKING BOM — Phase 3 : Export ใบเบิก (By Stock / By Project) · รายงาน PDF รายการที่ขาด
 *  • By Stock   : CODE 8 หลักจาก BOM ของโปรเจกต์ · จำนวน = ตาม BOM · เลขที่ MF253A
 *  • By Project : CODE ของงาน (ไม่ใช่ 8 หลัก) ใน Item stock ที่มีคำว่าโปรเจกต์ตัว C
 *                 (PP1024D01 → …PP1024C… เช่น 3694PP1024C00403) · จำนวน = Onhand รายคลัง · เลขที่ MF253P
 *  • 1 หน้า A4 = 1 ชีต (หัวฟอร์ม + รายการ + ลายเซ็น + ท้ายกระดาษ "Page x of n" และรหัสฟอร์ม)
 *                 หน้าแรกมี Rev ที่มุมขวาบน — รหัสฟอร์ม / Rev แก้ได้ที่ตั้งค่า Booking BOM
 *  • Preview ก่อนออกไฟล์ (ลบ / เพิ่ม / แก้จำนวน / เลือกคลัง) → Google Sheet แม่แบบ → PDF + Excel
 *  • คอลัมน์ G และ I เว้นว่างกว้างพิเศษสำหรับติด Barcode · ความสูงแถว 39.75–55 pt ตามความยาวข้อความ
 *  • Export แล้วปิดโปรเจกต์ (ปลดยอดจอง) ได้ทันที (Admin)
 *  ต้องมี BomBooking.gs, BomTemplate.gs, OrderService.gs
 * ═════════════════════════════════════════════════════════════════════
 */
var BOM_EXP_SHEET = 'BomExports';
var BOM_EXP_HEADER = ['TIME', 'PROJECT', 'KIND', 'DOC_NO', 'ITEMS', 'QTY', 'PDF', 'XLSX', 'BY', 'CLOSED'];
// template rows after the Rev row is inserted on top:
// 1 Rev · 2–6 heading · 7 table header · 8–30 items (23) · 31 gap · 32–36 signatures · then spacer + footer
var BOM_FORM = { ROW0: 8, TPL_ROWS: 23, MAX_ROWS: 300, COLS: 9, RATIO: 1.43, REV_H: 24, FOOT_H: 28 };
// column widths (px) — G and I are wide blank boxes for barcode stickers
var BOM_FORM_W = [30, 132, 250, 50, 50, 54, 170, 44, 170];
// row height (px): 53 px ≈ 39.75 pt (1 line) · 73 px ≈ 55 pt (2 lines)
var BOM_FORM_H = { min: 53, max: 73, wrap: 46 };
var BOM_KIND = {
  stock: { label: 'By Stock', file: 'ByStock', doc: 'MF253A', b64: function () { return typeof BOM_TPL_STOCK_B64 === 'string' ? BOM_TPL_STOCK_B64 : ''; } },
  project: { label: 'By Project', file: 'ByProject', doc: 'MF253P', b64: function () { return typeof BOM_TPL_PROJECT_B64 === 'string' ? BOM_TPL_PROJECT_B64 : ''; } }
};

// ── helpers ─────────────────────────────────────────────────────────
function bomKind_(k) { return BOM_KIND[k] ? k : 'stock'; }
function bomIs8_(k) { return /^\d{8}$/.test(String(k || '')); }
// PP1024D01 → PP1024C · PP1024D → PP1024C
function bomProjPrefix_(project) {
  var k = bomKey_(project), m = /^([A-Z]+\d+)D/.exec(k);
  return m ? m[1] + 'C' : k;
}
// stock rows → { KEY: { code, desc, unit, cost, onhand, locs: [{ loc, pk, qty }] } } (locations sorted by qty desc)
function bomStockIdx_(keep) {
  var out = {}, rows = [];
  try { rows = stockData_(); } catch (e) {}
  for (var i = 0; i < rows.length; i++) {
    var p = rows[i], t = p.indexOf('\t'); if (t < 0) continue;
    var k = ordCodeKey_(p.substring(0, t)); if (!k || !keep(k)) continue;
    var f = p.split('\t'), q = saNum_(f[3]) || 0, o = out[k];
    if (!o) o = out[k] = { code: String(f[0]).trim(), desc: f[1] || '', unit: f[4] || '', cost: saNum_(f[8]) || 0, onhand: 0, locs: [], by: '' };
    if (!o.by && f[7]) o.by = String(f[7]).trim();
    o.onhand = Math.round((o.onhand + q) * 1000) / 1000;
    var loc = String(f[5] || '').trim(), pk = String(f[6] || '').trim(), L = null;
    o.locs.forEach(function (x) { if (x.loc === loc && x.pk === pk) L = x; });
    if (L) L.qty = Math.round((L.qty + q) * 1000) / 1000; else o.locs.push({ loc: loc, pk: pk, qty: q });
    if (!o.cost && saNum_(f[8])) o.cost = saNum_(f[8]);
  }
  Object.keys(out).forEach(function (k) { out[k].locs.sort(function (a, b) { return b.qty - a.qty; }); });
  return out;
}
function bomExpFolder_() {
  var id = schedProp_('BOM_EXP_FOLDER_ID');
  if (id) { try { var f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {} }
  var f2 = ordFolder_().createFolder('BOM_Export');
  schedProp_('BOM_EXP_FOLDER_ID', f2.getId());
  return f2;
}
function bomTplId_(kind) {
  var K = BOM_KIND[kind], pid = 'BOM_TPL_' + kind.toUpperCase(), ver = 'v' + (typeof BOM_TEMPLATE_VER === 'number' ? BOM_TEMPLATE_VER : 1);
  var id = schedProp_(pid);
  if (id && schedProp_(pid + '_VER') === ver) { try { if (!DriveApp.getFileById(id).isTrashed()) return id; } catch (e) {} }
  var b64 = K.b64();
  if (!b64) throw new Error('ไม่พบแม่แบบใบเบิก ' + K.label + ' — ตรวจว่ามีไฟล์ BomTemplate.gs ในโปรเจกต์');
  var nid = ordConvertXlsx_(Utilities.newBlob(Utilities.base64Decode(b64), MimeType.MICROSOFT_EXCEL, K.file + '_253.xlsx'), 'BOM_' + K.file + '_Template');
  if (id && id !== nid) { try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {} }
  schedProp_(pid, nid); schedProp_(pid + '_VER', ver);
  return nid;
}
// last exports per project (small JSON in a property — read by getBom)
function bomExpLast_() { try { return JSON.parse(schedProp_('BOM_EXP_LAST') || '{}') || {}; } catch (e) { return {}; } }
function bomExpLog_(rec) {
  var s = bomSheet_(BOM_EXP_SHEET, BOM_EXP_HEADER);
  writeText_(s, s.getLastRow() + 1, 1, [rec.time, rec.project, rec.kind, rec.docNo, String(rec.items), String(rec.qty), rec.pdf || '', rec.xlsx || '', rec.by, rec.closed ? 'Y' : '']);
  var m = bomExpLast_(), L = m[rec.project] || [];
  L.unshift({ t: rec.time, k: rec.kind, n: rec.items, pdf: rec.pdf || '', xlsx: rec.xlsx || '' });
  m[rec.project] = L.slice(0, 3);
  var keys = Object.keys(m);
  if (keys.length > 60) { keys.sort(function (a, b) { return String(m[b][0].t).localeCompare(String(m[a][0].t)); }); keys.slice(60).forEach(function (k) { delete m[k]; }); }
  schedProp_('BOM_EXP_LAST', JSON.stringify(m));
  dataBump_('BOM_EXP_VER');
}

// ── public: rows for the preview ────────────────────────────────────
// kind = 'stock' | 'project' · prefix = override for By Project (default from the project number)
function bomExportPrep(token, project, kind, prefix) {
  var t0 = Date.now(), key = bomKey_(project); kind = bomKind_(kind);
  var p = bomProjects_().filter(function (x) { return x.project === key; })[0];
  if (!p) throw new Error('ไม่พบโปรเจกต์ ' + key);
  var rows = [], excluded = [], pre = '';
  if (kind === 'stock') {
    var items = bomItems_()[key] || [], snapIt = {};
    try { ((bomSnapGet_() || {}).projects || []).forEach(function (x) { if (x.project === key) x.items.forEach(function (i) { snapIt[ordCodeKey_(i.code)] = i; }); }); } catch (e) {}
    var want = {}; items.forEach(function (it) { want[ordCodeKey_(it.code)] = 1; });
    var st = bomStockIdx_(function (k) { return want[k]; });
    items.forEach(function (it) {
      var k = ordCodeKey_(it.code);
      if (!bomIs8_(k)) { var s0 = st[k], L0 = s0 && s0.locs[0];
        excluded.push({ code: s0 ? s0.code : it.code, material: (s0 && s0.desc) || it.material || '', qty: it.qty, store: L0 ? L0.loc : '', pk: L0 ? L0.pk : '', unit: (s0 && s0.unit) || '',
          onhand: s0 ? s0.onhand : 0, locs: s0 ? s0.locs : [], known: !!s0, by: s0 ? s0.by : '', why: 'ไม่ใช่ CODE 8 หลัก' }); return; }
      var s = st[k], L = s && s.locs[0], b = snapIt[k] || {};
      rows.push({ code: s ? s.code : it.code, material: (s && s.desc) || it.material || '', store: L ? L.loc : '', pk: L ? L.pk : '', qty: it.qty,
        unit: (s && s.unit) || '', onhand: s ? s.onhand : 0, by: s ? s.by : '', locs: s ? s.locs : [], known: !!s, st: b.st || '', short: b.short || 0, late: b.late || 0, eta: b.eta || '' });
    });
  } else {
    pre = bomKey_(prefix || bomProjPrefix_(key));
    if (pre.length < 4) throw new Error('Prefix ต้องยาวอย่างน้อย 4 ตัวอักษร');
    var st2 = bomStockIdx_(function (k) { return k.indexOf(pre) >= 0 && !bomIs8_(k); });
    Object.keys(st2).sort().forEach(function (k) {
      var s = st2[k];
      s.locs.forEach(function (L) {
        if (!(L.qty > 0)) { excluded.push({ code: s.code, material: s.desc, qty: 0, store: L.loc, pk: L.pk, unit: s.unit, onhand: s.onhand, locs: s.locs, known: true, by: s.by, why: 'Onhand ' + L.qty + ' (' + (L.loc || '-') + ')' }); return; }
        rows.push({ code: s.code, material: s.desc, store: L.loc, pk: L.pk, qty: L.qty, unit: s.unit, onhand: s.onhand, locs: s.locs, known: true, by: s.by });
      });
    });
  }
  perf_('bomExportPrep', t0);
  return { project: key, title: p.title, rev: p.rev, status: p.status, kind: kind, prefix: pre, docNo: BOM_KIND[kind].doc,
    date: Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy'), rows: rows, excluded: excluded,
    stockTime: (stockInfo_() || {}).time || '', last: bomExpLast_()[key] || [], admin: isAdmin_(token) };
}
// public: search Item stock to add a row in the preview (CODE prefix or words in the description)
function bomExportLookup(token, q) {
  q = String(q || '').trim().toUpperCase(); if (q.length < 3) return [];
  var kq = ordCodeKey_(q), words = q.split(/\s+/).filter(String), n = 0;
  var st = bomStockIdx_(function () { return true; }), out = [];
  var keys = Object.keys(st);
  keys.forEach(function (k) { if (kq && k.indexOf(kq) === 0) out.push(k); });
  if (out.length < 25) keys.forEach(function (k) {
    if (out.length >= 25 || (kq && k.indexOf(kq) === 0)) return;
    var h = (k + ' ' + st[k].desc).toUpperCase();
    if (words.every(function (w) { return h.indexOf(w) >= 0; })) out.push(k);
  });
  return out.slice(0, 25).map(function (k) { var s = st[k]; return { code: s.code, material: s.desc, unit: s.unit, onhand: s.onhand, by: s.by, locs: s.locs.slice(0, 8) }; });
}

// ── public: fill the requisition form → PDF / Excel ─────────────────
// req = { project, kind, docNo, date, job, fmt: 'pdf'|'xlsx'|'both', height: 'auto'|'min'|'max', close, rows: [{ code, material, store, pk, qty, unit }] }
function exportBomForm(token, req) {
  var t0 = Date.now(); req = req || {};
  var kind = bomKind_(req.kind), K = BOM_KIND[kind], key = bomKey_(req.project), admin = isAdmin_(token);
  var p = bomProjects_().filter(function (x) { return x.project === key; })[0];
  if (!p) throw new Error('ไม่พบโปรเจกต์ ' + key);
  if (req.close && !admin) throw new Error('การปิดโปรเจกต์ต้องใช้สิทธิ์ Admin');
  var rows = (req.rows || []).slice(0, BOM_FORM.MAX_ROWS + 1).map(function (r) {
    var q = Number(String(r.qty == null ? '' : r.qty).replace(/,/g, ''));
    return { code: ordStr_(r.code, 40), material: ordStr_(r.material, 160), store: ordStr_(r.store, 20), pk: ordStr_(r.pk, 20), qty: q, unit: ordStr_(r.unit, 12) };
  }).filter(function (r) { return r.code || r.material; });
  if (!rows.length) throw new Error('ไม่มีรายการให้ Export');
  if (rows.length > BOM_FORM.MAX_ROWS) throw new Error('ใบเบิก 1 ชุดมีได้ไม่เกิน ' + BOM_FORM.MAX_ROWS + ' รายการ');
  var bad = rows.filter(function (r) { return !(r.qty > 0); });
  if (bad.length) throw new Error('จำนวนต้องมากกว่า 0 — ตรวจ ' + bad.slice(0, 5).map(function (r) { return r.code || r.material; }).join(', '));
  var fmt = { pdf: 'pdf', xlsx: 'xlsx' }[req.fmt] || 'both';
  var o = { kind: kind, rows: rows, docNo: ordStr_(req.docNo || K.doc, 30), date: ordStr_(req.date || Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy'), 20),
    job: ordStr_(req.job || key, 40), height: { min: 1, max: 1 }[req.height] ? req.height : 'auto' };
  var stamp = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd_HHmm');
  var base = (key + '_' + K.file + '_' + stamp).replace(/[\\\/:*?"<>|]+/g, '_'), folder = bomExpFolder_(), files = [];
  var copy = DriveApp.getFileById(bomTplId_(kind)).makeCopy('_tmp_' + base, folder);
  try {
    var ss = SpreadsheetApp.openById(copy.getId()), sh = ss.getSheets()[0];
    ss.getSheets().forEach(function (x) { if (x.getSheetId() !== sh.getSheetId()) ss.deleteSheet(x); });
    bomFormBuild_(ss, sh, o, bomCfg_());
    SpreadsheetApp.flush();
    (fmt === 'both' ? ['pdf', 'xlsx'] : [fmt]).forEach(function (f) {
      var blob = bomFetchExport_(ss.getId(), f).setName(base + (f === 'pdf' ? '.pdf' : '.xlsx'));
      var file = folder.createFile(blob);
      try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
      var b = blob.getBytes();
      files.push({ fmt: f, name: blob.getName(), mime: blob.getContentType(), url: file.getUrl(), b64: b.length < 9 * 1024 * 1024 ? Utilities.base64Encode(b) : '' });
    });
  } finally { try { copy.setTrashed(true); } catch (e) {} }
  var qty = Math.round(rows.reduce(function (s, r) { return s + r.qty; }, 0) * 1000) / 1000, now = nowStr_(), closed = false;
  withLock_(function () {
    bomExpLog_({ time: now, project: key, kind: kind, docNo: o.docNo, items: rows.length, qty: qty, by: admin ? 'Admin' : 'User', closed: !!req.close,
      pdf: (files.filter(function (f) { return f.fmt === 'pdf'; })[0] || {}).url, xlsx: (files.filter(function (f) { return f.fmt === 'xlsx'; })[0] || {}).url });
    if (req.close) {
      var cur = bomProjects_().filter(function (x) { return x.project === key; })[0];
      if (cur && cur.status === 'active') {
        ss_().getSheetByName(BOM_PROJ_SHEET).getRange(cur.row, 7, 1, 3).setValues([['closed', cur.imported, now + ' (Export ' + K.label + ')']]);
        closed = true;
      }
    }
  });
  audit_([[admin ? 'Admin' : 'User', 'Export ใบเบิก BOM ' + K.label + ' (' + fmt.toUpperCase() + ')', key, rows.length + ' รายการ · ' + o.docNo + (closed ? ' · ปิดโปรเจกต์' : '')]]);
  var view = closed ? bomView_(bomSnapBuild_()) : null;
  perf_('exportBomForm', t0);
  return { files: files, closed: closed, view: view, last: bomExpLast_()[key] || [] };
}
function bomRowH_(o, r) {
  if (o.height === 'min') return BOM_FORM_H.min;
  if (o.height === 'max') return BOM_FORM_H.max;
  return r && String(r.material || '').length > BOM_FORM_H.wrap ? BOM_FORM_H.max : BOM_FORM_H.min;
}
function bomQtyTxt_(q) { return String(Math.round(q * 1000) / 1000); }
function bomFormW_() { return BOM_FORM_W.reduce(function (s, w) { return s + w; }, 0); }
// split the rows into A4 pages: every page = heading + items + signatures + footer, sized to the A4 aspect ratio
function bomPages_(o, fixedPx) {
  var budget = Math.round(bomFormW_() * BOM_FORM.RATIO) - fixedPx, pages = [], cur = [], used = 0;
  o.rows.forEach(function (r) {
    var h = bomRowH_(o, r);
    if (cur.length && used + h > budget) { pages.push(cur); cur = []; used = 0; }
    cur.push(r); used += h;
  });
  if (cur.length || !pages.length) pages.push(cur);
  return { pages: pages, budget: budget };
}
// base = converted template sheet → one sheet per page ("Page 1", "Page 2", …), base removed
function bomFormBuild_(ss, base, o, cfg) {
  var F = BOM_FORM;
  base.insertRowBefore(1); base.setRowHeight(1, F.REV_H);
  BOM_FORM_W.forEach(function (w, i) { base.setColumnWidth(i + 1, w); });
  var mc = base.getMaxColumns(); if (mc > F.COLS) base.deleteColumns(F.COLS + 1, mc - F.COLS);
  try { base.setFrozenRows(0); } catch (e) {}
  ordSet_(base, 'D4', 'เลขที่_____' + o.docNo + '__');
  ordSet_(base, 'A5', 'วันที่_____' + o.date + '_____');
  ordSet_(base, 'D5', 'เบอร์งาน__' + o.job + '__');
  var hs = function (a, b) { var s = 0; for (var r = a; r <= b; r++) s += base.getRowHeight(r); return s; };
  var tail = F.ROW0 + F.TPL_ROWS;                                    // gap row (31) + 5 signature rows (32–36)
  var P = bomPages_(o, hs(1, F.ROW0 - 1) + hs(tail, tail + 5) + F.FOOT_H), no = 0;
  P.pages.forEach(function (rows, pi) {
    var sh = base.copyTo(ss).setName('Page ' + (pi + 1));
    bomPageFill_(sh, o, rows, no, pi, P.pages.length, P.budget, cfg);
    no += rows.length;
  });
  ss.deleteSheet(base);
  return P.pages.length;
}
function bomPageFill_(sh, o, rows, no, pi, n, budget, cfg) {
  var F = BOM_FORM, r0 = F.ROW0, used = 0;
  rows.forEach(function (r) { used += bomRowH_(o, r); });
  var blanks = Math.max(0, Math.floor((budget - used) / BOM_FORM_H.min)), R = Math.max(1, rows.length + blanks);
  var spacer = Math.max(2, budget - used - blanks * BOM_FORM_H.min);
  if (R < F.TPL_ROWS) sh.deleteRows(r0 + R, F.TPL_ROWS - R);
  else if (R > F.TPL_ROWS) sh.insertRowsAfter(r0 + F.TPL_ROWS - 2, R - F.TPL_ROWS);     // inside the bordered block → keeps its format
  // items (blank rows fill the rest of the page)
  var vals = [];
  for (var i = 0; i < R; i++) {
    var x = rows[i];
    vals.push(x ? [String(no + i + 1), x.code, x.material, x.store, x.pk, bomQtyTxt_(x.qty), '', x.unit, ''] : ['', '', '', '', '', '', '', '', '']);
  }
  var rg = sh.getRange(r0, 1, R, F.COLS);
  rg.setNumberFormat('@').setValues(vals);
  rg.setVerticalAlignment('middle').setHorizontalAlignment('center').setFontWeight('normal').setWrap(false);
  rg.setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  [[1, 16], [2, 16], [3, 16], [4, 14], [5, 14], [6, 18], [7, 14], [8, 14], [9, 14]].forEach(function (c) { sh.getRange(r0, c[0], R, 1).setFontSize(c[1]); });
  sh.getRange(r0, 3, R, 1).setHorizontalAlignment('left').setWrap(true);
  sh.getRange(r0, 6, R, 1).setFontWeight('bold');
  for (var j = 0; j < R; j++) {
    var it = rows[j], len = it ? String(it.material || '').length : 0;
    sh.setRowHeight(r0 + j, it ? bomRowH_(o, it) : BOM_FORM_H.min);
    if (len > 90) sh.getRange(r0 + j, 3).setFontSize(13); else if (len > 70) sh.getRange(r0 + j, 3).setFontSize(14);
    if (it && String(it.code).length > 16) sh.getRange(r0 + j, 2).setFontSize(14);
  }
  // spacer (pushes the footer to the bottom of the page) + footer
  var sp = r0 + R + 6, ft = sp + 1, mr = sh.getMaxRows();
  if (mr < ft) sh.insertRowsAfter(mr, ft - mr);
  else if (mr > ft) sh.deleteRows(ft + 1, mr - ft);
  sh.setRowHeight(sp, spacer); sh.setRowHeight(ft, F.FOOT_H);
  sh.getRange(sp, 1, 2, F.COLS).breakApart().clearContent().setBorder(false, false, false, false, false, false);
  sh.getRange(ft, 1, 1, F.COLS - 1).merge();
  ordSet_(sh, 'A' + ft, 'Page ' + (pi + 1) + ' of ' + n, { h: 'center', v: 'bottom', size: 12 });
  ordSet_(sh, 'I' + ft, cfg.formCode || '', { h: 'right', v: 'bottom', size: 12 });
  ordSet_(sh, 'I1', pi === 0 ? cfg.formRev || '' : '', { h: 'right', v: 'top', size: 14 });
}
function bomFetchExport_(ssId, fmt) {
  var base = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export?';
  var q = fmt === 'pdf'                                               // all sheets · each sheet fitted to one A4 page
    ? 'format=pdf&size=A4&portrait=true&scale=4&gridlines=false&printtitle=false&sheetnames=false&pagenum=UNDEFINED' +
      '&fzr=false&horizontal_alignment=CENTER&vertical_alignment=TOP&top_margin=0.3&bottom_margin=0.3&left_margin=0.3&right_margin=0.3'
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

// ── public: shortage report (PDF) ───────────────────────────────────
// opt = { project: '' (all active) | 'PP1024D01', late: true → include "ของเข้าไม่ทัน" }
function exportBomShortReport(token, opt) {
  var t0 = Date.now(); opt = opt || {};
  var admin = isAdmin_(token), snap = bomSnapGet_() || bomSnapBuild_(), want = opt.project ? bomKey_(opt.project) : '';
  var withLate = opt.late !== false;
  var projs = (snap.projects || []).filter(function (p) { return p.status === 'active' && (!want || p.project === want); });
  if (want && !projs.length) throw new Error('ไม่พบโปรเจกต์ที่จองอยู่: ' + want);
  var bad = function (x) { return x.st === 'short' || (withLate && x.st === 'late'); };
  var codeSet = {};
  projs.forEach(function (p) { p.items.forEach(function (x) { if (bad(x)) codeSet[ordCodeKey_(x.code)] = 1; }); });
  var codes = Object.keys(codeSet), st = bomStockIdx_(function (k) { return codeSet[k]; });
  var cMap = {}; (snap.codes || []).forEach(function (c) { cMap[c.code] = c; });
  // Order Item already opened for these CODEs (not yet in Pending Order)
  var ords = {};
  try {
    for (var i = 0; i < codes.length && i < 150; i += 30) {
      var h = checkOrderDup(codes.slice(i, i + 30), '');
      Object.keys(h).forEach(function (k) { if ((h[k].orders || []).length) ords[k] = h[k].orders.filter(function (o) { return o.kind !== 'rejected'; }); });
    }
  } catch (e) { console.warn(e); }
  var fmt = function (n) { return n == null || n === '' ? '-' : Number(n).toLocaleString('en-US', { maximumFractionDigits: 2 }); };
  var money = function (n) { return '฿' + Number(Math.round(n || 0)).toLocaleString('en-US'); };
  var C = '#9f1239', A = '#cc785c', L = '#e6dfd8', M = '#6c6a64';
  var th = function (t, al) { return '<th style="background:#f5f0e8;border:1px solid ' + L + ';padding:4px 5px;font-size:8pt;text-align:' + (al || 'left') + ';">' + esc_(t) + '</th>'; };
  var td = function (t, al, s2) { return '<td style="border:1px solid ' + L + ';padding:4px 5px;font-size:8pt;vertical-align:top;text-align:' + (al || 'left') + ';' + (s2 || '') + '">' + t + '</td>'; };
  var kpi = function (v, l, col) { return '<td style="border:1px solid ' + L + ';padding:8px;text-align:center;width:20%;"><div style="font-size:15pt;font-weight:bold;color:' + (col || '#141413') + ';">' + v + '</div><div style="font-size:8pt;color:' + M + ';">' + esc_(l) + '</div></td>'; };
  var pill = function (s2) { return s2 === 'short' ? '<b style="color:' + C + ';">ขาด</b>' : '<b style="color:#92400e;">เข้าไม่ทัน</b>'; };
  var nShort = 0, nLate = 0, val = 0, projRows = [];
  projs.forEach(function (p) {
    var L2 = p.items.filter(bad);
    L2.forEach(function (x) { if (x.st === 'short') nShort++; else nLate++; var s = st[ordCodeKey_(x.code)]; val += (x.short || 0) * (s ? s.cost : 0); });
    projRows.push({ p: p, items: L2 });
  });
  var h = [];
  h.push('<html><head><meta charset="utf-8"><style>body{font-family:Sarabun,Tahoma,Arial,sans-serif;color:#141413;font-size:9pt;}h1{font-size:17pt;margin:0;}h2{font-size:12pt;margin:16px 0 6px;border-bottom:2px solid ' + A + ';padding-bottom:3px;}table{border-collapse:collapse;width:100%;}</style></head><body>');
  h.push('<table><tr><td style="background:' + A + ';color:#fff;padding:10px 12px;"><h1>📋 รายงานรายการที่ขาด — Booking BOM</h1><div style="font-size:9pt;">DAIHEN ELECTRIC CO., LTD. · ' + (want ? 'โปรเจกต์ ' + esc_(want) : 'ทุกโปรเจกต์ที่จองอยู่') + '</div></td></tr></table>');
  h.push('<p style="font-size:8.5pt;color:' + M + ';margin:6px 0;">ออกรายงาน ' + esc_(nowStr_()) + ' · คำนวณการจอง ' + esc_(snap.built || '-') + ' · Stock ' + esc_(snap.stockTime || '-') + ' · Pending Order ' + esc_(snap.poTime || '-') + ' · ตารางงาน ' + esc_(snap.schedTime || '-') + (withLate ? ' · รวม "ของเข้าไม่ทัน"' : ' · เฉพาะ "ขาด"') + '</p>');
  // 1) summary
  h.push('<h2>1. สรุปภาพรวม</h2><table><tr>' + kpi(fmt(projs.length), 'โปรเจกต์') + kpi(fmt(projRows.filter(function (r) { return r.items.length; }).length), 'โปรเจกต์ที่มีปัญหา', '#92400e') +
    kpi(fmt(nShort), 'รายการที่ขาด', C) + kpi(fmt(nLate), 'ของเข้าไม่ทัน', '#92400e') +
    (admin ? kpi(money(val), 'มูลค่าที่ต้องสั่งเพิ่ม (โดยประมาณ)', C) : kpi(fmt(codes.length), 'CODE ที่เกี่ยวข้อง')) + '</tr></table>');
  // 2) per project overview
  h.push('<h2>2. สรุปรายโปรเจกต์ (เรียงตามวันที่ต้องใช้)</h2><table><tr>' + th('โปรเจกต์') + th('วันที่ต้องใช้') + th('เหลือ', 'right') + th('รายการ', 'right') + th('พร้อม', 'right') + th('เข้าไม่ทัน', 'right') + th('ขาด', 'right') + '</tr>');
  projRows.forEach(function (r) { var p = r.p, c = p.count || {};
    h.push('<tr>' + td('<b>' + esc_(p.project) + '</b>' + (p.title ? '<br><span style="color:' + M + ';">' + esc_(p.title) + '</span>' : '')) + td(esc_(p.needDate || '-') + '<br><span style="color:' + M + ';">' + esc_(p.needLabel || '') + '</span>') +
      td(p.daysLeft == null ? '-' : p.daysLeft < 0 ? '<b style="color:' + C + ';">เลย ' + (-p.daysLeft) + ' วัน</b>' : p.daysLeft + ' วัน', 'right') + td(fmt(p.total), 'right') + td(fmt((c.ok || 0) + (c.incoming || 0)), 'right', 'color:#065f46;') +
      td(fmt(c.late || 0), 'right', c.late ? 'color:#92400e;font-weight:bold;' : '') + td(fmt(c.short || 0), 'right', c.short ? 'color:' + C + ';font-weight:bold;' : '') + '</tr>'); });
  h.push('</table>');
  // 3) per project detail
  h.push('<h2>3. รายละเอียดรายการที่ขาด / เข้าไม่ทัน รายโปรเจกต์</h2>');
  projRows.forEach(function (r) {
    var p = r.p;
    h.push('<div style="margin:12px 0 4px;font-weight:bold;font-size:10pt;color:' + A + ';">' + esc_(p.project) + (p.rev ? ' · Rev ' + esc_(p.rev) : '') + ' · ต้องใช้ ' + esc_(p.needDate || '-') + ' · ' + r.items.length + ' รายการ</div>');
    if (!r.items.length) { h.push('<p style="color:#065f46;margin:2px 0 6px;">✓ ไม่มีรายการขาด — พร้อมผลิต</p>'); return; }
    h.push('<table><tr>' + th('#', 'center') + th('CODE') + th('รายการ') + th('ต้องใช้', 'right') + th('จาก Stock', 'right') + th('เข้าทัน', 'right') + th('เข้าไม่ทัน / ETA', 'right') + th('ขาด', 'right') + th('ควรสั่งเพิ่ม', 'right') + th('สถานะ') + th('ใบสั่ง / PO ที่มี') + (admin ? th('มูลค่า', 'right') : '') + '</tr>');
    r.items.forEach(function (x, i) {
      var k = ordCodeKey_(x.code), s = st[k], cc = cMap[k] || {}, ob = (ords[k] || []).map(function (o) { return esc_(o.id + ' ' + (o.state || '') + (o.qty ? ' ×' + o.qty : '')); });
      var po = (cc.po || []).slice(0, 3).map(function (y) { return esc_((y.po || y.pr || '-') + ' ×' + fmt(y.qty) + (y.due ? ' (' + y.due + ')' : '')); });
      var sug = x.short || 0;
      h.push('<tr>' + td(String(i + 1), 'center') + td('<b>' + esc_(x.code) + '</b>' + ((s && s.by) || x.by ? '<br><span style="color:' + M + ';">ผู้สั่ง ' + esc_((s && s.by) || x.by) + '</span>' : '')) + td(esc_(x.material || '-') + ([x.type, x.rated, x.maker].filter(Boolean).length ? '<br><span style="color:' + M + ';">' + esc_([x.type, x.rated, x.maker].filter(Boolean).join(' · ')) + '</span>' : '')) +
        td(fmt(x.qty) + ' ' + esc_(x.unit || ''), 'right') + td(x.fromStock ? fmt(x.fromStock) : '-', 'right') + td(x.onTime ? fmt(x.onTime) : '-', 'right') +
        td(x.late ? fmt(x.late) + (x.eta ? '<br><span style="color:#92400e;">' + esc_(x.eta) + '</span>' : '') : '-', 'right') + td(x.short ? fmt(x.short) : '-', 'right', x.short ? 'color:' + C + ';font-weight:bold;' : '') +
        td(sug ? '<b>' + fmt(sug) + '</b>' : '-', 'right') + td(pill(x.st)) + td((ob.concat(po)).join('<br>') || '<span style="color:' + C + ';">ยังไม่มี</span>') +
        (admin ? td(s && s.cost && x.short ? money(x.short * s.cost) : '-', 'right') : '') + '</tr>');
    });
    h.push('</table>');
  });
  // 4) consolidated by code
  var agg = codes.map(function (k) {
    var c = cMap[k] || {}, s = st[k], need = 0, short = 0, late = 0, ps = [];
    projRows.forEach(function (r) { r.items.forEach(function (x) { if (ordCodeKey_(x.code) === k) { need += x.qty; short += x.short || 0; late += x.late || 0; ps.push(r.p.project + ' ×' + fmt(x.short || x.late)); } }); });
    return { code: k, desc: (s && s.desc) || c.desc || '', by: (s && s.by) || c.by || '', unit: (s && s.unit) || c.unit || '', onhand: c.onhand, incoming: c.incoming, booked: c.booked, balance: c.balance, short: short, late: late, ps: ps, cost: s ? s.cost : 0, ord: (ords[k] || []).length };
  }).sort(function (a, b) { return b.short - a.short || b.late - a.late || a.code.localeCompare(b.code); });
  h.push('<h2>4. สรุปรวมตาม CODE (ใช้วางแผนสั่งซื้อ)</h2><table><tr>' + th('CODE') + th('รายการ') + th('ผู้สั่งล่าสุด') + th('Onhand', 'right') + th('กำลังเข้า', 'right') + th('จองรวม', 'right') + th('คงเหลือหลังจอง', 'right') + th('ขาดรวม', 'right') + th('เข้าไม่ทันรวม', 'right') + th('โปรเจกต์') + th('ใบสั่งเปิดอยู่', 'center') + (admin ? th('มูลค่าที่ขาด', 'right') : '') + '</tr>');
  agg.forEach(function (a) {
    h.push('<tr>' + td('<b>' + esc_(a.code) + '</b>') + td(esc_(a.desc || '-')) + td(esc_(a.by || '-')) + td(fmt(a.onhand), 'right') + td(fmt(a.incoming), 'right') + td(fmt(a.booked), 'right') +
      td(fmt(a.balance), 'right', a.balance < 0 ? 'color:' + C + ';font-weight:bold;' : '') + td(a.short ? fmt(a.short) + ' ' + esc_(a.unit) : '-', 'right', a.short ? 'color:' + C + ';font-weight:bold;' : '') +
      td(a.late ? fmt(a.late) : '-', 'right') + td(esc_(a.ps.join(', '))) + td(a.ord ? '✓ ' + a.ord : '-', 'center') + (admin ? td(a.cost && a.short ? money(a.short * a.cost) : '-', 'right') : '') + '</tr>');
  });
  if (!agg.length) h.push('<tr>' + td('✓ ไม่มีรายการขาด', 'center', 'color:#065f46;') + '</tr>');
  h.push('</table>');
  h.push('<p style="font-size:8pt;color:' + M + ';margin-top:14px;">หมายเหตุ: จัดสรรต่อ CODE ตามวันที่ต้องใช้ของโปรเจกต์ (ใกล้ก่อน) — ใช้ Onhand ทุกคลังก่อน แล้วของ Pending Order ที่เข้าทันวัน แล้วของที่เข้าไม่ทัน · "ควรสั่งเพิ่ม" = จำนวนที่ขาดหลังนับของที่กำลังเข้าทั้งหมดแล้ว · "ใบสั่งเปิดอยู่" = Order Item ที่บันทึก/รออนุมัติ/อนุมัติแล้วแต่ยังไม่เข้าระบบ' + (admin ? ' · มูลค่า = จำนวนที่ขาด × LASTCOST (ใช้ภายในเท่านั้น)' : '') + '</p></body></html>');
  var name = 'BOM_Shortage_' + (want || 'All') + '_' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd_HHmm') + '.pdf';
  audit_([[admin ? 'Admin' : 'User', 'ออกรายงาน PDF รายการที่ขาด (BOM)', want, nShort + ' ขาด · ' + nLate + ' เข้าไม่ทัน']]);
  if (opt.html) { perf_('exportBomShortReport (html)', t0); return { name: name, html: h.join(''), short: nShort, late: nLate }; }   // browser print preview (zqPrintReport)
  var pdf = Utilities.newBlob(h.join(''), 'text/html', 'r.html').getAs('application/pdf').setName(name);
  perf_('exportBomShortReport', t0);
  return { name: name, mime: 'application/pdf', b64: Utilities.base64Encode(pdf.getBytes()), short: nShort, late: nLate };
}
