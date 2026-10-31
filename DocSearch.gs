/**
 * ═════════════════════════════════════════════════════════════════════
 *  DOCUMENT SEARCH — Specification sheets (PDF) on the file server (\\192.168.0.26\electrical-ds\...)
 *  • The file list (ID, AddressLink, FileName) is synced from a OneDrive "Anyone with the link" file
 *    (.csv or .xlsx) → sheet "FileList" in Helpdesk_ItemStock, the same way as Pending Order.
 *  • getDocIndex()   the whole list, parsed once per sync: RQN / DS1 numbers, project, Rev, Cancelled
 *                    → the browser keeps it (IndexedDB) and searches it instantly.
 *  • docLookup(q)    PO / CODE / item name → RQN + DS1 numbers from Pending Order and Order Item,
 *                    so a document can be found by PO or CODE although the file name has neither.
 *  A web page cannot open \\server paths itself: the page copies the file / folder path
 *  (Win+R → paste → Enter, or the File Explorer address bar).
 *  ต้องมี Code.gs (shareCandidates_, fetchWithCookies_, csvLine_, stockSS_), Workspace.gs, OrderService.gs
 * ═════════════════════════════════════════════════════════════════════
 */
var DOC_SHEET = 'FileList', DOC_HEADER = ['ID', 'AddressLink', 'FileName'];

function docInfo_() { try { return JSON.parse(schedProp_('DOC_SYNC_LAST') || 'null'); } catch (e) { return null; } }

// the shared file → rows of strings. CSV (UTF-8 or Thai windows-874) or .xlsx (converted through Drive once)
function docDownload_(u) {
  var tries = shareCandidates_(u), notes = [];
  for (var i = 0; i < tries.length; i++) {
    try {
      var r = fetchWithCookies_(tries[i]);
      if (!r.res || r.code !== 200) { notes.push('#' + (i + 1) + ' HTTP ' + r.trail); continue; }
      var blob = r.res.getBlob(), b = blob.getBytes();
      if (b.length > 1 && b[0] === 0x50 && b[1] === 0x4B) return { rows: docXlsx_(blob), raw: Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, b)) };
      var text = blob.getDataAsString('UTF-8');
      if (text.indexOf('�') >= 0) text = blob.getDataAsString('windows-874');
      text = text.replace(/^﻿/, '');
      if (/^\s*</.test(text)) { notes.push('#' + (i + 1) + ' ได้หน้าเว็บ (' + r.trail + ')'); continue; }
      return { rows: docCsv_(text), raw: Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, text)) };
    } catch (e) { notes.push('#' + (i + 1) + ' ' + e.message); }
  }
  throw new Error('ดาวน์โหลดไฟล์ไม่ได้ — ตรวจว่าแชร์แบบ "Anyone with the link" และลิงก์ยังไม่หมดอายุ [' + notes.join(' | ') + ']');
}
function docCsv_(text) {
  var lines = text.split(/\r?\n/), out = [];
  for (var i = 0, buf = ''; i < lines.length; i++) {
    buf = buf ? buf + ' ' + lines[i] : lines[i];
    if (!buf) continue;
    if ((buf.match(/"/g) || []).length % 2) continue;          // quoted field spans a line break
    out.push(csvLine_(buf)); buf = '';
  }
  return out;
}
function docXlsx_(blob) {
  var f = Drive.Files.create({ name: '_tmp_FileList_' + Date.now(), mimeType: 'application/vnd.google-apps.spreadsheet' },
    blob.setContentType('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'));
  try { return SpreadsheetApp.openById(f.id).getSheets()[0].getDataRange().getDisplayValues(); }
  finally { try { DriveApp.getFileById(f.id).setTrashed(true); } catch (e) {} }
}

function syncDocsRun_(why, force) { return syncMutex_('DOC_SYNC_LAST', why, function () { return syncDocsInner_(why, force); }); }
function syncDocsInner_(why, force) {
  var url = schedProp_('DOC_SRC_URL'), info = { time: nowStr_(), why: why }, t0 = Date.now();
  if (!url) throw new Error('ยังไม่ได้ตั้งค่าลิงก์ OneDrive ของรายการเอกสาร (FileList)');
  try {
    var got = docDownload_(url), prev = docInfo_();
    if (!force && prev && prev.ok && prev.hash === got.raw) {
      info.ok = true; info.unchanged = true; info.rows = prev.rows; info.hash = got.raw; info.msg = 'ไฟล์ไม่มีการเปลี่ยนแปลง';
      schedProp_('DOC_SYNC_LAST', JSON.stringify(info)); return info;
    }
    var all = got.rows, raw = got.raw; got = null;
    var head = (all[0] || []).map(function (h) { return String(h || '').toLowerCase().replace(/[^a-z]/g, ''); });
    var col = function (names) { for (var k = 0; k < names.length; k++) { var j = head.indexOf(names[k]); if (j >= 0) return j; } return -1; };
    var cI = col(['id']), cP = col(['addresslink', 'address', 'path', 'fullpath', 'link']), cN = col(['filename', 'name', 'file']);
    if (cP < 0) throw new Error('หัวคอลัมน์ไม่ตรงรูปแบบ ต้องมี AddressLink (ที่อยู่ไฟล์) · รูปแบบ: ID, AddressLink, FileName');
    var rows = [], seen = {};
    for (var i = 1; i < all.length; i++) {
      var f = all[i], p = String(f[cP] == null ? '' : f[cP]).replace(/[\t\r\n]+/g, ' ').trim();
      if (!p || seen[p.toLowerCase()]) continue;
      seen[p.toLowerCase()] = 1;
      var nm = cN >= 0 ? String(f[cN] || '').trim() : ''; if (!nm) nm = p.split(/[\\\/]/).pop();
      rows.push([cI >= 0 ? String(f[cI] || '').trim() : String(i), p, nm]);
    }
    all = null;
    if (!rows.length) throw new Error('ไฟล์ไม่มีข้อมูล');
    var ss = stockSS_(), sh = ss.getSheetByName(DOC_SHEET) || ss.insertSheet(DOC_SHEET), out = [DOC_HEADER].concat(rows);
    sh.clearContents();
    if (sh.getMaxRows() < out.length) sh.insertRowsAfter(sh.getMaxRows(), out.length - sh.getMaxRows());
    if (sh.getMaxColumns() < DOC_HEADER.length) sh.insertColumnsAfter(sh.getMaxColumns(), DOC_HEADER.length - sh.getMaxColumns());
    var rg = sh.getRange(1, 1, out.length, DOC_HEADER.length); rg.setNumberFormat('@'); rg.setValues(out);
    sh.getRange(1, 1, 1, DOC_HEADER.length).setFontWeight('bold'); sh.setFrozenRows(1);
    var extra = sh.getMaxRows() - out.length; if (extra > 200) sh.deleteRows(out.length + 101, extra - 100);
    schedProp_('DOC_HASH', raw);
    var idx = docIndex_(true);
    info.ok = true; info.rows = rows.length; info.hash = raw; info.secs = Math.round((Date.now() - t0) / 1000);
    info.msg = 'อัปเดต ' + rows.length.toLocaleString() + ' ไฟล์ (มี RQN ' + idx.nR + ' · มี DS1 ' + idx.nD + ') ใช้เวลา ' + info.secs + ' วินาที';
    audit_([['System', 'ซิงค์รายการเอกสาร (FileList) จาก OneDrive', '', info.why + ' · ' + info.msg]]);
  } catch (e) {
    info.ok = false; info.msg = e.message || String(e);
    if (info.why !== 'อัตโนมัติ') audit_([['System', 'ซิงค์รายการเอกสารไม่สำเร็จ', '', info.msg]]);
  }
  schedProp_('DOC_SYNC_LAST', JSON.stringify(info));
  return info;
}

// ── file name → numbers ─────────────────────────────────────────────
// RQN26-03780 · RQN20--03339 · RQN263-03798 (extra digit) · RQN2101287 (no dash) · RQN-03006 (no year)
// DS1C26-1161 · DS1C23-0798,0799 · DS1C20-0756 0757 0758 · DS1C20-0728-0729 · DS1C201100 · DS121-0544 · DS1SC21-1239
// key = year-number without leading zeros ("26-3780"); DS keys carry the unit letter ("C26-1161")
function docParse_(name) {
  var s = String(name || '').replace(/\.pdf$/i, ''), r = [], d = [], m;
  var reR = /RQN\s*(\d{2})?(?:\d(?=\s*-))?\s*-*\s*(\d{3,6})/gi;
  while ((m = reR.exec(s))) { var kr = (m[1] || '') + '-' + parseInt(m[2], 10); if (r.indexOf(kr) < 0) r.push(kr); }
  var reD = /DS\s*1?\s*S?([A-Z])?\s*(\d{2})(?:\d(?=\s*-))?\s*-?\s*(\d{3,5})/gi;
  while ((m = reD.exec(s))) {
    var u = (m[1] || 'C').toUpperCase(), yy = m[2], add = function (n) { var k = u + yy + '-' + parseInt(n, 10); if (d.indexOf(k) < 0) d.push(k); };
    add(m[3]);
    var rest = s.substring(reD.lastIndex), c;                       // 0757 0758 · ,0799 · -0729
    while ((c = /^(?:\s*[,&\-]\s*(\d{3,4})|\s+(\d{4}))(?=[\s,&\-]|$)/.exec(rest))) { add(c[1] || c[2]); rest = rest.substring(c[0].length); reD.lastIndex += c[0].length; }
  }
  var pj = /^\s*(PP\d{3,5}[A-Z]?\d{0,2})/i.exec(s), rev = /\bRev\.?\s*([A-Z0-9]{1,3})\b/i.exec(s);
  return { r: r, d: d, p: pj ? pj[1].toUpperCase() : '', rev: rev ? rev[1].toUpperCase() : '', x: /cancel/i.test(s) ? 1 : 0 };
}
// whole list for the browser: common root + folders once + one short row per file
function docIndex_(rebuild) {
  var hash = String(schedProp_('DOC_HASH') || ''), key = 'docidx_' + hash.replace(/[^\w]/g, '').substring(0, 16);
  if (!hash) return { hash: '', root: '', dirs: [], rows: [], n: 0, nR: 0, nD: 0 };
  if (!rebuild) { var hit = ordCacheGet_(key); if (hit) return hit; }
  var sh = stockSS_().getSheetByName(DOC_SHEET), n = sh ? sh.getLastRow() - 1 : 0;
  var src = n > 0 ? sh.getRange(2, 1, n, DOC_HEADER.length).getDisplayValues() : [];
  var dirs = [], dix = {}, rows = [], nR = 0, nD = 0, root = null;
  src.forEach(function (x) {
    var p = String(x[1] || ''), cut = Math.max(p.lastIndexOf('\\'), p.lastIndexOf('/')), dir = cut >= 0 ? p.substring(0, cut + 1) : '';
    if (root === null) root = dir;
    else while (root && dir.indexOf(root) !== 0) root = root.substring(0, root.replace(/[\\\/]$/, '').search(/[\\\/][^\\\/]*$/) + 1);
  });
  root = root || '';
  src.forEach(function (x) {
    var p = String(x[1] || ''), cut = Math.max(p.lastIndexOf('\\'), p.lastIndexOf('/')), dir = (cut >= 0 ? p.substring(0, cut + 1) : '').substring(root.length);
    if (!(dir in dix)) { dix[dir] = dirs.length; dirs.push(dir); }
    var nm = x[2] || p.substring(cut + 1), f = docParse_(nm);
    if (f.r.length) nR++; if (f.d.length) nD++;
    if (!f.p) { var m = /(?:^|[\\\/])(PP\d{3,5}[A-Z]?)[\\\/]/i.exec(dir); if (m) f.p = m[1].toUpperCase(); }
    rows.push([x[0], dix[dir], nm, f.r.join(' '), f.d.join(' '), f.p, f.rev, f.x, p.substring(cut + 1) === nm ? '' : p.substring(cut + 1)]);
  });
  var out = { hash: hash.substring(0, 16), root: root, dirs: dirs, rows: rows, n: rows.length, nR: nR, nD: nD, time: (docInfo_() || {}).time || '' };
  ordCachePut_(key, out, 21600);
  return out;
}
// public: the list (anyone can search; it holds file names and paths only)
function getDocIndex() { return docIndex_(false); }

// public: PO / CODE / name → the RQN and DS1 numbers of those purchases (Pending Order + Order Item)
function docLookup(q) {
  var toks = String(q || '').toUpperCase().split(/[\s,;]+/).map(function (t) { return t.replace(/[^A-Z0-9\-\/.]/g, ''); })
    .filter(function (t) { return t.length >= 4; }).slice(0, 6);
  var out = { hits: [], r: [], d: [] };
  if (!toks.length) return out;
  var key = function (v) { return String(v || '').toUpperCase().replace(/[^A-Z0-9]/g, ''); }, tk = toks.map(key).filter(String);
  var addR = function (v) { var p = docParse_(String(v || '')); p.r.forEach(function (k) { if (out.r.indexOf(k) < 0) out.r.push(k); }); return p.r.length; };
  var addD = function (v) { var p = docParse_(String(v || '')); p.d.forEach(function (k) { if (out.d.indexOf(k) < 0) out.d.push(k); }); return p.d.length; };
  var seen = {};
  // Pending Order: [Item_Id, Requisition_No, PO_NO, Specification_No, Status, Item_Name, ...]
  wsPendingRows_(function (code, f) {
    var po = key(f[2]), pr = key(f[1]);
    for (var i = 0; i < tk.length; i++) {
      var t = tk[i], via = '';
      if (po && (po === t || (t.length >= 5 && po.indexOf(t) >= 0))) via = 'PO';
      else if (code && code === t) via = 'CODE';
      else if (pr && t.length >= 7 && pr.indexOf(t) >= 0 && /^RQN/.test(t)) via = 'PR';
      if (!via) continue;
      var sk = via + '|' + f[1] + '|' + f[3] + '|' + code;
      if (!seen[sk] && (f[1] || f[3])) { seen[sk] = 1; out.hits.push({ via: via, po: f[2], pr: f[1], spec: f[3], code: f[0], name: f[5], status: f[4] }); }
      break;
    }
    return false;
  });
  out.hits.forEach(function (h) { addR(h.pr); addD(h.spec); });
  // Order Item: the form's RQN / Spec, found by an item CODE
  try {
    ordPubAll_().forEach(function (o) {
      var codes = (o.items || []).map(function (it) { return key(it.code); });
      if (!tk.some(function (t) { return codes.indexOf(t) >= 0; })) return;
      var hr = addR(o.rqn), hd = addD(o.spec);
      if (hr || hd) out.hits.push({ via: 'ORDER', order: o.id, pr: o.rqn, spec: o.spec, code: (o.items || []).map(function (it) { return it.code; }).filter(function (c) { return tk.indexOf(key(c)) >= 0; }).join(', ') });
    });
  } catch (e) {}
  out.hits = out.hits.slice(0, 40);
  return out;
}

// ── settings (Admin) ────────────────────────────────────────────────
function docAutoSync() { if (schedProp_('DOC_SRC_URL')) syncDocsRun_('อัตโนมัติ', false); }   // time trigger
function getDocSource(token) {
  requireAdmin_(token);
  var hrs = 0; ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'docAutoSync') hrs = Number(schedProp_('DOC_SYNC_HOURS') || 24); });
  var id = schedProp_('STOCK_SS_ID');
  return { url: schedProp_('DOC_SRC_URL') || '', hours: hrs, last: docInfo_(), sheetUrl: id ? 'https://docs.google.com/spreadsheets/d/' + id + '/edit' : '' };
}
function setDocSource(token, url, hours) {
  requireAdmin_(token);
  url = String(url || '').trim();
  if (url && !/^https:\/\/[^\s]+$/i.test(url)) throw new Error('ลิงก์ต้องขึ้นต้นด้วย https://');
  if (/_layouts\/15\/onedrive\.aspx|\/Forms\/AllItems\.aspx/i.test(url)) throw new Error('นี่คือลิงก์หน้าโฟลเดอร์ OneDrive ไม่ใช่ลิงก์แชร์ไฟล์ — คลิกขวาที่ไฟล์ → Share → Copy link แล้ววางลิงก์นั้นแทน');
  schedProp_('DOC_SRC_URL', url || null);
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'docAutoSync') ScriptApp.deleteTrigger(t); });
  hours = Number(hours) || 0;
  if (url && hours > 0) {
    syncTrigger_(ScriptApp.newTrigger('docAutoSync').timeBased(), hours, 5, 24);
    schedProp_('DOC_SYNC_HOURS', String(hours));
  }
  audit_([['Admin', 'ตั้งค่าซิงค์รายการเอกสาร', '', (url ? 'ตั้งลิงก์แล้ว' : 'ลบลิงก์') + ' · ' + syncLabel_(hours)]]);
  return getDocSource(token);
}
function syncDocsNow(token) { requireAdmin_(token); return syncDocsRun_('สั่งโดย Admin', true); }

// Run ONCE from the editor after saving the link, to authorize and check the result in the log.
function testDocSync() {
  if (!schedProp_('DOC_SRC_URL')) { Logger.log('✗ ยังไม่ได้บันทึกลิงก์ใน ตั้งค่าระบบ → ☁ ซิงค์รายการเอกสาร'); return; }
  var r = syncDocsRun_('ทดสอบจาก Editor', true);
  Logger.log((r.ok ? '✓ ' : '✗ ') + r.msg);
}
