/**
 * ESTIMATE COST (Admin only) — ต้นทุนจาก Shipment รายแผนก × โปรเจกต์
 *
 * Source: Q_Shipment exported from Access as .csv to OneDrive (~3 years, can be > 50 MB).
 * UrlFetchApp cannot take more than 50 MB in one response, so the file is read in byte ranges
 * (Range: bytes=…, 20 MB each) and parsed as it streams — nothing is converted by Drive and the
 * full file is never held in memory. Only rows whose CATEGORY is a project (1030P1 / P99901) are kept.
 *
 * Storage: spreadsheet Helpdesk_Shipment (its own file, so it never eats the cell limit of the others)
 *   sheet Ship      one row per project per chunk read: CATEGORY | lines | gzip+base64 of the lines (≤ 45k chars)
 *   sheet ShipMeta  A1… = gzip+base64 of { ver, idx: {cat: [rows]}, sum: [[dept, cat, n, total, last]] }
 * A project is read with one range per stored row (a few), never the whole sheet.
 * Line = CODE \t DESCRIPTION \t yyyymmdd \t DOCNUM \t CUSTNAME \t LOCATION \t QTY \t UNIT \t UNITCOST \t COST \t AUDTUSER
 */
var SHIP_SHEET = 'Ship', SHIP_META = 'ShipMeta', SHIP_CHUNK = 20 * 1024 * 1024, SHIP_CELL = 45000, SHIP_MAX_MS = 320000;

function shipInfo_() { try { return JSON.parse(schedProp_('SHIP_SYNC_LAST') || 'null'); } catch (e) { return null; } }
function shipSS_(create) {
  var id = schedProp_('SHIP_SS_ID');
  if (id) {
    var err;
    for (var n = 0; n < 3; n++) { try { return SpreadsheetApp.openById(id); } catch (e) { err = e; Utilities.sleep(1500 * (n + 1)); } }
    throw new Error('เปิดไฟล์ Helpdesk_Shipment ไม่ได้ (' + (err && err.message || err) + ') — ระบบจะลองใหม่รอบถัดไป ไม่ได้สร้างไฟล์ใหม่');
  }
  if (!create) return null;
  var ss = SpreadsheetApp.create('Helpdesk_Shipment');
  try { DriveApp.getFileById(ss.getId()).moveTo(DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID)); } catch (e) {}
  schedProp_('SHIP_SS_ID', ss.getId());
  return ss;
}

// ── project number ⇄ CATEGORY (the rule of the Access form, extended) ─────────
// PP103001 / PP1030D01 / PP1030-01 → 1030P1  (4 digits ≥ 1000: digits & "P" & Val(last 2))
// PP099901 / PP0990D01 → P99901 / P99001      (old style: "P" & Mid(x, 4, 5), the unit letter skipped)
// 1030P1 / 1030P01 / P99901 typed as they are → the same CATEGORY
function shipCat_(s) {
  s = String(s || '').toUpperCase().replace(/[\s\-_\/]/g, '');
  var m = /^(\d{4})P(\d{1,2})$/.exec(s); if (m) return m[1] + 'P' + Number(m[2]);
  if (/^P\d{5}$/.test(s)) return s;
  m = /^PP(\d{4})[A-Z]*(\d{2})$/.exec(s);
  if (m) return Number(m[1]) >= 1000 ? m[1] + 'P' + Number(m[2]) : 'P' + m[1].substr(1) + m[2];   // PP0990D01 → P99001
  return '';
}
function shipProj_(cat) {           // CATEGORY → the project number printed on the report
  var m = /^(\d{4})P(\d{1,2})$/.exec(cat); if (m) return 'PP' + m[1] + ('0' + m[2]).slice(-2);
  m = /^P(\d{5})$/.exec(cat); return m ? 'PP0' + m[1] : cat;
}
// DateShip "dd/mm/yyyy" (Format() in the query) → yyyymmdd · also yyyy-mm-dd and Buddhist years
function shipDate_(v) {
  var s = String(v || '').trim(), m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(s), y;
  if (m) { y = +m[3]; if (y > 2400) y -= 543; return y + ('0' + m[2]).slice(-2) + ('0' + m[1]).slice(-2); }
  m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s);
  if (m) { y = +m[1]; if (y > 2400) y -= 543; return y + ('0' + m[2]).slice(-2) + ('0' + m[3]).slice(-2); }
  return /^\d{8}$/.test(s) ? s : '';
}
function shipNum_(v) { var s = String(v == null ? '' : v).replace(/[^0-9.\-]/g, ''); var n = Number(s); return s && isFinite(n) ? n : 0; }
function shipB64_(s) { return Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(s, 'text/plain')).getBytes()); }
function shipUnB64_(b) { return Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(b), 'application/x-gzip')).getDataAsString(); }

// ── download in byte ranges ──────────────────────────────────────────────
// one request, following redirects by hand with a cookie jar (see fetchWithCookies_) and keeping the Range header
function shipGet_(url, from, to, jar) {
  for (var hop = 0; hop < 10; hop++) {
    var ck = Object.keys(jar).map(function (k) { return k + '=' + jar[k]; }).join('; ');
    var hd = { 'User-Agent': 'Mozilla/5.0 (HelpdeskSync)', Range: 'bytes=' + from + '-' + to }; if (ck) hd.Cookie = ck;
    var res = UrlFetchApp.fetch(url, { followRedirects: false, muteHttpExceptions: true, headers: hd });
    var code = res.getResponseCode(), h = res.getAllHeaders ? res.getAllHeaders() : res.getHeaders();
    var sc = h['Set-Cookie'] || h['set-cookie'];
    if (sc) [].concat(sc).forEach(function (c) { var p = String(c).split(';')[0], i = p.indexOf('='); if (i > 0) jar[p.substring(0, i).trim()] = p.substring(i + 1).trim(); });
    if (code >= 300 && code < 400) {
      var loc = h.Location || h.location; if (!loc) break;
      loc = [].concat(loc)[0]; if (/^\//.test(loc)) loc = url.match(/^https:\/\/[^\/]+/)[0] + loc;
      url = loc; continue;
    }
    var cr = String(h['Content-Range'] || h['content-range'] || ''), tm = /\/(\d+)\s*$/.exec(cr);
    return { res: res, code: code, url: url, total: tm ? Number(tm[1]) : null, etag: String(h.ETag || h.etag || '') };
  }
  return { res: null, code: 0 };
}
// the first range of the shared file: tries the direct-download forms of the link (shareCandidates_)
function shipFirst_(u) {
  var tries = shareCandidates_(u), notes = [];
  for (var i = 0; i < tries.length; i++) {
    try {
      var jar = {}, r = shipGet_(tries[i], 0, SHIP_CHUNK - 1, jar);
      if (!r.res || (r.code !== 200 && r.code !== 206)) { notes.push('#' + (i + 1) + ' HTTP ' + r.code); continue; }
      var head = r.res.getBlob().getDataAsString('ISO-8859-1').substring(0, 400);
      if (/^PK/.test(head)) throw new Error('ไฟล์นี้เป็น .xlsx — ข้อมูล Shipment ต้อง export เป็น .csv (ไฟล์ Excel ขนาดนี้ Google แปลงไม่ได้)');
      if (/^\s*</.test(head)) { notes.push('#' + (i + 1) + ' ได้หน้าเว็บ'); continue; }
      r.jar = jar; return r;
    } catch (e) { if (/\.xlsx/.test(e.message)) throw e; notes.push('#' + (i + 1) + ' ' + e.message); }
  }
  throw new Error('ดาวน์โหลดไฟล์ไม่ได้ — ตรวจว่าแชร์แบบ "Anyone with the link" และลิงก์ยังไม่หมดอายุ [' + notes.join(' | ') + ']');
}
// bytes taken by a piece of decoded text (to know where the next range starts)
function shipBytes_(s, enc) {
  if (enc !== 'UTF-8') return s.length;                    // windows-874: one byte per character
  var n = 0;
  for (var i = 0; i < s.length; i++) {
    var c = s.charCodeAt(i);
    if (c < 0x80) n++; else if (c < 0x800) n += 2; else if (c >= 0xD800 && c <= 0xDBFF) { n += 4; i++; } else n += 3;
  }
  return n;
}

// ── sync ─────────────────────────────────────────────────────────────────
function syncShipRun_(why, force) { return syncMutex_('SHIP_SYNC_LAST', why, function () { return syncShipInner_(why, force); }); }
function syncShipInner_(why, force) {
  var url = schedProp_('SHIP_SRC_URL'), info = { time: nowStr_(), why: why }, t0 = Date.now();
  if (!url) throw new Error('ยังไม่ได้ตั้งค่าลิงก์ OneDrive ของข้อมูล Shipment');
  try {
    var first = shipFirst_(url), prev = shipInfo_();
    var got = first, off = 0, enc = null, cols = null, hash = null, chunk = 0, kept = 0, skipped = 0, bytes = 0, tmp = null, ss = null;
    var idx = {}, sum = {}, out = [], depts = {}, oldest = '', newest = '', SA = { agg: {}, desc: {} };   // SA: monthly totals for สถิติและต้นทุน
    var flush = function (groups) {
      Object.keys(groups).forEach(function (cat) {
        var ls = groups[cat], parts = [ls];
        while (parts.length) {                              // a cell holds ≤ 50k characters: halve until it fits
          var p = parts.shift(), b = shipB64_(p.join('\n'));
          if (b.length > SHIP_CELL && p.length > 1) { var h = Math.ceil(p.length / 2); parts.unshift(p.slice(0, h), p.slice(h)); continue; }
          out.push([cat, p.length, b]);
        }
      });
      if (!out.length) return;
      if (!tmp) {
        ss = shipSS_(true);
        tmp = ss.getSheetByName(SHIP_SHEET + '_new') || ss.insertSheet(SHIP_SHEET + '_new');
        tmp.clearContents();
        tmp.getRange(1, 1, 1, 3).setValues([['CATEGORY', 'LINES', 'DATA']]);
      }
      var r0 = tmp.getLastRow() + 1;
      ensureRows_(tmp, r0 + out.length);
      var rg = tmp.getRange(r0, 1, out.length, 3); fmtText_(rg); rg.setValues(out);
      out.forEach(function (o, i) { (idx[o[0]] = idx[o[0]] || []).push(r0 + i); });
      out = [];
    };
    while (got) {
      var total = got.total, blob = got.res.getBlob();
      if (got.code === 200) total = null;                  // the server sent the whole file at once
      var t = enc === 'windows-874' ? blob.getDataAsString('windows-874') : blob.getDataAsString('UTF-8');
      if (!enc) {
        enc = t.substring(0, Math.max(0, t.length - 8)).indexOf('�') >= 0 ? 'windows-874' : 'UTF-8';
        if (enc !== 'UTF-8') t = blob.getDataAsString('windows-874');
        hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, t.substring(0, 2000000))) + '|' + (total || t.length) + '|' + got.etag;
        if (!force && prev && prev.ok && prev.hash === hash) {
          info.ok = true; info.unchanged = true; info.rows = prev.rows; info.projects = prev.projects; info.hash = hash;
          info.msg = 'ไฟล์ไม่มีการเปลี่ยนแปลง';
          schedProp_('SHIP_SYNC_LAST', JSON.stringify(info)); return info;
        }
      }
      var last = !total || off + SHIP_CHUNK >= total;
      // records: a quoted field may hold a line break, so a record ends at a newline with an even quote count
      var pos = 0, start = 0, q = 0, groups = {};
      var take = function (rec) {
        rec = rec.replace(/\r$/, '');
        if (!rec) return;
        var f = rec.indexOf('"') < 0 ? rec.split(',') : csvLine_(rec);
        if (!cols) {
          var hd = f.map(function (x) { return String(x || '').replace(/^﻿/, '').toUpperCase().replace(/[^A-Z]/g, ''); });
          var at = function (names) { for (var k = 0; k < names.length; k++) { var j = hd.indexOf(names[k]); if (j >= 0) return j; } return -1; };
          cols = { code: at(['FMTITEMNO', 'ITEMNO', 'CODE']), desc: at(['DESCRIPTION', 'ITEMDESC']), date: at(['DATESHIP', 'TRANSDATE', 'DATE']),
            doc: at(['DOCNUM']), dept: at(['CUSTNAME']), cat: at(['CATEGORY']), loc: at(['LOCATION']), qty: at(['SHIPQTY', 'QTY']),
            unit: at(['SHIPUNIT', 'UNIT']), uc: at(['UNITCOST']), cost: at(['SHIPCOST', 'COST']), user: at(['AUDTUSER']) };
          if (cols.code < 0 || cols.cat < 0 || cols.dept < 0 || cols.cost < 0)
            throw new Error('หัวคอลัมน์ไม่ตรงรูปแบบ Q_Shipment — ต้องมีอย่างน้อย FMTITEMNO, CUSTNAME, CATEGORY, SHIPCOST');
          return;
        }
        var g = function (k) { var j = cols[k]; return j < 0 || f[j] == null ? '' : String(f[j]).replace(/[\t\r\n]+/g, ' ').trim(); };
        var rc = g('cat').toUpperCase().replace(/\s/g, ''), cat = /^(\d{4}P\d{1,2}|P\d{5})$/.test(rc) ? shipCat_(rc) : '';
        if (!cat) { skipped++; return; }                    // stock / T / PH … categories are not projects
        var dept = g('dept').toUpperCase() || '-', d8 = shipDate_(g('date')), cost = shipNum_(g('cost'));
        (groups[cat] = groups[cat] || []).push([g('code'), g('desc'), d8, g('doc'), dept, g('loc'), shipNum_(g('qty')), g('unit'), shipNum_(g('uc')), cost, g('user')].join('\t'));
        var k = dept + '\t' + cat, s = sum[k] || (sum[k] = [dept, cat, 0, 0, '', 0]);
        s[2]++; s[3] += cost; if (d8 > s[4]) s[4] = d8; if (cost > 0) s[5] += cost;   // [5] = returns (positive cost)
        statAggAdd_(SA, g('code'), g('desc'), g('unit'), dept, d8, shipNum_(g('qty')), cost);
        depts[dept] = 1; kept++;
        if (d8 && (!oldest || d8 < oldest)) oldest = d8; if (d8 > newest) newest = d8;
      };
      while (true) {
        var nl = t.indexOf('\n', pos); if (nl < 0) break;
        var seg = t.substring(pos, nl); if (seg.indexOf('"') >= 0) q ^= (seg.split('"').length - 1) & 1;
        pos = nl + 1;
        if (q) continue;
        take(t.substring(start, nl)); start = pos;
      }
      if (last) { if (start < t.length && !q) take(t.substring(start)); }
      else if (!start) throw new Error('อ่านไฟล์ไม่ได้ — ไม่พบการขึ้นบรรทัดใหม่ในช่วง ' + Math.round(SHIP_CHUNK / 1048576) + ' MB');
      if (last) bytes = total || shipBytes_(t, enc);
      var used = last ? 0 : shipBytes_(t.substring(0, start), enc);
      t = null; blob = null;
      flush(groups); groups = null; chunk++;
      if (last) break;
      if (Date.now() - t0 > SHIP_MAX_MS) throw new Error('ไฟล์ใหญ่เกินกว่าจะอ่านจบใน 1 รอบ (อ่านได้ ' + Math.round((off + used) / 1048576) + ' MB) — ลดช่วงปีหรือคอลัมน์ในไฟล์ CSV');
      off += used;
      got = shipGet_(got.url, off, off + SHIP_CHUNK - 1, first.jar);
      if (!got.res || got.code !== 206) throw new Error('ดาวน์โหลดช่วงถัดไปไม่ได้ (HTTP ' + got.code + ') — เซิร์ฟเวอร์ไม่รองรับการอ่านเป็นช่วง');
    }
    if (!cols) throw new Error('ไฟล์ไม่มีข้อมูล');
    if (!kept) throw new Error('ไม่พบแถวที่ CATEGORY เป็นเลขโปรเจกต์ (เช่น 1030P1) — ตรวจไฟล์ที่ export');
    // meta + swap the new sheet in
    var meta = { ver: String(Date.now()), idx: idx, sum: Object.keys(sum).map(function (k) { var s = sum[k]; s[3] = Math.round(s[3] * 100) / 100; return s; }),
      oldest: oldest, newest: newest, time: info.time };
    var b = shipB64_(JSON.stringify(meta)), cells = [];
    for (var i = 0; i < b.length; i += SHIP_CELL) cells.push([b.substring(i, i + SHIP_CELL)]);
    withLock_(function () {
      var old = ss.getSheetByName(SHIP_SHEET); if (old) ss.deleteSheet(old);
      tmp.setName(SHIP_SHEET);
      var ms = ss.getSheetByName(SHIP_META) || ss.insertSheet(SHIP_META);
      ms.clearContents(); ensureRows_(ms, cells.length + 1);
      var rg = ms.getRange(1, 1, cells.length, 1); fmtText_(rg); rg.setValues(cells);
      try { var s1 = ss.getSheetByName('Sheet1') || ss.getSheetByName('ชีต1'); if (s1 && ss.getSheets().length > 2) ss.deleteSheet(s1); } catch (e) {}
      schedProp_('SHIP_VER', meta.ver);
    });
    shipMemo_ = null;
    info.ok = true; info.rows = kept; info.projects = Object.keys(idx).length; info.depts = Object.keys(depts).length; info.skipped = skipped;
    info.mb = Math.round(bytes / 104857.6) / 10; info.hash = hash; info.sec = Math.round((Date.now() - t0) / 1000);
    info.range = oldest && newest ? shipShow_(oldest) + ' – ' + shipShow_(newest) : '';
    info.msg = 'อัปเดต ' + kept.toLocaleString() + ' บรรทัด · ' + info.projects + ' โปรเจกต์ · ' + info.depts + ' แผนก (' + info.mb + ' MB, ' + info.sec + ' วินาที)';
    audit_([['System', 'ซิงค์ข้อมูล Shipment จาก OneDrive', '', info.why + ' · ' + info.msg]]);
    try { statAggSave_(SA); } catch (e) { console.warn('statAggSave_', e); }
    SA = null;
    if (Date.now() - t0 < 240000) { try { bomRefresh_(); info.msg += ' · คำนวณ Booking BOM ใหม่แล้ว'; } catch (e) { console.warn('ship → bomRefresh_', e); } }   // Booking BOM: what MF has issued
    if (Date.now() - t0 < 280000) { if (statRebuild_('Shipment')) info.msg += ' · สรุปสถิติใหม่แล้ว'; }   // สถิติและต้นทุน
  } catch (e) {
    info.ok = false; info.msg = e.message || String(e);
    try { if (tmp && ss) ss.deleteSheet(tmp); } catch (x) {}
    if (info.why !== 'อัตโนมัติ') audit_([['System', 'ซิงค์ข้อมูล Shipment ไม่สำเร็จ', '', info.msg]]);
  }
  schedProp_('SHIP_SYNC_LAST', JSON.stringify(info));
  return info;
}
function shipShow_(d8) { return d8 ? d8.substr(6, 2) + '/' + d8.substr(4, 2) + '/' + d8.substr(0, 4) : ''; }

// ── read ─────────────────────────────────────────────────────────────────
var shipMemo_ = null;
function shipMeta_() {
  var ver = String(schedProp_('SHIP_VER') || '');
  if (!ver) return null;
  if (shipMemo_ && shipMemo_.ver === ver) return shipMemo_;
  var C = CacheService.getScriptCache(), key = 'shipM_' + ver, b = null;
  try { var n = Number(C.get(key + '_n') || 0); if (n) { var o = C.getAll(Array.from({ length: n }, function (x, i) { return key + '_' + i; })), ps = []; for (var i = 0; i < n; i++) ps.push(o[key + '_' + i]); if (ps.every(Boolean)) b = ps.join(''); } } catch (e) {}
  if (!b) {
    var ss = shipSS_(false), ms = ss && ss.getSheetByName(SHIP_META);
    if (!ms || ms.getLastRow() < 1) return null;
    b = ms.getRange(1, 1, ms.getLastRow(), 1).getValues().map(function (r) { return r[0]; }).join('');
    try { var put = {}, k = 0; for (var j = 0; j < b.length; j += 90000) put[key + '_' + (k++)] = b.substring(j, j + 90000); put[key + '_n'] = String(k); C.putAll(put, 21600); } catch (e) {}
  }
  shipMemo_ = JSON.parse(shipUnB64_(b));
  return shipMemo_;
}
// the page: departments, projects (with line count / total / last ship date), report settings
function getEstMeta(token) { apiOnly_();
  requireAdmin_(token);
  var m = shipMeta_(), L = shipInfo_();
  return { ok: !!m, ver: m ? m.ver : '', sum: m ? m.sum.map(function (s) { return [s[0], s[1], shipProj_(s[1]), s[2], s[3], shipShow_(s[4]), s[4]]; }) : [],
    range: m && m.oldest ? shipShow_(m.oldest) + ' – ' + shipShow_(m.newest) : '', last: L, cfg: estCfg_(), hasUrl: !!schedProp_('SHIP_SRC_URL'),
    names: m ? shipNames_(m) : {} };
}
// CATEGORY → the full project number (with its unit letter, PP0990D01) and what the system knows of it:
// Production Schedule first (customer, MVA, kV), then Booking BOM (title). CATEGORY alone cannot tell the letter.
function shipNames_(m) {
  var want = {}, out = {};
  m.sum.forEach(function (s) { want[s[1]] = 1; });
  try {
    (schedCached_() || []).forEach(function (r) {
      var no = String(r[0] || '').trim().toUpperCase(), cat = shipCat_(no);
      if (!cat || !want[cat] || out[cat]) return;
      out[cat] = { no: no, cust: String(r[4] || r[3] || '').trim(), mva: String(r[1] || '').trim(), kv: String(r[2] || '').trim() };
    });
  } catch (e) { console.warn('shipNames_ schedule', e); }
  try {
    (typeof bomProjects_ === 'function' ? bomProjects_() : []).forEach(function (p) {
      var no = String(p.project || '').trim().toUpperCase(), cat = shipCat_(no);
      if (!cat || !want[cat]) return;
      if (!out[cat]) out[cat] = { no: no, cust: '', mva: '', kv: '' };
      if (!out[cat].title && p.title) out[cat].title = String(p.title);
    });
  } catch (e) { console.warn('shipNames_ bom', e); }
  return out;
}
// the stored lines of one CATEGORY (every department) → [CODE, DESCRIPTION, yyyymmdd, DOCNUM, CUSTNAME, LOCATION, QTY, UNIT, UNITCOST, COST, AUDTUSER]
// (also read by Booking BOM — BomShip.gs)
function shipLines_(m, cat) {
  var rows = m.idx[cat] || [], C = CacheService.getScriptCache(), key = 'shipP_' + m.ver + '_' + cat, txt = null;
  try { txt = C.get(key); if (txt) txt = shipUnB64_(txt); } catch (e) { txt = null; }
  if (txt == null) {
    var sh = shipSS_(false).getSheetByName(SHIP_SHEET), parts = [];
    // neighbouring rows are read as one range
    for (var i = 0; i < rows.length;) {
      var j = i; while (j + 1 < rows.length && rows[j + 1] === rows[j] + 1) j++;
      sh.getRange(rows[i], 3, j - i + 1, 1).getValues().forEach(function (r) { if (r[0]) parts.push(shipUnB64_(String(r[0]))); });
      i = j + 1;
    }
    txt = parts.join('\n');
    try { var z = shipB64_(txt); if (z.length < 95000) C.put(key, z, 21600); } catch (e) {}
  }
  var lines = txt ? txt.split('\n').map(function (l) {
    var f = l.split('\t');
    return [f[0], f[1], f[2], f[3], f[4], f[5], Number(f[6]) || 0, f[7], Number(f[8]) || 0, Number(f[9]) || 0, f[10] || ''];
  }) : [];
  return lines;
}
// the lines of one project (every department — the page filters), newest ship date first
function getEstCost(token, proj) { apiOnly_();
  requireAdmin_(token);
  var cat = shipCat_(proj);
  if (!cat) throw new Error('เลขโปรเจกต์ไม่ถูกต้อง — ใช้รูปแบบ PP103001 / PP1030D01 หรือ 1030P1');
  var m = shipMeta_(); if (!m) throw new Error('ยังไม่มีข้อมูล Shipment — ตั้งค่าลิงก์และกดซิงค์ที่ ตั้งค่าระบบ → ☁ แหล่งข้อมูล OneDrive');
  var lines = shipLines_(m, cat);
  lines.sort(function (a, b) { return a[2] < b[2] ? 1 : a[2] > b[2] ? -1 : 0; });
  return { cat: cat, project: shipProj_(cat), cols: ['CODE', 'DESCRIPTION', 'DATE', 'DOCNUM', 'CUSTNAME', 'LOCATION', 'QTY', 'UNIT', 'UNITCOST', 'COST', 'AUDTUSER'], lines: lines };
}

// ── report settings per department: title + 3 signature names ───────────
function estCfg_() { try { return JSON.parse(schedProp_('EST_CFG') || '{}') || {}; } catch (e) { return {}; } }
function saveEstCfg(token, dept, o) { apiOnly_();
  requireAdmin_(token);
  dept = String(dept || '').trim().toUpperCase(); if (!dept) throw new Error('ไม่ได้เลือกแผนก');
  o = o || {};
  var clean = function (s, n) { return String(s == null ? '' : s).replace(/[\r\n\t]+/g, ' ').trim().substring(0, n); };
  var all = estCfg_(), rec = { title: clean(o.title, 120), sign: [0, 1, 2].map(function (i) { return clean((o.sign || [])[i], 60); }) };
  if (!rec.title && !rec.sign.some(Boolean)) delete all[dept]; else all[dept] = rec;
  var s = JSON.stringify(all); if (s.length > 8500) throw new Error('การตั้งค่ารายงานยาวเกินไป');
  schedProp_('EST_CFG', s);
  audit_([['Admin', 'ตั้งค่ารายงาน Estimate Cost', dept, rec.title || '(ค่าเริ่มต้น)']]);
  return all;
}

// ── settings (Admin) ────────────────────────────────────────────────────
function shipAutoSync() { trustEntry_(); if (schedProp_('SHIP_SRC_URL')) syncShipRun_('อัตโนมัติ', false); }   // time trigger
function getShipSource(token) { apiOnly_();
  requireAdmin_(token);
  var hrs = trigHas_('shipAutoSync') ? Number(schedProp_('SHIP_SYNC_HOURS') || 24) : 0, id = schedProp_('SHIP_SS_ID');
  return { url: schedProp_('SHIP_SRC_URL') || '', hours: hrs, last: shipInfo_(), sheetUrl: id ? 'https://docs.google.com/spreadsheets/d/' + id + '/edit' : '' };
}
function setShipSource(token, url, hours) { apiOnly_();
  requireAdmin_(token);
  url = String(url || '').trim();
  if (url && !/^https:\/\/[^\s]+$/i.test(url)) throw new Error('ลิงก์ต้องขึ้นต้นด้วย https://');
  if (/_layouts\/15\/onedrive\.aspx|\/Forms\/AllItems\.aspx/i.test(url)) throw new Error('นี่คือลิงก์หน้าโฟลเดอร์ OneDrive ไม่ใช่ลิงก์แชร์ไฟล์ — คลิกขวาที่ไฟล์ .csv → Share → Copy link แล้ววางลิงก์นั้นแทน');
  schedProp_('SHIP_SRC_URL', url || null);
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'shipAutoSync') ScriptApp.deleteTrigger(t); });
  hours = Number(hours) || 0;
  if (url && hours > 0) {
    syncTrigger_(ScriptApp.newTrigger('shipAutoSync').timeBased(), hours, 4, 24);
    schedProp_('SHIP_SYNC_HOURS', String(hours));
  }
  audit_([['Admin', 'ตั้งค่าซิงค์ข้อมูล Shipment', '', (url ? 'ตั้งลิงก์แล้ว' : 'ลบลิงก์') + ' · ' + syncLabel_(hours)]]);
  return getShipSource(token);
}
function syncShipNow(token) { apiOnly_(); requireAdmin_(token); return syncShipRun_('สั่งโดย Admin', true); }

// Run ONCE from the editor after saving the link, to authorize and check the result in the log.
function testShipSync() { trustEntry_();
  if (!schedProp_('SHIP_SRC_URL')) { Logger.log('✗ ยังไม่ได้บันทึกลิงก์ใน ตั้งค่าระบบ → ☁ ซิงค์ข้อมูล Shipment'); return; }
  var r = syncShipRun_('ทดสอบจาก Editor', true);
  Logger.log((r.ok ? '✓ ' : '✗ ') + r.msg);
}
