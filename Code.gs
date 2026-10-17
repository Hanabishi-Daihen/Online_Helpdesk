/**
 * Online Helpdesk & Schedule Portal — Google Apps Script backend
 * Deploy: Web app → Execute as: Me → Who has access: Anyone
 * Set the admin passcode in Project Settings → Script properties: ADMIN_PASSCODE
 */
var CONFIG = {
  SPREADSHEET_ID: '1zJwAM1Cpfpz8-vaMBkjxEhkr2cu2gW6XzpD-en7Kqns',
  DRIVE_FOLDER_ID: '1qA3Mn-NISkI4yWdwDL_mVoP0UEObYRJ7',
  ISSUES: 'Issues',
  SCHEDULES: 'Q_SchedulesForUpdate',
  SETTINGS: 'Settings'
};
var ISSUE_HEADER = ['Issue ID', 'Project', 'Description', 'Reporter', 'Section', 'Category',
  'Files', 'Reported Date', 'Status', 'ChatLog', 'Priority', 'Timestamp', 'Reporter Email', 'Archived', 'Assignee', 'Closed Date', 'Related'];
var NCOLS = ISSUE_HEADER.length;   // 17

// Clickjacking guard: the page may only be framed when Script property ALLOW_EMBED = 1 (e.g. embedded in Google Sites).
function frameMode_() {
  return PropertiesService.getScriptProperties().getProperty('ALLOW_EMBED') === '1'
    ? HtmlService.XFrameOptionsMode.ALLOWALL : HtmlService.XFrameOptionsMode.DEFAULT;
}
function doGet(e) {
  var ap = (e && e.parameter && e.parameter.ordapprove) || '';
  if (ap) return orderApprovalPage_(ap, (e.parameter.k || ''));     // OrderApproval.gs
  var out = HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Online Helpdesk & Schedule Portal')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(frameMode_());
  var id = (e && e.parameter && e.parameter.issue) || '';
  var tr = (e && e.parameter && e.parameter.track) || '';
  out.append('<script>window.__INIT_ISSUE__=' + JSON.stringify(id).replace(/</g, '\\u003c') +
    ';window.__INIT_TRACK__=' + JSON.stringify(tr).replace(/</g, '\\u003c') + ';</script>');
  return out;
}

// ── helpers ─────────────────────────────────────────────────────────
var __ss = null;                                   // one openById per server call (each call is a fresh execution)
function ss_() { return __ss || (__ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)); }
// timing → Executions log (Apps Script editor → Executions); compare before / after
function perf_(name, t0) { try { console.log('[perf] ' + name + ' ' + (Date.now() - t0) + ' ms'); } catch (e) {} }
function sheet_(name, header) {
  var s = ss_().getSheetByName(name);
  if (!s) { s = ss_().insertSheet(name); if (header) s.appendRow(header); }
  return s;
}
// make sure the Issues sheet has all columns (Archived, Assignee, Closed Date, Related) + headers
function ensureIssueCols_(s) {
  if (s.getMaxColumns() < NCOLS) s.insertColumnsAfter(s.getMaxColumns(), NCOLS - s.getMaxColumns());
  if (s.getLastRow() >= 1) {
    var h = s.getRange(1, 1, 1, NCOLS).getValues()[0], dirty = false;
    for (var i = 13; i < NCOLS; i++) if (!h[i]) { h[i] = ISSUE_HEADER[i]; dirty = true; }
    if (dirty) s.getRange(1, 1, 1, NCOLS).setValues([h]);
  }
}
function nowStr_() { return Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss'); }
function getAppUrl() { return ScriptApp.getService().getUrl(); }
function parse_(v) { try { var a = JSON.parse(v); return Array.isArray(a) ? a : []; } catch (e) { return []; } }
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { throw new Error('ระบบกำลังบันทึกข้อมูลของผู้ใช้อื่นอยู่ กรุณาลองอีกครั้งในไม่กี่วินาที'); }
  if (typeof ordMemoReset_ === 'function') ordMemoReset_();
  try { return fn(); } finally { lock.releaseLock(); }
}
function isAdmin_(token) { return !!token && !!CacheService.getScriptCache().get('adm_' + token); }
function requireAdmin_(token) { if (!isAdmin_(token)) throw new Error('Admin only: สิทธิ์ Admin หมดอายุหรือยังไม่ได้ปลดล็อค'); }
function esc_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
// Plain-text format. Google Sheets "Tables" refuse a multi-column format change
// ("โปรดเลือกภายในคอลัมน์เดียว…"), so fall back to one column at a time.
function fmtText_(rg) {
  try { rg.setNumberFormat('@'); return; } catch (e) {}
  var sh = rg.getSheet(), r0 = rg.getRow(), c0 = rg.getColumn(), nr = rg.getNumRows(), nc = rg.getNumColumns();
  for (var j = 0; j < nc; j++) { try { sh.getRange(r0, c0 + j, nr, 1).setNumberFormat('@'); } catch (e2) {} }
}
function writeText_(sh, row, col, values) {
  var r = sh.getRange(row, col, 1, values.length);
  fmtText_(r);                      // keep text as-is (no formula / date auto-conversion)
  r.setValues([values]);
}
function rowToIssue_(r) {
  return {
    id: String(r[0]).trim(), project: r[1], description: r[2], reporter: r[3],
    section: r[4] || 'MF', category: r[5] || 'อื่นๆ', files: parse_(r[6]),
    reportedDate: r[7], status: r[8] || 'New', chat: parse_(r[9]),
    priority: r[10] || 'Medium', timestamp: r[11], reporterEmail: r[12] || undefined,
    archived: String(r[13] || '').trim() === '1',
    assignee: r[14] || '', closedDate: r[15] || '',
    related: String(r[16] || '').split(/[,\s]+/).filter(Boolean)
  };
}
function issueToRow_(i) {
  return [i.id, i.project, i.description, i.reporter, i.section, i.category,
    JSON.stringify(i.files || []), i.reportedDate, i.status, JSON.stringify(i.chat || []),
    i.priority, i.timestamp, i.reporterEmail || '', i.archived ? '1' : '',
    i.assignee || '', i.closedDate || '', (i.related || []).join(',')];
}

// Move base64 data-URL attachments into the Drive folder; return Drive links.
// Allowlist: only these extensions are accepted, and the MIME type is taken from the extension (never from the client).
var FILE_ALLOW = {
  pdf: 'application/pdf', png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', webp: 'image/webp',
  bmp: 'image/bmp', heic: 'image/heic', heif: 'image/heif', tif: 'image/tiff', tiff: 'image/tiff',
  txt: 'text/plain', csv: 'text/csv', log: 'text/plain',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', xls: 'application/vnd.ms-excel',
  xlsm: 'application/vnd.ms-excel.sheet.macroEnabled.12',
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', doc: 'application/msword',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', ppt: 'application/vnd.ms-powerpoint',
  zip: 'application/zip', '7z': 'application/x-7z-compressed', rar: 'application/vnd.rar',
  dwg: 'application/acad', dxf: 'application/dxf', step: 'application/step', stp: 'application/step',
  igs: 'model/iges', iges: 'model/iges', mp4: 'video/mp4', mov: 'video/quicktime'
};
var FILE_MAX_BYTES = 20 * 1024 * 1024;
function fileExt_(nm) { var m = /\.([A-Za-z0-9]{1,5})$/.exec(nm || ''); return m ? m[1].toLowerCase() : ''; }
function extForMime_(mime) {                          // pasted images often arrive without a usable file name
  for (var k in FILE_ALLOW) if (FILE_ALLOW[k] === mime && k !== 'jpeg' && k !== 'tif') return k;
  return '';
}
function safeFileUrl_(u) { return /^https?:\/\/[^\s"'<>]+$/i.test(String(u || '')); }   // blocks javascript:/data: links
function persistFiles_(files) {
  return (files || []).map(function (f) {
    if (!f || !f.url) return f;
    if (f.url.indexOf('data:') !== 0) {               // already stored: keep only plain http(s) links
      if (!safeFileUrl_(f.url)) throw new Error('ลิงก์ไฟล์แนบไม่ถูกต้อง: ' + String(f.name || '').substring(0, 60));
      return { name: String(f.name || 'file').substring(0, 120), type: String(f.type || '').substring(0, 120), size: f.size,
        url: f.url, thumbnail: safeFileUrl_(f.thumbnail) ? f.thumbnail : null };
    }
    var nm = String(f.name || 'file').replace(/[\\/:*?"<>|\u0000-\u001f]+/g, '_').substring(0, 120) || 'file';
    var comma = f.url.indexOf(',');
    if (comma < 0) throw new Error('ไฟล์ "' + nm + '" ไม่ถูกต้อง');
    var ext = fileExt_(nm);
    if (!ext) { ext = extForMime_(f.url.substring(5, comma).split(';')[0].toLowerCase()); if (ext) nm = nm.substring(0, 110) + '.' + ext; }
    if (!FILE_ALLOW[ext]) throw new Error('ไม่อนุญาตให้แนบไฟล์ชนิดนี้: ' + nm + ' (อนุญาต: ' + Object.keys(FILE_ALLOW).join(', ') + ')');
    if ((f.url.length - comma) * 0.75 > FILE_MAX_BYTES) throw new Error('ไฟล์ "' + nm + '" ใหญ่เกิน 20MB');
    f.name = nm;
    var mime = FILE_ALLOW[ext];
    var blob = Utilities.newBlob(Utilities.base64Decode(f.url.substring(comma + 1)), mime, f.name);
    var file = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID).createFile(blob);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    var id = file.getId();
    return {
      name: f.name, type: mime, size: f.size,
      url: 'https://drive.google.com/file/d/' + id + '/view',
      thumbnail: f.thumbnail ? 'https://drive.google.com/thumbnail?id=' + id + '&sz=w1000' : null
    };
  });
}

// ── reads (public) ──────────────────────────────────────────────────
// ── issues version: bumped on every write; the 60 s poll asks for it first (one cache read)
// and only re-reads the sheet when it changed. The 5-minute bucket still picks up edits made
// directly in the sheet within 5 minutes.
function issVer_() {
  var c = CacheService.getScriptCache(), v = c.get('iss_ver');
  if (!v) { v = PropertiesService.getScriptProperties().getProperty('ISS_VER') || '0'; c.put('iss_ver', v, 21600); }
  return v + ':' + Math.floor(Date.now() / 3e5);
}
function issBump_() {
  var v = String(Date.now());
  PropertiesService.getScriptProperties().setProperty('ISS_VER', v); CacheService.getScriptCache().put('iss_ver', v, 21600);
}
function getIssuesVer() { return issVer_(); }
function getIssues(token) {
  var t0 = Date.now();
  var s = ss_().getSheetByName(CONFIG.ISSUES);
  if (!s || s.getLastRow() < 2) return [];
  ensureIssueCols_(s);
  var rows = s.getRange(2, 1, s.getLastRow() - 1, NCOLS).getDisplayValues();
  if (!isAdmin_(token)) rows.forEach(function (r) { r[12] = ''; });   // reporter email: admin only
  perf_('getIssues', t0);
  return rows;
}
function getSchedules() {
  var s = ss_().getSheetByName(CONFIG.SCHEDULES);
  if (!s || s.getLastRow() < 2) return [];
  return s.getRange(2, 1, s.getLastRow() - 1, 11).getDisplayValues();
}
function getSettings() {
  var s = ss_().getSheetByName(CONFIG.SETTINGS);
  if (!s) return {};
  try { return JSON.parse(s.getRange('A1').getValue()) || {}; } catch (e) { return {}; }
}

// ── issue writes (public, same openness as the original app) ────────
function nextId_(sh) {
  var year = String(new Date().getFullYear()), max = 0;
  var last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) {
    var id = String(r[0]);
    if (id.indexOf('ISS-' + year) === 0) max = Math.max(max, parseInt(id.substring(8), 10) || 0);
  });
  return 'ISS-' + year + ('00' + (max + 1)).slice(-3);
}

function createIssue(issue) {
  var t0 = Date.now();
  clean_(issue); rateLimit_(issue);
  issue.files = persistFiles_(issue.files);                     // Drive uploads happen outside the lock
  (issue.chat || []).forEach(function (m) { if (m && m.files) m.files = persistFiles_(m.files); });
  var res = withLock_(function () {
    issue.archived = false; issue.assignee = ''; issue.closedDate = ''; issue.related = [];   // admin-only fields
    issue.status = 'New';                                       // a new case always starts as New
    if (!isAdmin_(issue._token)) (issue.chat || []).forEach(function (m) { if (m && String(m.sender || '').trim().toLowerCase() === 'admin') m.sender = 'User'; });
    delete issue._token;
    var sh = sheet_(CONFIG.ISSUES, ISSUE_HEADER), oldId = issue.id;
    ensureIssueCols_(sh);
    var last = sh.getLastRow();
    var ids = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return String(r[0]).trim(); }) : [];
    if (!oldId || ids.indexOf(oldId) >= 0) issue.id = nextId_(sh);
    writeText_(sh, Math.max(last, 1) + 1, 1, issueToRow_(issue));
    issBump_();
    return { oldId: oldId, issue: issue };
  });
  try { notifyNew_(issue); } catch (e) { console.warn(e); }      // mail after the lock — never blocks other users
  perf_('createIssue', t0);
  return res;
}

function mergeChat_(oldChat, newChat) {
  var key = function (m) { return [m.sender, m.time, m.message, m.type || '', m.newStatus || ''].join('|'); };
  var out = oldChat.slice(), map = {};
  out.forEach(function (m, i) { map[key(m)] = i; });
  newChat.forEach(function (m) {
    var k = key(m);
    if (k in map) { if (m.read) out[map[k]].read = true; }
    else { m.message = String(m.message || '').substring(0, 2000); if (m.files) m.files = persistFiles_(m.files); map[k] = out.length; out.push(m); }
  });
  return out;
}

// Non-admin callers may not post messages as the team (that would also trigger "reply" e-mails),
// and may not add System / status-change entries (the status itself is admin-only).
// Messages already stored (same key as in mergeChat_) pass through untouched so "read" flags still work.
function userChat_(list, oldChat) {
  var key = function (m) { return [m.sender, m.time, m.message, m.type || '', m.newStatus || ''].join('|'); }, have = {};
  (oldChat || []).forEach(function (m) { have[key(m)] = 1; });
  return (Array.isArray(list) ? list : []).filter(function (m) {
    if (!m || typeof m !== 'object') return false;
    if (have[key(m)]) return true;
    return m.sender !== 'System' && m.type !== 'status' && !m.newStatus;
  }).map(function (m) {
    if (!have[key(m)] && String(m.sender || '').trim().toLowerCase() === 'admin') m.sender = 'User';
    return m;
  });
}
function nz_(v) { return v == null ? '' : (typeof v === 'object' ? JSON.stringify(v) : String(v)); }
// Merge one incoming issue into the stored one (pure function — no sheet access).
function mergeOne_(cur, inc, admin) {
  if (!admin) return mergeUser_(cur, inc);
  var m = {
    id: cur.id, project: inc.project, description: inc.description, reporter: inc.reporter,
    section: inc.section, category: inc.category,
    files: cur.files.length ? cur.files : persistFiles_(inc.files),
    reportedDate: cur.reportedDate, status: inc.status, priority: inc.priority,
    timestamp: cur.timestamp, reporterEmail: inc.reporterEmail || cur.reporterEmail,
    chat: mergeChat_(cur.chat, inc.chat || []),
    archived: admin ? !!inc.archived : cur.archived,
    assignee: admin ? String(inc.assignee || '') : cur.assignee,
    related: admin ? (inc.related || []) : cur.related,
    closedDate: ''
  };
  // Conflict guard: inc._base = the values this client last saw from the server.
  // Fields the client did not touch keep the latest server value; if both sides changed the same field, the first save wins.
  var b = inc._base;
  if (b && typeof b === 'object') {
    var F = ['project', 'description', 'section', 'category', 'status', 'priority'], conf = [];
    if (admin) F = F.concat(['assignee', 'archived', 'related']);
    F.forEach(function (f) {
      var n = nz_(inc[f]), bb = nz_(b[f]), c = nz_(cur[f]);
      if (n === bb) m[f] = cur[f];
      else if (c !== bb && c !== n) { m[f] = cur[f]; conf.push(f); }
    });
    if (conf.length) m._conflicts = conf;
  }
  // real close date: set once when the case first becomes Resolved/Closed, cleared when re-opened
  if (m.status === 'Resolved' || m.status === 'Closed') m.closedDate = cur.closedDate || (admin && inc.closedDate) || nowStr_();
  return m;
}

// Non-admin update: the case itself is read-only — only new chat messages (and "read" flags) are accepted.
// Status, description, priority, reporter, reporter e-mail, files, etc. always keep the stored values,
// so nobody can take over another person's case (or redirect its notification e-mails).
function mergeUser_(cur, inc) {
  var m = JSON.parse(JSON.stringify(cur));
  m.chat = mergeChat_(cur.chat, userChat_(inc.chat || [], cur.chat));
  var F = ['project', 'description', 'reporter', 'section', 'category', 'status', 'priority', 'reporterEmail'], tried = [];
  F.forEach(function (f) { if (inc[f] !== undefined && nz_(inc[f]) !== nz_(cur[f]) && !(f === 'reporterEmail' && !inc[f])) tried.push(f); });
  if (tried.length) m._denied = tried;
  return m;
}

// One read + one write per call (fast even for bulk Archive / bulk status change).
function upsertIssues(list, token) {
  var t0 = Date.now(), notify = [];
  var out = withLock_(function () {
    var admin = isAdmin_(token);
    var sh = sheet_(CONFIG.ISSUES, ISSUE_HEADER), result = [], appendRows = [], dirty = {}, audits = [];
    ensureIssueCols_(sh);
    var last = sh.getLastRow();
    var all = last > 1 ? sh.getRange(2, 1, last - 1, NCOLS).getDisplayValues() : [];
    var idx = {};
    all.forEach(function (r, i) { idx[String(r[0]).trim()] = i; });
    list.forEach(function (inc) {
      clean_(inc);
      if (!admin) { inc.archived = false; }           // only admin can archive / restore / assign / link
      var key = String(inc.id).trim();
      if (!(key in idx)) {                             // row removed manually → recreate (admin only)
        if (!admin) throw new Error('ไม่พบเคส ' + key + ' — กรุณารีเฟรชหน้า (ผู้ใช้ทั่วไปสร้างเคสได้จากแบบฟอร์มแจ้งปัญหาเท่านั้น)');
        inc.files = persistFiles_(inc.files);
        idx[key] = all.length; all.push(issueToRow_(inc)); appendRows.push(idx[key]);
        result.push(inc); return;
      }
      var i = idx[key], cur = rowToIssue_(all[i]);
      var m = mergeOne_(cur, inc, admin);
      all[i] = issueToRow_(m);
      if (appendRows.indexOf(i) < 0) dirty[i] = true;
      if (m._denied) { audits.push(['User', 'พยายามแก้ไขเคส (ถูกปฏิเสธ)', m.id, m._denied.join(', ')]); delete m._denied; }
      notify.push([cur, m]); auditDiff_(cur, m, admin, audits); if (m._conflicts) audits.push([admin ? 'Admin' : 'User', 'ชนกันกับผู้อื่น (ใช้ค่าเดิม)', m.id, m._conflicts.join(', ')]);
      result.push(admin ? m : JSON.parse(JSON.stringify(m, function (k, v) { return k === 'reporterEmail' ? undefined : v; })));
    });
    var rows = Object.keys(dirty).map(Number).sort(function (a, b) { return a - b; });
    if (rows.length && rows.length <= 4) {
      rows.forEach(function (i) {
        var full = all[i], row = i + 2;
        writeText_(sh, row, 2, full.slice(1, 7));        // B:G
        writeText_(sh, row, 9, full.slice(8, 11));       // I:K
        writeText_(sh, row, 13, full.slice(12, NCOLS));  // M:Q
      });
    } else if (rows.length) {                            // bulk: write one block
      var lo = rows[0], hi = rows[rows.length - 1], blk = all.slice(lo, hi + 1);
      var rg = sh.getRange(lo + 2, 1, blk.length, NCOLS);
      fmtText_(rg); rg.setValues(blk);
    }
    if (appendRows.length) {
      var first = appendRows[0], blk2 = all.slice(first);
      var rg2 = sh.getRange(first + 2, 1, blk2.length, NCOLS);
      fmtText_(rg2); rg2.setValues(blk2);
    }
    audit_(audits);
    if (rows.length || appendRows.length) issBump_();
    return result;
  });
  notify.forEach(function (p) {                          // mails after the lock — a mail problem never blocks saving
    try { notifyChange_(p[0], p[1]); } catch (e) { console.warn(e); }
  });
  perf_('upsertIssues', t0);
  return out;
}

// Admin: connection / permission self-test (run from Settings → ตรวจสอบการเชื่อมต่อ, or from the editor)
function diagnose(token) {
  requireAdmin_(token);
  return diagnoseRun_();
}
function diagnoseRun_() {
  var out = [];
  try {
    var ss = ss_();
    out.push('✓ เปิดสเปรดชีตได้: ' + ss.getName());
    out.push('แท็บ: ' + ss.getSheets().map(function (x) { return x.getName(); }).join(', '));
    var s = ss.getSheetByName(CONFIG.ISSUES);
    if (!s) out.push('✗ ไม่พบแท็บ "' + CONFIG.ISSUES + '"');
    else {
      out.push('Issues: ' + s.getLastRow() + ' แถวข้อมูล, ตารางกว้าง ' + s.getMaxColumns() + ' คอลัมน์');
      try { ensureIssueCols_(s); out.push('✓ คอลัมน์ครบ ' + NCOLS + ' คอลัมน์'); } catch (e) { out.push('✗ เพิ่มคอลัมน์ไม่ได้: ' + e.message); }
      try { var c = s.getRange(1, 1); c.setValue(c.getValue()); out.push('✓ ทดสอบเขียนลงชีต: ผ่าน'); } catch (e) { out.push('✗ เขียนลงชีตไม่ได้: ' + e.message); }
      try { s.getRange(1, 1, 1, 2).setNumberFormat('@'); out.push('✓ ตั้งรูปแบบข้อความหลายคอลัมน์: ผ่าน'); } catch (e) { out.push('⚠ แท็บ Issues น่าจะเป็น "ตาราง (Table)" ของ Google Sheets — ระบบใช้วิธีสำรองตั้งรูปแบบทีละคอลัมน์ให้แล้ว (' + e.message + ')'); }
      try {
        var p = s.getProtections(SpreadsheetApp.ProtectionType.SHEET).concat(s.getProtections(SpreadsheetApp.ProtectionType.RANGE));
        if (p.length) out.push('⚠ แท็บ Issues มีการป้องกัน (Protected) ' + p.length + ' รายการ');
      } catch (e) {}
    }
  } catch (e) { out.push('✗ เปิดสเปรดชีตไม่ได้: ' + e.message + ' (ตรวจ SPREADSHEET_ID และสิทธิ์แก้ไขของบัญชีที่ Deploy)'); }
  try { DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID).getName(); out.push('✓ โฟลเดอร์ไฟล์แนบ: เข้าถึงได้'); } catch (e) { out.push('✗ เข้าโฟลเดอร์ไฟล์แนบไม่ได้: ' + e.message); }
  return out;
}

// ── admin ───────────────────────────────────────────────────────────
function adminLogin(pass) {
  var real = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSCODE') || 'control';
  if (String(pass) !== real) { Utilities.sleep(800); return ''; }
  var token = Utilities.getUuid();
  CacheService.getScriptCache().put('adm_' + token, '1', 21600);
  return token;
}

function saveAllIssues(token, issues) {
  requireAdmin_(token);
  try { backupRun_('ก่อนเขียนทับ Issues'); } catch (e) { console.warn(e); }
  audit_([['Admin', 'เขียนทับ Issues ทั้งชีต', '', (issues || []).length + ' รายการ']]);
  return withLock_(function () {
    var sh = sheet_(CONFIG.ISSUES, ISSUE_HEADER);
    var rows = issues.map(function (i) {
      i.files = persistFiles_(i.files);
      (i.chat || []).forEach(function (m) { if (m.files) m.files = persistFiles_(m.files); });
      return issueToRow_(i);
    });
    ensureIssueCols_(sh);
    sh.clearContents();
    var all = [ISSUE_HEADER].concat(rows);
    var r = sh.getRange(1, 1, all.length, NCOLS);
    fmtText_(r); r.setValues(all);
    issBump_();
    return true;
  });
}

// NOTE: overwrites the whole sheet. If Q_SchedulesForUpdate is a QUERY()/ImportRange formula sheet, do not use this.
function saveAllSchedules(token, schedules) {
  requireAdmin_(token);
  try { backupRun_('ก่อนเขียนทับ Schedules'); } catch (e) { console.warn(e); }
  audit_([['Admin', 'เขียนทับ Schedules ทั้งชีต', '', (schedules || []).length + ' รายการ']]);
  return withLock_(function () {
    var sh = sheet_(CONFIG.SCHEDULES);
    var header = ['ID / ชื่องาน (Project)', 'MVA', 'Rating Volt', 'ผู้สั่งซื้อ (Purchaser)', 'ลูกค้า (Customer)',
      'NewDelivery', 'หมายเหตุ', 'ActualPlanDelivery', 'Process', 'Start', 'Finish'];
    var rows = schedules.map(function (s) {
      return [s.id || s.title || '', s.mva || '', s.ratingVolt || '', s.purchaser || '', s.customer || '',
        s.newDelivery || '', s.notes || '', s.actualPlanDelivery || '', s.process || '', s.start || '', s.finish || ''];
    });
    sh.clearContents();
    var all = [header].concat(rows);
    sh.getRange(1, 1, all.length, 11).setValues(all);
    return true;
  });
}

function saveSettings(token, patch) {
  requireAdmin_(token);
  return withLock_(function () {
    var cur = getSettings();
    ['ui', 'admin', 'emailTemplates'].forEach(function (k) { if (patch[k]) cur[k] = patch[k]; });
    audit_([['Admin', 'แก้ไขการตั้งค่า', '', Object.keys(patch || {}).join(', ')]]);
    var js = JSON.stringify(cur);
    if (js.length > 49000 && cur.ui && cur.ui.logoDataUrl) { delete cur.ui.logoDataUrl; js = JSON.stringify(cur); }
    sheet_(CONFIG.SETTINGS).getRange('A1').setNumberFormat('@').setValue(js);
    return js.length <= 49000;
  });
}

// ── validation / anti-spam ──────────────────────────────────────────
function clean_(i) {
  var L = function (v, n) { return String(v == null ? '' : v).substring(0, n); };
  i.project = L(i.project, 200); i.description = L(i.description, 5000);
  i.reporter = L(i.reporter, 100); i.reporterEmail = L(i.reporterEmail, 120);
  i.assignee = L(i.assignee, 80); i.closedDate = L(i.closedDate, 30);
  i.related = (Array.isArray(i.related) ? i.related : []).map(String).filter(function (x) { return /^[A-Za-z0-9-]{3,24}$/.test(x); }).slice(0, 20);
  if ((i.files || []).length > 5) throw new Error('แนบไฟล์ได้สูงสุด 5 ไฟล์ต่อเคส');
  (i.files || []).forEach(function (f) { if (f && f.url && f.url.length > 28000000) throw new Error('ไฟล์ "' + (f.name || '') + '" มีขนาดใหญ่เกินไป (สูงสุด 20MB)'); });
}
function rateLimit_(i) {                       // max 10 new issues / hour per reporter
  var c = CacheService.getScriptCache(), k = 'rl_' + String(i.reporterEmail || i.reporter).toLowerCase().substring(0, 80);
  var n = Number(c.get(k) || 0);
  if (n >= 10) throw new Error('ส่งเคสถี่เกินไป กรุณารอสักครู่');
  c.put(k, String(n + 1), 3600);
}

// ═════════════════════════════════════════════════════════════════════
//  E-MAIL  —  templates are DATA (EMAIL_DEFAULTS). Admin → Settings → "แก้ไขรูปแบบอีเมล" saves overrides to
//  Settings!B1 (JSON: {templates:{key:{...}}, brand:{...}}). Layout/design lives in emailHtml_().
// ═════════════════════════════════════════════════════════════════════
var EMAIL_CFG = {
  brand: '#cc785c', brandDark: '#a9583e', ink: '#141413', muted: '#6c6a64', line: '#e6dfd8', bg: '#f5f0e8',
  company: 'DAIHEN ELECTRIC CO., LTD.', product: 'Online Helpdesk', senderName: 'DAIHEN Helpdesk'
};
var EMAIL_BASE = JSON.parse(JSON.stringify(EMAIL_CFG));          // pristine copy (brand overrides are applied per run)
var EMAIL_FOOTER_DEFAULT = 'อีเมลฉบับนี้ส่งโดยอัตโนมัติจากระบบแจ้งปัญหาออนไลน์ กรุณาอย่าตอบกลับอีเมลนี้โดยตรง หากต้องการสอบถามเพิ่มเติม ให้ใช้ปุ่มด้านบนเพื่อเปิดเคสและส่งข้อความถึงทีมงาน';
var STATUS_TH = { 'New': 'ใหม่', 'Waiting': 'รอดำเนินการ', 'In Progress': 'กำลังแก้ไข', 'Resolved': 'แก้ไขแล้ว', 'Closed': 'ปิดงาน' };
var STATUS_ORDER = ['New', 'Waiting', 'In Progress', 'Resolved', 'Closed'];
var STATUS_COL = { 'New': '#3b82f6', 'Waiting': '#d97706', 'In Progress': '#7c3aed', 'Resolved': '#059669', 'Closed': '#6b7280' };
var PRIORITY_TH = { 'Low': 'ต่ำ', 'Medium': 'ปานกลาง', 'High': 'สูง', 'Urgent': 'ด่วนมาก' };
var PRIORITY_COL = { 'Low': ['#e0f2fe', '#075985'], 'Medium': ['#fef3c7', '#92400e'], 'High': ['#ffedd5', '#9a3412'], 'Urgent': ['#ffe4e6', '#9f1239'] };
var STATUS_INTRO = {
  'New': 'เคสของคุณถูกบันทึกเข้าระบบแล้ว และกำลังรอทีมงานตรวจสอบ',
  'Waiting': 'เคสของคุณอยู่ในสถานะรอดำเนินการ ทีมงานจะเริ่มดำเนินการโดยเร็วที่สุด',
  'In Progress': 'ทีมงานเริ่มดำเนินการแก้ไขปัญหาของคุณแล้ว',
  'Resolved': 'ปัญหาของคุณได้รับการแก้ไขเรียบร้อยแล้ว หากยังพบปัญหาเดิม สามารถตอบกลับผ่านหน้าติดตามสถานะได้',
  'Closed': 'เคสนี้ปิดงานเรียบร้อยแล้ว ขอบคุณที่แจ้งให้เราทราบ'
};
var EMAIL_DEFAULTS = {
  new_admin:         { enabled: true, subject: '[Helpdesk] เคสใหม่ {{id}}{{urgentTag}} — {{project}}',
                       heading: 'มีเคสแจ้งปัญหาใหม่', intro: '{{reporter}} แจ้งปัญหาในแผนก {{section}} กรุณาตรวจสอบและรับเรื่อง',
                       button: 'เปิดเคสในระบบ', button2: 'ลิงก์หน้าติดตามสถานะ' },
  new_reporter:      { enabled: true, subject: '[Helpdesk] เราได้รับเรื่องของคุณแล้ว ({{id}})',
                       heading: 'เราได้รับเรื่องของคุณแล้ว', intro: 'ขอบคุณที่แจ้งปัญหา ทีมงานจะตรวจสอบและแจ้งความคืบหน้าทางอีเมลนี้ คุณสามารถติดตามสถานะได้ทุกเมื่อจากปุ่มด้านล่าง',
                       button: 'ติดตามสถานะเคส' },
  status_reporter:   { enabled: true, subject: '[Helpdesk] {{id}} — สถานะเปลี่ยนเป็น {{statusLabel}}',
                       heading: 'อัปเดตสถานะเคสของคุณ', intro: '{{statusIntro}}', button: 'ดูความคืบหน้า' },
  status_admin:      { enabled: false, subject: '[Helpdesk] {{id}} เปลี่ยนสถานะเป็น {{statusLabel}}',
                       heading: 'เคสเปลี่ยนสถานะ', intro: '{{id}} เปลี่ยนจาก {{oldStatusLabel}} เป็น {{statusLabel}}', button: 'เปิดเคสในระบบ' },
  assigned_reporter: { enabled: true, subject: '[Helpdesk] {{id}} — มีผู้รับผิดชอบเคสของคุณแล้ว',
                       heading: 'มีผู้รับผิดชอบเคสของคุณแล้ว', intro: '{{assignee}} รับผิดชอบเคสนี้ และจะติดตามจนกว่าปัญหาจะได้รับการแก้ไข', button: 'ดูความคืบหน้า' },
  reply_reporter:    { enabled: true, subject: '[Helpdesk] {{id}} — ทีมงานตอบกลับเคสของคุณ',
                       heading: 'ทีมงานตอบกลับเคสของคุณ', intro: 'มีข้อความใหม่จากทีมงานเกี่ยวกับเคส {{id}}', button: 'เปิดดูและตอบกลับ' },
  reminder_admin:    { enabled: true, subject: '[Helpdesk] เคสค้างเกิน {{days}} วัน: {{count}} รายการ',
                       heading: 'สรุปเคสที่ค้างอยู่', intro: 'มี {{count}} เคสที่ยังไม่ปิดและค้างเกิน {{days}} วัน', button: 'เปิดแดชบอร์ด' }
};

Object.keys(EMAIL_DEFAULTS).forEach(function (k) {
  EMAIL_DEFAULTS[k].showCard = k !== 'reminder_admin';
  EMAIL_DEFAULTS[k].showStepper = (k === 'status_reporter' || k === 'status_admin' || k === 'new_reporter');
});

// Labels shown in the admin editor
var EMAIL_META = {
  new_admin:         { label: 'เคสใหม่ → ผู้ดูแล', desc: 'ส่งทันทีที่มีคนแจ้งปัญหาใหม่ ถึงอีเมลผู้ดูแล (To) และ CC ที่ตั้งไว้ใน Email Notifications' },
  new_reporter:      { label: 'ตอบรับเคส → ผู้แจ้ง', desc: 'ส่งให้ผู้แจ้งที่กรอกอีเมลไว้ เพื่อยืนยันว่าได้รับเรื่องแล้ว พร้อมลิงก์ติดตามสถานะ' },
  status_reporter:   { label: 'เปลี่ยนสถานะ → ผู้แจ้ง', desc: 'ส่งทุกครั้งที่สถานะเคสเปลี่ยน (ผู้ดูแลได้รับสำเนาด้วย)' },
  assigned_reporter: { label: 'มอบหมายผู้รับผิดชอบ → ผู้แจ้ง', desc: 'ส่งเมื่อมีการระบุ/เปลี่ยนผู้รับผิดชอบเคส โดยสถานะไม่ได้เปลี่ยน' },
  reply_reporter:    { label: 'ทีมงานตอบกลับ → ผู้แจ้ง', desc: 'ส่งเมื่อ Admin ตอบข้อความในเคส (ข้อความจะแสดงในกล่องสีฟ้า)' },
  reminder_admin:    { label: 'สรุปเคสค้าง (รายวัน)', desc: 'ส่งตามตัวตั้งเวลา installDailyReminder สรุปเคสที่ค้างเกินจำนวนวันที่ตั้งไว้' },
  status_admin:      { label: 'เปลี่ยนสถานะ → ผู้ดูแล', inactive: true,
                       desc: 'ยังไม่ถูกใช้งานในระบบ — ผู้ดูแลได้รับอีเมลสถานะผ่านเทมเพลต "เปลี่ยนสถานะ → ผู้แจ้ง" อยู่แล้ว' }
};
var EMAIL_VARS = [
  ['id', 'เลขเคส'], ['project', 'หัวข้อปัญหา'], ['description', 'รายละเอียด'], ['reporter', 'ชื่อผู้แจ้ง'],
  ['section', 'แผนก'], ['category', 'หมวดหมู่'], ['priorityLabel', 'ความสำคัญ (ไทย)'], ['priority', 'ความสำคัญ (Urgent...)'],
  ['statusLabel', 'สถานะ (ไทย)'], ['status', 'สถานะ (New...)'], ['statusIntro', 'ข้อความอธิบายสถานะ'],
  ['assignee', 'ผู้รับผิดชอบ'], ['reportedDate', 'วันที่แจ้ง'], ['closedDate', 'วันที่ปิดงาน'],
  ['urgentTag', '" (ด่วนมาก)" เฉพาะเคส Urgent'], ['link', 'ลิงก์เปิดเคส (Admin)'], ['trackLink', 'ลิงก์ติดตามสถานะ']
];

var __emailStore = null;
// Overrides live in Settings!B1 (kept apart from A1 so a big logo can never push them out). Falls back to the legacy A1 key.
var EMAIL_ROSTER_DEFAULT = ['premwit@daihen.co.th', 'thutchanon@daihen.co.th', 'panudech@daihen.co.th', 'panuwat@daihen.co.th', 'teerapol@daihen.co.th', 'nuttapol@daihen.co.th'];
function rosterDefault_() {
  return EMAIL_ROSTER_DEFAULT.map(function (e) { var n = e.split('@')[0]; return { email: e, name: n.charAt(0).toUpperCase() + n.slice(1) }; });
}
function cleanRoster_(arr) {
  var seen = {}, out = [];
  (Array.isArray(arr) ? arr : []).forEach(function (r) {
    var e = String((r && r.email) || '').trim().toLowerCase();
    if (!/^[^@\s,;]+@[^@\s,;]+$/.test(e) || seen[e] || out.length >= 50) return;
    seen[e] = 1;
    out.push({ email: e, name: String((r && r.name) || '').replace(/[\r\n]+/g, ' ').trim().substring(0, 40) || e.split('@')[0] });
  });
  return out;
}
function emailStore_() {
  if (__emailStore) return __emailStore;
  var out = { templates: {}, brand: {}, roster: null };
  try {
    var sh = ss_().getSheetByName(CONFIG.SETTINGS), raw = sh ? String(sh.getRange('B1').getValue() || '') : '';
    if (raw) { var o = JSON.parse(raw); out.templates = o.templates || {}; out.brand = o.brand || {}; if (Array.isArray(o.roster)) out.roster = cleanRoster_(o.roster); }
    else out.templates = settings_().emailTemplates || {};
  } catch (e) { console.warn(e); }
  if (!out.roster) out.roster = rosterDefault_();
  return (__emailStore = out);
}
function emailList_(s) { return String(s || '').split(/[,;\s]+/).filter(function (e) { return /^[^@\s]+@[^@\s]+$/.test(e); }); }
function cleanBrand_(b) {
  b = b || {}; var o = {};
  ['brand', 'brandDark'].forEach(function (k) { if (/^#[0-9a-fA-F]{6}$/.test(String(b[k] || ''))) o[k] = String(b[k]); });
  ['product', 'company', 'senderName', 'footer'].forEach(function (k) {
    var v = String(b[k] == null ? '' : b[k]).replace(/[\r\n]+/g, ' ').trim();
    if (v) o[k] = v.substring(0, k === 'footer' ? 500 : 80);
  });
  return o;
}
function cleanTpl_(s) {
  var o = {};
  if (!s || typeof s !== 'object') return o;
  ['subject', 'heading', 'intro', 'button', 'button2', 'footer'].forEach(function (f) {
    if (s[f] != null && String(s[f]) !== '') o[f] = String(s[f]).substring(0, (f === 'footer' || f === 'intro') ? 1500 : 400);
  });
  ['enabled', 'showStepper', 'showCard'].forEach(function (f) { if (typeof s[f] === 'boolean') o[f] = s[f]; });
  return o;
}
function applyBrand_(draftBrand) {
  Object.keys(EMAIL_BASE).forEach(function (k) { EMAIL_CFG[k] = EMAIL_BASE[k]; });
  var b = cleanBrand_(draftBrand !== undefined && draftBrand !== null ? draftBrand : emailStore_().brand);
  Object.keys(b).forEach(function (k) { EMAIL_CFG[k] = b[k]; });
}

var __settingsCache = null;
function settings_() { if (!__settingsCache) __settingsCache = getSettings(); return __settingsCache; }
function emails_(k) { return String((settings_().admin || {})[k] || '').split(/[,;\s]+/).filter(function (e) { return /@/.test(e); }); }
function tpl_(key, draftOv) {                    // draftOv = unsaved override from the editor (preview / test send)
  var o = {}, d = EMAIL_DEFAULTS[key] || {};
  var ov = (draftOv !== undefined && draftOv !== null) ? cleanTpl_(draftOv) : ((emailStore_().templates || {})[key] || {});
  Object.keys(d).forEach(function (k) { o[k] = d[k]; });
  Object.keys(ov).forEach(function (k) { if (ov[k] !== undefined && ov[k] !== null && ov[k] !== '') o[k] = ov[k]; });
  return o;
}
function baseUrl_() { return ScriptApp.getService().getUrl(); }
function link_(id) { return baseUrl_() + '?issue=' + encodeURIComponent(id); }
function trackLink_(id) { return baseUrl_() + '?track=' + encodeURIComponent(id); }
function fill_(str, v) { return String(str == null ? '' : str).replace(/\{\{(\w+)\}\}/g, function (_, k) { return v[k] == null ? '' : String(v[k]); }); }

function vars_(i, extra) {
  var v = {
    id: i.id, project: i.project, description: i.description, reporter: i.reporter, section: i.section, category: i.category,
    priority: i.priority, priorityLabel: PRIORITY_TH[i.priority] || i.priority,
    status: i.status, statusLabel: STATUS_TH[i.status] || i.status, statusIntro: STATUS_INTRO[i.status] || '',
    assignee: i.assignee || 'ยังไม่มีผู้รับผิดชอบ', reportedDate: i.reportedDate, closedDate: i.closedDate || '',
    link: link_(i.id), trackLink: trackLink_(i.id), urgentTag: i.priority === 'Urgent' ? ' (ด่วนมาก)' : ''
  };
  Object.keys(extra || {}).forEach(function (k) { v[k] = extra[k]; });
  return v;
}

// ── HTML building blocks (table layout + inline CSS = works in Gmail / Outlook / mobile) ──
function pill_(text, bg, fg) {
  return '<span style="display:inline-block;padding:3px 10px;border-radius:999px;background:' + bg + ';color:' + fg + ';font-size:12pt;font-weight:700;">' + esc_(text) + '</span>';
}
function button_(label, url, primary) {
  var C = EMAIL_CFG;
  return '<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="display:inline-block;margin:4px 6px 4px 0;"><tr>' +
    '<td bgcolor="' + (primary ? C.brand : '#ffffff') + '" style="border-radius:10px;' + (primary ? '' : 'border:1px solid ' + C.line + ';') + '">' +
    '<a href="' + esc_(url) + '" target="_blank" style="display:inline-block;padding:13px 24px;font-size:16pt;font-weight:700;color:' + (primary ? '#ffffff' : C.ink) +
    ';text-decoration:none;border-radius:10px;font-family:inherit;">' + esc_(label) + '</a></td></tr></table>';
}
function stepper_(status) {
  var idx = STATUS_ORDER.indexOf(status), C = EMAIL_CFG, bar = '', dots = '';
  STATUS_ORDER.forEach(function (s, n) {
    var done = n < idx || status === 'Closed', now = n === idx && status !== 'Closed';
    var col = done ? '#059669' : (now ? C.brand : '#e6dfd8');
    bar += '<td height="5" bgcolor="' + col + '" style="font-size:0;line-height:0;' + (n ? 'border-left:2px solid #ffffff;' : '') + '">&nbsp;</td>';
    dots += '<td align="center" width="20%" style="padding-top:8px;font-size:11pt;line-height:1.3;color:' + (done || now ? C.ink : '#9a968d') + ';font-weight:' + (now ? '700' : '500') + ';">' +
      (done ? '✓ ' : (now ? '● ' : '')) + esc_(STATUS_TH[s]) + '</td>';
  });
  return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:18px 0 6px;"><tr>' + bar + '</tr><tr>' + dots + '</tr></table>';
}
function infoRow_(label, valueHtml) {
  return '<tr><td style="padding:8px 12px 8px 0;width:150px;font-size:16pt;color:' + EMAIL_CFG.muted + ';vertical-align:top;">' + esc_(label) +
    '</td><td style="padding:8px 0;font-size:16pt;line-height:1.5;color:' + EMAIL_CFG.ink + ';vertical-align:top;">' + valueHtml + '</td></tr>';
}
function caseCard_(i) {
  var C = EMAIL_CFG, pc = PRIORITY_COL[i.priority] || PRIORITY_COL.Medium, sc = STATUS_COL[i.status] || '#6b7280';
  var rows = infoRow_('ความสำคัญ', pill_(PRIORITY_TH[i.priority] || i.priority, pc[0], pc[1])) +
    infoRow_('สถานะ', pill_(STATUS_TH[i.status] || i.status, sc, '#ffffff')) +
    infoRow_('แผนก / หมวด', esc_(i.section) + ' · ' + esc_(i.category)) +
    infoRow_('ผู้แจ้ง', esc_(i.reporter)) +
    infoRow_('แจ้งเมื่อ', esc_(i.reportedDate)) +
    infoRow_('ผู้รับผิดชอบ', i.assignee ? esc_(i.assignee) : '<span style="color:#9a968d;">ยังไม่มีผู้รับผิดชอบ</span>') +
    (i.closedDate ? infoRow_('ปิดงานเมื่อ', esc_(i.closedDate)) : '');
  var desc = String(i.description || ''); if (desc.length > 600) desc = desc.substring(0, 600) + '…';
  return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border:1px solid ' + C.line + ';border-radius:14px;background:#faf9f5;margin:18px 0;"><tr><td style="padding:18px 20px;">' +
    '<div style="font-size:14pt;font-weight:700;color:' + C.brand + ';letter-spacing:.04em;">#' + esc_(i.id) + '</div>' +
    '<div style="font-size:18pt;font-weight:700;color:' + C.ink + ';line-height:1.4;margin:4px 0 10px;">' + esc_(i.project) + '</div>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-top:1px solid ' + C.line + ';">' + rows + '</table>' +
    '<div style="margin-top:10px;padding:12px 14px;background:#ffffff;border:1px solid ' + C.line + ';border-left:4px solid ' + C.brand + ';border-radius:8px;font-size:16pt;line-height:1.65;color:#3d3d3a;white-space:pre-line;">' + esc_(desc) + '</div>' +
    '</td></tr></table>';
}
function emailHtml_(o) {
  var C = EMAIL_CFG, btns = (o.buttons || []).map(function (b, n) { return button_(b.label, b.url, n === 0); }).join('');
  return '<!DOCTYPE html><html lang="th"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>' + esc_(o.title || '') + '</title></head>' +
    '<body style="margin:0;padding:0;background:' + C.bg + ';font-family:\'Segoe UI\',\'Sarabun\',Tahoma,Arial,sans-serif;">' +
    '<div style="display:none;max-height:0;overflow:hidden;opacity:0;color:transparent;">' + esc_(o.preheader || '') + '</div>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="' + C.bg + '"><tr><td align="center" style="padding:24px 12px;">' +
    '<table role="presentation" width="680" cellpadding="0" cellspacing="0" border="0" style="width:100%;max-width:680px;background:#ffffff;border-radius:18px;overflow:hidden;border:1px solid ' + C.line + ';">' +
    '<tr><td bgcolor="' + C.brand + '" style="padding:20px 28px;"><table role="presentation" width="100%"><tr>' +
    '<td style="font-size:18pt;font-weight:800;color:#ffffff;">✽ ' + esc_(C.product) + '</td>' +
    '<td align="right" style="font-size:11pt;color:#ffe9e0;letter-spacing:.06em;">' + esc_(C.company) + '</td></tr></table></td></tr>' +
    '<tr><td style="padding:28px 28px 8px;">' +
    '<div style="font-size:18pt;font-weight:800;color:' + C.ink + ';line-height:1.35;">' + esc_(o.heading) + '</div>' +
    '<div style="font-size:16pt;color:#3d3d3a;line-height:1.7;margin-top:8px;">' + esc_(o.intro) + '</div>' +
    (o.stepper && o.issue ? stepper_(o.issue.status) : '') +
    (o.message ? '<div style="margin-top:16px;padding:14px 16px;background:#eff6ff;border:1px solid #bfdbfe;border-radius:12px;font-size:14pt;line-height:1.7;color:#1e3a8a;white-space:pre-line;"><b style="display:block;font-size:16pt;margin-bottom:4px;">' + esc_(o.messageLabel || 'ข้อความ') + '</b>' + esc_(o.message) + '</div>' : '') +
    (o.issue && o.showCard !== false ? caseCard_(o.issue) : '') +
    (o.extraHtml || '') +
    '<div style="margin:6px 0 4px;">' + btns + '</div>' +
    '</td></tr>' +
    '<tr><td style="padding:18px 28px 26px;border-top:1px solid ' + C.line + ';font-size:11pt;line-height:1.7;color:' + C.muted + ';">' +
    esc_(o.footer || EMAIL_CFG.footer || EMAIL_FOOTER_DEFAULT) +
    '<br>© ' + esc_(C.company) + '</td></tr></table></td></tr></table></body></html>';
}
function emailText_(o, v) {
  var L = [o.heading, '', o.intro, ''];
  if (o.issue) L.push('เคส: #' + o.issue.id + ' — ' + o.issue.project, 'สถานะ: ' + (STATUS_TH[o.issue.status] || o.issue.status), 'ความสำคัญ: ' + (PRIORITY_TH[o.issue.priority] || o.issue.priority), '');
  if (o.message) L.push((o.messageLabel || 'ข้อความ') + ': ' + o.message, '');
  (o.buttons || []).forEach(function (b) { L.push(b.label + ': ' + b.url); });
  return L.join('\n');
}

// Build one e-mail: {subject, html, text, enabled, extraTo, extraCc}. draft = {tpl, brand} from the admin editor (unsaved).
function renderEmail_(key, issue, extra, draft) {
  applyBrand_(draft ? draft.brand : undefined);
  var t = tpl_(key, draft ? draft.tpl : undefined), v = vars_(issue, extra), isAdminMail = /_admin$/.test(key);
  var buttons = [{ label: fill_(t.button, v), url: isAdminMail ? v.link : v.trackLink }];
  if (t.button2 && isAdminMail) buttons.push({ label: fill_(t.button2, v), url: v.trackLink });
  var o = {
    title: fill_(t.subject, v), preheader: fill_(t.intro, v), heading: fill_(t.heading, v), intro: fill_(t.intro, v), issue: issue,
    stepper: t.showStepper === true, showCard: t.showCard !== false,
    message: v.message || '', messageLabel: v.messageLabel || '', buttons: buttons, footer: t.footer ? fill_(t.footer, v) : ''
  };
  return { subject: fill_(t.subject, v).replace(/[\r\n]+/g, ' '), html: emailHtml_(o), text: emailText_(o, v), enabled: t.enabled !== false,
    extraTo: emailList_(t.extraTo), extraCc: emailList_(t.extraCc) };
}

function logMail_(ev, id, to, subject, result, cc) {
  try {
    var s = sheet_('EmailLog', ['Time', 'Event', 'Case', 'To', 'Subject', 'Result', 'Cc']);
    if (s.getMaxColumns() < 7) s.insertColumnsAfter(s.getMaxColumns(), 7 - s.getMaxColumns());
    if (!s.getRange(1, 7).getValue()) s.getRange(1, 7).setValue('Cc');
    s.appendRow([nowStr_(), ev, id, to, subject, result, cc || '']);
    if (s.getLastRow() > 1500) s.deleteRows(2, 500);
  } catch (e) { console.warn(e); }
}
// Recipients: Settings → Email Notifications (adminEmails = To, ccEmails = Cc) are always the base;
// the case's own e-mail (Issues column "Reporter Email"), when filled in, is added on top.
function uniq_(list) {
  var seen = {}, out = [];
  (list || []).forEach(function (x) { x = String(x || '').trim(); var k = x.toLowerCase(); if (/@/.test(x) && !seen[k]) { seen[k] = 1; out.push(x); } });
  return out;
}
function recipients_(issueEmail) {
  var to = uniq_(emails_('adminEmails').concat(String(issueEmail || '').split(/[,;\s]+/))), have = {};
  to.forEach(function (x) { have[x.toLowerCase()] = 1; });
  var cc = uniq_(emails_('ccEmails')).filter(function (x) { return !have[x.toLowerCase()]; });
  return { to: to, cc: cc };
}
function sendMail_(ev, id, to, cc, m) {
  if (!m.enabled) { return; }
  // Recipients come ONLY from Settings → Email Notifications (adminEmails / ccEmails) + the case's reporter e-mail.
  to = uniq_(to || []);
  var have = {}; to.forEach(function (x) { have[x.toLowerCase()] = 1; });
  cc = uniq_(cc || []).filter(function (x) { return !have[x.toLowerCase()]; });
  var ccs = cc.join(',');
  if (!to.length) { logMail_(ev, id, '', m.subject, 'ข้าม: ไม่มีอีเมลผู้รับ', ccs); return; }
  try {
    if (MailApp.getRemainingDailyQuota() < 1) { logMail_(ev, id, to.join(','), m.subject, 'ไม่ส่ง: โควตาอีเมลรายวันหมด', ccs); return; }
    var opts = { to: to.join(','), cc: cc.join(','), subject: m.subject, htmlBody: m.html, body: m.text, name: EMAIL_CFG.senderName };
    var rp = emails_('adminEmails')[0]; if (rp) opts.replyTo = rp;
    MailApp.sendEmail(opts);
    logMail_(ev, id, to.join(','), m.subject, 'ส่งแล้ว', ccs);
  } catch (e) { console.warn(e); logMail_(ev, id, to.join(','), m.subject, 'ผิดพลาด: ' + e.message, ccs); }
}

// ── events ──
function notifyNew_(i) {
  var r = recipients_('');
  sendMail_('new_admin', i.id, r.to, r.cc, renderEmail_('new_admin', i));
  if (/@/.test(i.reporterEmail || '')) {            // confirmation to the person who filed it (admins already got new_admin)
    var rr = recipients_(i.reporterEmail), adm = {};
    emails_('adminEmails').forEach(function (x) { adm[x.toLowerCase()] = 1; });
    sendMail_('new_reporter', i.id, rr.to.filter(function (x) { return !adm[x.toLowerCase()]; }), [], renderEmail_('new_reporter', i));
  }
}
// cur = stored issue before the change, m = issue after the change
function notifyChange_(cur, m) {
  var statusChanged = cur.status !== m.status, assignChanged = !!m.assignee && cur.assignee !== m.assignee;
  var ex = { oldStatusLabel: STATUS_TH[cur.status] || cur.status };
  if (statusChanged) {
    var rs = recipients_(m.reporterEmail);
    sendMail_('status_reporter', m.id, rs.to, rs.cc, renderEmail_('status_reporter', m, ex));
  } else if (assignChanged) {
    var ra = recipients_(m.reporterEmail);
    sendMail_('assigned_reporter', m.id, ra.to, ra.cc, renderEmail_('assigned_reporter', m));
  }
  var replies = m.chat.slice(cur.chat.length).filter(function (c) { return c.sender === 'Admin' && !c.type; });
  if (replies.length) {
    var msg = replies.map(function (c) { return c.message; }).join('\n\n');
    var rq = recipients_(m.reporterEmail);
    sendMail_('reply_reporter', m.id, rq.to, rq.cc, renderEmail_('reply_reporter', m, { message: msg, messageLabel: 'ข้อความจากทีมงาน' }));
  }
}
function dateOf_(s) {
  var m = /(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(String(s)); if (!m) return null;
  var y = +m[3]; if (y > 2400) y -= 543;
  return new Date(y, +m[2] - 1, +m[1]);
}
function dailyReminder() {                     // run by a daily trigger
  var days = Number((settings_().admin || {}).reminderDays) || 7, now = Date.now(), list = [];
  getIssues('').forEach(function (r) {
    var d = dateOf_(r[7]);
    if (String(r[13]) === '1') return;            // skip archived
    if (r[8] !== 'Resolved' && r[8] !== 'Closed' && d && (now - d.getTime()) / 864e5 >= days)
      list.push({ id: r[0], project: r[1], status: r[8], priority: r[10], age: Math.floor((now - d.getTime()) / 864e5), assignee: r[14] || '' });
  });
  if (!list.length) return;
  sendMail_('reminder_admin', '', emails_('adminEmails'), emails_('ccEmails'), buildReminder_(list, days));
}
function buildReminder_(list, days, draft) {   // shared by the daily trigger and the admin preview
  applyBrand_(draft ? draft.brand : undefined);
  list.sort(function (a, b) { return b.age - a.age; });
  var C = EMAIL_CFG, rows = list.slice(0, 40).map(function (x) {
    var pc = PRIORITY_COL[x.priority] || PRIORITY_COL.Medium;
    return '<tr><td style="padding:9px 8px;border-bottom:1px solid ' + C.line + ';font-size:16pt;"><a href="' + esc_(link_(x.id)) + '" style="color:' + C.brand + ';font-weight:700;text-decoration:none;">' + esc_(x.id) + '</a><br><span style="color:#3d3d3a;">' + esc_(x.project) + '</span></td>' +
      '<td style="padding:9px 8px;border-bottom:1px solid ' + C.line + ';font-size:14pt;">' + pill_(STATUS_TH[x.status] || x.status, '#f5f0e8', '#3d3d3a') + '<br><span style="color:' + C.muted + ';">' + esc_(x.assignee || 'ยังไม่มีผู้รับผิดชอบ') + '</span></td>' +
      '<td align="right" style="padding:9px 8px;border-bottom:1px solid ' + C.line + ';font-size:16pt;font-weight:700;color:' + pc[1] + ';">' + x.age + ' วัน</td></tr>';
  }).join('');
  var table = '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:16px 0;border:1px solid ' + C.line + ';border-radius:12px;"><tr bgcolor="#faf9f5"><td style="padding:8px;font-size:12pt;color:' + C.muted + ';">เคส</td><td style="padding:8px;font-size:12pt;color:' + C.muted + ';">สถานะ / ผู้รับผิดชอบ</td><td align="right" style="padding:8px;font-size:12pt;color:' + C.muted + ';">อายุเคส</td></tr>' + rows + '</table>' +
    (list.length > 40 ? '<div style="font-size:12pt;color:' + C.muted + ';">แสดง 40 รายการแรกจากทั้งหมด ' + list.length + ' รายการ</div>' : '');
  var t = tpl_('reminder_admin', draft ? draft.tpl : undefined), v = { days: days, count: list.length }, subject = fill_(t.subject, v);
  var html = emailHtml_({ title: subject, preheader: fill_(t.intro, v), heading: fill_(t.heading, v), intro: fill_(t.intro, v), extraHtml: table,
    footer: t.footer ? fill_(t.footer, v) : '', buttons: [{ label: fill_(t.button, v), url: baseUrl_() }] });
  return { enabled: t.enabled !== false, subject: subject, html: html, text: subject + '\n' + baseUrl_(),
    extraTo: emailList_(t.extraTo), extraCc: emailList_(t.extraCc) };
}

// Admin: send a sample of every e-mail design to one address (Settings → ส่งอีเมลทดสอบ)
function sampleIssue_() {
  return { id: 'ISS-2026999', project: 'เครื่องพิมพ์ฉลากบาร์โค้ดไม่ตอบสนอง', reporter: 'สมชาย ใจดี', reporterEmail: '', section: 'MF', category: 'ระบบไอที',
    priority: 'Urgent', status: 'In Progress', reportedDate: '09/10/2026 09:15:00', closedDate: '', assignee: 'Neung',
    description: 'เครื่องพิมพ์ฉลากที่สายการผลิต 2 ไม่ตอบสนองตั้งแต่เช้า ลองรีสตาร์ตแล้วยังไม่หาย กระทบการติดฉลากตู้ควบคุม\nรบกวนตรวจสอบด่วนครับ', chat: [] };
}
function sendTestEmails(token, to) {
  requireAdmin_(token);
  to = String(to || '').trim();
  if (!/^[^@\s]+@[^@\s]+$/.test(to)) throw new Error('รูปแบบอีเมลไม่ถูกต้อง');
  var i = sampleIssue_(), out = [], items = [
    ['new_admin', i, {}], ['new_reporter', i, {}], ['status_reporter', i, { oldStatusLabel: 'ใหม่' }],
    ['assigned_reporter', i, {}], ['reply_reporter', i, { message: 'รับเรื่องแล้วครับ ช่างจะเข้าไปตรวจสอบเครื่องพิมพ์ภายในวันนี้', messageLabel: 'ข้อความจากทีมงาน' }]];
  items.forEach(function (it) {
    var m = renderEmail_(it[0], it[1], it[2]); m.enabled = true; m.extraTo = []; m.extraCc = []; m.subject = '[ตัวอย่าง] ' + m.subject;
    sendMail_('test:' + it[0], i.id, [to], [], m); out.push('✓ ' + it[0]);
  });
  var c = Object.assign({}, i, { status: 'Resolved', closedDate: '09/10/2026 14:30:00' });
  var m2 = renderEmail_('status_reporter', c, { oldStatusLabel: 'กำลังแก้ไข' }); m2.extraTo = []; m2.extraCc = []; m2.subject = '[ตัวอย่าง] ' + m2.subject; sendMail_('test:status_resolved', c.id, [to], [], m2); out.push('✓ status_reporter (Resolved)');
  out.push('ส่งไปที่ ' + to + ' — ดูประวัติได้ที่แท็บ EmailLog');
  return out;
}
// ── Admin e-mail editor API (Settings → แก้ไขรูปแบบอีเมล) ─────────────────
function getEmailConfig(token) {
  requireAdmin_(token);
  var general = EMAIL_VARS.slice();
  var keys = Object.keys(EMAIL_DEFAULTS).map(function (k) {
    var meta = EMAIL_META[k] || {}, vars = general.slice();
    if (k === 'reminder_admin') vars = [['days', 'จำนวนวันที่ค้าง'], ['count', 'จำนวนเคสที่ค้าง']];
    else if (k === 'status_reporter' || k === 'status_admin') vars.push(['oldStatusLabel', 'สถานะเดิม (ไทย)']);
    return { key: k, label: meta.label || k, desc: meta.desc || '', inactive: !!meta.inactive,
      adminMail: /_admin$/.test(k), reminder: k === 'reminder_admin', vars: vars, defaults: EMAIL_DEFAULTS[k] };
  });
  return { keys: keys, overrides: emailStore_(), defaultBrand: EMAIL_BASE, defaultFooter: EMAIL_FOOTER_DEFAULT, roster: emailStore_().roster };
}
function saveEmailTemplates(token, payload) {
  requireAdmin_(token);
  return withLock_(function () {
    var T = {}, src = (payload && payload.templates) || {};
    Object.keys(EMAIL_DEFAULTS).forEach(function (k) {
      var o = cleanTpl_(src[k]);
      if (Object.keys(o).length) T[k] = o;
    });
    var js = JSON.stringify({ templates: T, brand: cleanBrand_(payload && payload.brand), roster: emailStore_().roster });
    if (js.length > 45000) throw new Error('ข้อความเทมเพลตยาวเกินไป (' + js.length + ' ตัวอักษร)');
    sheet_(CONFIG.SETTINGS).getRange('B1').setNumberFormat('@').setValue(js);
    __emailStore = null;
    audit_([['Admin', 'แก้ไขแม่แบบอีเมล', '', Object.keys(T).join(', ') || '(คืนค่าเริ่มต้น)']]);
    return true;
  });
}
// Recipient roster (single list used by every To/Cc picker; selection itself is stored in Settings → admin.adminEmails / ccEmails)
function getEmailRoster(token) { requireAdmin_(token); return emailStore_().roster; }
function saveEmailRoster(token, roster) {
  requireAdmin_(token);
  return withLock_(function () {
    __emailStore = null;
    var st = emailStore_(), R = cleanRoster_(roster);
    var js = JSON.stringify({ templates: st.templates, brand: st.brand, roster: R });
    if (js.length > 45000) throw new Error('ข้อมูลยาวเกินไป');
    sheet_(CONFIG.SETTINGS).getRange('B1').setNumberFormat('@').setValue(js);
    __emailStore = null;
    audit_([['Admin', 'แก้ไขรายชื่ออีเมล', '', R.length + ' รายชื่อ']]);
    return R;
  });
}
// Builds a sample mail for the editor: draft = {tpl, brand} (unsaved), opts = {status}
function buildSample_(key, draft, opts) {
  if (!EMAIL_DEFAULTS[key]) throw new Error('ไม่รู้จักเทมเพลต: ' + key);
  draft = draft || {}; opts = opts || {};
  var d = { tpl: (draft.tpl && typeof draft.tpl === 'object') ? draft.tpl : {}, brand: draft.brand || {} };
  if (key === 'reminder_admin') {
    return buildReminder_([
      { id: 'ISS-2026012', project: 'Wiring ผิดตำแหน่งที่ตู้ควบคุม', status: 'Waiting', priority: 'High', age: 12, assignee: '' },
      { id: 'ISS-2026031', project: 'เครื่องพิมพ์ฉลากบาร์โค้ดไม่ตอบสนอง', status: 'In Progress', priority: 'Urgent', age: 9, assignee: 'Neung' },
      { id: 'ISS-2026044', project: 'ขอปรับปรุงชั้นวางอุปกรณ์', status: 'New', priority: 'Low', age: 7, assignee: '' }], 7, d);
  }
  var i = sampleIssue_();
  if (opts.status && STATUS_TH[opts.status]) i.status = opts.status;
  if (i.status === 'Resolved' || i.status === 'Closed') i.closedDate = '09/10/2026 14:30:00';
  var extra = { oldStatusLabel: 'ใหม่' };
  if (key === 'reply_reporter') { extra.message = 'รับเรื่องแล้วครับ ช่างจะเข้าไปตรวจสอบเครื่องพิมพ์ภายในวันนี้'; extra.messageLabel = 'ข้อความจากทีมงาน'; }
  return renderEmail_(key, i, extra, d);
}
function previewEmail(token, key, draft, opts) {
  requireAdmin_(token);
  var m = buildSample_(key, draft, opts);
  return { subject: m.subject, html: m.html, enabled: m.enabled, extraTo: m.extraTo, extraCc: m.extraCc };
}
function sendTestEmail(token, key, draft, to, opts) {
  requireAdmin_(token);
  to = String(to || '').trim();
  if (!/^[^@\s]+@[^@\s]+$/.test(to)) throw new Error('รูปแบบอีเมลไม่ถูกต้อง');
  var m = buildSample_(key, draft, opts);
  m.enabled = true; m.extraTo = []; m.extraCc = []; m.subject = '[ตัวอย่าง] ' + m.subject;
  sendMail_('test:' + key, 'SAMPLE', [to], [], m);
  return 'ส่งตัวอย่างไปที่ ' + to + ' แล้ว (ดูผลที่แท็บ EmailLog)';
}
// ══ Templates for users: issue templates (Settings!C1) + report template (Settings!D1) ══
function readJsonCell_(a1) {
  try { var sh = ss_().getSheetByName(CONFIG.SETTINGS), raw = sh ? String(sh.getRange(a1).getValue() || '') : ''; return raw ? JSON.parse(raw) : null; }
  catch (e) { return null; }
}
function writeJsonCell_(a1, obj) {
  var js = JSON.stringify(obj);
  if (js.length > 45000) throw new Error('ข้อมูลยาวเกินไป (' + js.length + ' ตัวอักษร)');
  sheet_(CONFIG.SETTINGS).getRange(a1).setNumberFormat('@').setValue(js);
}
function cleanIssueTpls_(list) {
  var P = { Low: 1, Medium: 1, High: 1, Urgent: 1 }, L = function (v, n) { return String(v == null ? '' : v).substring(0, n); };
  return (Array.isArray(list) ? list : []).map(function (t) {
    return { label: L(t && t.label, 60).trim(), title: L(t && t.title, 200), desc: L(t && t.desc, 2000), cat: L(t && t.cat, 80), pri: P[t && t.pri] ? t.pri : 'Medium', section: L(t && t.section, 80) };
  }).filter(function (t) { return t.label; }).slice(0, 40);
}
function cleanReportTpl_(o) {
  o = o || {}; var r = {}, L = function (v, n) { return String(v == null ? '' : v).replace(/[\r\n]+/g, ' ').substring(0, n); };
  ['company', 'title', 'subtitle', 'sign1', 'sign2', 'sign3', 'note', 'docNo'].forEach(function (k) { if (o[k] != null) r[k] = L(o[k], k === 'note' ? 600 : 120); });
  ['showKpi', 'showAnalysis', 'showTable', 'land', 'full', 'logOn', 'showLogo', 'showSign', 'showReporter'].forEach(function (k) { if (typeof o[k] === 'boolean') r[k] = o[k]; });
  return r;
}
function getPublicTemplates() { return { issue: readJsonCell_('C1'), report: readJsonCell_('D1'), categories: readJsonCell_('E1') }; }
function cleanNameList_(list) {
  var seen = {}, out = [];
  (Array.isArray(list) ? list : []).forEach(function (v) {
    var t = String(v == null ? '' : v).replace(/[\r\n,]+/g, ' ').trim().substring(0, 60);
    if (t && !seen[t.toLowerCase()] && out.length < 50) { seen[t.toLowerCase()] = 1; out.push(t); }
  });
  return out;
}
function saveCategories(token, list) {
  requireAdmin_(token);
  return withLock_(function () { var L = cleanNameList_(list); if (!L.length) throw new Error('ต้องมีหมวดหมู่อย่างน้อย 1 รายการ'); writeJsonCell_('E1', L); audit_([['Admin', 'แก้ไขหมวดหมู่ปัญหา', '', L.join(', ')]]); return L; });
}   // read by every user
function saveIssueTemplates(token, list) {
  requireAdmin_(token);
  return withLock_(function () { var L = cleanIssueTpls_(list); writeJsonCell_('C1', L); audit_([['Admin', 'แก้ไขแม่แบบปัญหา', '', L.length + ' รายการ']]); return L; });
}
function saveReportTemplate(token, tpl) {
  requireAdmin_(token);
  return withLock_(function () { var R = cleanReportTpl_(tpl); writeJsonCell_('D1', R); audit_([['Admin', 'แก้ไขแม่แบบรายงาน', '', '']]); return R; });
}

// ══ Audit log (sheet AuditLog) ══
function audit_(rows) {
  if (!rows || !rows.length) return;
  try {
    var s = sheet_('AuditLog', ['Time', 'By', 'Action', 'Case', 'Detail']), t = nowStr_();
    var data = rows.map(function (r) { return [t, r[0], r[1], r[2] || '', String(r[3] || '').substring(0, 500)]; });
    var at = s.getLastRow() + 1; s.getRange(at, 1, data.length, 5).setNumberFormat('@').setValues(data);
    if (s.getLastRow() > 5000) s.deleteRows(2, 1000);
  } catch (e) { console.warn(e); }
}
function auditDiff_(cur, m, admin, out) {
  var by = admin ? 'Admin' : 'User';
  if (cur.status !== m.status) out.push([by, 'เปลี่ยนสถานะ', m.id, cur.status + ' → ' + m.status]);
  if ((cur.assignee || '') !== (m.assignee || '')) out.push([by, 'เปลี่ยนผู้รับผิดชอบ', m.id, (cur.assignee || '-') + ' → ' + (m.assignee || '-')]);
  if (!!cur.archived !== !!m.archived) out.push([by, m.archived ? 'ย้ายเข้า Archive' : 'นำออกจาก Archive', m.id, '']);
  if (cur.priority !== m.priority) out.push([by, 'เปลี่ยนความสำคัญ', m.id, cur.priority + ' → ' + m.priority]);
  var ed = ['project', 'description', 'section', 'category'].filter(function (k) { return String(cur[k] || '') !== String(m[k] || ''); });
  if (ed.length) out.push([by, 'แก้ไขรายละเอียด', m.id, ed.join(', ')]);
  var n = (m.chat || []).length - (cur.chat || []).length;
  if (n > 0) out.push([by, 'ข้อความใหม่', m.id, n + ' ข้อความ']);
}
function getAuditLog(token, limit) {
  requireAdmin_(token);
  var s = ss_().getSheetByName('AuditLog'); if (!s || s.getLastRow() < 2) return [];
  var n = Math.min(Number(limit) || 100, 500), last = s.getLastRow(), from = Math.max(2, last - n + 1);
  return s.getRange(from, 1, last - from + 1, 5).getDisplayValues().reverse();
}

// ══ Backup (daily copy of the whole spreadsheet into Drive, keep 30 days) ══
var BACKUP_KEEP_DAYS = 30;
function backupFolder_() {
  var P = PropertiesService.getScriptProperties(), id = P.getProperty('BACKUP_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var parent = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID), it = parent.getFoldersByName('_Helpdesk_Backups');
  var f = it.hasNext() ? it.next() : parent.createFolder('_Helpdesk_Backups');
  P.setProperty('BACKUP_FOLDER_ID', f.getId());
  return f;
}
function backupRun_(why) {
  var folder = backupFolder_(), stamp = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd_HHmm');
  var copy = DriveApp.getFileById(CONFIG.SPREADSHEET_ID).makeCopy('Helpdesk_Backup_' + stamp, folder);
  var cut = Date.now() - BACKUP_KEEP_DAYS * 864e5, it = folder.getFiles(), removed = 0;
  while (it.hasNext()) { var f = it.next(); if (f.getId() !== copy.getId() && f.getDateCreated().getTime() < cut) { f.setTrashed(true); removed++; } }
  var info = { time: nowStr_(), name: copy.getName(), url: copy.getUrl(), why: why || 'อัตโนมัติรายวัน', removed: removed };
  PropertiesService.getScriptProperties().setProperty('LAST_BACKUP', JSON.stringify(info));
  audit_([['System', 'สำรองข้อมูล', '', info.why + ' · ' + info.name]]);
  return info;
}
function dailyBackup() { backupRun_('อัตโนมัติรายวัน'); }   // run by a daily trigger
function backupNow(token) { requireAdmin_(token); return backupRun_('สั่งสำรองโดย Admin'); }
function listBackups(token) {
  requireAdmin_(token);
  var it = backupFolder_().getFiles(), out = [];
  while (it.hasNext()) { var f = it.next(); out.push({ name: f.getName(), url: f.getUrl(), t: f.getDateCreated().getTime() }); }
  return out.sort(function (a, b) { return b.t - a.t; }).slice(0, 40);
}


// ══ Schedules auto-sync from a OneDrive / SharePoint "Anyone with the link" CSV ══
var SCHED_HEADER = ['ProjectNo', 'MVA', 'RatingVolt', 'Purchaser', 'Customer', 'DeliveryByPO', 'NewDelivery', 'ActualPlanDelivery', 'Process', 'Start-date', 'Finish-date'];
function schedProp_(k, v) { var P = PropertiesService.getScriptProperties(); if (v === undefined) return P.getProperty(k); if (v === null) P.deleteProperty(k); else P.setProperty(k, v); }
function schedSyncInfo_() { try { return JSON.parse(schedProp_('SCHED_SYNC_LAST') || 'null'); } catch (e) { return null; } }
// SharePoint/OneDrive "Anyone" links answer with a redirect that sets a FedAuth cookie; UrlFetchApp's automatic
// redirect-following drops cookies and lands on the login page. So redirects are followed manually with a cookie jar,
// trying the direct-download forms of the link first.
function shareCandidates_(u) {
  u = String(u || '').trim();
  var out = [], m = /^(https:\/\/[^\/]+)\/:\w:\/(?:[a-z]\/)?(personal\/[^\/]+|sites\/[^\/]+)\/([^?#\/]+)/i.exec(u);
  if (m) {
    out.push(m[1] + '/' + m[2] + '/_layouts/15/download.aspx?share=' + m[3]);
    out.push(m[1] + '/:u:/g/' + m[2] + '/' + m[3] + '?download=1');
  }
  var base = u.replace(/([?&])nav=[^&]*/g, '$1').replace(/[?&]+$/, '');
  out.push(base + (base.indexOf('?') >= 0 ? '&' : '?') + 'download=1');
  return out;
}
function fetchWithCookies_(url) {
  var jar = {}, trail = [];
  for (var hop = 0; hop < 10; hop++) {
    var ck = Object.keys(jar).map(function (k) { return k + '=' + jar[k]; }).join('; ');
    var hd = { 'User-Agent': 'Mozilla/5.0 (HelpdeskSync)' }; if (ck) hd['Cookie'] = ck;
    var res = UrlFetchApp.fetch(url, { followRedirects: false, muteHttpExceptions: true, headers: hd });
    var code = res.getResponseCode(), h = res.getAllHeaders ? res.getAllHeaders() : res.getHeaders();
    var sc = h['Set-Cookie'] || h['set-cookie'];
    if (sc) [].concat(sc).forEach(function (c) { var p = String(c).split(';')[0], i = p.indexOf('='); if (i > 0) jar[p.substring(0, i).trim()] = p.substring(i + 1).trim(); });
    trail.push(code);
    if (code >= 300 && code < 400) {
      var loc = h['Location'] || h['location']; if (!loc) break;
      loc = [].concat(loc)[0];
      if (/^\//.test(loc)) loc = url.match(/^https:\/\/[^\/]+/)[0] + loc;
      url = loc; continue;
    }
    return { res: res, code: code, trail: trail.join('→') };
  }
  return { res: null, code: 0, trail: trail.join('→') };
}
function downloadShared_(u) {
  var tries = shareCandidates_(u), notes = [];
  for (var i = 0; i < tries.length; i++) {
    try {
      var r = fetchWithCookies_(tries[i]);
      if (!r.res || r.code !== 200) { notes.push('#' + (i + 1) + ' HTTP ' + r.trail); continue; }
      var blob = r.res.getBlob(), text = blob.getDataAsString('UTF-8');
      if (text.indexOf('\uFFFD') >= 0) text = blob.getDataAsString('windows-874');
      text = text.replace(/^\uFEFF/, '');
      if (/^\s*</.test(text)) { notes.push('#' + (i + 1) + ' ได้หน้าเว็บ (' + r.trail + ')'); continue; }
      return text;
    } catch (e) { notes.push('#' + (i + 1) + ' ' + e.message); }
  }
  throw new Error('ดาวน์โหลดไฟล์ไม่ได้ — ตรวจว่าแชร์แบบ "Anyone with the link" และลิงก์ยังไม่หมดอายุ [' + notes.join(' | ') + ']');
}
function syncSchedulesRun_(why, force) {
  var url = schedProp_('SCHED_SRC_URL'), info = { time: nowStr_(), why: why };
  if (!url) throw new Error('ยังไม่ได้ตั้งค่าลิงก์ OneDrive');
  try {
    var text = downloadShared_(url);
    var hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, text));
    var prev = schedSyncInfo_();
    if (!force && prev && prev.ok && prev.hash === hash) {
      info.ok = true; info.unchanged = true; info.rows = prev.rows; info.projects = prev.projects; info.hash = hash;
      info.msg = 'ไฟล์ไม่มีการเปลี่ยนแปลง';
      schedProp_('SCHED_SYNC_LAST', JSON.stringify(info)); return info;
    }
    var data = Utilities.parseCsv(text);
    var h = data[0] || [], norm = function (x) { return String(x || '').toLowerCase().replace(/[^a-z]/g, ''); };
    if (norm(h[0]) !== 'projectno' || norm(h[8]) !== 'process' || h.length < 11) throw new Error('หัวคอลัมน์ไม่ตรงรูปแบบ ต้องเป็น: ' + SCHED_HEADER.join(', '));
    var dateCols = [5, 6, 7, 9, 10], projects = {}, zero = 0;
    var rows = data.slice(1).map(function (r) {
      var o = []; for (var c = 0; c < 11; c++) o.push(String(r[c] == null ? '' : r[c]).trim());
      dateCols.forEach(function (c) { var m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(o[c]); if (m && +m[3] < 1950) { o[c] = ''; zero++; } });
      if (o[0]) projects[o[0]] = 1;
      return o;
    }).filter(function (o) { return o.some(function (v) { return v; }); });
    if (!rows.length) throw new Error('ไฟล์ไม่มีข้อมูล');
    withLock_(function () {
      var sh = sheet_(CONFIG.SCHEDULES), all = [SCHED_HEADER].concat(rows);
      sh.clearContents();
      ensureRows_(sh, all.length);
      var rg = sh.getRange(1, 1, all.length, 11); fmtText_(rg); rg.setValues(all);
      var extra = sh.getMaxRows() - all.length; if (extra > 200) sh.deleteRows(all.length + 101, extra - 100);
    });
    info.ok = true; info.rows = rows.length; info.projects = Object.keys(projects).length; info.zeroDates = zero; info.hash = hash;
    info.msg = 'อัปเดต ' + rows.length + ' แถว (' + info.projects + ' Project)';
    audit_([['System', 'ซิงค์ตารางงานจาก OneDrive', '', info.why + ' · ' + info.msg]]);
  } catch (e) {
    info.ok = false; info.msg = e.message || String(e);
    if (info.why !== 'อัตโนมัติ') audit_([['System', 'ซิงค์ตารางงานไม่สำเร็จ', '', info.msg]]);
  }
  schedProp_('SCHED_SYNC_LAST', JSON.stringify(info));
  return info;
}
function scheduleAutoSync() { if (schedProp_('SCHED_SRC_URL')) syncSchedulesRun_('อัตโนมัติ', false); }   // time trigger
function getScheduleSource(token) {
  requireAdmin_(token);
  var hrs = 0; ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'scheduleAutoSync') hrs = Number(schedProp_('SCHED_SYNC_HOURS') || 1); });
  return { url: schedProp_('SCHED_SRC_URL') || '', hours: hrs, last: schedSyncInfo_() };
}
function setScheduleSource(token, url, hours) {
  requireAdmin_(token);
  url = String(url || '').trim();
  if (url && !/^https:\/\/[^\s]+$/i.test(url)) throw new Error('ลิงก์ต้องขึ้นต้นด้วย https://');
  if (/_layouts\/15\/onedrive\.aspx|\/Forms\/AllItems\.aspx/i.test(url)) throw new Error('นี่คือลิงก์หน้าโฟลเดอร์ OneDrive (จากแถบที่อยู่เบราว์เซอร์) ไม่ใช่ลิงก์แชร์ไฟล์ — คลิกขวาที่ไฟล์ .csv → Share → Copy link แล้ววางลิงก์นั้นแทน');
  schedProp_('SCHED_SRC_URL', url || null);
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'scheduleAutoSync') ScriptApp.deleteTrigger(t); });
  hours = Number(hours) || 0;
  if (url && hours > 0) {
    var tb = ScriptApp.newTrigger('scheduleAutoSync').timeBased();
    if (hours >= 24) tb.everyDays(1).atHour(6).create(); else tb.everyHours([1, 2, 4, 6, 8, 12].indexOf(hours) >= 0 ? hours : 1).create();
    schedProp_('SCHED_SYNC_HOURS', String(hours));
  }
  audit_([['Admin', 'ตั้งค่าซิงค์ตารางงาน', '', (url ? 'ตั้งลิงก์แล้ว' : 'ลบลิงก์') + ' · ' + (hours ? 'ทุก ' + hours + ' ชม.' : 'ปิดอัตโนมัติ')]]);
  return getScheduleSource(token);
}
function syncSchedulesNow(token) { requireAdmin_(token); return syncSchedulesRun_('สั่งโดย Admin', true); }

// ══ System health (Admin → สุขภาพระบบ) ══
function installTriggers(token) {
  requireAdmin_(token);
  var have = {};
  ScriptApp.getProjectTriggers().forEach(function (t) { have[t.getHandlerFunction()] = 1; });
  var made = [];
  if (!have.dailyReminder) { ScriptApp.newTrigger('dailyReminder').timeBased().everyDays(1).atHour(8).create(); made.push('dailyReminder (08:00)'); }
  if (!have.dailyBackup) { ScriptApp.newTrigger('dailyBackup').timeBased().everyDays(1).atHour(2).create(); made.push('dailyBackup (02:00)'); }
  audit_([['Admin', 'ติดตั้ง Trigger', '', made.join(', ') || 'มีครบแล้ว']]);
  return made;
}
function getSystemHealth(token) {
  requireAdmin_(token);
  var h = { checkedAt: nowStr_(), warn: [] };
  try { h.mailQuota = MailApp.getRemainingDailyQuota(); } catch (e) { h.mailQuota = null; }
  var trig = {}; try { ScriptApp.getProjectTriggers().forEach(function (t) { trig[t.getHandlerFunction()] = 1; }); } catch (e) {}
  h.triggers = { dailyReminder: !!trig.dailyReminder, dailyBackup: !!trig.dailyBackup };
  try { h.lastBackup = JSON.parse(PropertiesService.getScriptProperties().getProperty('LAST_BACKUP') || 'null'); } catch (e) { h.lastBackup = null; }
  var book = ss_(), is = book.getSheetByName(CONFIG.ISSUES);
  h.issueRows = is ? Math.max(0, is.getLastRow() - 1) : 0;
  h.cells = 0; book.getSheets().forEach(function (x) { h.cells += x.getMaxRows() * x.getMaxColumns(); });
  h.mailFail = []; h.mailSent7 = 0;
  var el = book.getSheetByName('EmailLog');
  if (el && el.getLastRow() > 1) {
    var n = Math.min(400, el.getLastRow() - 1), rows = el.getRange(el.getLastRow() - n + 1, 1, n, 6).getDisplayValues(), cut = Date.now() - 7 * 864e5;
    rows.forEach(function (r) {
      var d = dateOf_(r[0]); if (!d || d.getTime() < cut) return;
      if (/^ส่งแล้ว/.test(r[5])) h.mailSent7++; else if (!/^ข้าม/.test(r[5])) h.mailFail.push(r);
    });
    h.mailFail = h.mailFail.reverse().slice(0, 10);
  }
  if (h.mailQuota !== null && h.mailQuota < 20) h.warn.push('โควตาอีเมลวันนี้เหลือน้อย (' + h.mailQuota + ')');
  if (!h.triggers.dailyReminder || !h.triggers.dailyBackup) h.warn.push('Trigger รายวันยังติดตั้งไม่ครบ');
  if (!h.lastBackup) h.warn.push('ยังไม่เคยสำรองข้อมูล');
  else { var lb = dateOf_(h.lastBackup.time); if (lb && Date.now() - lb.getTime() > 2 * 864e5) h.warn.push('สำรองข้อมูลล่าสุดเกิน 2 วันแล้ว'); }
  if (h.mailFail.length) h.warn.push('มีอีเมลส่งไม่สำเร็จ ' + h.mailFail.length + ' ฉบับใน 7 วัน');
  if (h.cells > 8e6) h.warn.push('สเปรดชีตใกล้เต็ม (' + h.cells + ' / 10,000,000 เซลล์)');
  h.schedSync = schedSyncInfo_(); h.schedSrc = !!schedProp_('SCHED_SRC_URL');
  if (h.schedSrc && h.schedSync && !h.schedSync.ok) h.warn.push('ซิงค์ตารางงานจาก OneDrive ล่าสุดไม่สำเร็จ: ' + h.schedSync.msg);
  h.audit = getAuditLog(token, 30);
  return h;
}
function installDailyReminder() {              // run ONCE manually from the editor
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'dailyReminder') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('dailyReminder').timeBased().everyDays(1).atHour(8).create();
}

// ── Schedules CSV import (chunked: temp sheet → swap in one final call) ──
function importScheduleRows(token, header, rows, first, last, mode) {
  requireAdmin_(token);
  return withLock_(function () {
    var TMP = 'Schedules_Import_Tmp', book = ss_(), tmp = book.getSheetByName(TMP);
    if (first && tmp) { book.deleteSheet(tmp); tmp = null; }
    if (!tmp) tmp = book.insertSheet(TMP);
    if (rows.length) {
      var at = tmp.getLastRow() + 1;
      ensureRows_(tmp, at + rows.length);
      var r = tmp.getRange(at, 1, rows.length, 11);
      fmtText_(r); r.setValues(rows);
    }
    if (!last) return { staged: tmp.getLastRow() };
    var all = tmp.getLastRow() ? tmp.getRange(1, 1, tmp.getLastRow(), 11).getValues() : [];
    var main = sheet_(CONFIG.SCHEDULES);
    if (mode === 'replace' || main.getLastRow() < 1) {
      main.clearContents();
      all.unshift(header.slice(0, 11));
      ensureRows_(main, all.length);
      var w = main.getRange(1, 1, all.length, 11);
      fmtText_(w); w.setValues(all);
    } else if (all.length) {
      var start = main.getLastRow() + 1;
      ensureRows_(main, start + all.length);
      var a = main.getRange(start, 1, all.length, 11);
      fmtText_(a); a.setValues(all);
    }
    book.deleteSheet(tmp);
    return { rows: all.length };
  });
}
function ensureRows_(sh, needed) {
  var max = sh.getMaxRows();
  if (max < needed) sh.insertRowsAfter(max, needed - max);
}
// ── Run this ONCE from the Apps Script editor (select it → ▶ Run) to authorize external requests
//    and to see exactly what OneDrive returns for each download method (View → Logs / Execution log).
function testScheduleSync() {
  var url = schedProp_('SCHED_SRC_URL');
  if (!url) { Logger.log('✗ ยังไม่ได้บันทึกลิงก์ในหน้า ตารางงานการผลิต → การ์ด ☁ ซิงค์ตารางงานจาก OneDrive'); return; }
  Logger.log('ลิงก์ที่บันทึกไว้: ' + url);
  shareCandidates_(url).forEach(function (u, i) {
    try {
      var r = fetchWithCookies_(u), head = r.res ? r.res.getContentText().substring(0, 160).replace(/\s+/g, ' ') : '';
      Logger.log('#' + (i + 1) + ' ' + u + '\n   สถานะ: ' + r.trail + '\n   ข้อมูลต้นไฟล์: ' + head);
    } catch (e) { Logger.log('#' + (i + 1) + ' ' + u + '\n   ผิดพลาด: ' + e.message); }
  });
  var res = syncSchedulesRun_('ทดสอบจาก Editor', true);
  Logger.log((res.ok ? '✓ ' : '✗ ') + res.msg);
}

// ══════════════════════════════════════════════════════════════════════
//  ITEM STOCK — separate spreadsheet (auto-created in the attachments folder),
//  synced from a OneDrive "Anyone with the link" CSV (Q_ItemOnhand). Search runs on the
//  server and returns one page at a time, so the browser never loads 190k+ rows.
// ══════════════════════════════════════════════════════════════════════
var STOCK_SHEET = 'Stock';
var STOCK_HEADER = ['ITEMNO', 'DESC', 'COMMENT1', 'QTYONHAND', 'STOCKUNIT', 'LOCATION', 'PICKINGSEQ',
  'LastOfRequest_By', 'LASTCOST', 'LastOfRequisition_No', '_packed'];
var STOCK_NCOL = 10, STOCK_W = 11;                 // 10 data columns + hidden packed search column
var STOCK_COST_COL = 8, STOCK_REQ_COL = 9;         // LASTCOST (admin / cost code) · Requisition No (admin only)
var STOCK_COST_CODE_DEFAULT = 'Daihencost';
var STOCK_ALL_MAX = 1500;                          // ≤ this many matches → return all rows at once
var STOCK_CHUNK = 15000;

function stockSS_() {
  var id = schedProp_('STOCK_SS_ID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  var ss = SpreadsheetApp.create('Helpdesk_ItemStock');
  try { DriveApp.getFileById(ss.getId()).moveTo(DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID)); } catch (e) {}
  schedProp_('STOCK_SS_ID', ss.getId());
  return ss;
}
function stockInfo_() { try { return JSON.parse(schedProp_('STOCK_SYNC_LAST') || 'null'); } catch (e) { return null; } }
function stockMeta_() { try { return JSON.parse(schedProp_('STOCK_META') || 'null'); } catch (e) { return null; } }

// one CSV line → fields (Access export; multi-line records are joined before this)
function csvLine_(s) {
  var out = [], cur = '', q = false;
  for (var i = 0; i < s.length; i++) {
    var ch = s.charAt(i);
    if (q) { if (ch === '"') { if (s.charAt(i + 1) === '"') { cur += '"'; i++; } else q = false; } else cur += ch; }
    else if (ch === '"') q = true;
    else if (ch === ',') { out.push(cur); cur = ''; }
    else cur += ch;
  }
  out.push(cur);
  return out;
}
function stockNum_(v) {                             // "?4380750.07" → "4380750.07"
  var s = String(v || '').replace(/[^0-9.\-]/g, '');
  return s === '' || s === '-' || s === '.' ? '' : s;
}
function stockRow_(f) {
  var r = [];
  for (var c = 0; c < STOCK_NCOL; c++) r.push(String(f[c] == null ? '' : f[c]).replace(/[\t\r\n]+/g, ' ').trim());
  r[3] = stockNum_(r[3]); r[8] = stockNum_(r[8]);
  r.push(r.join('\t'));
  return r;
}

function syncStockRun_(why, force) {
  var url = schedProp_('STOCK_SRC_URL'), info = { time: nowStr_(), why: why }, t0 = Date.now();
  if (!url) throw new Error('ยังไม่ได้ตั้งค่าลิงก์ OneDrive ของ Item stock');
  try {
    var text = downloadShared_(url);
    var hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, text));
    var prev = stockInfo_();
    if (!force && prev && prev.ok && prev.hash === hash) {
      info.ok = true; info.unchanged = true; info.rows = prev.rows; info.hash = hash; info.msg = 'ไฟล์ไม่มีการเปลี่ยนแปลง';
      schedProp_('STOCK_SYNC_LAST', JSON.stringify(info)); return info;
    }
    var lines = text.split(/\r?\n/); text = null;
    var head = csvLine_(lines[0] || '').map(function (x) { return String(x).trim().toUpperCase(); });
    if (head[0] !== 'ITEMNO' || head[3] !== 'QTYONHAND' || head[8] !== 'LASTCOST' || head.length < STOCK_NCOL)
      throw new Error('หัวคอลัมน์ไม่ตรงรูปแบบ ต้องเป็น: ' + STOCK_HEADER.slice(0, STOCK_NCOL).join(', '));
    var rows = [], locs = {}, codes = {}, bys = {};
    for (var i = 1, buf = ''; i < lines.length; i++) {
      buf = buf ? buf + ' ' + lines[i] : lines[i];
      if (!buf) continue;
      if ((buf.match(/"/g) || []).length % 2) continue;      // quoted field spans a line break → join next line
      var r = stockRow_(csvLine_(buf)); buf = '';
      if (!r[0]) continue;
      rows.push(r); codes[r[0]] = 1; if (r[5]) locs[r[5]] = (locs[r[5]] || 0) + 1; if (r[7]) bys[r[7]] = (bys[r[7]] || 0) + 1;
    }
    lines = null;
    if (!rows.length) throw new Error('ไฟล์ไม่มีข้อมูล');

    var ss = stockSS_(), tmpName = STOCK_SHEET + '_new', tmp = ss.getSheetByName(tmpName);
    if (tmp) ss.deleteSheet(tmp);
    tmp = ss.insertSheet(tmpName);
    var need = rows.length + 1;
    if (tmp.getMaxRows() < need) tmp.insertRowsAfter(tmp.getMaxRows(), need - tmp.getMaxRows());
    if (tmp.getMaxColumns() > STOCK_W) tmp.deleteColumns(STOCK_W + 1, tmp.getMaxColumns() - STOCK_W);
    else if (tmp.getMaxColumns() < STOCK_W) tmp.insertColumnsAfter(tmp.getMaxColumns(), STOCK_W - tmp.getMaxColumns());
    tmp.getRange(1, 1, need, STOCK_W).setNumberFormat('@');
    tmp.getRange(1, 1, 1, STOCK_W).setValues([STOCK_HEADER]).setFontWeight('bold');
    for (var s = 0; s < rows.length; s += STOCK_CHUNK) {
      var blk = rows.slice(s, s + STOCK_CHUNK);
      tmp.getRange(s + 2, 1, blk.length, STOCK_W).setValues(blk);
      SpreadsheetApp.flush();
    }
    tmp.setFrozenRows(1); tmp.hideColumns(STOCK_W);
    var old = ss.getSheetByName(STOCK_SHEET);
    if (old) ss.deleteSheet(old);
    tmp.setName(STOCK_SHEET);
    ss.getSheets().forEach(function (x) { if (x.getName() !== STOCK_SHEET && ss.getSheets().length > 1) ss.deleteSheet(x); });

    var locList = Object.keys(locs).sort().map(function (k) { return [k, locs[k]]; });
    schedProp_('STOCK_META', JSON.stringify({ rows: rows.length, items: Object.keys(codes).length, locations: locList, bys: stockUnitList_(bys), time: info.time }));
    schedProp_('STOCK_HASH', hash);
    try { stockWarm_(rows.map(function (r) { return r[STOCK_NCOL]; }), hash); stockBundle_(); } catch (e) {}   // prime fast caches
    info.ok = true; info.rows = rows.length; info.items = Object.keys(codes).length; info.hash = hash;
    info.secs = Math.round((Date.now() - t0) / 1000);
    info.msg = 'อัปเดต ' + rows.length.toLocaleString() + ' แถว (' + info.items.toLocaleString() + ' รายการ) ใช้เวลา ' + info.secs + ' วินาที';
    audit_([['System', 'ซิงค์ Item stock จาก OneDrive', '', info.why + ' · ' + info.msg]]);
  } catch (e) {
    info.ok = false; info.msg = e.message || String(e);
    if (info.why !== 'อัตโนมัติ') audit_([['System', 'ซิงค์ Item stock ไม่สำเร็จ', '', info.msg]]);
  }
  schedProp_('STOCK_SYNC_LAST', JSON.stringify(info));
  return info;
}
// ── Fast data cache ──────────────────────────────────────────────────
// Reading 190k cells from the sheet costs ~3–6 s per search. Instead keep the packed rows
// gzip-compressed in CacheService (chunks ≤ 90KB, 6 h) → one getAll() + unzip ≈ 0.3–0.8 s.
var STOCK_CACHE_TTL = 21600, STOCK_CHUNK_B64 = 90000, _stockMem = null;
function stockWarm_(packed, hash) {
  var b64 = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(packed.join('\n'), 'text/plain')).getBytes());
  var put = {}, n = Math.ceil(b64.length / STOCK_CHUNK_B64);
  for (var i = 0; i < n; i++) put['stkd_' + hash.substring(0, 10) + '_' + i] = b64.substring(i * STOCK_CHUNK_B64, (i + 1) * STOCK_CHUNK_B64);
  var cache = CacheService.getScriptCache();
  for (var keys = Object.keys(put), k = 0; k < keys.length; k += 30) {
    var part = {}; keys.slice(k, k + 30).forEach(function (x) { part[x] = put[x]; });
    cache.putAll(part, STOCK_CACHE_TTL);
  }
  cache.put('stkd_' + hash.substring(0, 10) + '_n', String(n), STOCK_CACHE_TTL);
  _stockMem = { hash: hash, rows: packed };
  return n;
}
function stockData_() {
  var hash = schedProp_('STOCK_HASH') || '';
  if (_stockMem && _stockMem.hash === hash) return _stockMem.rows;
  var cache = CacheService.getScriptCache(), pre = 'stkd_' + hash.substring(0, 10) + '_', n = Number(cache.get(pre + 'n') || 0);
  if (hash && n) {
    var keys = []; for (var i = 0; i < n; i++) keys.push(pre + i);
    var got = cache.getAll(keys), parts = [], ok = true;
    for (var j = 0; j < n; j++) { if (got[keys[j]] == null) { ok = false; break; } parts.push(got[keys[j]]); }
    if (ok) {
      try {
        var txt = Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(parts.join('')), 'application/x-gzip')).getDataAsString();
        _stockMem = { hash: hash, rows: txt.split('\n') }; return _stockMem.rows;
      } catch (e) {}
    }
  }
  var sh = stockSS_().getSheetByName(STOCK_SHEET);           // cache miss → read sheet once, then re-prime
  if (!sh || sh.getLastRow() < 2) return [];
  var rows = sh.getRange(2, STOCK_W, sh.getLastRow() - 1, 1).getValues().map(function (r) { return String(r[0]); });
  try { if (hash) stockWarm_(rows, hash); } catch (e) { _stockMem = { hash: hash, rows: rows }; }
  return rows;
}
// time trigger every ~5 h keeps the cache warm so nobody hits the slow path (installed by setStockSource)
function stockKeepWarm() { try { stockData_(); stockBundle_(); } catch (e) {} }

// ── Browser-side search bundle ───────────────────────────────────────
// Public columns only (0–7: CODE…LastOfRequest_By). LASTCOST / Requisition No never leave the server.
var STOCK_PUB_COLS = 8;
function stockBundle_() {
  var hash = schedProp_('STOCK_HASH') || '', cache = CacheService.getScriptCache(), pre = 'stkb_' + hash.substring(0, 10) + '_';
  var n = Number(cache.get(pre + 'n') || 0);
  if (hash && n) {
    var keys = []; for (var i = 0; i < n; i++) keys.push(pre + i);
    var got = cache.getAll(keys), parts = [];
    for (var j = 0; j < n; j++) { if (got[keys[j]] == null) { parts = null; break; } parts.push(got[keys[j]]); }
    if (parts) return { hash: hash, b64: parts.join('') };
  }
  var pub = stockData_().map(function (p) { return p.split('\t').slice(0, STOCK_PUB_COLS).join('\t'); });
  var b64 = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(pub.join('\n'), 'text/plain')).getBytes());
  try {
    var put = {}, m = Math.ceil(b64.length / STOCK_CHUNK_B64);
    for (var k = 0; k < m; k++) put[pre + k] = b64.substring(k * STOCK_CHUNK_B64, (k + 1) * STOCK_CHUNK_B64);
    var ks = Object.keys(put);
    for (var x = 0; x < ks.length; x += 30) { var part = {}; ks.slice(x, x + 30).forEach(function (y) { part[y] = put[y]; }); cache.putAll(part, STOCK_CACHE_TTL); }
    cache.put(pre + 'n', String(m), STOCK_CACHE_TTL);
  } catch (e) {}
  return { hash: hash, b64: b64 };
}
// public: gzip+base64 of public columns; the page stores it in IndexedDB keyed by hash
function getStockBundle() { return stockBundle_(); }

// public: LASTCOST / Requisition No for the rows on screen (keys = ["CODE|LOCATION", ...], max 200)
function getStockCosts(token, costCode, keys) {
  var admin = isAdmin_(token), cost = admin || stockCostOk_(costCode);
  if (!cost && !admin) return { cost: false, admin: false, map: {} };
  var want = {}, codes = {};
  (keys || []).slice(0, 200).forEach(function (k) { k = String(k); want[k] = 1; codes[k.split('|')[0]] = 1; });
  var map = {}, rows = stockData_();
  for (var i = 0; i < rows.length; i++) {
    var p = rows[i];
    if (!codes[p.substring(0, p.indexOf('\t'))]) continue;   // cheap code check before splitting
    var f = p.split('\t'), k = f[0] + '|' + f[5];
    if (!want[k]) continue;
    if (!map[k]) map[k] = [cost ? f[STOCK_COST_COL] : '', admin ? f[STOCK_REQ_COL] : ''];
  }
  return { cost: cost, admin: admin, map: map };
}

function stockAutoSync() { if (schedProp_('STOCK_SRC_URL')) syncStockRun_('อัตโนมัติ', false); }   // time trigger
function getStockSource(token) {
  requireAdmin_(token);
  var hrs = 0; ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'stockAutoSync') hrs = Number(schedProp_('STOCK_SYNC_HOURS') || 8); });
  var id = schedProp_('STOCK_SS_ID');
  return { url: schedProp_('STOCK_SRC_URL') || '', hours: hrs, last: stockInfo_(), sheetUrl: id ? 'https://docs.google.com/spreadsheets/d/' + id + '/edit' : '' };
}
function setStockSource(token, url, hours) {
  requireAdmin_(token);
  url = String(url || '').trim();
  if (url && !/^https:\/\/[^\s]+$/i.test(url)) throw new Error('ลิงก์ต้องขึ้นต้นด้วย https://');
  if (/_layouts\/15\/onedrive\.aspx|\/Forms\/AllItems\.aspx/i.test(url)) throw new Error('นี่คือลิงก์หน้าโฟลเดอร์ OneDrive ไม่ใช่ลิงก์แชร์ไฟล์ — คลิกขวาที่ไฟล์ .csv → Share → Copy link แล้ววางลิงก์นั้นแทน');
  schedProp_('STOCK_SRC_URL', url || null);
  ScriptApp.getProjectTriggers().forEach(function (t) { var h = t.getHandlerFunction(); if (h === 'stockAutoSync' || h === 'stockKeepWarm') ScriptApp.deleteTrigger(t); });
  if (url) ScriptApp.newTrigger('stockKeepWarm').timeBased().everyHours(4).create();
  hours = Number(hours) || 0;
  if (url && hours > 0) {
    var tb = ScriptApp.newTrigger('stockAutoSync').timeBased();
    if (hours >= 24) tb.everyDays(1).atHour(5).create(); else tb.everyHours([1, 2, 4, 6, 8, 12].indexOf(hours) >= 0 ? hours : 8).create();
    schedProp_('STOCK_SYNC_HOURS', String(hours));
  }
  audit_([['Admin', 'ตั้งค่าซิงค์ Item stock', '', (url ? 'ตั้งลิงก์แล้ว' : 'ลบลิงก์') + ' · ' + (hours ? 'ทุก ' + hours + ' ชม.' : 'ปิดอัตโนมัติ')]]);
  return getStockSource(token);
}
function syncStockNow(token) { requireAdmin_(token); return syncStockRun_('สั่งโดย Admin', true); }

function stockUnitList_(u) { return Object.keys(u).sort(function (a, b) { return u[b] - u[a] || a.localeCompare(b); }).map(function (k) { return [k, u[k]]; }); }
// public: summary for the page header
function getStockMeta() {
  var m = stockMeta_() || {}, l = stockInfo_();
  m.hash = (schedProp_('STOCK_HASH') || '').substring(0, 10);
  if (m.rows && !m.bys) {                            // data synced before this list existed → build once
    try {
      var sh = stockSS_().getSheetByName(STOCK_SHEET), u = {};
      sh.getRange(2, 8, sh.getLastRow() - 1, 1).getValues().forEach(function (r) { var k = String(r[0]).trim(); if (k) u[k] = (u[k] || 0) + 1; });
      m.bys = stockUnitList_(u); schedProp_('STOCK_META', JSON.stringify(m));
    } catch (e) { m.bys = []; }
  }
  return { hash: m.hash || '', rows: m.rows || 0, items: m.items || 0, locations: m.locations || [], bys: m.bys || [], updated: m.time || '', lastOk: l ? !!l.ok : null };
}

// public search. q: words (all must match), field: all|code|desc, loc: LOCATION filter ('' = all)
function stockCostCode_() { return schedProp_('STOCK_COST_CODE') || STOCK_COST_CODE_DEFAULT; }
function stockCostOk_(code) { var c = String(code || ''); return !!c && c === stockCostCode_(); }
// public: user enters the cost code to reveal LASTCOST (case-sensitive)
function verifyStockCostCode(code) {
  var ok = stockCostOk_(code);
  if (!ok) Utilities.sleep(800);                     // slow down guessing
  return { ok: ok };
}
function getStockCostCode(token) { requireAdmin_(token); return { code: stockCostCode_() }; }
function setStockCostCode(token, code) {
  requireAdmin_(token);
  code = String(code || '').trim();
  if (code.length < 4 || code.length > 40) throw new Error('รหัสต้องยาว 4–40 ตัวอักษร');
  schedProp_('STOCK_COST_CODE', code);
  audit_([['Admin', 'เปลี่ยนรหัสดูราคา Item stock', '', '']]);
  return { code: code };
}

function searchStock(token, q, field, loc, page, size, by, sort, costCode) {
  var admin = isAdmin_(token), cost = admin || stockCostOk_(costCode);
  size = Math.min(Math.max(Number(size) || 50, 10), 100); page = Math.max(1, Number(page) || 1);
  q = String(q || '').toLowerCase().replace(/\s+/g, ' ').trim().substring(0, 80);
  field = { code: 'code', desc: 'desc' }[field] || 'all'; loc = String(loc || '').trim(); by = String(by || '').trim();
  sort = /^(code|desc|qty|unit|loc|cost)_(asc|desc)$/.test(String(sort || '')) ? String(sort) : '';
  var key = 'stk_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5,
    [schedProp_('STOCK_HASH') || '', q, field, loc, page, size, admin ? 1 : 0, by, sort, cost ? 1 : 0].join('|')));
  var cache = CacheService.getScriptCache(), hit = cache.get(key);
  if (hit) return JSON.parse(hit);

  var vals = stockData_();
  if (!vals.length) return { total: 0, page: 1, pages: 1, rows: [], admin: admin, cost: cost, empty: true };
  var terms = q ? q.split(' ') : [], hits = [];
  for (var i = 0; i < vals.length; i++) {
    var p = vals[i]; if (!p) continue;
    var f = p.split('\t');
    if (loc && f[5] !== loc) continue;
    if (by && f[7] !== by) continue;            // filter: LastOfRequest_By
    var rank = 2;
    if (terms.length) {
      var code = f[0].toLowerCase(), desc = (f[1] + ' ' + f[2]).toLowerCase();
      var hay = field === 'code' ? code : field === 'desc' ? desc : code + ' ' + desc, ok = true;
      for (var t = 0; t < terms.length; t++) if (hay.indexOf(terms[t]) < 0) { ok = false; break; }
      if (!ok) continue;
      if (field !== 'desc') rank = code === q ? 0 : code.indexOf(q) === 0 ? 1 : 2;
    }
    hits.push([rank, i, f]);
  }
  if (sort === 'cost_asc' || sort === 'cost_desc') { if (!cost) sort = ''; }
  if (sort) {
    var sp = sort.split('_'), col = { code: 0, desc: 1, qty: 3, unit: 4, loc: 5, cost: 8 }[sp[0]], dir = sp[1] === 'desc' ? -1 : 1;
    hits.sort(function (a, b) {
      var x = a[2][col], y = b[2][col], d;
      if (col === 3 || col === 8) { x = x === '' ? -Infinity : Number(x); y = y === '' ? -Infinity : Number(y); d = x - y; }
      else d = x < y ? -1 : x > y ? 1 : 0;
      return d * dir || a[1] - b[1];
    });
  } else if (terms.length) hits.sort(function (a, b) { return a[0] - b[0] || a[1] - b[1]; });
  // small result → send every match once; the browser pages locally (next page = instant)
  var all = hits.length <= STOCK_ALL_MAX;
  var pages = Math.max(1, Math.ceil(hits.length / size)); if (page > pages) page = pages;
  var rows = (all ? hits : hits.slice((page - 1) * size, page * size)).map(function (h) {
    var f = h[2].slice(0, STOCK_NCOL);
    if (!cost) f[STOCK_COST_COL] = '';
    if (!admin) f[STOCK_REQ_COL] = '';
    return f;
  });
  var res = { total: hits.length, page: page, pages: pages, rows: rows, admin: admin, cost: cost, all: all };
  try { var js = JSON.stringify(res); if (js.length < 95000) cache.put(key, js, 600); } catch (e) {}
  return res;
}

// Run ONCE from the editor (▶ Run) after saving the link, to authorize and see timing in the log.
function testStockSync() {
  if (!schedProp_('STOCK_SRC_URL')) { Logger.log('✗ ยังไม่ได้บันทึกลิงก์ในหน้า Item stock → การ์ด ☁ ซิงค์ Item stock'); return; }
  var r = syncStockRun_('ทดสอบจาก Editor', true);
  Logger.log((r.ok ? '✓ ' : '✗ ') + r.msg);
}
// ══════════════════════════════════════════════════════════════════════
//  PENDING ORDER — PR / PO tracking, synced from a OneDrive "Anyone with the link" CSV.
//  Stored in the Helpdesk_ItemStock spreadsheet (sheet "PendingOrder"). The whole table is
//  small enough to send to the browser once (gzip, cached) and filtered there.
// ══════════════════════════════════════════════════════════════════════
var PO_SHEET = 'PendingOrder';
var PO_HEADER = ['Item_Id', 'Requisition_No', 'PO_NO', 'Specification_No', 'Status', 'Item_Name', 'Request_Date',
  'EXPARRIVAL', 'PO_Issue_Date', 'PO_Delivery', 'Record_By', 'Value', 'Request_By', 'Quantity', 'Unit'];
var PO_DATE_COLS = [6, 7, 8, 9];                   // Request_Date, EXPARRIVAL, PO_Issue_Date, PO_Delivery
var PO_QTY_COL = 13;

function poInfo_() { try { return JSON.parse(schedProp_('PO_SYNC_LAST') || 'null'); } catch (e) { return null; } }
function poKey_(x) { return String(x || '').toLowerCase().replace(/[^a-z0-9]/g, ''); }

// Dates in the export may be d/m/y or m/d/y (Access). Decide per column from values whose day part > 12.
function poDateOrder_(rows, c) {
  var dm = 0, md = 0;
  for (var i = 0; i < rows.length; i++) {
    var m = /^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/.exec(rows[i][c]); if (!m) continue;
    if (+m[1] > 12) dm++; else if (+m[2] > 12) md++;
  }
  return md > dm ? 'md' : 'dm';
}
function poDate_(v, order) {
  v = String(v || '').trim(); if (!v) return '';
  var m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(v), d, mo, y;
  if (m) { y = +m[1]; mo = +m[2]; d = +m[3]; }
  else {
    m = /^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/.exec(v); if (!m) return v;
    d = order === 'md' ? +m[2] : +m[1]; mo = order === 'md' ? +m[1] : +m[2]; y = +m[3];
    if (y < 100) y += 2000; else if (y > 2400) y -= 543;
  }
  if (y < 1950 || mo < 1 || mo > 12 || d < 1 || d > 31) return '';
  return (d < 10 ? '0' : '') + d + '/' + (mo < 10 ? '0' : '') + mo + '/' + y;
}

function syncPendingRun_(why, force) {
  var url = schedProp_('PO_SRC_URL'), info = { time: nowStr_(), why: why }, t0 = Date.now();
  if (!url) throw new Error('ยังไม่ได้ตั้งค่าลิงก์ OneDrive ของ Pending Order');
  try {
    var text = downloadShared_(url);
    var hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, text));
    var prev = poInfo_();
    if (!force && prev && prev.ok && prev.hash === hash) {
      info.ok = true; info.unchanged = true; info.rows = prev.rows; info.hash = hash; info.msg = 'ไฟล์ไม่มีการเปลี่ยนแปลง';
      schedProp_('PO_SYNC_LAST', JSON.stringify(info)); return info;
    }
    var lines = text.split(/\r?\n/); text = null;
    var head = csvLine_(lines[0] || '').map(poKey_), idx = PO_HEADER.map(function (h) { return head.indexOf(poKey_(h)); });
    if (idx[0] < 0 || (idx[1] < 0 && idx[2] < 0))
      throw new Error('หัวคอลัมน์ไม่ตรงรูปแบบ ต้องมีอย่างน้อย Item_Id และ Requisition_No หรือ PO_NO · รูปแบบเต็ม: ' + PO_HEADER.join(', '));
    var rows = [];
    for (var i = 1, buf = ''; i < lines.length; i++) {
      buf = buf ? buf + ' ' + lines[i] : lines[i];
      if (!buf) continue;
      if ((buf.match(/"/g) || []).length % 2) continue;      // quoted field spans a line break
      var f = csvLine_(buf); buf = '';
      var r = idx.map(function (c) { return c < 0 ? '' : String(f[c] == null ? '' : f[c]).replace(/[\t\r\n]+/g, ' ').trim(); });
      if (!r.some(function (x) { return x; })) continue;
      r[PO_QTY_COL] = stockNum_(r[PO_QTY_COL]) || r[PO_QTY_COL];
      rows.push(r);
    }
    lines = null;
    if (!rows.length) throw new Error('ไฟล์ไม่มีข้อมูล');
    PO_DATE_COLS.forEach(function (c) { var o = poDateOrder_(rows, c); rows.forEach(function (r) { r[c] = poDate_(r[c], o); }); });

    var ss = stockSS_(), sh = ss.getSheetByName(PO_SHEET) || ss.insertSheet(PO_SHEET), all = [PO_HEADER].concat(rows);
    sh.clearContents();
    if (sh.getMaxRows() < all.length) sh.insertRowsAfter(sh.getMaxRows(), all.length - sh.getMaxRows());
    if (sh.getMaxColumns() < PO_HEADER.length) sh.insertColumnsAfter(sh.getMaxColumns(), PO_HEADER.length - sh.getMaxColumns());
    var rg = sh.getRange(1, 1, all.length, PO_HEADER.length); rg.setNumberFormat('@'); rg.setValues(all);
    sh.getRange(1, 1, 1, PO_HEADER.length).setFontWeight('bold'); sh.setFrozenRows(1);
    var extra = sh.getMaxRows() - all.length; if (extra > 200) sh.deleteRows(all.length + 101, extra - 100);

    schedProp_('PO_HASH', hash);
    try { pendingBundle_(true); } catch (e) {}
    var prs = {}, pos = {};
    rows.forEach(function (r) { if (r[1]) prs[r[1]] = 1; if (r[2]) pos[r[2]] = 1; });
    info.ok = true; info.rows = rows.length; info.hash = hash; info.secs = Math.round((Date.now() - t0) / 1000);
    info.msg = 'อัปเดต ' + rows.length.toLocaleString() + ' แถว (PR ' + Object.keys(prs).length + ' · PO ' + Object.keys(pos).length + ') ใช้เวลา ' + info.secs + ' วินาที';
    audit_([['System', 'ซิงค์ Pending Order จาก OneDrive', '', info.why + ' · ' + info.msg]]);
  } catch (e) {
    info.ok = false; info.msg = e.message || String(e);
    if (info.why !== 'อัตโนมัติ') audit_([['System', 'ซิงค์ Pending Order ไม่สำเร็จ', '', info.msg]]);
  }
  schedProp_('PO_SYNC_LAST', JSON.stringify(info));
  return info;
}

// gzip+base64 of the TSV rows, chunked in CacheService (6 h) so most page loads skip the sheet read.
function pendingBundle_(rebuild) {
  var hash = schedProp_('PO_HASH') || '', cache = CacheService.getScriptCache(), pre = 'pob_' + hash.substring(0, 10) + '_';
  if (!hash) return { hash: '', b64: '' };
  if (!rebuild) {
    var n = Number(cache.get(pre + 'n') || 0);
    if (n) {
      var keys = []; for (var i = 0; i < n; i++) keys.push(pre + i);
      var got = cache.getAll(keys), parts = [];
      for (var j = 0; j < n; j++) { if (got[keys[j]] == null) { parts = null; break; } parts.push(got[keys[j]]); }
      if (parts) return { hash: hash.substring(0, 10), b64: parts.join('') };
    }
  }
  var sh = stockSS_().getSheetByName(PO_SHEET);
  var rows = sh && sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, PO_HEADER.length).getDisplayValues() : [];
  var b64 = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(rows.map(function (r) { return r.join('\t'); }).join('\n'), 'text/plain')).getBytes());
  try {
    var put = {}, m = Math.ceil(b64.length / STOCK_CHUNK_B64);
    for (var k = 0; k < m; k++) put[pre + k] = b64.substring(k * STOCK_CHUNK_B64, (k + 1) * STOCK_CHUNK_B64);
    var ks = Object.keys(put);
    for (var x = 0; x < ks.length; x += 30) { var part = {}; ks.slice(x, x + 30).forEach(function (y) { part[y] = put[y]; }); cache.putAll(part, STOCK_CACHE_TTL); }
    cache.put(pre + 'n', String(m), STOCK_CACHE_TTL);
  } catch (e) {}
  return { hash: hash.substring(0, 10), b64: b64 };
}
// public: summary + data hash (the page keeps the data in IndexedDB keyed by hash)
function getPendingMeta() {
  var l = poInfo_();
  return { hash: (schedProp_('PO_HASH') || '').substring(0, 10), rows: l && l.ok ? l.rows : (l && l.rows) || 0,
    updated: l && l.ok ? l.time : '', lastOk: l ? !!l.ok : null };
}
function getPendingBundle() { return pendingBundle_(false); }

function pendingAutoSync() { if (schedProp_('PO_SRC_URL')) syncPendingRun_('อัตโนมัติ', false); }   // time trigger
function getPendingSource(token) {
  requireAdmin_(token);
  var hrs = 0; ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'pendingAutoSync') hrs = Number(schedProp_('PO_SYNC_HOURS') || 4); });
  var id = schedProp_('STOCK_SS_ID');
  return { url: schedProp_('PO_SRC_URL') || '', hours: hrs, last: poInfo_(), sheetUrl: id ? 'https://docs.google.com/spreadsheets/d/' + id + '/edit' : '' };
}
function setPendingSource(token, url, hours) {
  requireAdmin_(token);
  url = String(url || '').trim();
  if (url && !/^https:\/\/[^\s]+$/i.test(url)) throw new Error('ลิงก์ต้องขึ้นต้นด้วย https://');
  if (/_layouts\/15\/onedrive\.aspx|\/Forms\/AllItems\.aspx/i.test(url)) throw new Error('นี่คือลิงก์หน้าโฟลเดอร์ OneDrive ไม่ใช่ลิงก์แชร์ไฟล์ — คลิกขวาที่ไฟล์ .csv → Share → Copy link แล้ววางลิงก์นั้นแทน');
  schedProp_('PO_SRC_URL', url || null);
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'pendingAutoSync') ScriptApp.deleteTrigger(t); });
  hours = Number(hours) || 0;
  if (url && hours > 0) {
    var tb = ScriptApp.newTrigger('pendingAutoSync').timeBased();
    if (hours >= 24) tb.everyDays(1).atHour(6).create(); else tb.everyHours([1, 2, 4, 6, 8, 12].indexOf(hours) >= 0 ? hours : 4).create();
    schedProp_('PO_SYNC_HOURS', String(hours));
  }
  audit_([['Admin', 'ตั้งค่าซิงค์ Pending Order', '', (url ? 'ตั้งลิงก์แล้ว' : 'ลบลิงก์') + ' · ' + (hours ? 'ทุก ' + hours + ' ชม.' : 'ปิดอัตโนมัติ')]]);
  return getPendingSource(token);
}
function syncPendingNow(token) { requireAdmin_(token); return syncPendingRun_('สั่งโดย Admin', true); }

// Run ONCE from the editor after saving the link, to authorize and check the result in the log.
function testPendingSync() {
  if (!schedProp_('PO_SRC_URL')) { Logger.log('✗ ยังไม่ได้บันทึกลิงก์ในหน้า Pending Order → การ์ด ☁ ซิงค์ Pending Order'); return; }
  var r = syncPendingRun_('ทดสอบจาก Editor', true);
  Logger.log((r.ok ? '✓ ' : '✗ ') + r.msg);
}
