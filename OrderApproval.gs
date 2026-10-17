/**
 * ═════════════════════════════════════════════════════════════════════
 *  ORDER APPROVAL — ขั้นตอนอนุมัติใบสั่ง  DES → CHK → APP
 *  • Admin กำหนดรายชื่อผู้ตรวจสอบ (CHK) และผู้อนุมัติ (APP) พร้อมอีเมล  (Order Item → แท็บ DatabaseItemOrder)
 *  • ผู้ออกใบสั่ง (DES) กด "ส่งขออนุมัติ" → เลือก CHK / APP → ระบบส่งอีเมลพร้อมลิงก์ให้ CHK
 *  • CHK ผ่าน  → ส่งอีเมลให้ APP + แจ้งผลกลับ DES / CHK
 *  • APP อนุมัติ → แจ้งผลกลับทุกคนที่เกี่ยวข้อง (DES, CHK, APP + CC ที่ตั้งไว้)
 *  • ไม่อนุมัติ (ขั้นใดก็ได้) → แจ้งผลกลับทุกคนพร้อมเหตุผล — DES แก้ไขแล้วส่งใหม่ได้
 *  • แก้ไขใบสั่งหลังส่งขออนุมัติ / อนุมัติแล้ว → สถานะอนุมัติถูกล้าง ต้องส่งใหม่ (ลิงก์เดิมใช้ไม่ได้)
 *  ลิงก์ในอีเมลมีรหัสสุ่มเฉพาะขั้นนั้น (ใช้ได้ครั้งเดียว) — ห้ามส่งต่ออีเมลให้ผู้อื่น
 *  ต้องมี Code.gs และ OrderService.gs (ใช้ helper ร่วมกัน) + ไฟล์ HTML "OrderApprovalPage"
 * ═════════════════════════════════════════════════════════════════════
 */
var OA_ROLES = { chk: 'CHK (ผู้ตรวจสอบ)', app: 'APP (ผู้อนุมัติ)' };
var OA_STATE_TH = { chk: 'รอ CHK ตรวจสอบ', app: 'รอ APP อนุมัติ', approved: 'อนุมัติแล้ว', rejected: 'ไม่อนุมัติ', '': 'ยังไม่ส่งขออนุมัติ' };
var OA_EMAIL_RE = /^[^@\s,;<>"]+@[^@\s,;<>"]+\.[^@\s,;<>"]+$/;

// ── approvers config (Script property ORD_APPROVERS) ─────────────────
function oaCleanPeople_(list) {
  var seen = {};
  return (Array.isArray(list) ? list : []).map(function (p) {
    return { name: ordStr_(p && p.name, 60).replace(/[\r\n]+/g, ' ').trim(), email: ordStr_(p && p.email, 120).trim() };
  }).filter(function (p) {
    var k = p.name.toLowerCase();
    if (!p.name || !OA_EMAIL_RE.test(p.email) || seen[k]) return false;
    seen[k] = 1; return true;
  }).slice(0, 50);
}
function oaCfg_() {
  var c = {};
  try { c = JSON.parse(schedProp_('ORD_APPROVERS') || '{}') || {}; } catch (e) {}
  return { chk: oaCleanPeople_(c.chk), app: oaCleanPeople_(c.app), cc: emailList_(c.cc).slice(0, 20).join(', ') };
}
function oaNames_() { var c = oaCfg_(); return { chk: c.chk.map(function (p) { return p.name; }), app: c.app.map(function (p) { return p.name; }) }; }
function oaPerson_(role, name) {
  var n = String(name || '').trim().toLowerCase();
  return oaCfg_()[role].filter(function (p) { return p.name.toLowerCase() === n; })[0] || null;
}
function getOrderApprovers(token) { requireAdmin_(token); return oaCfg_(); }
function saveOrderApprovers(token, cfg) {
  requireAdmin_(token);
  cfg = cfg || {};
  var c = { chk: oaCleanPeople_(cfg.chk), app: oaCleanPeople_(cfg.app), cc: emailList_(cfg.cc).slice(0, 20).join(', ') };
  schedProp_('ORD_APPROVERS', JSON.stringify(c));
  audit_([['Admin', 'ตั้งค่าผู้อนุมัติใบสั่ง (CHK/APP)', '', 'CHK ' + c.chk.length + ' คน · APP ' + c.app.length + ' คน']]);
  return c;
}

// ── approval object helpers ─────────────────────────────────────────
// approval = { state, req:{name,email}, chk:{name,email,at,comment}, app:{...}, k, sent, history:[{at,role,name,action,comment}] }
function oaOf_(o) { var a = o.approval; return a && typeof a === 'object' ? a : { state: '', history: [] }; }
function oaHist_(a, role, name, action, comment) {
  a.history = (a.history || []).concat([{ at: nowStr_(), role: role, name: name || '', action: action, comment: ordStr_(comment, 500) }]).slice(-30);
}
// what the browser may see (never the step key or e-mail addresses)
function oaPublic_(a) {
  if (!a || typeof a !== 'object' || (!a.state && !(a.history || []).length)) return null;
  var p = function (x) { return x ? { name: x.name || '', at: x.at || '', comment: x.comment || '' } : null; };
  return { state: a.state || '', stateLabel: OA_STATE_TH[a.state || ''] || '', req: a.req ? { name: a.req.name || '' } : null,
    chk: p(a.chk), app: p(a.app), sent: a.sent || '', history: a.history || [] };
}
function oaUrl_(id, k) { return baseUrl_() + '?ordapprove=' + encodeURIComponent(id) + '&k=' + encodeURIComponent(k); }
function oaKeyOk_(a, k) { return !!a.k && typeof k === 'string' && k.length === a.k.length && k === a.k; }
function oaSave_(o) { writeText_(ordSheet_(), o.row, 1, ordObjToRow_(o)); }

// ── e-mail ──────────────────────────────────────────────────────────
function oaCard_(o, a) {
  var C = EMAIL_CFG, row = function (l, v) {
    return '<tr><td style="padding:6px 10px;color:' + C.muted + ';font-size:12pt;white-space:nowrap;vertical-align:top;">' + esc_(l) +
      '</td><td style="padding:6px 10px;color:' + C.ink + ';font-size:13pt;font-weight:600;">' + esc_(v || '-') + '</td></tr>';
  };
  var items = (o.items || []).map(function (i, n) {
    return '<tr><td style="padding:6px 8px;border-top:1px solid ' + C.line + ';font-size:12pt;">' + (n + 1) + '</td>' +
      '<td style="padding:6px 8px;border-top:1px solid ' + C.line + ';font-size:12pt;font-family:monospace;">' + esc_(i.code) + '</td>' +
      '<td style="padding:6px 8px;border-top:1px solid ' + C.line + ';font-size:12pt;">' + esc_(i.material) + (i.spec ? '<br><span style="color:' + C.muted + ';">' + esc_(i.spec) + '</span>' : '') + '</td>' +
      '<td style="padding:6px 8px;border-top:1px solid ' + C.line + ';font-size:12pt;text-align:right;white-space:nowrap;">' + esc_(i.qty) + ' ' + esc_(i.unit) + '</td></tr>';
  }).join('');
  var step = function (label, x, on) {
    var done = x && x.at, col = done ? '#059669' : on ? '#d97706' : '#9ca3af';
    return '<td align="center" style="padding:8px;font-size:12pt;color:' + col + ';font-weight:700;">' + esc_(label) + '<br><span style="font-weight:400;color:' + C.ink + ';">' +
      esc_(x && x.name || '-') + '</span><br><span style="font-weight:400;font-size:10pt;color:' + C.muted + ';">' + esc_(done ? '✓ ' + x.at : on ? 'รอดำเนินการ' : '') + '</span></td>';
  };
  return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:18px;border:1px solid ' + C.line + ';border-radius:12px;">' +
    row('ใบสั่ง', o.id) + row('Spec No.', o.spec) + row('RQN', o.rqn) + row("DES'D (ผู้ออก)", o.des) + row('วันที่ออก', o.issueDate) +
    row('Delivery Date', o.delivery) + row('Remark', o.remark) + row('สถานะอนุมัติ', OA_STATE_TH[a.state || '']) + '</table>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:12px;border:1px solid ' + C.line + ';border-radius:12px;"><tr>' +
    step('DES', { name: o.des, at: a.sent }, false) + step('CHK', a.chk, a.state === 'chk') + step('APP', a.app, a.state === 'app') + '</tr></table>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:12px;border:1px solid ' + C.line + ';border-radius:12px;">' +
    '<tr><td style="padding:6px 8px;font-size:11pt;color:' + C.muted + ';">#</td><td style="padding:6px 8px;font-size:11pt;color:' + C.muted + ';">CODE</td>' +
    '<td style="padding:6px 8px;font-size:11pt;color:' + C.muted + ';">MATERIAL / SPEC</td><td style="padding:6px 8px;font-size:11pt;color:' + C.muted + ';text-align:right;">Q\'TY</td></tr>' +
    items + '</table>';
}
function oaMail_(ev, o, a, to, cc, heading, intro, message, button) {
  applyBrand_();
  var subject = '[Order ' + o.id + '] ' + heading + ' — ' + (o.spec || '') + (o.remark ? ' · ' + o.remark : '');
  var html = emailHtml_({ title: subject, preheader: intro, heading: heading, intro: intro, message: message || '',
    messageLabel: 'ข้อความ', extraHtml: oaCard_(o, a), buttons: button ? [button] : [] });
  var text = [heading, '', intro, '', 'ใบสั่ง: ' + o.id + ' · ' + o.spec + ' · ' + o.rqn, message ? 'ข้อความ: ' + message : '',
    button ? button.label + ': ' + button.url : ''].join('\n');
  sendMail_(ev, o.id, to, cc, { enabled: true, subject: subject.replace(/[\r\n]+/g, ' ').substring(0, 240), html: html, text: text });
}
function oaEveryone_(a) {
  return uniq_([a.req && a.req.email, a.chk && a.chk.email, a.app && a.app.email]);
}

// ── public: DES submits (or re-submits) an order for approval ───────
// req = { chk: name, app: name, email: requester e-mail, message }
function submitOrderApproval(id, req) {
  req = req || {};
  var email = ordStr_(req.email, 120).trim(), msg = ordStr_(req.message, 500);
  if (!OA_EMAIL_RE.test(email)) throw new Error('กรุณากรอกอีเมลของผู้ออกใบสั่ง (DES) ให้ถูกต้อง เพื่อรับผลการอนุมัติ');
  var chk = oaPerson_('chk', req.chk), app = oaPerson_('app', req.app);
  if (!chk) throw new Error('กรุณาเลือกผู้ตรวจสอบ (CHK) จากรายชื่อ — หากไม่มีรายชื่อ ให้ Admin ตั้งค่าก่อน');
  if (!app) throw new Error('กรุณาเลือกผู้อนุมัติ (APP) จากรายชื่อ');
  var rc = CacheService.getScriptCache(), rk = 'oa_rl_' + id, n = Number(rc.get(rk) || 0);
  if (n >= 5) throw new Error('ส่งขออนุมัติใบสั่งนี้ถี่เกินไป กรุณารอสักครู่');
  var res = withLock_(function () {
    var o = ordFind_(String(id || ''));
    if (!o) throw new Error('ไม่พบใบสั่ง ' + id + ' — กรุณาบันทึกก่อน');
    var a = oaOf_(o);
    if (a.state === 'approved') throw new Error('ใบสั่ง ' + o.id + ' อนุมัติแล้ว — หากต้องการส่งใหม่ ให้แก้ไขและบันทึกใบสั่งก่อน');
    var again = a.state === 'chk' || a.state === 'app';
    a = { state: 'chk', req: { name: o.des, email: email }, chk: { name: chk.name, email: chk.email }, app: { name: app.name, email: app.email },
      k: Utilities.getUuid(), sent: nowStr_(), history: a.history || [] };
    oaHist_(a, 'DES', o.des, again ? 'ส่งขออนุมัติใหม่' : 'ส่งขออนุมัติ', msg);
    o.chk = ''; o.app = '';                           // filled in automatically when each person signs off
    o.approval = a; o.updated = nowStr_();
    oaSave_(o);
    return o;
  });
  rc.put(rk, String(n + 1), 3600);
  var a = res.approval;
  audit_([['User', 'ส่งขออนุมัติใบสั่ง', res.id, 'CHK ' + a.chk.name + ' · APP ' + a.app.name]]);
  oaMail_('order_chk_request', res, a, [a.chk.email], [a.req.email], 'ขอให้ตรวจสอบใบสั่ง (CHK)',
    'เรียน คุณ' + a.chk.name + ' — คุณ' + res.des + ' ส่งใบสั่ง ' + res.id + ' ให้ตรวจสอบ กรุณากดปุ่มด้านล่างเพื่อดูรายละเอียดและ ผ่าน / ไม่ผ่าน',
    msg, { label: 'ตรวจสอบใบสั่ง', url: oaUrl_(res.id, a.k) });
  return { order: ordPub_(res), orders: ordPubAll_() };
}

// ── approval page (opened from the e-mail link) ─────────────────────
function orderApprovalPage_(id, k) {
  var t = HtmlService.createHtmlOutputFromFile('OrderApprovalPage')   // HTML file name must differ from this .gs file
    .setTitle('อนุมัติใบสั่ง ' + String(id).substring(0, 20))
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
  t.append('<script>window.__OA__=' + JSON.stringify({ id: String(id).substring(0, 40), k: String(k).substring(0, 64) }).replace(/</g, '\\u003c') + ';</script>');
  return t;
}
function oaView_(o, a, k) {
  var role = oaKeyOk_(a, k) && (a.state === 'chk' || a.state === 'app') ? a.state : '';
  var x = ordPub_(o);
  return { order: x, role: role, roleLabel: OA_ROLES[role] || '', you: role ? a[role].name : '', state: a.state || '',
    stateLabel: OA_STATE_TH[a.state || ''] || '' };
}
function getOrderApprovalView(id, k) {
  var o = ordFind_(String(id || ''));
  if (!o) throw new Error('ไม่พบใบสั่ง ' + id + ' (อาจถูกลบไปแล้ว)');
  var a = oaOf_(o), v = oaView_(o, a, String(k || ''));
  if (!v.role) Utilities.sleep(500);
  return v;
}
// action = 'approve' | 'reject'
function actOrderApproval(id, k, action, comment) {
  comment = ordStr_(comment, 500).trim();
  if (action !== 'approve' && action !== 'reject') throw new Error('คำสั่งไม่ถูกต้อง');
  if (action === 'reject' && !comment) throw new Error('กรุณาระบุเหตุผลที่ไม่อนุมัติ');
  var r = withLock_(function () {
    var o = ordFind_(String(id || ''));
    if (!o) throw new Error('ไม่พบใบสั่ง ' + id);
    var a = oaOf_(o), role = a.state;
    if (!(role === 'chk' || role === 'app') || !oaKeyOk_(a, String(k || ''))) {
      Utilities.sleep(800);
      throw new Error('ลิงก์นี้ใช้ไม่ได้แล้ว (ดำเนินการไปแล้ว ใบสั่งถูกแก้ไข หรือถูกส่งใหม่) — สถานะปัจจุบัน: ' + (OA_STATE_TH[a.state || ''] || '-'));
    }
    var who = a[role], R = role.toUpperCase();
    who.at = nowStr_(); who.comment = comment;
    if (action === 'reject') {
      a.state = 'rejected'; a.k = '';
      oaHist_(a, R, who.name, 'ไม่อนุมัติ', comment);
    } else if (role === 'chk') {
      o.chk = who.name; a.state = 'app'; a.k = Utilities.getUuid();
      oaHist_(a, R, who.name, 'ตรวจสอบผ่าน', comment);
    } else {
      o.app = who.name; a.state = 'approved'; a.k = '';
      oaHist_(a, R, who.name, 'อนุมัติ', comment);
    }
    o.approval = a; o.updated = nowStr_();
    oaSave_(o);
    return { o: o, a: a, role: role, who: who };
  });
  var o = r.o, a = r.a, cc = emails_('adminEmails').concat(emailList_(oaCfg_().cc));
  audit_([[r.role.toUpperCase() + ': ' + r.who.name, action === 'reject' ? 'ไม่อนุมัติใบสั่ง' : r.role === 'chk' ? 'CHK ตรวจสอบผ่าน' : 'APP อนุมัติใบสั่ง', o.id, comment]]);
  if (action === 'reject') {
    oaMail_('order_rejected', o, a, oaEveryone_(a), cc, 'ใบสั่ง ' + o.id + ' ไม่อนุมัติ',
      'คุณ' + r.who.name + ' (' + r.role.toUpperCase() + ') ไม่อนุมัติใบสั่งนี้ — ผู้ออกใบสั่งสามารถแก้ไขแล้วส่งขออนุมัติใหม่ได้', comment);
  } else if (r.role === 'chk') {
    oaMail_('order_app_request', o, a, [a.app.email], [a.req.email, a.chk.email], 'ขออนุมัติใบสั่ง (APP)',
      'เรียน คุณ' + a.app.name + ' — ใบสั่ง ' + o.id + ' ผ่านการตรวจสอบจาก คุณ' + a.chk.name + ' แล้ว กรุณากดปุ่มด้านล่างเพื่อ อนุมัติ / ไม่อนุมัติ',
      comment, { label: 'อนุมัติใบสั่ง', url: oaUrl_(o.id, a.k) });
    oaMail_('order_checked', o, a, [a.req.email], [a.chk.email], 'ใบสั่ง ' + o.id + ' ผ่านการตรวจสอบ (CHK) แล้ว',
      'คุณ' + a.chk.name + ' ตรวจสอบผ่านแล้ว — ระบบส่งต่อให้ คุณ' + a.app.name + ' อนุมัติ', comment);
  } else {
    oaMail_('order_approved', o, a, oaEveryone_(a), cc, 'ใบสั่ง ' + o.id + ' อนุมัติแล้ว ✓',
      'ใบสั่งนี้ผ่านครบทุกขั้นตอน: DES ' + o.des + ' → CHK ' + a.chk.name + ' → APP ' + a.app.name + ' — ส่งออก PDF / Excel ได้จากหน้า Order Item', comment);
  }
  return oaView_(o, a, '');
}
