/**
 * หน่วยงาน — the values that say what belongs to our unit, kept in one place (ตั้งค่าระบบ → 🏢 หน่วยงาน):
 *   by          orderer code in Item stock / Pending Order (LastOfRequest_By / Request_By), e.g. DS13 — also the
 *               Booking BOM "ผู้สั่ง" and the Stock Alert unit
 *   mfDept      Shipment department MF issues our parts to, e.g. CONTROL WIRING (Booking BOM × Shipment, สถิติ)
 *   svPrefix    CODE prefixes of service jobs, never counted (SV) — comma separated
 *   projLetter  letter after the project number in a By Project CODE (4556PP1030C00201 → C)
 *   codeLen     length of a stock CODE watched by Stock Alert (8)
 */
var UNIT_DEF = { by: 'DS13', mfDept: 'CONTROL WIRING', svPrefix: 'SV', projLetter: 'C', codeLen: 8 }, __unit = null;
function unitCfg_() {
  if (__unit) return __unit;
  var c = {}; try { c = JSON.parse(schedProp_('UNIT_CFG') || '{}') || {}; } catch (e) { errLog_(e, 'unitCfg_'); }
  var by = c.by || ''; if (!by) { try { by = bomCfg_().orderBy; } catch (e) { by = ''; } }
  var len = Number(c.codeLen);
  __unit = { by: String(by || UNIT_DEF.by).trim().toUpperCase(), mfDept: String(c.mfDept || UNIT_DEF.mfDept).trim().toUpperCase(),
    svPrefix: String(c.svPrefix != null ? c.svPrefix : UNIT_DEF.svPrefix).toUpperCase(), projLetter: String(c.projLetter || UNIT_DEF.projLetter).trim().toUpperCase().substring(0, 2),
    codeLen: len >= 4 && len <= 20 ? Math.round(len) : UNIT_DEF.codeLen };
  return __unit;
}
function unitNorm_(s) { return String(s || '').toUpperCase().replace(/[^A-Z0-9]/g, ''); }
function unitDeptOk_(s) { return unitNorm_(s) === unitNorm_(unitCfg_().mfDept); }
function unitSvSkip_(code) { var k = String(code || '').trim().toUpperCase(); return unitCfg_().svPrefix.split(/[,\s]+/).some(function (p) { return p && k.indexOf(p) === 0; }); }
function unitProjC_(code) { var L = unitCfg_().projLetter.replace(/[^A-Z]/g, '') || 'C'; return new RegExp('^\\d{4}PP\\d{4}' + L + '\\d{3,}$', 'i').test(String(code || '').trim()); }

function getUnitCfg(token) { apiOnly_(); var u = unitCfg_(); return { cfg: u, def: UNIT_DEF, admin: isAdmin_(token) }; }
function saveUnitCfg(token, c) { apiOnly_();
  requireAdmin_(token); c = c || {};
  var by = String(c.by || '').trim().toUpperCase().replace(/[^A-Z0-9,\/ -]/g, '').substring(0, 30);
  var dept = String(c.mfDept || '').trim().toUpperCase().replace(/\s+/g, ' ').substring(0, 60);
  var sv = String(c.svPrefix || '').toUpperCase().replace(/[^A-Z0-9, ]/g, '').substring(0, 40);
  var L = String(c.projLetter || '').trim().toUpperCase().replace(/[^A-Z]/g, '').substring(0, 2), len = Number(c.codeLen);
  if (!by) throw new Error('ใส่รหัสผู้สั่งของหน่วยงาน (เช่น DS13)');
  if (!dept) throw new Error('ใส่ชื่อแผนกที่ MF เบิกเข้า (เช่น CONTROL WIRING)');
  if (!L) throw new Error('ใส่ตัวอักษรของ CODE By Project (เช่น C)');
  if (!(len >= 4 && len <= 20)) throw new Error('ความยาว CODE ต้องอยู่ระหว่าง 4–20 หลัก');
  var v = { by: by, mfDept: dept, svPrefix: sv, projLetter: L, codeLen: Math.round(len) };
  schedProp_('UNIT_CFG', JSON.stringify(v)); __unit = null;
  // the older settings that hold the orderer follow it (Booking BOM ผู้สั่ง, Stock Alert unit)
  try { var b = JSON.parse(schedProp_('BOM_CFG') || '{}') || {}; b.orderBy = by; schedProp_('BOM_CFG', JSON.stringify(b)); } catch (e) { errLog_(e, 'saveUnitCfg bom'); }
  try { var s = JSON.parse(schedProp_('SA_CFG') || '{}') || {}; s.by = by; schedProp_('SA_CFG', JSON.stringify(s)); saBump_(); } catch (e) { errLog_(e, 'saveUnitCfg sa'); }
  audit_([['Admin', 'ตั้งค่าหน่วยงาน', '', 'ผู้สั่ง ' + by + ' · MF → ' + dept + ' · ไม่นับ ' + (sv || '-') + ' · By Project …' + L + '… · CODE ' + v.codeLen + ' หลัก']]);
  return getUnitCfg(token);
}
