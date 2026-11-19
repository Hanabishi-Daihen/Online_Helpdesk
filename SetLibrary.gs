/**
 * ชุดอุปกรณ์ (Set Library) — parts that are always used together: a relay and its socket and clip, an MCB and its
 * auxiliary switch, a motor circuit breaker and its auxiliary contact… The page shows each part's Stock and how many
 * sets Stock can make, and drafts an Order Item for the whole set or only what is short.
 *
 * Sheet SetLibrary (main spreadsheet): ID | NAME | CAT | NOTE | PARTS (JSON) | SOURCE | UPDATED | BY
 *   PARTS = [{ role, per, unit, optional, opts: [{ code, kw, name }] }]
 *   one part = one place in the set; opts = models that can be used there (a CODE from Item stock, or only a model
 *   keyword `kw` when the shop has no CODE for it yet). The first part is the main item.
 * Everyone signed in reads; only Admin creates, edits, deletes, and imports the catalog suggestions (SET_SEEDS).
 */
var SETLIB_SHEET = 'SetLibrary', SETLIB_HEAD = ['ID', 'NAME', 'CAT', 'NOTE', 'PARTS', 'SOURCE', 'UPDATED', 'BY'], SETLIB_CK = 'setlib_v1';

function setLibSheet_() { return sheet_(SETLIB_SHEET, SETLIB_HEAD); }
function setLibRead_() {
  var C = CacheService.getScriptCache(), hit = C.get(SETLIB_CK);
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  var sh = setLibSheet_(), n = sh.getLastRow() - 1, out = [];
  if (n > 0) sh.getRange(2, 1, n, 8).getValues().forEach(function (r) {
    if (!r[0]) return;
    var parts = []; try { parts = JSON.parse(String(r[4] || '[]')) || []; } catch (e) { parts = []; }
    out.push({ id: String(r[0]), name: String(r[1] || ''), cat: String(r[2] || ''), note: String(r[3] || ''), parts: parts, source: String(r[5] || ''), updated: String(r[6] || ''), by: String(r[7] || '') });
  });
  out.sort(function (a, b) { return a.cat < b.cat ? -1 : a.cat > b.cat ? 1 : a.name < b.name ? -1 : a.name > b.name ? 1 : 0; });
  try { var s = JSON.stringify(out); if (s.length < 95000) C.put(SETLIB_CK, s, 21600); } catch (e) {}
  return out;
}
function getSetLib(token) { apiOnly_(); return { sets: setLibRead_(), admin: isAdmin_(token) }; }

function setLibStr_(v, n) { return String(v == null ? '' : v).replace(/[\r\n\t]+/g, ' ').trim().substring(0, n); }
function setLibClean_(s) {
  s = s || {};
  var name = setLibStr_(s.name, 80);
  if (!name) throw new Error('ใส่ชื่อชุด');
  var parts = (Array.isArray(s.parts) ? s.parts : []).slice(0, 12).map(function (p) {
    p = p || {};
    var per = Number(p.per); if (!(per > 0)) per = 1; per = Math.min(1000, Math.round(per * 1000) / 1000);
    var opts = (Array.isArray(p.opts) ? p.opts : []).slice(0, 8).map(function (o) {
      o = o || {};
      return { code: setLibStr_(o.code, 40).toUpperCase().replace(/\s+/g, ''), kw: setLibStr_(o.kw, 40), name: setLibStr_(o.name, 120) };
    }).filter(function (o) { return o.code || o.kw; });
    return { role: setLibStr_(p.role, 30) || 'อุปกรณ์', per: per, unit: setLibStr_(p.unit, 10).toUpperCase() || 'PCS', optional: !!p.optional, opts: opts };
  }).filter(function (p) { return p.opts.length; });
  if (!parts.length) throw new Error('ชุดต้องมีอุปกรณ์อย่างน้อย 1 รายการ (ใส่ CODE หรือรุ่น)');
  parts[0].optional = false;                                                   // the main item is never optional
  return { id: setLibStr_(s.id, 40), name: name, cat: setLibStr_(s.cat, 30) || 'อื่น ๆ', note: setLibStr_(s.note, 300), source: setLibStr_(s.source, 600), parts: parts };
}
// one set or a list (catalog import) · same id = replace
function saveSetLib(token, sets) { apiOnly_();
  requireAdmin_(token);
  var list = (Array.isArray(sets) ? sets : [sets]).slice(0, 60).map(setLibClean_), who = __admName || (typeof __user !== 'undefined' && __user && __user.n) || 'Admin';
  withLock_(function () {
    var sh = setLibSheet_(), n = sh.getLastRow() - 1, ids = n > 0 ? sh.getRange(2, 1, n, 1).getValues().map(function (r) { return String(r[0]); }) : [];
    list.forEach(function (s, i) {
      if (!s.id) s.id = 'S' + Date.now().toString(36).toUpperCase() + i;
      var row = [s.id, s.name, s.cat, s.note, JSON.stringify(s.parts), s.source, nowStr_(), who], at = ids.indexOf(s.id);
      if (JSON.stringify(s.parts).length > 45000) throw new Error('ชุด ' + s.name + ' ใหญ่เกินไป');
      var rg; if (at >= 0) rg = sh.getRange(at + 2, 1, 1, 8); else { rg = sh.getRange(sh.getLastRow() + 1, 1, 1, 8); ids.push(s.id); }
      fmtText_(rg); rg.setValues([row]);
    });
    CacheService.getScriptCache().remove(SETLIB_CK);
  });
  audit_([['Admin', list.length > 1 ? 'นำเข้าชุดอุปกรณ์' : 'บันทึกชุดอุปกรณ์', list.length > 1 ? list.length + ' ชุด' : list[0].id, list.map(function (s) { return s.name; }).slice(0, 5).join(', ')]]);
  return getSetLib(token);
}
function deleteSetLib(token, id) { apiOnly_();
  requireAdmin_(token);
  id = String(id || '');
  var name = '';
  withLock_(function () {
    var sh = setLibSheet_(), n = sh.getLastRow() - 1, rows = n > 0 ? sh.getRange(2, 1, n, 2).getValues() : [];
    for (var i = 0; i < rows.length; i++) if (String(rows[i][0]) === id) { name = String(rows[i][1]); sh.deleteRow(i + 2); break; }
    CacheService.getScriptCache().remove(SETLIB_CK);
  });
  if (!name) throw new Error('ไม่พบชุดนี้ (อาจถูกลบไปแล้ว)');
  audit_([['Admin', 'ลบชุดอุปกรณ์', id, name]]);
  return getSetLib(token);
}

// ── catalog suggestions (manufacturer catalogs / distributor listings, collected 11/2026) ──
// kw = model text searched in Item stock (CODE + description); '=SC-0' = that model exactly (not SC-03 / SC-05)
// ok: 1 = confirmed in a catalog, 0 = check before use
function getSetSeeds() { apiOnly_(); return SET_SEEDS; }
var SET_SEEDS = [
  { cat: 'Relay + Socket', name: 'OMRON MY4N (4 ขั้ว) + Socket', ok: 1, note: 'MY4N / MY4N-GS ใช้ Socket 14 ขา · PYF14A-E กำลังเลิกผลิต ทดแทนด้วย PYFZ-14-E · Clip PYC-A1 กันหลุด',
    src: 'https://www.ia.omron.com/data_pdf/cat/my-gs_ds_e_8_1_csm1051922.pdf | https://hken.rs-online.com/web/p/relay-sockets/0403285',
    parts: [{ role: 'หลัก', kw: ['MY4N', 'MY4IN'] }, { role: 'Socket', kw: ['PYFZ-14', 'PYF14A'] }, { role: 'Clip', kw: ['PYC-A1'], optional: 1 }] },
  { cat: 'Relay + Socket', name: 'OMRON MY2N (2 ขั้ว) + Socket', ok: 1, note: 'Socket 8 ขา PYF08A-E (ทดแทนด้วย PYFZ-08-E) · Clip PYC-A1',
    src: 'https://www.tme.com/il/en/details/pyf08a-e/electromagnetic-relays-accessories/omron/ | https://cn.element14.com/en-CN/omron-industrial-automation/pyca1/clip-my-ly-relay-2pole-pk2/dp/103923',
    parts: [{ role: 'หลัก', kw: ['MY2N', 'MY2IN'] }, { role: 'Socket', kw: ['PYFZ-08', 'PYF08A'] }, { role: 'Clip', kw: ['PYC-A1'], optional: 1 }] },
  { cat: 'Relay + Socket', name: 'OMRON LY2N + Socket', ok: 0, note: 'Socket PTF08A-E · Clip: ผู้จำหน่ายระบุ PYC-A1 ใช้กับ PTF08A-E/PTF14A-E — ควรตรวจกับ Datasheet',
    src: 'https://uk.rs-online.com/web/p/relay-sockets/0803376 | https://cn.element14.com/en-CN/omron-industrial-automation/pyca1/clip-my-ly-relay-2pole-pk2/dp/103923',
    parts: [{ role: 'หลัก', kw: ['LY2N'] }, { role: 'Socket', kw: ['PTF08A'] }, { role: 'Clip', kw: ['PYC-A1'], optional: 1 }] },
  { cat: 'Relay + Socket', name: 'OMRON LY4N + Socket', ok: 0, note: 'Socket PTF14A-E · Clip PYC-A1 (ควรตรวจ)',
    src: 'https://cn.element14.com/en-CN/omron-industrial-automation/pyca1/clip-my-ly-relay-2pole-pk2/dp/103923',
    parts: [{ role: 'หลัก', kw: ['LY4N'] }, { role: 'Socket', kw: ['PTF14A'] }, { role: 'Clip', kw: ['PYC-A1'], optional: 1 }] },
  { cat: 'Relay + Socket', name: 'OMRON G2R-2-S + Socket', ok: 0, note: 'Socket P2RF-08-E · Clip P2CM-S ระบุสำหรับ Socket แบบ Screwless — ควรตรวจกับรุ่นที่ใช้',
    src: 'https://docs.rs-online.com/dd86/A700000012687690.pdf',
    parts: [{ role: 'หลัก', kw: ['G2R-2-S'] }, { role: 'Socket', kw: ['P2RF-08'] }, { role: 'Clip', kw: ['P2CM'], optional: 1 }] },
  { cat: 'Relay + Socket', name: 'IDEC RR2KP (Latching) + Socket', ok: 1, note: 'Socket 11 ขา SR3P-05 · Hold-down spring SR3P-06F3 หรือ SR3B-02F1',
    src: 'https://au.element14.com/idec/sr3p-06f3/hold-down-clip-relay-socket/dp/4815478 | https://wolfautomation.com/sr3b-02f1-hold-down-pullover-spring',
    parts: [{ role: 'หลัก', kw: ['RR2KP'] }, { role: 'Socket', kw: ['SR3P-05'] }, { role: 'Hold-down', kw: ['SR3P-06F3', 'SR3B-02F1'], optional: 1 }] },
  { cat: 'Relay + Socket', name: 'IDEC RH2B + Socket', ok: 1, note: 'Socket SH2B-05 · Pullover spring SY4S-02F1',
    src: 'https://manualmachine.com/idec/sockets/16898656-selection-guide',
    parts: [{ role: 'หลัก', kw: ['RH2B'] }, { role: 'Socket', kw: ['SH2B-05'] }, { role: 'Hold-down', kw: ['SY4S-02F1'], optional: 1 }] },
  { cat: 'Relay + Socket', name: 'IDEC RU4S + Socket', ok: 1, note: 'Socket SY4S-05 (สกรู) หรือ SU4S-11L (Spring clamp)',
    src: 'https://manualmachine.com/idec/ruseries/16824373-catalog',
    parts: [{ role: 'หลัก', kw: ['RU4S'] }, { role: 'Socket', kw: ['SY4S-05', 'SU4S-11L'] }] },
  { cat: 'Relay + Socket', name: 'IDEC RJ2S + Socket', ok: 1, note: 'Socket 8 ขา SJ2S-05B',
    src: 'https://images.kleen-ritecorp.com/Documents/Manuals/IDEC/IDEC2052/IDEC Socket Data Sheet.pdf',
    parts: [{ role: 'หลัก', kw: ['RJ2S'] }, { role: 'Socket', kw: ['SJ2S-05B'] }] },
  { cat: 'Relay + Socket', name: 'Finder 62.33 + Socket', ok: 1, note: 'Socket 92.03 (92.03SMA มี Clip โลหะ 092.71 ในชุด)',
    src: 'https://findernet.com/en/uk/series/92-series-sockets-for-62-series-relays | https://www.digikey.co.uk/en/products/detail/finder-relays-inc/092-71/10055752',
    parts: [{ role: 'หลัก', kw: ['62.33'] }, { role: 'Socket', kw: ['92.03'] }, { role: 'Clip', kw: ['092.71'], optional: 1 }] },
  { cat: 'Relay + Socket', name: 'Finder 55.34 + Socket', ok: 1, note: 'Socket 94.04 · Clip โลหะ 094.71 (94.04SMA) หรือพลาสติก 094.91.3 (94.04SPA)',
    src: 'https://findernet.com/en/usa/series/94-series-sockets-for-55-and-85-series-relays',
    parts: [{ role: 'หลัก', kw: ['55.34'] }, { role: 'Socket', kw: ['94.04'] }, { role: 'Clip', kw: ['094.71', '094.91'], optional: 1 }] },
  { cat: 'Relay + Socket', name: 'Schneider RXM4 / RXM2 + Socket', ok: 1, note: 'Socket RXZE2M114(M) · Clamp โลหะ RXZ400 หรือพลาสติก RXZR335',
    src: 'https://www.se.com/us/en/product/RXZ400/maintaining-clamp-harmony-for-rxm-sockets-metal',
    parts: [{ role: 'หลัก', kw: ['RXM4', 'RXM2'] }, { role: 'Socket', kw: ['RXZE2M114'] }, { role: 'Clamp', kw: ['RXZ400', 'RXZR335'], optional: 1 }] },
  { cat: 'Timer + Socket', name: 'OMRON H3CR-A8 + Socket', ok: 1, note: 'Socket 8 ขา P2CF-08-E (ราง DIN)',
    src: 'https://uk.rs-online.com/p/panel-meter-accessories/0329468',
    parts: [{ role: 'หลัก', kw: ['H3CR-A8'] }, { role: 'Socket', kw: ['P2CF-08'] }] },
  { cat: 'Timer + Socket', name: 'OMRON H3Y-4 + Socket', ok: 0, note: 'ใช้ Socket 14 ขาแบบ PYF / PY เดียวกับ MY4 (ควรตรวจ)',
    src: 'https://www.ia.omron.com/products/family/811/lineup.html',
    parts: [{ role: 'หลัก', kw: ['H3Y-4'] }, { role: 'Socket', kw: ['PYFZ-14', 'PYF14A'] }] },
  { cat: 'Timer + Socket', name: 'Autonics MS4SC + Socket', ok: 0, note: 'Socket 8 ขา PG-08 (หลังตู้) หรือ PS-08 (ราง DIN) — ยังไม่พบเอกสารยืนยัน ควรตรวจ',
    src: 'https://baa.vn/document/ps-and-pg-series-autonics-de-ro-le-4992-catalog1.pdf',
    parts: [{ role: 'หลัก', kw: ['MS4SC'] }, { role: 'Socket', kw: ['PS-08', 'PG-08', 'TP48X'] }] },
  { cat: 'MCB', name: 'EATON MCB PLS6 + อุปกรณ์เสริม', ok: 1, note: 'Aux switch ZP-IHK (1NO+1NC) หรือ ZP-WHK (1CO) ติดด้านซ้าย · Shunt trip ZP-ASA',
    src: 'https://www.eaton.com/mx/en-us/skuPage.286052.html | https://www.eaton.com/dk/da-dk/skuPage.286053.html | https://pdf.directindustry.com/pdf/eaton/accessories-rcds-mcbs-combined-rcd-mcb-devices-motor-starters-power-limiters/5067-564558.html',
    parts: [{ role: 'หลัก', kw: ['PLS6'] }, { role: 'Aux switch', kw: ['ZP-IHK', 'ZP-WHK'], optional: 1 }, { role: 'Shunt trip', kw: ['ZP-ASA'], optional: 1 }] },
  { cat: 'Motor CB', name: 'EATON PKZM0 + Aux contact', ok: 1, note: 'NHI11-PKZ0 ติดด้านข้าง หรือ NHI-E-11-PKZ0 ติดด้านหน้า (ความกว้าง 45 mm เท่าเดิม)',
    src: 'https://www.eaton.com/tr/tr-tr/skuPage.072896.html | https://www.eaton.com/sg/en-us/skuPage.082882.html',
    parts: [{ role: 'หลัก', kw: ['PKZM0'] }, { role: 'Aux contact', kw: ['NHI11-PKZ0', 'NHI-E-11-PKZ0'] }] },
  { cat: 'Motor CB', name: 'Schneider GV2ME + Aux contact', ok: 1, note: 'GVAE11 ติดด้านหน้า (1NO+1NC) · GVAN11 ติดด้านข้าง (ควรตรวจ)',
    src: 'https://www.se.com/be/en/product/GVAE11/auxiliary-contact-blocktesys-deca1no+1ncfront-mountingfor-gv2',
    parts: [{ role: 'หลัก', kw: ['GV2ME'] }, { role: 'Aux contact', kw: ['GVAE11', 'GVAN11'] }] },
  { cat: 'Motor CB', name: 'ABB MS116 / MS132 + Aux contact', ok: 1, note: 'HKF1-11 ติดด้านหน้า 1NO+1NC',
    src: 'https://www.standardelectricsupply.com/userfiles/documents/ad/a/abb_hkf1_11_data_sheet.pdf',
    parts: [{ role: 'หลัก', kw: ['MS116', 'MS132'] }, { role: 'Aux contact', kw: ['HKF1-11'] }] },
  // ── magnetic contactors: thermal overload relay, auxiliary contact, covers ──
  { cat: 'Magnetic contactor', name: 'FUJI SC-03 + Thermal relay', ok: 1, note: 'Thermal TR-0N/3 ติดใต้ Contactor · Aux SZ-A11 (หน้า 1NO+1NC) หรือ SZ-AS1 (ข้าง) · ฝาครอบขั้ว SZ-T1 · Fuji ประกาศเลิกผลิต SC-03/0/05/4-1 แล้ว (ทดแทนด้วย SC09X…)',
    src: 'https://onlinecomponents.com/productfiles/mf-fji/aeh561.pdf | https://americas.fujielectric.com/wp-content/uploads/2025/09/Discontinuation-notice-for-Orange-Line-and-Odyssey-complete-1.pdf',
    parts: [{ role: 'หลัก', kw: ['SC-03'] }, { role: 'Thermal relay', kw: ['TR-0N', 'TR-ON'] }, { role: 'Aux contact', kw: ['SZ-A11', 'SZ-AS1'], optional: 1 }, { role: 'ฝาครอบขั้ว', kw: ['SZ-T1'], optional: 1 }] },
  { cat: 'Magnetic contactor', name: 'FUJI SC-0 + Thermal relay', ok: 1, note: 'Thermal TR-0N/3 · Aux SZ-A11 / SZ-AS1 · ฝาครอบขั้ว SZ-T1',
    src: 'https://onlinecomponents.com/productfiles/mf-fji/aeh561.pdf',
    parts: [{ role: 'หลัก', kw: ['=SC-0'] }, { role: 'Thermal relay', kw: ['TR-0N', 'TR-ON'] }, { role: 'Aux contact', kw: ['SZ-A11', 'SZ-AS1'], optional: 1 }, { role: 'ฝาครอบขั้ว', kw: ['SZ-T1'], optional: 1 }] },
  { cat: 'Magnetic contactor', name: 'FUJI SC-05 + Thermal relay', ok: 1, note: 'Thermal TR-0N/3 · Aux SZ-A11 / SZ-AS1 · ฝาครอบขั้ว SZ-T2',
    src: 'https://onlinecomponents.com/productfiles/mf-fji/aeh561.pdf',
    parts: [{ role: 'หลัก', kw: ['SC-05'] }, { role: 'Thermal relay', kw: ['TR-0N', 'TR-ON'] }, { role: 'Aux contact', kw: ['SZ-A11', 'SZ-AS1'], optional: 1 }, { role: 'ฝาครอบขั้ว', kw: ['SZ-T2'], optional: 1 }] },
  { cat: 'Magnetic contactor', name: 'FUJI SC-4-0 / SC-4-1 + Thermal relay', ok: 0, note: 'Thermal TR-5-1N/3 · ฝาครอบขั้ว SZ-T3 · Aux SZ-A11 / SZ-AS1 (ควรตรวจกับรุ่นที่ใช้)',
    src: 'https://onlinecomponents.com/productfiles/mf-fji/aeh561.pdf',
    parts: [{ role: 'หลัก', kw: ['SC-4-1', 'SC-4-0'] }, { role: 'Thermal relay', kw: ['TR-5-1N'] }, { role: 'Aux contact', kw: ['SZ-A11', 'SZ-AS1'], optional: 1 }, { role: 'ฝาครอบขั้ว', kw: ['SZ-T3'], optional: 1 }] },
  { cat: 'Magnetic contactor', name: 'FUJI SC-E02…E05 + Thermal relay', ok: 1, note: 'Thermal TK-E02 สำหรับ SC-E02/E03/E04/E05 (TK-E2 สำหรับ SC-E1/E2) · Aux หน้า SZ-A11T / SZ-A22T',
    src: 'https://www.fujielectric.com/fcs/support_contact/support/standard/box/pdf/TUV/TUV.pdf | https://cdn.automationdirect.com/static/specs/cutsheet/TK-E02-260_cutsheet.pdf',
    parts: [{ role: 'หลัก', kw: ['SC-E02', 'SC-E03', 'SC-E04', 'SC-E05'] }, { role: 'Thermal relay', kw: ['TK-E02'] }, { role: 'Aux contact', kw: ['SZ-A11', 'SZ-A22'], optional: 1 }] },
  { cat: 'Magnetic contactor', name: 'Mitsubishi S-T10 / S-T12 / S-T20 + Thermal relay', ok: 1, note: 'Thermal TH-T18 · Aux UN-AX2 / UN-AX4 · Surge absorber UN-SA (ตามแรงดัน Coil)',
    src: 'https://th.misumi-ec.com/en/vona2/detail/222005725443 | https://www.mitsubishielectric.co.jp/fa/products/lvd/lvsw/support/QR/pdf/st10st12st20.pdf',
    parts: [{ role: 'หลัก', kw: ['S-T10', 'S-T12', 'S-T20'] }, { role: 'Thermal relay', kw: ['TH-T18'] }, { role: 'Aux contact', kw: ['UN-AX2', 'UN-AX4'], optional: 1 }, { role: 'Surge absorber', kw: ['UN-SA'], optional: 1 }] },
  { cat: 'Magnetic contactor', name: 'Schneider LC1D09…D38 + Thermal relay', ok: 0, note: 'Thermal LRD (ต่อใต้ LC1D09–D38) · RC suppressor LAD4RC… ตามแรงดัน Coil (LAD4RCE 24–48VAC) · Aux หน้า LADN11 (ควรตรวจ)',
    src: 'https://www.se.com/il/en/product/LRD08/tesys-lrd-thermal-overload-relays-2-5-4-a-class-10a | https://www.se.com/us/en/faqs/FA342734',
    parts: [{ role: 'หลัก', kw: ['LC1D09', 'LC1D12', 'LC1D18', 'LC1D25', 'LC1D32', 'LC1D38'] }, { role: 'Thermal relay', kw: ['LRD'] }, { role: 'Aux contact', kw: ['LADN11'], optional: 1 }, { role: 'Suppressor', kw: ['LAD4RC'], optional: 1 }] },
  // ── terminals: end plate / cover, end clamps, partition, markers (per 10 terminals) ──
  { cat: 'Terminal', name: 'TOGI PT30 · ชุดละ 10 ตัว', ok: 1, note: 'ตามแผ่นข้อมูล PT series: PT-30 + End plate PT-30L2 + ฝาครอบ CA-36 + ป้ายชื่อ AM-10',
    src: 'https://www.togi.co.jp/en/pdf_file/terminalblock/series/PT_series.pdf',
    parts: [{ role: 'หลัก', kw: ['=PT30', '=PT-30'], per: 10 }, { role: 'End plate', kw: ['PT30L2', 'PT-30L2'] }, { role: 'ฝาครอบ', kw: ['CA-36'], optional: 1 }, { role: 'ป้ายชื่อ', kw: ['AM-10'], optional: 1 }] },
  { cat: 'Terminal', name: 'TOGI PT40 · ชุดละ 10 ตัว', ok: 0, note: 'End plate PT40L2 · ฝาครอบ CA-40 (ไม่พบในเอกสาร ใช้ตามชุด PT50 — ควรตรวจ)',
    src: 'https://www.togi.co.jp/en/pdf_file/terminalblock/series/Accessories - 1.pdf',
    parts: [{ role: 'หลัก', kw: ['=PT40', '=PT-40'], per: 10 }, { role: 'End plate', kw: ['PT40L2', 'PT-40L2'] }, { role: 'ฝาครอบ', kw: ['CA-40'], optional: 1 }] },
  { cat: 'Terminal', name: 'Phoenix UK 5 N · ชุดละ 10 ตัว', ok: 1, note: 'End cover D-UK 4/10 · End clamp E/UK 2 ตัว · Partition ATP-UK (คั่นกลุ่ม) · Marker ZB 5 (ควรตรวจ)',
    src: 'https://au.element14.com/phoenix-contact/3003020/end-cover-d-uk4-10/dp/3042250 | https://au.element14.com/phoenix-contact/e-uk/end-clamp-din-rail-mount/dp/179044',
    parts: [{ role: 'หลัก', kw: ['UK 5 N', 'UK5N', 'UK 5N'], per: 10 }, { role: 'End cover', kw: ['D-UK 4/10', 'D-UK4/10'] }, { role: 'End clamp', kw: ['E/UK'], per: 2 }, { role: 'Partition', kw: ['ATP-UK'], optional: 1 }, { role: 'Marker', kw: ['ZB 5', 'ZB5'], per: 10, optional: 1 }] },
  { cat: 'Terminal', name: 'Phoenix UK 6 N · ชุดละ 10 ตัว', ok: 1, note: 'End cover D-UK 4/10 (ใช้กับ UK 3 N / 5 N / 6 N / 10 N) · End clamp E/UK 2 ตัว · Marker ZB 6',
    src: 'https://www.ahlsell.se/en/products/electric/automation-29-45/29-terminals-labelling-ducts/terminal-blocks/end-trays/2911301',
    parts: [{ role: 'หลัก', kw: ['UK 6 N', 'UK6N', 'UK 6N'], per: 10 }, { role: 'End cover', kw: ['D-UK 4/10', 'D-UK4/10'] }, { role: 'End clamp', kw: ['E/UK'], per: 2 }, { role: 'Marker', kw: ['ZB 6', 'ZB6'], per: 10, optional: 1 }] },
  { cat: 'Terminal', name: 'USST · ชุดละ 10 ตัว', ok: 0, note: 'ไม่พบเอกสารผู้ผลิตของ USST — ระบบค้นจากชื่อใน Stock · ตรวจ End cover / End clamp ที่ใช้จริงแล้วแก้ชุด',
    src: 'ยังไม่พบเอกสาร (ตรวจรุ่นกับผู้ขาย)',
    parts: [{ role: 'หลัก', kw: ['USST'], per: 10 }, { role: 'End cover', kw: ['D-USST'] }, { role: 'End clamp', kw: ['E/UK', 'CLIPFIX'], per: 2 }] },
  // ── test terminals ──
  { cat: 'Test terminal', name: 'Test terminal Type B (CT / PT) · 1 กลุ่ม', ok: 1, note: 'สัดส่วนจากใบ Estimate cost ที่ใช้จริง (CTT 8 · PTT 4 · Short bar 8 · Cover 4) = ต่อกลุ่ม CTT 2 · PTT 1 · Short bar 2 · Cover 1',
    src: 'ใบ Estimate cost CONTROL WIRING (B-CTT / B-PTT / B-SB / Cover B-C)',
    parts: [{ role: 'หลัก', kw: ['B-CTT'], per: 2 }, { role: 'PT terminal', kw: ['B-PTT'] }, { role: 'Short bar', kw: ['B-SB'], per: 2 }, { role: 'Cover', kw: ['=B-C'] }] },
  { cat: 'Test terminal', name: 'GE PK-2 Test block + Test plug', ok: 0, note: 'PK-2 = Test block 4 / 6 ขั้ว มีฝาครอบในตัว และ Contact ลัดวงจร CT เมื่อถอดฝา · Test plug ใช้ตอนทดสอบ (ตรวจรุ่น Plug กับเอกสาร GE)',
    src: 'https://www.gevernova.com/grid-solutions/products/brochures/pk2.pdf',
    parts: [{ role: 'หลัก', kw: ['PK-2', 'PK2'] }, { role: 'Test plug', kw: ['PK-2 PLUG', 'TEST PLUG'], optional: 1 }] },
  { cat: 'Terminal', name: 'TOGI PT50 · ชุดละ 10 ตัว', ok: 1, note: '1 ชุด = Terminal 10 ตัว + End plate 1 + ฝาครอบ 1 เส้น (ตามใบ Estimate cost ที่ใช้จริง)',
    src: 'ใบ Estimate cost CONTROL WIRING (PT50 / PT50L2 / CA-40)',
    parts: [{ role: 'หลัก', kw: ['PT50'], per: 10 }, { role: 'End plate', kw: ['PT50L2'] }, { role: 'ฝาครอบ', kw: ['CA-40'], optional: 1 }] },
  { cat: 'Terminal', name: 'Phoenix URTK/S · ชุดละ 10 ตัว', ok: 0, note: 'End cover D-URTK/S · End bracket CLIPFIX 35 (ควรตรวจรหัสฝาปิดกับรุ่นที่ใช้)',
    src: 'https://www.phoenixcontact.com/ (URTK/S 0311087 · D-URTK/S-BEN 0308029 · CLIPFIX 35 3022218)',
    parts: [{ role: 'หลัก', kw: ['URTK/S'], per: 10 }, { role: 'End cover', kw: ['D-URTK'] }, { role: 'End bracket', kw: ['CLIPFIX'], per: 2 }] }
];
