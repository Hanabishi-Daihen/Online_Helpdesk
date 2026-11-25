/**
 * Booking BOM × Shipment — has MF already issued a booked project's parts?
 *
 * MF issues the parts of a BOM to department CONTROL WIRING; the official record is the Shipment file
 * (Estimate Cost). For every open BOM project the BOM's CODEs are compared with what Shipment issued to the
 * project's CATEGORY after the BOM was imported:
 *   full = CODEs issued ≥ the BOM qty · pct = full / CODEs in the BOM
 * pct ≥ 80 % → "เบิกแล้ว ยังไม่ปิด" (alert on the page, the menu badge and My Day); still open 14 days after the
 * first alert → closed automatically, which returns its bookings to Stock.
 *
 * Units of one project number share a CATEGORY (PP1048A02 and PP1048B02 → 1048P2). A Shipment line whose
 * DOCNUM names the full project number (…/PP1048B02-…) goes to that unit only (or to none when that unit has no
 * open BOM); other lines are shared, CODE by CODE, in need-date order — each unit takes what it still needs.
 * Returns (positive COST) are taken back from the unit that received the CODE last.
 */
var BOM_SHIP_PCT = 0.8, BOM_SHIP_DAYS = 14, BOM_SHIP_DEPT = /CONTROL\s*WIRING/i;

// imported "dd/MM/yyyy HH:mm:ss" → yyyymmdd
function bomShipD8_(s) { var m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(String(s || '')); return m ? m[3] + ('0' + m[2]).slice(-2) + ('0' + m[1]).slice(-2) : ''; }

// ps = [{ project, imported, t (need date ms or null), items: [{ code, qty }] }] (open projects)
// → { PROJECT: { n, full, part, pct, last (dd/mm/yyyy), lines, tagged, shared, by: { CODE: issued } } } · null = no Shipment data
function bomShipMatch_(ps) {
  var m = null; try { m = shipMeta_(); } catch (e) { console.warn('bomShipMatch_ meta', e); }
  if (!m) return null;
  var byCat = {}, out = {};
  ps.forEach(function (p) { var c = shipCat_(p.project); if (c) (byCat[c] = byCat[c] || []).push(p); });
  Object.keys(byCat).forEach(function (cat) {
    var G = byCat[cat].slice().sort(function (a, b) { return (a.t == null ? 9e15 : a.t) - (b.t == null ? 9e15 : b.t) || String(a.project).localeCompare(String(b.project)); });
    var keys = {}, st = {};
    G.forEach(function (p) {
      var k = bomKey_(p.project), need = {};
      (p.items || []).forEach(function (it) { var c = ordCodeKey_(it.code), q = Number(it.qty) || 0; if (c && q > 0) need[c] = (need[c] || 0) + q; });
      keys[k] = 1; st[k] = { p: p, d8: bomShipD8_(p.imported), need: need, got: {}, last: '', lines: 0, tagged: 0 };
    });
    var lines = []; if (m.idx[cat]) { try { lines = shipLines_(m, cat); } catch (e) { console.warn('bomShipMatch_ ' + cat, e); } }
    var any = /^PP\d{4}[A-Z]{1,2}\d{2}/, tagOf = function (f) {
      var docs = String(f[3] || '').toUpperCase().split(/[\/\s,;]+/);
      for (var i = 0; i < docs.length; i++) { var x = any.exec(docs[i].replace(/[^A-Z0-9]/g, '')); if (x && shipCat_(x[0]) === cat) return bomKey_(x[0]); }
      return '';
    };
    // lines naming the unit first (they are certain), then the shared ones; each in date order
    var L = lines.filter(function (f) { return BOM_SHIP_DEPT.test(f[4] || ''); }).map(function (f) { return { f: f, tag: tagOf(f) }; });
    L.sort(function (a, b) { return (a.tag ? 0 : 1) - (b.tag ? 0 : 1) || (a.f[2] < b.f[2] ? -1 : a.f[2] > b.f[2] ? 1 : 0); });
    L.forEach(function (x) {
      var f = x.f, tag = x.tag, code = ordCodeKey_(f[0]), d8 = String(f[2] || ''), q = Number(f[6]) || 0, cost = Number(f[9]) || 0;
      q = cost > 0 ? -Math.abs(q) : Math.abs(q);                       // a return has a positive cost (as the Estimate Cost page)
      if (!code || !q) return;
      if (tag && !keys[tag]) return;                                    // issued to another unit of this project number
      var take = function (s, n) { s.got[code] = Math.round(((s.got[code] || 0) + n) * 1000) / 1000; s.lines++; if (d8 > s.last) s.last = d8; };
      if (tag) { var s0 = st[tag]; if (s0.d8 && d8 < s0.d8) return; s0.tagged++; take(s0, q); return; }
      var el = G.map(function (p) { return st[bomKey_(p.project)]; }).filter(function (s) { return s.need[code] && (!s.d8 || d8 >= s.d8); });
      if (!el.length) return;
      if (q > 0) {
        for (var j = 0; j < el.length && q > 0; j++) {
          var left = el[j].need[code] - (el[j].got[code] || 0); if (left <= 0) continue;
          var u = Math.min(left, q); take(el[j], u); q -= u;
        }
        if (q > 0) take(el[el.length - 1], q);                          // more than the BOMs need: the last unit
      } else {
        for (var r = el.length - 1; r >= 0 && q < 0; r--) {
          var has = el[r].got[code] || 0; if (has <= 0) continue;
          var b = Math.max(q, -has); take(el[r], b); q -= b;
        }
      }
    });
    G.forEach(function (p) {
      var s = st[bomKey_(p.project)], codes = Object.keys(s.need), full = 0, part = 0;
      codes.forEach(function (c) { var g = s.got[c] || 0; if (g >= s.need[c] - 1e-9) full++; else if (g > 0) part++; });
      out[bomKey_(p.project)] = { n: codes.length, full: full, part: part, pct: codes.length ? Math.round(full / codes.length * 1000) / 1000 : 0,
        last: s.last ? shipShow_(s.last) : '', lines: s.lines, tagged: s.tagged, shared: G.length > 1, by: s.got, cat: cat };
    });
  });
  return out;
}

// called by bomCompute_ for the open projects: attaches p.ship, keeps the first-alert dates, closes after 14 days.
// → names of the projects it closed
function bomShipStep_(projs, itemsBy, sched, cfg, today) {
  var open = projs.filter(function (p) { return p.status === 'active'; });
  var res = bomShipMatch_(open.map(function (p) { var n = bomNeed_(p, sched, cfg); return { project: p.project, imported: p.imported, t: n.t, items: itemsBy[p.project] || [] }; }));
  if (!res) return [];
  var seen = {}; try { seen = JSON.parse(schedProp_('BOM_SHIP_SEEN') || '{}') || {}; } catch (e) {}
  var keep = {}, closed = [], sh = null, now = Date.now();
  open.forEach(function (p) {
    var k = bomKey_(p.project), r = res[k]; if (!r) return;
    p.ship = r;
    if (r.pct < BOM_SHIP_PCT || !r.n) return;
    var t0 = Number(seen[k]) || now; keep[k] = t0;
    r.alert = bomFmt_(t0); r.closeOn = bomFmt_(t0 + BOM_SHIP_DAYS * 864e5);
    if (now - t0 >= BOM_SHIP_DAYS * 864e5) {
      sh = sh || ss_().getSheetByName(BOM_PROJ_SHEET);
      var why = nowStr_() + ' (อัตโนมัติ: Shipment เบิกแล้ว ' + Math.round(r.pct * 100) + '%)';
      sh.getRange(p.row, 7, 1, 3).setValues([['closed', p.imported, why]]); p.status = 'closed'; p.closed = why; closed.push(p.project);
      delete keep[k];
    }
  });
  if (JSON.stringify(keep) !== JSON.stringify(seen)) schedProp_('BOM_SHIP_SEEN', JSON.stringify(keep));
  if (closed.length) audit_([['System', 'ปิดโปรเจกต์ BOM อัตโนมัติ (Shipment เบิกแล้วเกิน ' + BOM_SHIP_DAYS + ' วัน)', '', closed.join(', ')]]);
  return closed;
}
