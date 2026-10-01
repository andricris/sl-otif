/* SL & OTIF core — shared by the browser app and the Node test harness.
   - keeps the raw worksheet AS-IS (headers + formatting + dates untouched)
   - APPENDS helper columns written as LIVE FORMULAS (DATEVALUE-based)
   - adds a FORMULA-DRIVEN SUMMARY (SUMIFS) so every result is traceable
   Each formula also carries a cached `result`, so values show immediately. */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.SLOTIF = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";
  var MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  var NUMACC = '_(* #,##0_);_(* \\(#,##0\\);_(* "-"??_);_(@_)';
  var DATEFMT = "yyyy-mm-dd";
  var HFILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF2CC" } };
  var THIN = { style: "thin", color: { argb: "FFBFBFBF" } };
  var BORDER = { left: THIN, right: THIN, top: THIN, bottom: THIN };

  var FIELDS = {
    pcode: ["principal code","principal _ code"], pname: ["principal name","principal _ name"],
    periode: ["periode","period"], contract: ["contract"], po: ["po number"], note: ["po note"],
    line: ["line no","line _ no"], rel: ["rel no","rel _ no"], podate: ["po date","po _ date"],
    prod: ["product code","product _ code"], desc: ["description"], poqty: ["po qty"],
    hna: ["hna um"], povalue: ["po value"], plan: ["planned receipt date","planned receipt _ date"],
    status: ["po line status","po line _ status"], lpb: ["lpb number"], arrdate: ["arrival date"],
    arrqty: ["arrival qty"], arrvalue: ["arrival value"], no: ["no"]
  };

  function norm(s){ return String(s == null ? "" : s).replace(/\s+/g, " ").trim().toLowerCase(); }
  function cellVal(cell){
    if (!cell) return null;
    var v = cell.value;
    if (v == null) return null;
    if (typeof v === "object" && !(v instanceof Date)) {
      if (v.richText) return v.richText.map(function (t){ return t.text; }).join("");
      if (v.text != null && v.hyperlink != null) return v.text;
      if (v.result != null) return v.result;
      if (v.error) return null;
    }
    return v;
  }
  function toDate(v){
    if (v == null || v === "") return null;
    if (v instanceof Date) return isNaN(v.getTime()) ? null : new Date(Date.UTC(v.getFullYear(), v.getMonth(), v.getDate()));
    if (typeof v === "number") return null;
    var s = String(v).trim();
    var m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
    if (m) return new Date(Date.UTC(+m[1], +m[2]-1, +m[3]));
    m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);
    if (m) return new Date(Date.UTC(+m[3], +m[2]-1, +m[1]));
    return null;
  }
  function num(v){ if (v == null || v === "") return 0; if (typeof v === "number") return v;
    if (typeof v === "object" && v.result != null) v = v.result;
    var n = parseFloat(String(v).replace(/,/g, "").trim()); return isNaN(n) ? 0 : n; }
  function addDays(d, n){ return new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate() + n)); }
  function dayDiff(a, b){ return Math.round((Date.UTC(a.getUTCFullYear(),a.getUTCMonth(),a.getUTCDate()) - Date.UTC(b.getUTCFullYear(),b.getUTCMonth(),b.getUTCDate()))/86400000); }
  function colLetter(n){ var s = ""; while (n > 0){ var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; }
  function clone(o){ return o == null ? o : JSON.parse(JSON.stringify(o)); }
  function styleOf(cell){ var s = cell.style || {}; return { font: clone(s.font), fill: clone(s.fill), border: clone(s.border), alignment: clone(s.alignment), numFmt: s.numFmt }; }
  function applyStyle(cell, s){ if (!s) return; if (s.font) cell.font = s.font; if (s.fill) cell.fill = s.fill;
    if (s.border) cell.border = s.border; if (s.alignment) cell.alignment = s.alignment; if (s.numFmt) cell.numFmt = s.numFmt; }

  function pickSheet(wb){
    var found = null;
    wb.worksheets.forEach(function (s){ if (norm(s.name) === "worksheet") found = s; });
    return found || wb.worksheets[wb.worksheets.length - 1];
  }

  function parse(wb, tol){
    var ws = pickSheet(wb), hrow = 5, c, r;
    for (r = 1; r <= Math.min(ws.rowCount, 25); r++){
      var vals = [];
      for (c = 1; c <= ws.columnCount; c++){ var cl = ws.getRow(r).getCell(c); if (cl.value != null) vals.push(norm(cellVal(cl))); }
      if (vals.indexOf("po qty") >= 0 && vals.some(function (v){ return v.indexOf("principal") >= 0; })){ hrow = r; break; }
    }
    var cmap = {};
    for (c = 1; c <= ws.columnCount; c++){
      var h = norm(cellVal(ws.getRow(hrow).getCell(c)));
      for (var k in FIELDS){ if (cmap[k] == null && FIELDS[k].indexOf(h) >= 0) cmap[k] = c; }
    }
    var poCol = cmap.po || 5, last = hrow;
    for (r = hrow + 1; r <= ws.rowCount; r++){ var pv = cellVal(ws.getRow(r).getCell(poCol)); if (pv != null && pv !== "") last = r; }
    var rows = [];
    for (r = hrow + 1; r <= last; r++){
      var rr = ws.getRow(r); var g = function (k){ return cmap[k] ? cellVal(rr.getCell(cmap[k])) : null; };
      var plan = toDate(g("plan")), arr = toDate(g("arrdate"));
      var gap = (!plan || !arr) ? -1 : dayDiff(addDays(plan, tol), arr);
      rows.push({ row: r, plan: plan, arr: arr, gap: gap, on_time: gap >= 0, poqty: num(g("poqty")),
        arrqty: num(g("arrqty")), pname: g("pname"), tol_plus: plan ? addDays(plan, tol) : null });
    }
    return { ws: ws, hrow: hrow, last: last, cmap: cmap, rows: rows,
      planL: colLetter(cmap.plan), arrL: colLetter(cmap.arrdate),
      poL: colLetter(cmap.poqty), arrqL: colLetter(cmap.arrqty) };
  }

  function appendHelpers(ws, hrow, last, rows, tol, planL, arrL){
    var start = ws.columnCount + 1;
    var periodL = colLetter(start), tolL = colLetter(start + 1), gapL = colLetter(start + 2), remL = colLetter(start + 3);
    var heads = ["PERIODE", "Planned receipt + " + tol + " hari", "GAP", "Remarks"];
    var hstyle = styleOf(ws.getRow(hrow).getCell(1));
    for (var i = 0; i < heads.length; i++){
      var hc = ws.getRow(hrow).getCell(start + i);
      hc.value = heads[i]; applyStyle(hc, hstyle);
      hc.font = clone(hstyle.font) || {}; hc.font.bold = true;
      hc.fill = clone(HFILL); hc.border = clone(BORDER);
      hc.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      ws.getColumn(start + i).width = i === 0 ? 11 : 14;
    }
    for (var j = 0; j < rows.length; j++){
      var x = rows[j], R = x.row, left = styleOf(ws.getRow(R).getCell(start - 1));
      var put = function (col, formula, result, fmt){
        var cell = ws.getRow(R).getCell(col);
        cell.value = { formula: formula, result: result };
        applyStyle(cell, left); cell.border = clone(BORDER);
        cell.alignment = { horizontal: "center", vertical: "middle" };
        if (fmt) cell.numFmt = fmt;
      };
      var mName = x.plan ? MONTHS[x.plan.getUTCMonth()] : "";
      put(start,     'IF(' + planL + R + '="","",TEXT(DATEVALUE(' + planL + R + '),"MMM"))', mName);
      put(start + 1, 'IF(' + planL + R + '="","",DATEVALUE(' + planL + R + ')+' + tol + ')', x.tol_plus || "", DATEFMT);
      put(start + 2, 'IF(OR(' + planL + R + '="",' + arrL + R + '=0),-1,' + tolL + R + '-DATEVALUE(' + arrL + R + '))', x.gap);
      put(start + 3, 'IF(' + gapL + R + '>=0,"On Time","Late")', x.on_time ? "On Time" : "Late");
    }
    return { periodL: periodL, tolL: tolL, gapL: gapL, remL: remL, start: start };
  }

  function buildSummary(wb, order, agg, grand, title, year, wsName, poL, arrL, periodL, remL, first, last){
    var existing = null;
    wb.worksheets.forEach(function (s){ if (norm(s.name) === "summary") existing = s; });
    if (existing) wb.removeWorksheet(existing.id);
    var sm = wb.addWorksheet("SUMMARY");
    [13.1,16.3,19.3,14.1,20.3,19.1,27.7,9.1,9.1].forEach(function (w, i){ sm.getColumn(i + 1).width = w; });
    sm.getCell("A1").value = title; sm.getCell("A1").font = { bold: true };
    sm.getCell("A2").value = year;
    var labels = { B3:"Column Labels", H3:"%", I3:"%", B4:"Late", D4:"On Time",
      F4:"Total Sum of PO QTY", G4:"Total Sum of ARRIVAL QTY", H4:"FULFIL", I4:"OTIF",
      A5:"Row Labels", B5:"Sum of PO QTY", C5:"Sum of ARRIVAL QTY", D5:"Sum of PO QTY",
      E5:"Sum of ARRIVAL QTY", H5:"%", I5:"%" };
    Object.keys(labels).forEach(function (a){ sm.getCell(a).value = labels[a]; });
    for (var row = 3; row <= 5; row++){
      for (var col = 2; col <= 9; col++){
        var cc = sm.getRow(row).getCell(col);
        cc.border = clone(BORDER);
        if (cc.value != null && cc.value !== "%") cc.font = { bold: true };
        cc.alignment = { horizontal: cc.value === "%" ? "center" : "left" };
      }
    }
    var q = "'" + String(wsName).replace(/'/g, "''") + "'!";
    var rng = function (L){ return q + "$" + L + "$" + first + ":$" + L + "$" + last; };
    var fmt = function (L){ return NUMACC; };

    function rowFormula(r, monthCell, labelWhenNoMonth){
      var crit = monthCell ? ", $A" + monthCell : "";
      var c2 = monthCell ? ", $" + periodL + "blank" : "";
      var sumifs = function (valL, remark){
        return monthCell
          ? "SUMIFS(" + rng(valL) + ", " + rng(periodL) + ", $A" + monthCell + ", " + rng(remL) + ', "' + remark + '")'
          : "SUMIFS(" + rng(valL) + ", " + rng(remL) + ', "' + remark + '")';
      };
      return { latePO: sumifs(poL, "Late"), lateArr: sumifs(arrL, "Late"),
               onPO: sumifs(poL, "On Time"), onArr: sumifs(arrL, "On Time") };
    }

    var r = 6;
    order.forEach(function (m){
      var d = agg[m], F = d.lateL + d.ontL, G = d.lateS + d.ontS;
      var f = rowFormula(r, r);
      sm.getRow(r).getCell(1).value = m;
      sm.getRow(r).getCell(2).value = { formula: f.latePO, result: d.lateL };
      sm.getRow(r).getCell(3).value = { formula: f.lateArr, result: d.lateS };
      sm.getRow(r).getCell(4).value = { formula: f.onPO, result: d.ontL };
      sm.getRow(r).getCell(5).value = { formula: f.onArr, result: d.ontS };
      sm.getRow(r).getCell(6).value = { formula: "B" + r + "+D" + r, result: F };
      sm.getRow(r).getCell(7).value = { formula: "C" + r + "+E" + r, result: G };
      sm.getRow(r).getCell(8).value = { formula: "G" + r + "/F" + r, result: F ? G / F : 0 };
      sm.getRow(r).getCell(9).value = { formula: "E" + r + "/F" + r, result: F ? d.ontS / F : 0 };
      r++;
    });
    var f = rowFormula(r, null);
    var F = grand.lateL + grand.ontL, G = grand.lateS + grand.ontS;
    sm.getRow(r).getCell(1).value = "Grand Total";
    sm.getRow(r).getCell(2).value = { formula: f.latePO, result: grand.lateL };
    sm.getRow(r).getCell(3).value = { formula: f.lateArr, result: grand.lateS };
    sm.getRow(r).getCell(4).value = { formula: f.onPO, result: grand.ontL };
    sm.getRow(r).getCell(5).value = { formula: f.onArr, result: grand.ontS };
    sm.getRow(r).getCell(6).value = { formula: "B" + r + "+D" + r, result: F };
    sm.getRow(r).getCell(7).value = { formula: "C" + r + "+E" + r, result: G };
    sm.getRow(r).getCell(8).value = { formula: "G" + r + "/F" + r, result: F ? G / F : 0 };
    sm.getRow(r).getCell(9).value = { formula: "E" + r + "/F" + r, result: F ? grand.ontS / F : 0 };
    var lastRow = r;
    for (var rr = 6; rr <= lastRow; rr++){
      for (var c2 = 1; c2 <= 9; c2++){
        var cell = sm.getRow(rr).getCell(c2);
        cell.border = clone(BORDER);
        if (c2 <= 7){ cell.numFmt = NUMACC; } else { cell.numFmt = "0.0%"; cell.alignment = { horizontal: "center" }; }
      }
    }
    sm.getRow(lastRow).getCell(1).alignment = { horizontal: "left" };
    try { sm.orderNo = 0; wb.worksheets.forEach(function (s, i){ if (s !== sm) s.orderNo = i + 1; }); } catch (e) {}
    return sm;
  }

  /** Main entry: mutate `wb` in place, return stats. */
  function build(wb, tol){
    tol = tol || 7;
    var p = parse(wb, tol), rows = p.rows;
    var agg = {}, order;
    rows.forEach(function (x){
      if (!x.plan) return;
      var m = MONTHS[x.plan.getUTCMonth()];
      var d = agg[m] || (agg[m] = { lateL:0, lateS:0, ontL:0, ontS:0 });
      if (x.on_time){ d.ontL += x.poqty; d.ontS += x.arrqty; } else { d.lateL += x.poqty; d.lateS += x.arrqty; }
    });
    order = MONTHS.filter(function (m){ return agg[m]; });
    var grand = { lateL:0, lateS:0, ontL:0, ontS:0 };
    order.forEach(function (m){ for (var k in grand) grand[k] += agg[m][k]; });
    var title = String((rows.filter(function (x){ return x.pname; })[0] || {}).pname || "PRINCIPAL").trim();
    var yc = {}, year = new Date().getFullYear();
    rows.forEach(function (x){ if (x.plan){ var y = x.plan.getUTCFullYear(); yc[y] = (yc[y] || 0) + 1; } });
    var ys = Object.keys(yc).sort(function (a, b){ return yc[b] - yc[a]; });
    if (ys.length) year = +ys[0];

    var h = appendHelpers(p.ws, p.hrow, p.last, rows, tol, p.planL, p.arrL);
    buildSummary(wb, order, agg, grand, title, year, p.ws.name, p.poL, p.arrqL, h.periodL, h.remL, p.hrow + 1, p.last);
    try { wb.calcProperties.fullCalcOnLoad = true; } catch (e) {}
    var F = grand.lateL + grand.ontL, G = grand.lateS + grand.ontS;
    return { title: title, year: year, order: order, agg: agg, grand: grand,
      nlines: rows.length, fulfill: F ? G / F : 0, otif: F ? grand.ontS / F : 0, F: F, G: G };
  }

  function outName(fileName){
    var b = String(fileName).replace(/\.xlsx?$/i, "").replace(/[^0-9A-Za-z ]+/g, " ").replace(/\s+/g, " ").trim();
    return "2026_" + b.toUpperCase() + "_SL & OTIF.xlsx";
  }

  return { build: build, outName: outName, MONTHS: MONTHS };
});
