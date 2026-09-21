/**
 * ALEL Production Plan — read-only data proxy
 * ===========================================
 * Serves the planning workbook to the public dashboard WITHOUT changing
 * the workbook's sharing settings. The script runs as YOU, and you only
 * need *view* access to the workbook — no edit rights required.
 *
 * ── Deploy (2 minutes) ────────────────────────────────────────────────────
 *  1. Open https://script.google.com  →  New project
 *  2. Delete the sample code, paste this whole file
 *  3. Click  Deploy  →  New deployment
 *       • Type            : Web app
 *       • Description     : Production Plan proxy
 *       • Execute as      : Me            ← important
 *       • Who has access  : Anyone        ← important
 *     Click Deploy, then Authorize (choose your account → Advanced →
 *     "Go to <project> (unsafe)" → Allow).
 *  4. Copy the /exec URL it gives you.
 *  5. Send that URL back and it gets wired into the dashboard
 *     (production-plan.js → PROXY_URL).
 *
 * Read-only: this script only calls getDataRange().getValues(). It never
 * writes to the workbook.
 */

var SHEET_ID = "1LYws_M4TANxF0O0N3lJ7RixKxsEQSOsTJfa4h5V8pro";
var GROUP_TAB = "SKU Plan 3M";
var MONA = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || "plan";
    if (action === "ping") return json({ ok: true, pong: true });
    return json(buildPlan());
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function buildPlan() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var tab = pickPlanTab(ss);
  if (!tab) return { ok: false, error: "No monthly plan tab found" };

  var sh = ss.getSheetByName(tab);
  var values = sh.getDataRange().getValues();
  if (values.length < 3) return { ok: false, error: "Tab too small: " + tab };

  var hdr = values[0].map(function (h) { return String(h == null ? "" : h).trim(); });
  var r1 = values[1].map(function (h) { return String(h == null ? "" : h).trim(); });

  function find(re, not) {
    for (var i = 0; i < hdr.length; i++) {
      if (re.test(hdr[i]) && !(not && not.test(hdr[i]))) return i;
    }
    return -1;
  }
  var iCode = find(/^code$/i), iName = find(/item\s*name/i);
  var iFc = find(/forecast/i, /value/i), iOpen = find(/opening/i, /value/i);
  var iReq = find(/^req\.?\s*qty$/i), iProd = find(/total\s*production/i), iDel = find(/delivered\s*quantity/i);

  var dateCols = [], dateIdx = [];
  for (var c = 0; c < r1.length; c++) {
    var n = Number(r1[c]);
    if (n > 40000 && n < 60000) { dateCols.push(xlDate(n)); dateIdx.push(c); }
  }

  var rows = [];
  for (var r = 2; r < values.length; r++) {
    var row = values[r];
    var name = String(row[iName] || "").trim();
    if (!name || /^total/i.test(name)) continue;
    var daily = {};
    for (var k = 0; k < dateIdx.length; k++) {
      var q = num(row[dateIdx[k]]);
      if (q) daily[dateCols[k]] = q;
    }
    rows.push({
      code: String(row[iCode] || "").trim(), name: name,
      fc: num(row[iFc]), open: num(row[iOpen]), req: num(row[iReq]),
      prod: num(row[iProd]), del: num(row[iDel]), daily: daily,
    });
  }

  return { ok: true, tab: tab, generatedAt: new Date().toISOString(), dateCols: dateCols, rows: rows, groups: groupMap(ss) };
}

function pickPlanTab(ss) {
  var now = new Date();
  for (var b = 0; b <= 6; b++) {
    var d = new Date(now.getFullYear(), now.getMonth() - b, 1);
    var m = MONA[d.getMonth()], y = String(d.getFullYear()).slice(2);
    var cands = [
      m + y + "-Forecast vs Production", m + y + "-Forecast vs production",
      m + "-Forecast vs Production", m + "-Forecast vs production",
      m + y + " Forecast", m + " Forecast",
    ];
    for (var i = 0; i < cands.length; i++) {
      if (ss.getSheetByName(cands[i])) return cands[i];
    }
  }
  return null;
}

function groupMap(ss) {
  var out = {};
  try {
    var sh = ss.getSheetByName(GROUP_TAB);
    if (!sh) return out;
    var v = sh.getDataRange().getValues();
    for (var r = 4; r < v.length; r++) {
      var c = String(v[r][1] == null ? "" : v[r][1]).trim();
      var g = String(v[r][7] == null ? "" : v[r][7]).trim();
      if (c && g) out[c] = g;
    }
  } catch (e) { /* optional */ }
  return out;
}

function num(v) {
  if (v == null || v === "") return 0;
  var n = parseFloat(String(v).replace(/[%,]/g, ""));
  return isNaN(n) ? 0 : n;
}
function xlDate(n) {
  var d = new Date(Date.UTC(1899, 11, 30) + n * 86400000);
  return d.toISOString().slice(0, 10);
}
