/**
 * ============================================================================
 *  ALEL 5S Audit — EDIT / DELETE backend   (Google Apps Script Web App)
 *  File: apps-script/5s-edit.gs
 * ----------------------------------------------------------------------------
 *  Ki kore:
 *    • update  — ekta 5S audit row er value change kore (Total/Max/Score% nijei
 *                abar hiseb kore), ar "5S Edit Log" sheet e audit-trail rakhe
 *    • delete  — ekta 5S audit row delete kore (log e rekhe dey)
 *    • verify  — JSONP GET; dashboard save ta sotti hoyeche kina check kore
 *
 *  SETUP (ekbar, ~3 minute):
 *    1) script.google.com → New project → ei code ta paste korun (Code.gs)
 *    2) Save (💾) → Run ekbar `pingTest` → Google permission "Allow"
 *    3) Deploy → New deployment → type: Web app
 *         Execute as        : Me (your email)
 *         Who has access    : Anyone
 *       → Deploy → copy the /exec URL
 *    4) Oi /exec URL ta dashboard e paste korun:
 *         https://akij-light-engineering-ltd.github.io/ALEL-All-Entry/5s-dashboard.html
 *       → ✏️ Edit → setup box e URL paste → "Save URL"
 *
 *  NOTE: ei script ta tomar 5S FORM er script theke ALADA rakha bhalo —
 *        tahole form er /exec URL kono rokome change hobe na.
 * ============================================================================
 */

// ------------------------- CONFIG -------------------------------------------
var SHEET_ID   = '1iVlAdl9bcmZE1Uk_24z7NUl4FtivFTa2h7wOI1T8T0w';  // "ALEL All Entry Central Sheet"
var TAB_NAME   = '5S Form';        // je tab e 5S audit data ache
var LOG_TAB    = '5S Edit Log';    // audit trail (script nijei banabe)
var KEY_COLS   = ['Email', 'Audit Date', 'Floor / Area'];  // row identify korar key
var SCORE_RE   = /^\d\.\d\s/;      // "1.1 ..." = score column
var REMARK_RE  = /Remarks/i;
var MAX_PER_Q  = 3;                // protiti question er max score

// Edit permission (khali array = sob allowed, tabe log e record thakbe).
// Specific manush der jonno niche email gulo boshie din:
var EDITORS = [
  // 'md.marufhossain@akijlightengineering.com',
  // 'jhumour@akijlightengineering.com'
];

// ------------------------- WEB ENTRY ----------------------------------------
function doGet(e) {
  var p = (e && e.parameter) || {};
  var action = String(p.action || 'ping');
  var res;
  try {
    if (action === 'verify')      res = doVerify_(p);
    else if (action === 'ping')   res = doPing_();
    else                          res = { ok: false, error: 'Unknown GET action: ' + action };
  } catch (err) {
    res = { ok: false, error: String(err) };
  }
  return respond_(res, e);
}

function doPost(e) {
  var res;
  try {
    var body = (e && e.postData && e.postData.contents) || '{}';
    var p = JSON.parse(body);
    var action = String(p.action || '');
    if (action === 'update')      res = doUpdate_(p);
    else if (action === 'delete') res = doDelete_(p);
    else                          res = { ok: false, error: 'Unknown POST action: ' + action };
  } catch (err) {
    res = { ok: false, error: String(err) };
  }
  return respond_(res, e);
}

/** JSONP support (dashboard theke response PORA jay) */
function respond_(obj, e) {
  var cb = e && e.parameter && e.parameter.callback;
  var text = JSON.stringify(obj);
  if (cb) {
    return ContentService.createTextOutput(cb + '(' + text + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JSON);
}

// ------------------------- HELPERS ------------------------------------------
function sheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(TAB_NAME);
  if (!sh) throw new Error('Tab pelam na: ' + TAB_NAME);
  return sh;
}
function logSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(LOG_TAB);
  if (!sh) {
    sh = ss.insertSheet(LOG_TAB);
    sh.appendRow(['When', 'Action', 'Row', 'By (name)', 'By (email)', 'Key (email | date | floor)', 'Changes']);
    sh.getRange('A1:G1').setFontWeight('bold');
  }
  return sh;
}
function norm_(v) { return String(v == null ? '' : v).trim().toLowerCase(); }
function day_(v) {
  if (v instanceof Date) {
    var t = new Date(v.getTime() - v.getTimezoneOffset() * 60000);
    return t.toISOString().slice(0, 10);
  }
  var m = String(v).match(/(\d{4})-(\d{2})-(\d{2})/);
  return m ? (m[1] + '-' + m[2] + '-' + m[3]) : String(v).trim();
}
function allowed_(p) {
  if (!EDITORS || !EDITORS.length) return true;
  var em = norm_((p.by && p.by.email) || '');
  if (!em) return false;
  for (var i = 0; i < EDITORS.length; i++) if (norm_(EDITORS[i]) === em) return true;
  return false;
}

/** Row khoja: row number age (key mile kina check kore), na mille key diye scan. */
function findRow_(values, head, key, rowNo) {
  var ci = {}; for (var i = 0; i < head.length; i++) ci[head[i]] = i;
  function match(r) {
    for (var k = 0; k < KEY_COLS.length; k++) {
      var col = KEY_COLS[k]; if (ci[col] === undefined) continue;
      var want = (col === 'Audit Date') ? day_(key && key.date) : norm_((key || {})[col === 'Email' ? 'email' : 'floor']);
      var have = (col === 'Audit Date') ? day_(r[ci[col]]) : norm_(r[ci[col]]);
      if (want !== '' && want !== have) return false;
    }
    return true;
  }
  var n = Number(rowNo || 0);
  if (n >= 2 && n <= values.length) {
    if (match(values[n - 1])) return n - 1;            // 0-based index
  }
  if (key) {
    for (var r = 1; r < values.length; r++) if (match(values[r])) return r;
  }
  return -1;
}

function rowToObj_(head, row) {
  var o = {};
  for (var i = 0; i < head.length; i++) o[head[i]] = (row[i] === undefined || row[i] === null) ? '' : String(row[i]);
  return o;
}

// ------------------------- ACTIONS ------------------------------------------
function doPing_() {
  var sh = sheet_();
  return { ok: true, info: 'ALEL 5S edit backend', tab: TAB_NAME, rows: sh.getLastRow(), editors: EDITORS };
}

function doVerify_(p) {
  var sh = sheet_();
  var values = sh.getDataRange().getValues();
  var head = values[0].map(function (h) { return String(h).trim(); });
  var idx = findRow_(values, head, p.key || null, p.row);
  if (idx < 1) return { ok: false, error: 'Row pelam na' };
  return { ok: true, row: idx + 1, values: rowToObj_(head, values[idx]) };
}

function doUpdate_(p) {
  if (!allowed_(p)) return { ok: false, error: 'Ei email er edit permission nei (EDITORS list dekho)' };
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    var sh = sheet_();
    var values = sh.getDataRange().getValues();
    var head = values[0].map(function (h) { return String(h).trim(); });
    var idx = findRow_(values, head, p.key || null, p.row);
    if (idx < 1) return { ok: false, error: 'Row khuje pelam na — Refresh kore abar chesta koro' };

    var row = values[idx].slice();
    var v = p.values || {};
    var changed = [];
    Object.keys(v).forEach(function (k) {
      var key = String(k).trim();
      if (key === 'sheet') return;
      var c = head.indexOf(key);
      if (c < 0) return;                                  // ei column nei → skip
      var oldV = row[c], newV = v[k];
      if (String(oldV) !== String(newV)) {
        changed.push(key + ': "' + oldV + '" -> "' + newV + '"');
        row[c] = newV;
      }
    });

    // Total / Max / Score % abar hiseb
    var total = 0, cnt = 0;
    for (var c2 = 0; c2 < head.length; c2++) {
      if (SCORE_RE.test(head[c2]) && !REMARK_RE.test(head[c2])) {
        cnt++;
        total += Number(row[c2]) || 0;
      }
    }
    var max = cnt * MAX_PER_Q;
    var tC = head.indexOf('Total Score'), mC = head.indexOf('Max Score'), pC = head.indexOf('Score %');
    if (tC >= 0) row[tC] = total;
    if (mC >= 0) row[mC] = max;
    if (pC >= 0) row[pC] = max ? Math.round(total / max * 100) : 0;

    sh.getRange(idx + 1, 1, 1, row.length).setValues([row]);

    logSheet_().appendRow([
      new Date(), 'UPDATE', idx + 1,
      (p.by && p.by.name) || '', (p.by && p.by.email) || '',
      [day_(row[head.indexOf('Email')]), day_(row[head.indexOf('Audit Date')]), row[head.indexOf('Floor / Area')]].join(' | '),
      changed.join(' ; ') || '(no value change)'
    ]);

    return { ok: true, row: idx + 1, total: total, max: max, pct: max ? Math.round(total / max * 100) : 0, changed: changed.length };
  } catch (err) {
    return { ok: false, error: String(err) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function doDelete_(p) {
  if (!allowed_(p)) return { ok: false, error: 'Ei email er edit permission nei (EDITORS list dekho)' };
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    var sh = sheet_();
    var values = sh.getDataRange().getValues();
    var head = values[0].map(function (h) { return String(h).trim(); });
    var idx = findRow_(values, head, p.key || null, p.row);
    if (idx < 1) return { ok: false, error: 'Row khuje pelam na' };

    var snap = rowToObj_(head, values[idx]);
    logSheet_().appendRow([
      new Date(), 'DELETE', idx + 1,
      (p.by && p.by.name) || '', (p.by && p.by.email) || '',
      [snap['Email'], day_(snap['Audit Date']), snap['Floor / Area']].join(' | '),
      'Total=' + snap['Total Score'] + ' Score%=' + snap['Score %'] + ' Auditor=' + snap['Auditor Name']
    ]);
    sh.deleteRow(idx + 1);
    return { ok: true, deletedRow: idx + 1 };
  } catch (err) {
    return { ok: false, error: String(err) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ------------------------- TEST ---------------------------------------------
/** Editor e ekbar Run korun — permission allow hoye jabe. */
function pingTest() {
  Logger.log(JSON.stringify(doPing_()));
}
