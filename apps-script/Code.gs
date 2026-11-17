/**
 * ALEL Opex Hub — Shared Login Backend
 * =====================================
 * A Google Apps Script web app that stores a SINGLE, shared list of
 * users and access requests for every device that opens the hub.
 *
 * What it does:
 *   - list          -> returns { users, requests, approvals }
 *   - request       -> submit an access request or a password-reset request
 *   - upsert        -> create / update a user (admin action)
 *   - resetpass     -> set a user's password hash (admin action)
 *   - removeRequest -> remove a request from the queue (admin action)
 *   - deluser       -> permanently remove a user by email (admin action)
 *
 * Deploy steps (one-time, done by the administrator):
 *   1. Create a new Google Sheet (leave it open / shared, this is the data store).
 *   2. In that sheet: Extensions -> Apps Script, replace Code.gs with this file.
 *   3. Run `initSheets()` once (in the editor click Run). Authorize when asked.
 *   4. Deploy -> New deployment -> type: Web app.
 *      - Execute as: Me
 *      - Who has access: Anyone
 *   5. Copy the /exec URL and put it into:
 *        index.html            -> BACKEND_URL
 *        operation-bulletin.html -> BACKEND_URL
 *   6. Commit & push. Done — login now shared across all devices.
 */

// ---------------------------------------------------------------------------
// CONFIG
// ---------------------------------------------------------------------------
// File keys in the bound spreadsheet. Change to match your own column order
// if you customise initSheets().
var USERS_SHEET    = "Users";
var REQUESTS_SHEET = "Requests";
var REMARKS_SHEET  = "Remarks";
var APPROV_TAB     = "SMV Approvals";
var OPS_TAB        = "Manual Ops";
var PROP_USERS     = "ALEL_USERS";
var PROP_REQS      = "ALEL_REQUESTS";

// Approvals + manually typed operations are written into the SAME spreadsheet
// that feeds the Operation Bulletin generator, so the generator can pick them
// up without any extra fetch. Falls back to the bound spreadsheet on error.
var PG_SHEET_ID    = "1GpALH7TsJ0eW-Ko8ejzV8oQF073dhe7Eb9930o0qvSo";

function pgSheet_(tab) {
  var ss = null;
  try { ss = SpreadsheetApp.openById(PG_SHEET_ID); } catch (e) { ss = null; }
  if (!ss) ss = ss_();
  var sh = ss.getSheetByName(tab);
  if (!sh) sh = ss.insertSheet(tab);
  if (sh.getLastRow() === 0) {
    if (tab === APPROV_TAB) sh.getRange(1, 1, 1, 6).setValues([["Code", "P", "C", "A", "By", "At"]]);
    if (tab === OPS_TAB)    sh.getRange(1, 1, 1, 7).setValues([["Code", "Operation", "SMV", "Manpower", "Line", "By", "At"]]);
  }
  return sh;
}


// ---------------------------------------------------------------------------
// WEB APP ENTRY POINTS
// ---------------------------------------------------------------------------

/** GET (used for a quick health check). */
function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, info: "ALEL Opex Hub shared backend" }))
    .setMimeType(ContentService.MimeType.JSON);
}

/** POST — every client action comes here. */
function doPost(e) {
  try {
    var body = (e && e.postData && e.postData.contents) || "";
    var payload = JSON.parse(body || "{}");
    return respond(handle(payload));
  } catch (err) {
    return respond({ ok: false, error: String(err) });
  }
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------------------
// ACTION ROUTER
// ---------------------------------------------------------------------------
function handle(p) {
  var action = p.action || "";
  switch (action) {
    case "list":          return doList();
    case "request":       return doRequest(p);
    case "upsert":        return doUpsert(p);
    case "setaccess":     return doSetAccess(p);
    case "resetpass":     return doResetPass(p);
    case "removeRequest": return doRemoveRequest(p);
    case "deluser":       return doDelUser(p);
    case "remark":        return doRemark(p);
    case "approve":       return doApprove(p);
    case "addprocess":    return doAddProcess(p);
    case "delprocess":    return doDelProcess(p);
    default:
      return { ok: false, error: "Unknown action: " + action };
  }
}

// ---------------------------------------------------------------------------
// ACTIONS
// ---------------------------------------------------------------------------

function doList() {
  return { ok: true, users: readUsers(), requests: readRequests(),
           approvals: readApprovals(), remarks: readRemarks(), processes: readProcesses() };
}

/** 3-stage approval state for one SKU: P = prepared, C = checked, A = approved. */
function doApprove(p) {
  var code = String(p.code || "").trim();
  if (!code) return { ok: false, error: "No code" };
  var rows = readApprovals().filter(function (r) { return r.code !== code; });
  rows.push({
    code: code,
    p: Number(p.p || 0) ? 1 : 0,
    c: Number(p.c || 0) ? 1 : 0,
    a: Number(p.a || 0) ? 1 : 0,
    by: String(p.by || ""),
    at: String(p.at || new Date().toLocaleString())
  });
  writeApprovals(rows);
  return { ok: true };
}

/** One operation line typed from the page for a SKU that has no time study. */
function doAddProcess(p) {
  var code = String(p.code || "").trim();
  var op   = String(p.op || "").trim();
  if (!code || !op) return { ok: false, error: "Code and operation are required" };
  var sh = pgSheet_(OPS_TAB);
  sh.appendRow([
    code, op,
    Number(p.smv || 0) || "",
    Number(p.mp || 0) || "",
    String(p.line || ""),
    String(p.by || ""),
    String(p.at || new Date().toLocaleString())
  ]);
  var n = readProcesses().filter(function (r) { return r.code === code; }).length;
  return { ok: true, count: n };
}

function doDelProcess(p) {
  var code = String(p.code || "").trim();
  var op   = String(p.op || "").trim();
  if (!code) return { ok: false, error: "No code" };
  var rows = readProcesses().filter(function (r) {
    return !(r.code === code && (!op || r.op === op));
  });
  writeProcesses(rows);
  return { ok: true };
}

function readApprovals() {
  var sh = pgSheet_(APPROV_TAB);
  var data = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    out.push({ code: String(row[0]).replace(/\.0$/, ""), p: Number(row[1] || 0) ? 1 : 0,
               c: Number(row[2] || 0) ? 1 : 0, a: Number(row[3] || 0) ? 1 : 0,
               by: String(row[4] || ""), at: String(row[5] || "") });
  }
  return out;
}

function writeApprovals(rows) {
  var sh = pgSheet_(APPROV_TAB);
  sh.clearContents();
  var out = [["Code", "P", "C", "A", "By", "At"]];
  rows.forEach(function (r) { out.push([r.code, r.p, r.c, r.a, r.by || "", r.at || ""]); });
  sh.getRange(1, 1, out.length, 6).setValues(out);
}

function readProcesses() {
  var sh = pgSheet_(OPS_TAB);
  var data = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0] || !row[1]) continue;
    out.push({ code: String(row[0]).replace(/\.0$/, ""), op: String(row[1]),
               smv: Number(row[2] || 0), mp: Number(row[3] || 0),
               line: String(row[4] || ""), by: String(row[5] || ""), at: String(row[6] || "") });
  }
  return out;
}

function writeProcesses(rows) {
  var sh = pgSheet_(OPS_TAB);
  sh.clearContents();
  var out = [["Code", "Operation", "SMV", "Manpower", "Line", "By", "At"]];
  rows.forEach(function (r) {
    out.push([r.code, r.op, r.smv || "", r.mp || "", r.line || "", r.by || "", r.at || ""]);
  });
  sh.getRange(1, 1, out.length, 7).setValues(out);
}


/** Save (or update) the remarks for one bulletin, keyed by its item code. */
function doRemark(p) {
  var code = String(p.code || "").trim();
  if (!code) return { ok: false, error: "No code" };
  var text = String(p.text || "");
  var by   = String(p.by || "");
  var role = String(p.role || "");
  var at   = String(p.at || new Date().toLocaleString());
  var rows = readRemarks().filter(function (r) { return r.code !== code; });
  rows.push({ code: code, text: text, by: by, role: role, at: at });
  writeRemarks(rows);
  return { ok: true };
}

function doRequest(p) {
  var type  = p.type;                 // "access" | "reset"
  var name  = String(p.name  || "").trim();
  var email = String(p.email || "").trim().toLowerCase();
  var sec   = String(p.section || "").trim();
  var when  = String(p.when || new Date().toLocaleString());

  if (type !== "access" && type !== "reset") return { ok: false, error: "Unknown request type" };
  if (!email) return { ok: false, error: "No email" };

  var reqs = readRequests();

  if (type === "access") {
    if (userExists(email)) return { ok: false, error: "Email already has an account" };
    var dupe = reqs.filter(function (r) { return r.type === "access" && r.email === email && !r.done; });
    if (dupe.length) return { ok: false, error: "Request already pending" };
    reqs.push({ type: "access", name: name, email: email, section: sec, when: when, done: false });
  } else {
    if (!userExists(email)) return { ok: false, error: "No account with this email" };
    reqs.push({ type: "reset", email: email, when: when, done: false });
  }

  writeRequests(reqs);
  return { ok: true };
}

/** Admin: create or update a user. Pass `hash` = SHA-256 of the password. */
function doUpsert(p) {
  var email = String(p.email || "").trim().toLowerCase();
  if (!email) return { ok: false, error: "No email" };
  var users = readUsers();
  var existing = null;
  for (var i = 0; i < users.length; i++) if (users[i].id === email) existing = users[i];
  var perms = (p.perms && p.perms.length) ? p.perms : (existing && existing.perms && existing.perms.length ? existing.perms : []);
  var rec = {
    id: email,
    name: String(p.name || (existing && existing.name) || email).trim(),
    role: String(p.role || (existing && existing.role) || "user"),
    section: String(p.section || (existing && existing.section) || ""),
    hash: String(p.hash || (existing && existing.hash) || ""),
    perms: perms
  };
  var found = false;
  users = users.map(function (u) {
    if (u.id === email) { found = true; return rec; }
    return u;
  });
  if (!found) users.push(rec);
  writeUsers(users);
  return { ok: true, user: rec };
}

/** Admin: update a user's role and/or module permissions (does not touch password). */
function doSetAccess(p) {
  var email = String(p.email || "").trim().toLowerCase();
  if (!email) return { ok: false, error: "No email" };
  var users = readUsers();
  var found = false;
  users = users.map(function (u) {
    if (u.id === email) {
      found = true;
      if (p.role !== undefined && p.role !== null) u.role = String(p.role);
      if (p.section !== undefined && p.section !== null) u.section = String(p.section);
      if (p.perms !== undefined && p.perms !== null) u.perms = Array.isArray(p.perms) ? p.perms : [];
      return u;
    }
    return u;
  });
  if (!found) return { ok: false, error: "No such user" };
  writeUsers(users);
  return { ok: true };
}

/** Admin: set a user's password hash. */
function doResetPass(p) {
  var email = String(p.email || "").trim().toLowerCase();
  var hash  = String(p.hash || "");
  if (!email || !hash) return { ok: false, error: "Missing email or hash" };
  var users = readUsers();
  var found = false;
  users = users.map(function (u) {
    if (u.id === email) { found = true; u.hash = hash; }
    return u;
  });
  if (!found) return { ok: false, error: "No such user" };
  writeUsers(users);
  return { ok: true };
}

/** Admin: remove a pending request by email. */
function doRemoveRequest(p) {
  var email = String(p.email || "").trim().toLowerCase();
  if (!email) return { ok: false, error: "No email" };
  var reqs = readRequests().filter(function (r) { return r.email !== email; });
  writeRequests(reqs);
  return { ok: true };
}

/** Admin: permanently remove a user by email. */
function doDelUser(p) {
  var email = String(p.email || "").trim().toLowerCase();
  if (!email) return { ok: false, error: "No email" };
  var users = readUsers().filter(function (u) { return u.id !== email; });
  writeUsers(users);
  return { ok: true, removed: email, remaining: users.length };
}

// ---------------------------------------------------------------------------
// STORAGE (Google Sheet backed, with fallback to Script Properties)
// ---------------------------------------------------------------------------

function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function ensureSheets_() {
  var ss = ss_();
  if (!ss.getSheetByName(USERS_SHEET)) ss.insertSheet(USERS_SHEET);
  if (!ss.getSheetByName(REQUESTS_SHEET)) ss.insertSheet(REQUESTS_SHEET);
  if (!ss.getSheetByName(REMARKS_SHEET)) ss.insertSheet(REMARKS_SHEET);
}

/**
 * Must be run once in the editor to create the Users / Requests sheets.
 * Run after authorizing the script (Run button) — then deploy.
 * Idempotent: adds any missing seed user; never overwrites existing users.
 * All seed passwords are: ALEL@2026  (change them from the hub Admin panel).
 */
function initSheets() {
  ensureSheets_();
  var users = readUsers();
  var SEED = [
    { id: "md.marufhossain@akijlightengineering.com", name: "Md. Maruf Hossain",     role: "admin",    section: "",   hash: "18e005af46f178e2ff6083db381ef31888176f834c02c0470bdb8fbdbf2a4ece" },
    { id: "anoy@akijlightengineering.com",            name: "Anoy Kumar Das",        role: "prepared", section: "",   hash: "18e005af46f178e2ff6083db381ef31888176f834c02c0470bdb8fbdbf2a4ece" },
    { id: "jhumour@akijlightengineering.com",          name: "Jhumour Rani",          role: "checker",  section: "GSS", hash: "18e005af46f178e2ff6083db381ef31888176f834c02c0470bdb8fbdbf2a4ece" },
    { id: "kajal04@akijlightengineering.com",          name: "Kajal Kanti",           role: "checker",  section: "LED", hash: "18e005af46f178e2ff6083db381ef31888176f834c02c0470bdb8fbdbf2a4ece" },
    { id: "almamun@akijlightengineering.com",          name: "Abdullah Al-Mamun",     role: "checker",  section: "HAP", hash: "18e005af46f178e2ff6083db381ef31888176f834c02c0470bdb8fbdbf2a4ece" },
    { id: "head.plant@akijlightengineering.com",       name: "Md. Moshfequr Rahman",  role: "approver", section: "",   hash: "18e005af46f178e2ff6083db381ef31888176f834c02c0470bdb8fbdbf2a4ece" }
  ];
  var missing = SEED.filter(function (s) {
    return !users.some(function (u) { return u.id === s.id; });
  });
  users = users.concat(missing);
  writeUsers(users);
  return "Sheets initialised (" + users.length + " user(s), +" + missing.length + " added)";
}

function readUsers() {
  ensureSheets_();
  var sheet = ss_().getSheetByName(USERS_SHEET);
  if (sheet.getLastRow() === 0) { initSheets(); sheet = ss_().getSheetByName(USERS_SHEET); }
  var data  = sheet.getDataRange().getValues();
  var users = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    users.push({ id: String(row[0]), name: String(row[1]), role: String(row[2]), section: String(row[3] || ""), hash: String(row[4]), perms: parsePerms(row[5]) });
  }
  return users;
}

function parsePerms(v) {
  if (!v) return [];
  try { var a = JSON.parse(String(v)); return Array.isArray(a) ? a : []; } catch (e) { return []; }
}

function writeUsers(users) {
  ensureSheets_();
  var sheet = ss_().getSheetByName(USERS_SHEET);
  sheet.clearContents();
  var rows = users.map(function (u) {
    return [u.id, u.name, u.role, u.section || "", u.hash, (u.perms && u.perms.length) ? JSON.stringify(u.perms) : ""];
  });
  if (rows.length) sheet.getRange(1, 1, rows.length, 6).setValues(rows);
}

function readRequests() {
  ensureSheets_();
  var sheet = ss_().getSheetByName(REQUESTS_SHEET);
  var data  = sheet.getDataRange().getValues();
  var reqs  = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (!row[1]) continue; // email is column 2
    reqs.push({ type: String(row[0]), email: String(row[1]), name: String(row[2] || ""), section: String(row[3] || ""), when: String(row[4] || ""), done: row[5] === "DONE" });
  }
  return reqs;
}

function writeRequests(reqs) {
  ensureSheets_();
  var sheet = ss_().getSheetByName(REQUESTS_SHEET);
  sheet.clearContents();
  var rows = reqs.map(function (r) {
    return [r.type, r.email, r.name || "", r.section || "", r.when || "", r.done ? "DONE" : ""];
  });
  if (rows.length) sheet.getRange(1, 1, rows.length, 6).setValues(rows);
}

function readRemarks() {
  ensureSheets_();
  var sheet = ss_().getSheetByName(REMARKS_SHEET);
  var data  = sheet.getDataRange().getValues();
  var rows  = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue; // code is column 1
    rows.push({ code: String(row[0]), text: String(row[1] || ""), by: String(row[2] || ""), role: String(row[3] || ""), at: String(row[4] || "") });
  }
  return rows;
}

function writeRemarks(rows) {
  ensureSheets_();
  var sheet = ss_().getSheetByName(REMARKS_SHEET);
  sheet.clearContents();
  var out = rows.map(function (r) {
    return [r.code, r.text || "", r.by || "", r.role || "", r.at || ""];
  });
  if (out.length) sheet.getRange(1, 1, out.length, 5).setValues(out);
}

function userExists(email) {
  return readUsers().some(function (u) { return u.id === email; });
}

// ---------------------------------------------------------------------------
// OPTIONAL: browser add-ons are not needed; the sheet is the single source.
// ---------------------------------------------------------------------------
