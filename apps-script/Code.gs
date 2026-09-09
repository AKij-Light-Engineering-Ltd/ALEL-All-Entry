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
var PROP_USERS     = "ALEL_USERS";
var PROP_REQS      = "ALEL_REQUESTS";

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
    case "resetpass":     return doResetPass(p);
    case "removeRequest": return doRemoveRequest(p);
    default:
      return { ok: false, error: "Unknown action: " + action };
  }
}

// ---------------------------------------------------------------------------
// ACTIONS
// ---------------------------------------------------------------------------

function doList() {
  return { ok: true, users: readUsers(), requests: readRequests(), approvals: [] };
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
  var rec   = { id: email, name: String(p.name || email).trim(), role: String(p.role || "user"), section: String(p.section || "") || "", hash: String(p.hash || "") };
  var found = false;
  users = users.map(function (u) {
    if (u.id === email) { found = true; return rec; }
    return u;
  });
  if (!found) users.push(rec);
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
  var data  = sheet.getDataRange().getValues();
  var users = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    users.push({ id: String(row[0]), name: String(row[1]), role: String(row[2]), section: String(row[3] || ""), hash: String(row[4]) });
  }
  return users;
}

function writeUsers(users) {
  ensureSheets_();
  var sheet = ss_().getSheetByName(USERS_SHEET);
  sheet.clearContents();
  var rows = users.map(function (u) {
    return [u.id, u.name, u.role, u.section || "", u.hash];
  });
  if (rows.length) sheet.getRange(1, 1, rows.length, 5).setValues(rows);
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

function userExists(email) {
  return readUsers().some(function (u) { return u.id === email; });
}

// ---------------------------------------------------------------------------
// OPTIONAL: browser add-ons are not needed; the sheet is the single source.
// ---------------------------------------------------------------------------
