/**
 * GLA Alumni Get-Together 2026 — Apps Script backend
 * IMPORTANT:
 * 1) Keep the web app restricted to your account / trusted organization.
 * 2) Set ADMIN_EMAILS before deployment.
 * 3) This script reads the two configured spreadsheet tabs by sheet ID (gid).
 * 4) It creates/uses a "Check-in Log" tab in the master spreadsheet.
 */

const CONFIG = {
  EVENT_NAME: "GLA Alumni Get-Together 2026",
  MASTER_SPREADSHEET_ID: "11qrL9LWiF1zj_JBaXIJQyXFURZS_FEbADsO7NqFPle0",
  MASTER_SHEET_GID: 1277655877,
  RESPONSES_SPREADSHEET_ID: "1IpauR_Cp6Rvf2IPFLNmcy7zQ4hKlPnkshXuovmWFdXQ",
  RESPONSES_SHEET_GID: 904347459,
  CHECKIN_LOG_SHEET_NAME: "Check-in Log",
  // Replace with the Google account email(s) allowed to view personal records and mark attendance.
  // Example: ["your.name@gmail.com"]
  ADMIN_EMAILS: ["REPLACE_WITH_YOUR_EMAIL"],
  // Dashboard polls every 60 seconds. Google Apps Script quotas still apply.
  DEFAULT_REFRESH_SECONDS: 60
};

function doGet() {
  assertAuthorized_();
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle(CONFIG.EVENT_NAME + " | Command Centre")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function getDashboardData() {
  assertAuthorized_();
  const master = readSheetByGid_(CONFIG.MASTER_SPREADSHEET_ID, CONFIG.MASTER_SHEET_GID);
  const responses = readSheetByGid_(CONFIG.RESPONSES_SPREADSHEET_ID, CONFIG.RESPONSES_SHEET_GID);
  const masterRows = master.rows;
  const responseRows = responses.rows;
  const masterHeaders = master.headers;
  const responseHeaders = responses.headers;

  const masterKeyIndex = findKeyIndex_(masterHeaders);
  const responseKeyIndex = findKeyIndex_(responseHeaders);
  const masterPhotoIndex = findPhotoIndex_(masterHeaders);
  const responsePhotoIndex = findPhotoIndex_(responseHeaders);

  const masterByKey = new Map();
  masterRows.forEach((r, i) => {
    const key = normalize_(masterKeyIndex >= 0 ? r[masterKeyIndex] : "");
    if (key && !masterByKey.has(key)) masterByKey.set(key, { row: r, rowNumber: i + 2 });
  });

  const responseByKey = new Map();
  const duplicateKeys = new Set();
  responseRows.forEach((r, i) => {
    const key = normalize_(responseKeyIndex >= 0 ? r[responseKeyIndex] : "");
    if (!key) return;
    if (responseByKey.has(key)) duplicateKeys.add(key);
    else responseByKey.set(key, { row: r, rowNumber: i + 2 });
  });

  const profiles = [];
  const responseKeySet = new Set();
  const matchedKeys = new Set();
  let mismatched = 0;
  let incomplete = 0;

  responseRows.forEach((r, i) => {
    const keyRaw = responseKeyIndex >= 0 ? r[responseKeyIndex] : "";
    const key = normalize_(keyRaw);
    const matched = !!key && masterByKey.has(key);
    if (key) responseKeySet.add(key);
    if (matched) matchedKeys.add(key);

    const photoValue = responsePhotoIndex >= 0 ? r[responsePhotoIndex] : "";
    const profile = rowToObject_(responseHeaders, r);
    const masterRecord = matched ? rowToObject_(masterHeaders, masterByKey.get(key).row) : null;
    const compare = matched ? compareRecords_(masterRecord, profile, masterHeaders, responseHeaders) : [];
    const isMismatch = compare.some(x => !x.match);
    if (isMismatch) mismatched++;
    if (r.some(v => v === "" || v === null)) incomplete++;

    profiles.push({
      id: key || ("response-row-" + (i + 2)),
      key: String(keyRaw || ""),
      source: matched ? "Matched" : (key ? "New respondent" : "Missing identifier"),
      duplicate: key ? duplicateKeys.has(key) : false,
      mismatch: isMismatch,
      incomplete: r.some(v => v === "" || v === null),
      photoUrl: photoUrl_(photoValue),
      responseRow: i + 2,
      masterRow: matched ? masterByKey.get(key).rowNumber : null,
      response: profile,
      master: masterRecord,
      differences: compare,
      submittedAt: findTimestamp_(responseHeaders, r)
    });
  });

  const pending = masterRows.filter(r => {
    const key = normalize_(masterKeyIndex >= 0 ? r[masterKeyIndex] : "");
    return key && !responseKeySet.has(key);
  }).map(r => rowToObject_(masterHeaders, r));

  const checkins = readCheckins_();
  const checkedInKeys = new Set(checkins.map(c => normalize_(c.key)));
  profiles.forEach(p => p.checkedIn = checkedInKeys.has(normalize_(p.key)));

  const batchCounts = countByField_(profiles.filter(p => p.source === "Matched" || p.source === "New respondent"), responseHeaders, ["batch", "graduation year", "passing year", "year of passing"]);
  const departmentCounts = countByField_(profiles, responseHeaders, ["department", "course", "program", "faculty"]);
  const timeline = buildTimeline_(profiles);

  return {
    eventName: CONFIG.EVENT_NAME,
    generatedAt: new Date().toISOString(),
    refreshSeconds: CONFIG.DEFAULT_REFRESH_SECONDS,
    masterHeaders, responseHeaders,
    counts: {
      masterTotal: masterRows.length,
      totalResponses: responseRows.length,
      uniqueResponseIdentifiers: responseKeySet.size,
      matched: matchedKeys.size,
      pending: pending.length,
      newRespondents: profiles.filter(p => p.source === "New respondent").length,
      missingIdentifier: profiles.filter(p => p.source === "Missing identifier").length,
      duplicateResponseRows: profiles.filter(p => p.duplicate).length,
      mismatched,
      incomplete,
      checkedIn: checkins.length,
      notCheckedIn: Math.max(0, profiles.length - checkedInKeys.size)
    },
    batchCounts, departmentCounts, timeline,
    profiles,
    pending,
    checkins,
    keyFieldMaster: masterKeyIndex >= 0 ? masterHeaders[masterKeyIndex] : null,
    keyFieldResponses: responseKeyIndex >= 0 ? responseHeaders[responseKeyIndex] : null,
    warnings: buildWarnings_(masterKeyIndex, responseKeyIndex, masterRows.length, responseRows.length)
  };
}

function markCheckIn(identifier, notes) {
  assertAuthorized_();
  if (!identifier) throw new Error("Select an alumnus before checking in.");
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = getDashboardData();
    const key = normalize_(identifier);
    const already = data.checkins.find(c => normalize_(c.key) === key);
    if (already) return { ok: false, message: "Already checked in at " + already.time + ".", existing: already };
    const ss = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
    let sheet = ss.getSheetByName(CONFIG.CHECKIN_LOG_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.CHECKIN_LOG_SHEET_NAME);
      sheet.appendRow(["Identifier", "Alumni Name", "Check-in Time", "Checked In By", "Notes"]);
      sheet.setFrozenRows(1);
    }
    const p = data.profiles.find(x => normalize_(x.key) === key);
    if (!p) throw new Error("Alumni not found in current responses. Check in only after verifying their identity and registration.");
    const email = Session.getActiveUser().getEmail() || "Authorized dashboard user";
    const name = getNameFromProfile_(p.response);
    sheet.appendRow([identifier, name, new Date(), email, notes || ""]);
    return { ok: true, message: name + " checked in successfully." };
  } finally {
    lock.releaseLock();
  }
}

function exportCsv() {
  assertAuthorized_();
  const data = getDashboardData();
  const headers = ["Identifier", "Status", "Checked In", "Photo URL", "Response Row", "Master Row"];
  const rows = data.profiles.map(p => [p.key, p.source, p.checkedIn ? "Yes" : "No", p.photoUrl, p.responseRow, p.masterRow || ""]);
  return [headers].concat(rows).map(row => row.map(csvEscape_).join(",")).join("\r\n");
}

function assertAuthorized_() {
  const email = (Session.getActiveUser().getEmail() || "").toLowerCase();
  const allowed = CONFIG.ADMIN_EMAILS.map(x => String(x).toLowerCase());
  if (!email || allowed.includes("replace_with_your_email") || !allowed.includes(email)) {
    throw new Error("Access not configured. Set CONFIG.ADMIN_EMAILS to your Google account email, then deploy the web app with restricted access.");
  }
}

function readSheetByGid_(spreadsheetId, gid) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheets().find(s => s.getSheetId() === Number(gid));
  if (!sheet) throw new Error("Could not find tab with gid " + gid + " in spreadsheet " + spreadsheetId + ". Check the tab ID in the URL.");
  const values = sheet.getDataRange().getDisplayValues();
  if (!values.length) return { headers: [], rows: [] };
  const headers = values[0].map((h, i) => String(h || ("Column " + (i + 1))).trim());
  return { headers, rows: values.slice(1).filter(r => r.some(v => String(v).trim() !== "")) };
}

function findKeyIndex_(headers) {
  const exactHints = ["roll number", "roll no", "student id", "alumni id", "registration id", "email address", "email", "email id", "phone number", "mobile number", "contact number"];
  const normalizedHeaders = headers.map(h => normalize_(h));
  for (const hint of exactHints) {
    const i = normalizedHeaders.findIndex(h => h === normalize_(hint));
    if (i >= 0) return i;
  }
  return normalizedHeaders.findIndex(h => /(roll|alumni.?id|student.?id|registration.?id|email|mobile|phone|contact)/.test(h));
}
function findPhotoIndex_(headers) {
  return headers.findIndex(h => /(photo|picture|image|upload|drive link)/i.test(h));
}
function findTimestamp_(headers, row) {
  const i = headers.findIndex(h => /(timestamp|submitted at|submission time|date submitted)/i.test(h));
  return i >= 0 ? row[i] : "";
}
function normalize_(v) {
  return String(v == null ? "" : v).trim().toLowerCase().replace(/\s+/g, "");
}
function rowToObject_(headers, row) {
  const o = {};
  headers.forEach((h, i) => o[h] = row[i] == null ? "" : row[i]);
  return o;
}
function photoUrl_(value) {
  const s = String(value || "").trim();
  if (!s) return "";
  const m = s.match(/https?:\/\/[^\s,]+/);
  if (!m) return "";
  return m[0].replace(/[)\]]+$/, "");
}
function compareRecords_(master, response, masterHeaders, responseHeaders) {
  const masterMap = new Map(masterHeaders.map(h => [normalize_(h), master[h]]));
  const result = [];
  responseHeaders.forEach(h => {
    const k = normalize_(h);
    if (!masterMap.has(k) || !String(response[h] || "").trim()) return;
    const a = String(masterMap.get(k) || "").trim();
    const b = String(response[h] || "").trim();
    if (a && b) result.push({ field: h, masterValue: a, responseValue: b, match: normalize_(a) === normalize_(b) });
  });
  return result;
}
function countByField_(profiles, headers, hints) {
  const i = headers.findIndex(h => hints.some(hint => normalize_(h).includes(normalize_(hint))));
  if (i < 0) return [];
  const field = headers[i];
  const counts = {};
  profiles.forEach(p => {
    const value = p.response[field] || "Not specified";
    counts[value] = (counts[value] || 0) + 1;
  });
  return Object.keys(counts).map(label => ({ label, count: counts[label] })).sort((a,b) => b.count-a.count).slice(0, 20);
}
function buildTimeline_(profiles) {
  const counts = {};
  profiles.forEach(p => {
    const raw = p.submittedAt;
    if (!raw) return;
    const d = new Date(raw);
    const label = isNaN(d.getTime()) ? String(raw).slice(0, 10) : Utilities.formatDate(d, Session.getScriptTimeZone(), "dd MMM");
    counts[label] = (counts[label] || 0) + 1;
  });
  return Object.keys(counts).map(label => ({ label, count: counts[label] }));
}
function buildWarnings_(masterKeyIndex, responseKeyIndex, masterTotal, responseTotal) {
  const w = [];
  if (masterKeyIndex < 0) w.push("Could not auto-detect a unique identifier in the master tab. Set the key column explicitly in Code.gs.");
  if (responseKeyIndex < 0) w.push("Could not auto-detect a unique identifier in the response tab. Set the key column explicitly in Code.gs.");
  if (!masterTotal) w.push("The master tab appears empty or has only a header row.");
  if (!responseTotal) w.push("The response tab appears empty or has only a header row.");
  return w;
}
function readCheckins_() {
  const ss = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.CHECKIN_LOG_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getDataRange().getValues().slice(1).filter(r => r[0]).map(r => ({
    key: String(r[0]), name: String(r[1] || ""), time: r[2] instanceof Date ? r[2].toISOString() : String(r[2] || ""),
    by: String(r[3] || ""), notes: String(r[4] || "")
  }));
}
function getNameFromProfile_(obj) {
  const k = Object.keys(obj || {}).find(h => /^(name|full name|alumni name|student name)$/i.test(h.trim()));
  return k ? obj[k] : (Object.values(obj || {}).slice(0, 2).join(" "));
}
function csvEscape_(v) {
  const s = String(v == null ? "" : v);
  return '"' + s.replace(/"/g, '""') + '"';
}
