// PracticeTracker — Google Apps Script backend (stores data in this Google Sheet)
// Paste this whole file into Extensions > Apps Script, then Deploy > Manage deployments > pencil > New version.

const API_TOKEN = "PUT-YOUR-EXISTING-SECRET-WORD-HERE";   // must match SYNC_TOKEN in the HTML file
const SHEET_NAME = "Sessions";
const ARCHIVE_NAME = "Archive";
const HEADERS = ["key", "subject", "group", "date", "topic", "submitted",
                 "done", "pending", "absent", "teacher", "updatedAt", "json"];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange("A:D").setNumberFormat("@");   // keep dates as plain text
  }
  return sh;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// When each subject was last reset (so every device can clear its old local data)
function getEpochs_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty("epochs") || "{}"); }
  catch (e) { return {}; }
}

// Read all sessions
function doGet(e) {
  try {
    if (!e.parameter || e.parameter.token !== API_TOKEN) return out_({ ok: false, error: "bad token" });
    const sh = sheet_();
    const last = sh.getLastRow();
    const sessions = [];
    if (last > 1) {
      const rows = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
      rows.forEach(r => {
        try { sessions.push(JSON.parse(r[HEADERS.length - 1])); } catch (err) {}
      });
    }
    return out_({ ok: true, sessions: sessions, epochs: getEpochs_() });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

// Save one session, or reset a subject
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const body = JSON.parse(e.postData.contents);
    if (body.token !== API_TOKEN) return out_({ ok: false, error: "bad token" });

    if (body.action === "reset") return reset_(body.subject);

    const s = body.session;
    if (!s || !s.key) return out_({ ok: false, error: "no session" });

    // Ignore stale saves from a device that still holds data from before a reset
    const epoch = Number(getEpochs_()[s.subject]) || 0;
    if (epoch && (Number(s.updatedAt) || 0) < epoch) return out_({ ok: true, ignored: true });

    const sh = sheet_();
    const recs = Object.values(s.records || {});
    const count = st => recs.filter(r => r.status === st).length;
    const row = [s.key, s.subject, s.group, s.date, s.topic || "", s.submitted ? "Yes" : "No",
                 count("Done"), count("Pending"), count("Absent"),
                 s.teacher || "", s.updatedAt || Date.now(), JSON.stringify(s)];

    const last = sh.getLastRow();
    let target = 0;
    if (last > 1) {
      const keys = sh.getRange(2, 1, last - 1, 1).getValues();
      for (let i = 0; i < keys.length; i++) {
        if (String(keys[i][0]) === s.key) { target = i + 2; break; }
      }
    }
    if (target) sh.getRange(target, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

// Remove all rows of a subject ("Chemistry", "Physics" or "ALL").
// Removed rows are copied to an "Archive" sheet first, so nothing is lost by accident.
// Delete the Archive sheet yourself if you want them gone for good.
function reset_(subject) {
  const subjects = subject === "ALL" ? ["Chemistry", "Physics"] : [subject];
  if (subjects.some(s => s !== "Chemistry" && s !== "Physics")) return out_({ ok: false, error: "bad subject" });

  const sh = sheet_();
  const last = sh.getLastRow();
  let archived = 0;
  if (last > 1) {
    const rows = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
    const keep = [], gone = [];
    rows.forEach(r => (subjects.indexOf(String(r[1])) >= 0 ? gone : keep).push(r));

    if (gone.length) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let ar = ss.getSheetByName(ARCHIVE_NAME);
      if (!ar) {
        ar = ss.insertSheet(ARCHIVE_NAME);
        ar.appendRow(HEADERS.concat(["archivedAt"]));
        ar.setFrozenRows(1);
        ar.getRange("A:D").setNumberFormat("@");
      }
      const stamp = new Date().toISOString();
      const out = gone.map(r => r.concat([stamp]));
      ar.getRange(ar.getLastRow() + 1, 1, out.length, HEADERS.length + 1).setValues(out);
      archived = gone.length;

      sh.getRange(2, 1, last - 1, HEADERS.length).clearContent();
      if (keep.length) sh.getRange(2, 1, keep.length, HEADERS.length).setValues(keep);
    }
  }

  const epochs = getEpochs_();
  const now = Date.now();
  subjects.forEach(s => { epochs[s] = now; });
  PropertiesService.getScriptProperties().setProperty("epochs", JSON.stringify(epochs));
  return out_({ ok: true, archived: archived, epochs: epochs });
}
