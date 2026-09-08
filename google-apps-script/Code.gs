/**
 * Proof of Privacy — Investor Interest Form backend.
 *
 * Setup:
 * 1. Open the target Google Sheet (create a new one if you don't have one yet).
 * 2. Extensions > Apps Script.
 * 3. Delete any placeholder code and paste this file in.
 * 4. Change SHEET_NAME below if you want the responses tab to be named differently.
 * 5. Deploy > New deployment > type "Web app".
 *      Execute as: Me
 *      Who has access: Anyone
 * 6. Authorize when prompted (it'll warn "Google hasn't verified this app" —
 *    that's expected for a script you wrote yourself; click Advanced > Go to
 *    [project name] (unsafe) > Allow).
 * 7. Copy the Web app URL (ends in /exec) and paste it into
 *    GOOGLE_SHEET_WEBHOOK_URL in index.html.
 *
 * The Sheet's own sharing/visibility (who can open it in Drive) is completely
 * separate from this — set that in the Sheet's normal Share dialog. Setting
 * "Who has access: Anyone" here only lets anonymous visitors trigger this
 * doPost function (append one row); it does not grant them any ability to
 * open, read, or browse the spreadsheet itself.
 */

const SHEET_NAME = "Responses";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    const name = (data.name || "").toString().trim();
    const email = (data.email || "").toString().trim();
    const telegram = (data.telegram || "").toString().trim();
    const fund = (data.fund || "").toString().trim();
    const teams = Array.isArray(data.teams) ? data.teams : [];

    if (!name || !isValidEmail(email) || !telegram || !fund || teams.length === 0) {
      return jsonResponse({ ok: false, error: "Missing or invalid required fields." });
    }

    const sheet = getOrCreateSheet();
    sheet.appendRow([
      new Date(),
      name,
      email,
      telegram,
      fund,
      teams.join(", "),
    ]);

    return jsonResponse({ ok: true });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["Timestamp", "Investor Name", "Email", "Telegram Handle", "Fund / Company", "Selected Teams"]);
  }
  return sheet;
}

function isValidEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
