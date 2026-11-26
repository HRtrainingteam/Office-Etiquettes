/**
 * SMC Office Etiquettes Peer Rating — Google Apps Script
 * This receiver belongs ONLY to the Office-Etiquettes repository.
 * Paste into the Apps Script project attached to the results spreadsheet,
 * save, then update the existing web-app deployment to a new version.
 */
const SHEET_NAME = "Peer Ratings";
const EMPLOYEES = [
  "Harsh Arora",
  "Madhur Garg",
  "Rahul Kumar",
  "Chirag Papreja",
  "Vijay Kumar Gupta",
  "Mritunjay Kumar Tiwari",
  "Samridhi Gupta",
  "Kanhaiya Kumar Singh",
  "Anjali Jaiswal",
  "Charu",
  "Sandeep Kaushik",
  "Sunil Gaur",
  "Sandeep Kumar Rajak",
  "Saurabh Kumar Gupta",
  "Atul Bisht",
  "Vikas",
  "Santosh Adhikari",
  "Rajesh Kumar Bhatt",
  "Anjali Mishra",
  "Rajiv Joshi"
];
const CATEGORIES = [
  "Punctuality",
  "Respectful Behaviour",
  "Acts with Integrity",
  "Appropriate Dressing & Professional Appearance",
  "Team Collaboration"
];
const HEADERS = [
  "Submitted At",
  "Submitter Name",
  "Employee Code",
  "Rated Employee",
  ...CATEGORIES
];

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("No submission data was received.");
    }

    const data = JSON.parse(e.postData.contents);
    const submitter = String(data.submitter || "").trim();
    const employeeCode = String(data.employeeCode || "").trim();

    if (!EMPLOYEES.includes(submitter)) {
      throw new Error("Unknown submitter: " + submitter);
    }
    if (!employeeCode) {
      throw new Error("Employee code is missing.");
    }
    if (!Array.isArray(data.ratings) || data.ratings.length !== EMPLOYEES.length - 1) {
      throw new Error("Expected ratings for exactly 19 colleagues.");
    }

    const expectedTargets = EMPLOYEES.filter(name => name !== submitter);
    const receivedTargets = data.ratings.map(item => String(item.employee || "").trim());

    if (new Set(receivedTargets).size !== receivedTargets.length) {
      throw new Error("Duplicate colleague found in the ratings.");
    }
    if (expectedTargets.some(name => !receivedTargets.includes(name)) ||
        receivedTargets.some(name => !expectedTargets.includes(name))) {
      throw new Error("The ratings must include every colleague except the submitter.");
    }

    const submittedAt = new Date();
    const rows = data.ratings.map(item => {
      const scores = CATEGORIES.map(category => {
        const score = Number(item.scores && item.scores[category]);
        if (!Number.isInteger(score) || score < 1 || score > 5) {
          throw new Error("Invalid score for " + item.employee + " — " + category);
        }
        return score;
      });
      return [submittedAt, submitter, employeeCode, item.employee, ...scores];
    });

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      throw new Error("No active spreadsheet is linked to this Apps Script project.");
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      let sheet = spreadsheet.getSheetByName(SHEET_NAME);
      if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);

      if (sheet.getLastRow() === 0) {
        sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      }

      const firstRow = sheet.getLastRow() + 1;
      sheet.getRange(firstRow, 1, rows.length, HEADERS.length).setValues(rows);
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, rowsSaved: rows.length }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error(error);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(error.message || error) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
