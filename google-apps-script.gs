/**
 * Al Ansari Exchange Customer Feedback — Google Apps Script Web App (v2 questionnaire)
 *
 * doPost: receives feedback from the form, appends a row to Sheet1.
 * doGet:  ?action=branches returns the branch list from Sheet2 (col A,
 *         skipping the header row) as JSON for the form's dropdown.
 *
 * Sheet1 headers (row 1), in this exact order:
 *   Timestamp | Branch | Civil ID | Mobile | Happy | Respect | Speed |
 *   Visit Again | NPS | Improvement
 *
 * Sheet2 layout: A1 = "Branch" (header), A2..An = one branch name per row.
 *
 * After editing this code: Deploy -> Manage deployments -> pencil icon ->
 * Version: "New version" -> Deploy. The /exec URL stays the same.
 */

const SHEET_ID = "1WNnheZhnFfKlM4lqdOPlMZi6OizfeVB6VYzxLjIczpQ";
const SHEET_NAME = "Sheet1";
const BRANCH_SHEET = "Sheet2";

function doPost(e) {
  const data = JSON.parse(e.postData.contents);

  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  sheet.appendRow([
    new Date(),
    data.branch || "Unknown Branch",
    "'" + (data.civilId || ""),
    "'" + (data.mobile || ""),
    Number(data.happy) || "",
    Number(data.respect) || "",
    Number(data.speed) || "",
    Number(data.visitAgain) || "",
    isNaN(Number(data.nps)) || data.nps === "" || data.nps === null || data.nps === undefined ? "" : Number(data.nps),
    data.improve || ""
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  if (e && e.parameter && e.parameter.action === "branches") {
    let branches = [];
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(BRANCH_SHEET);
    if (sheet) {
      const last = sheet.getLastRow();
      if (last >= 2) {
        branches = sheet.getRange(2, 1, last - 1, 1).getValues()
          .map(function(row) { return String(row[0]).trim(); })
          .filter(function(name) { return name.length > 0; });
      }
    }
    return ContentService
      .createTextOutput(JSON.stringify({ branches: branches }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService
    .createTextOutput("BEC Feedback endpoint is live OK")
    .setMimeType(ContentService.MimeType.TEXT);
}
