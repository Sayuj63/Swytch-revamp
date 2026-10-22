/**
 * Swytch Studio — Contact form → Google Sheet
 * ------------------------------------------------------------
 * Appends every website contact-form submission as a new row.
 *
 * SETUP (one time, ~2 minutes):
 *   1. Create a new Google Sheet (sheet.new). Name it e.g. "Swytch Leads".
 *   2. In that sheet: Extensions → Apps Script.
 *   3. Delete whatever code is there, paste THIS whole file, Save.
 *   4. Click Deploy → New deployment → gear icon → Web app.
 *        - Description:    Swytch leads endpoint
 *        - Execute as:     Me (your Google account)
 *        - Who has access: Anyone
 *      Click Deploy, then Authorize access (allow your own account).
 *   5. Copy the "Web app URL" (ends in /exec).
 *   6. Paste that URL into swytch-contact-typeform.html where it says
 *      SHEET_ENDPOINT = '...'  (one line, clearly marked).
 *
 * That's it. New leads will append to the "Leads" tab automatically.
 * To get an .xlsx: File → Download → Microsoft Excel (.xlsx).
 */

// Which tab to write to (created automatically if it doesn't exist).
var SHEET_NAME = 'Leads';

// Column order. Keys must match the fields the form POSTs.
var COLUMNS = ['timestamp', 'name', 'brand', 'need', 'budget', 'details', 'email', 'phone'];
var HEADERS = ['Timestamp', 'Name', 'Brand', 'Looking for', 'Budget', 'Details', 'Email', 'Phone'];

function doPost(e) {
  try {
    var data = {};
    if (e && e.postData && e.postData.contents) {
      // Form sends JSON as text/plain to avoid a CORS preflight.
      data = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      data = e.parameter;
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }

    // Write the header row once (first time the tab is empty).
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    }

    var row = COLUMNS.map(function (key) {
      if (key === 'timestamp') {
        // Local time in IST, readable in the sheet.
        return Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd HH:mm:ss');
      }
      return data[key] != null ? String(data[key]) : '';
    });

    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Lets you open the /exec URL in a browser to confirm it's live.
function doGet() {
  return ContentService
    .createTextOutput('Swytch leads endpoint is live.')
    .setMimeType(ContentService.MimeType.TEXT);
}
