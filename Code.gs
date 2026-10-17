/**
 * Google Apps Script web app — saves landing-page form responses to a Sheet.
 *
 * Setup:
 *   1. In your Google Sheet: Extensions > Apps Script, replace Code.gs with this, save.
 *   2. Deploy > New deployment > Web app.
 *        Execute as: Me
 *        Who has access: Anyone
 *   3. Copy the /exec URL — you'll paste it into index.html as SCRIPT_URL.
 */

var SHEET_NAME = 'Responses';
var HEADERS = ['Timestamp', 'Name', 'Email', 'Phone'];

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  try {
    var params = (e && e.parameter) || {};
    var name  = String(params.name  || '').trim();
    var email = String(params.email || '').trim();
    var phone = String(params.phone || '').trim();

    if (!name || !email || !phone) {
      return jsonResponse(false, 'Missing required field(s): name, email, phone');
    }
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      return jsonResponse(false, 'Invalid email address');
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);

    // Auto-create the tab with bold headers if it doesn't exist yet
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    }

    sheet.appendRow([new Date(), name, email, phone]);

    return jsonResponse(true, 'Response saved');
  } catch (err) {
    return jsonResponse(false, String(err && err.message ? err.message : err));
  }
}

function jsonResponse(success, message) {
  return ContentService
    .createTextOutput(JSON.stringify({ success: success, message: message }))
    .setMimeType(ContentService.MimeType.JSON);
}