/*
Google Sheets lead capture endpoint for the Web Forge landing page.

Setup:
1. Create a Google Sheet with a tab named "Leads".
2. Open Extensions → Apps Script.
3. Paste this code.
4. Replace SHEET_ID with your spreadsheet ID.
5. Deploy → New deployment → Web app.
6. Execute as: Me. Who has access: Anyone.
7. Copy the Web App URL and paste it into index.html as GOOGLE_SCRIPT_URL.
*/

const SHEET_ID = 'PASTE_YOUR_GOOGLE_SHEET_ID_HERE';
const SHEET_NAME = 'Leads';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents || '{}');
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME) || SpreadsheetApp.openById(SHEET_ID).insertSheet(SHEET_NAME);

    ensureHeaders_(sheet);

    sheet.appendRow([
      new Date(),
      data.name || '',
      data.email || '',
      data.phone || '',
      data.service || '',
      data.existingWebsite || '',
      data.goal || '',
      data.industry || '',
      data.budget || '',
      data.timeline || '',
      data.message || '',
      data.leadScore || '',
      data.submittedAt || '',
      'New'
    ]);

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doGet() {
  return json_({ ok: true, message: 'Lead endpoint is live.' });
}

function ensureHeaders_(sheet) {
  const headers = [
    'Received At',
    'Name',
    'Email',
    'Phone',
    'Service',
    'Existing Website',
    'Goal',
    'Industry',
    'Budget',
    'Timeline',
    'Message',
    'Lead Score',
    'Submitted At',
    'Status'
  ];

  const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const isEmpty = firstRow.every(cell => !cell);

  if (isEmpty) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
