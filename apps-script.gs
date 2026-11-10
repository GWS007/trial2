// Paste this into Google Sheet -> Extensions -> Apps Script, then Deploy as a Web app
// (Execute as: Me, Who has access: Anyone).
//
// Row 1 of the sheet should be:
// Name | Roll No. | Document 1 | Document 2 | Document 3 | Comment | Time

function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const d = JSON.parse(e.postData.contents);

  // stop values starting with = + - @ from being treated as formulas
  const safe = v => /^[=+\-@]/.test(String(v)) ? "'" + v : String(v ?? "");

  sheet.appendRow([
    safe(d.name),
    safe(d.roll),
    safe(d.d1),
    safe(d.d2),
    safe(d.d3),
    safe(d.comt),
    new Date()
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
