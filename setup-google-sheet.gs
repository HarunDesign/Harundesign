// HarunDesign: receives website briefs into Google Sheets, saves uploads to Drive, and emails you.
// Setup: Google Sheet > Extensions > Apps Script > paste this > Deploy > New deployment >
// Type: Web app > Execute as: Me > Who has access: Anyone > Deploy > copy the URL into SHEET_URL in index.html.
const OWNER = 'harundesign.web@gmail.com';
const FOLDER = 'HarunDesign Client Uploads';
const HEAD = ['Date', 'Name', 'Email', 'WhatsApp', 'Business', 'Project type', 'Budget', 'Deadline', 'Reference links', 'Details', 'Uploaded files'];

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sheet.getLastRow() === 0) { sheet.appendRow(HEAD); sheet.getRange(1, 1, 1, HEAD.length).setFontWeight('bold'); sheet.setFrozenRows(1); }
    const it = DriveApp.getFoldersByName(FOLDER);
    const root = it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER);
    const sub = root.createFolder((d.name || 'Client') + ' - ' + new Date().toISOString().slice(0, 10));
    const links = (d.files || []).map(f => {
      const file = sub.createFile(Utilities.newBlob(Utilities.base64Decode(f.data), f.type, f.name));
      return file.getUrl();
    });
    if (links.length) sheet.getRange(sheet.getLastRow() + 1, 1); // keep row order stable
    sheet.appendRow([new Date(), d.name, d.email, d.phone, d.business, d.project_type, d.budget, d.deadline, d.reference_links, d.details, links.join('\n')]);
    MailApp.sendEmail(OWNER, 'New project brief from ' + d.name,
      'Name: ' + d.name + '\nEmail: ' + d.email + '\nWhatsApp: ' + d.phone + '\nBusiness: ' + d.business +
      '\nProject: ' + d.project_type + '\nBudget: ' + d.budget + '\nDeadline: ' + d.deadline +
      '\nLinks: ' + d.reference_links + '\n\n' + d.details + '\n\nUploads:\n' + (links.join('\n') || 'none'));
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  }
}
