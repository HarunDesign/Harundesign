// HarunDesign: project briefs + client reviews -> Google Sheets, uploads -> Drive, email alerts.
// SETUP: open your Google Sheet > Extensions > Apps Script > replace ALL code with this file > Save.
// Then Deploy > Manage deployments > pencil icon > Version: New version > Deploy. (Keeps the same URL.)
// REVIEWS: new reviews arrive in a tab called "Reviews" with Status = Pending.
// Set Status to Approved to show a review on the website, or Rejected to hide it.
const SHEET_ID = '';   // only needed if this script was not created from inside the sheet
const OWNER = 'harundesign.web@gmail.com';
const FOLDER = 'HarunDesign Client Uploads';
const HEAD = ['Date', 'Name', 'Email', 'WhatsApp', 'Business', 'Project type', 'Budget', 'Deadline', 'Reference links', 'Details', 'Uploaded files'];
const REVIEW_TAB = 'Reviews';
const REVIEW_HEAD = ['Date', 'Name', 'Business', 'Rating', 'Review', 'Status'];

function ss_() {
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No spreadsheet found. Create this script from Extensions > Apps Script inside the sheet, or set SHEET_ID.');
  return ss;
}
function briefSheet_() { return ss_().getSheets()[0]; }
function reviewSheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName(REVIEW_TAB);
  if (!sh) {
    sh = ss.insertSheet(REVIEW_TAB);
    sh.appendRow(REVIEW_HEAD);
    sh.getRange(1, 1, 1, REVIEW_HEAD.length).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('F2:F1000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Pending', 'Approved', 'Rejected'], true).build());
    sh.setColumnWidth(5, 420);
  }
  return sh;
}
const json_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

// Open the /exec URL in a browser: "HarunDesign form endpoint is live". With ?action=reviews it returns approved reviews.
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'reviews') {
    try {
      const rows = reviewSheet_().getDataRange().getValues().slice(1);
      const reviews = rows.filter(r => String(r[5]).trim() === 'Approved' && r[1] && r[4])
        .map(r => ({ name: String(r[1]), business: String(r[2]), rating: Number(r[3]) || 0, text: String(r[4]) }))
        .reverse();
      return json_({ ok: true, reviews: reviews });
    } catch (err) { return json_({ ok: false, error: String(err) }); }
  }
  return ContentService.createTextOutput('HarunDesign form endpoint is live');
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.type === 'review') return saveReview_(d);
    const sheet = briefSheet_();
    if (sheet.getLastRow() === 0) { sheet.appendRow(HEAD); sheet.getRange(1, 1, 1, HEAD.length).setFontWeight('bold'); sheet.setFrozenRows(1); }
    let links = [];
    if (d.files && d.files.length) {
      const it = DriveApp.getFoldersByName(FOLDER);
      const root = it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER);
      const sub = root.createFolder((d.name || 'Client') + ' - ' + new Date().toISOString().slice(0, 10));
      links = d.files.map(f => sub.createFile(Utilities.newBlob(Utilities.base64Decode(f.data), f.type, f.name)).getUrl());
    }
    sheet.appendRow([new Date(), d.name, d.email, d.phone, d.business, d.project_type, d.budget, d.deadline, d.reference_links, d.details, links.join('\n')]);
    return json_({ ok: true });
  } catch (err) {
    try { MailApp.sendEmail(OWNER, 'Website form: Google Sheet step failed', String(err)); } catch (x) {}
    return json_({ ok: false, error: String(err) });
  }
}

function saveReview_(d) {
  const name = String(d.name || '').trim().slice(0, 80);
  const text = String(d.text || '').trim().slice(0, 800);
  if (!name || text.length < 30) return json_({ ok: false, error: 'Invalid review' });
  const rating = Math.max(1, Math.min(5, Number(d.rating) || 5));
  const business = String(d.business || '').trim().slice(0, 100);
  // Stop spreadsheet formulas being injected through the form
  const safe = v => /^[=+\-@]/.test(v) ? "'" + v : v;
  reviewSheet_().appendRow([new Date(), safe(name), safe(business), rating, safe(text), 'Pending']);
  try { MailApp.sendEmail(OWNER, 'New review to approve from ' + name, name + ' (' + business + ') gave ' + rating + '/5:\n\n' + text + '\n\nTo publish it, open your sheet > Reviews tab > set Status to Approved.'); } catch (x) {}
  return json_({ ok: true });
}
