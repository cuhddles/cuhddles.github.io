// Cuhddles waitlist -> Google Sheet
// Paste into Extensions > Apps Script of the waitlist Google Sheet, then Deploy > New deployment > Web app
// (Execute as: Me, Who has access: Anyone). Copy the Web app URL into SHEET_URL in index.html.
function doPost(e) {
  var p = (e && e.parameter) || {};
  if (p.website) return out({ok: true}); // honeypot: bots fill this hidden field
  var email = String(p.email || '').trim().toLowerCase().slice(0, 200);
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) return out({ok: false, error: 'bad_email'});
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sheet.getLastRow() === 0) sheet.appendRow(['Signed up', 'Email', 'TikTok', 'Source']);
    var emails = sheet.getLastRow() > 1 ? sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues().map(function (r) { return r[0]; }) : [];
    if (emails.indexOf(email) === -1) {
      var clean = function (v) { v = String(v || '').slice(0, 60); return /^[=+\-@]/.test(v) ? "'" + v : v; };
      sheet.appendRow([new Date(), email, clean(p.handle), clean(p.source)]);
    }
  } finally { lock.releaseLock(); }
  return out({ok: true});
}
function doGet() { return out({ok: true, service: 'cuhddles waitlist'}); }
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
