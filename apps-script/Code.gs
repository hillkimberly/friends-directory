/**
 * The Friends Directory — Apps Script backend
 * -------------------------------------------------
 * doPost : the intake bot sends a listing → we append a row (Status = Pending).
 * doGet  : the pages request data →
 *            ?action=listings  (default) → Approved listings
 *            ?action=grants                → Active grants
 *          Returns JSONP when a ?callback= is supplied (the pages use this so they
 *          can read the data cross-origin from Netlify), plain JSON otherwise.
 *
 * SETUP
 *   1. Open/create a Google Sheet. Extensions ▸ Apps Script ▸ paste this file.
 *   2. Run doGet once from the editor to authorize and auto-create both tabs.
 *   3. Deploy ▸ New deployment ▸ Web app → Execute as: Me · Access: Anyone.
 *   4. Copy the /exec URL into API_URL in directory.html, grants.html, directory-intake.html.
 *
 * IMPORTANT: after ANY edit to this file, redeploy — Manage deployments ▸ edit ▸
 * New version — or the live URL keeps serving the old code.
 *
 * Publish a listing: set its Status cell to Approved (or Featured to pin + highlight).
 */

var SHEET_TAB  = 'Directory Listings';
var GRANTS_TAB = 'Grants';

var HEADERS = ['Status','Submitted At','Business','Owner','Category','City','Country',
               'Serves','Languages','Website','Referred By','Description','Contact'];

var GRANT_HEADERS = ['Status','Amount','Title','Funder','For Who','Deadline','Link'];

function getSheet_()       { return ensureTab_(SHEET_TAB, HEADERS); }
function getGrantsSheet_() { return ensureTab_(GRANTS_TAB, GRANT_HEADERS); }

function ensureTab_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// Prevent formula injection: a leading = + - @ becomes text, not a formula.
function safe_(v) {
  var s = (v == null ? '' : String(v));
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    getSheet_().appendRow([
      safe_(d.status || 'Pending'),
      safe_(d.submittedAt || new Date().toISOString()),
      safe_(d.business), safe_(d.owner), safe_(d.category),
      safe_(d.city), safe_(d.country), safe_(d.serves),
      safe_(d.languages), safe_(d.website), safe_(d.referredBy),
      safe_(d.description), safe_(d.contact)
    ]);
    return out_({ ok: true }, null);
  } catch (err) {
    return out_({ ok: false, error: String(err) }, null);
  }
}

function doGet(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  var action = p.action || 'listings';
  var data = (action === 'grants') ? grantsData_() : listingsData_();
  return out_(data, p.callback);
}

function listingsData_() {
  var values = getSheet_().getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    var status = String(r[0] || '').toLowerCase();
    if (status !== 'approved' && status !== 'live' && status.indexOf('featured') === -1) continue;
    out.push({
      business: r[2], owner: r[3], category: r[4],
      city: r[5], country: r[6], serves: r[7],
      languages: String(r[8] || '').split(',').map(function (s) { return s.trim(); }).filter(String),
      website: r[9], referredBy: r[10], description: r[11], contact: r[12],
      featured: /featured|\u2605/i.test(String(r[0] || ''))
    });
  }
  return out;
}

function grantsData_() {
  var values = getGrantsSheet_().getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    if (String(r[0] || '').toLowerCase() !== 'active') continue;
    out.push({ amount: r[1], title: r[2], funder: r[3], forWho: r[4], deadline: r[5], link: r[6] });
  }
  return out;
}

function out_(obj, callback) {
  var body = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + body + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}
