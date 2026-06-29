/**
 * Group 5 — Carrier Perspective field survey: shared backend.
 *
 * ── SETUP (one person, ~3 min) ───────────────────────────────────────────
 * 1. Create a new Google Sheet (sheets.new). This is where all data lands.
 * 2. Extensions ▸ Apps Script. Delete any sample code, paste THIS whole file.
 * 3. Click Save (disk icon).
 * 4. Deploy ▸ New deployment ▸ (gear) Web app.
 *       Description : g5-survey
 *       Execute as  : Me
 *       Who has access : Anyone            <-- important
 *    Deploy ▸ Authorize access (allow the permissions it asks for).
 * 5. Copy the "Web app URL" (ends in /exec).
 * 6. Open the HTML page ▸ ⚙️ Shared upload setup ▸ paste the URL ▸ Save URL.
 *    Then send that same HTML file to the other three teammates.
 *
 * Photos are saved to a Drive folder called "G5 Survey Photos" and linked
 * from the sheet. The header row is created automatically on first upload.
 * ─────────────────────────────────────────────────────────────────────────
 */

var SHEET_NAME  = 'observations';
var PHOTO_FOLDER = 'G5 Survey Photos';

var HEADERS = ['id','ts','date','observer','location','gps','arrive','depart','carrier',
  'vehicleType','powertrain','vehicleSize','packages','people','packageSize','equipment',
  'establishment','parking','distance','issues','challenges','notes','photoUrls'];

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function getFolder_() {
  var it = DriveApp.getFoldersByName(PHOTO_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_FOLDER);
}

/** Live per-member counts for the dashboard (GET ...?action=summary). */
function doGet(e) {
  var sh = getSheet_();
  var last = sh.getLastRow();
  var byMember = {}, total = 0;
  if (last > 1) {
    var col = HEADERS.indexOf('observer') + 1;
    var names = sh.getRange(2, col, last - 1, 1).getValues();
    names.forEach(function (r) {
      var n = String(r[0] || '').trim();
      if (n) { byMember[n] = (byMember[n] || 0) + 1; total++; }
    });
  }
  return json_({ ok: true, total: total, byMember: byMember });
}

/** Receive one observation (POST, body = {action:'add', record:{...}}). */
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var rec = body.record || {};
    var lock = LockService.getScriptLock();
    lock.waitLock(20000); // serialise concurrent uploads from 4 phones
    try {
      var sh = getSheet_();

      // de-dupe: same id already stored? return its photoUrls, don't add twice
      var existing = findRowById_(sh, rec.id);
      if (existing.row) {
        return json_({ ok: true, duplicate: true, photoUrls: existing.photoUrls });
      }

      // save photos (base64 data URLs) to Drive, collect links
      var photoUrls = [];
      var photos = rec.photos || [];
      if (photos.length) {
        var folder = getFolder_();
        for (var i = 0; i < photos.length; i++) {
          var url = savePhoto_(folder, photos[i], rec.id, i);
          if (url) photoUrls.push(url);
        }
      }

      sh.appendRow(HEADERS.map(function (h) {
        if (h === 'gps')        return rec.gps ? rec.gps.lat + ',' + rec.gps.lng : '';
        if (h === 'equipment')  return (rec.equipment || []).join('; ');
        if (h === 'issues')     return (rec.issues || []).join('; ');
        if (h === 'challenges') return (rec.challenges || []).join('; ');
        if (h === 'photoUrls')  return photoUrls.join('  ');
        return rec[h] != null ? rec[h] : '';
      }));

      return json_({ ok: true, photoUrls: photoUrls });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function savePhoto_(folder, dataUrl, id, i) {
  try {
    var m = /^data:(image\/\w+);base64,(.+)$/.exec(dataUrl);
    if (!m) return '';
    var blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], id + '_' + (i + 1) + '.jpg');
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) { return ''; }
}

function findRowById_(sh, id) {
  var last = sh.getLastRow();
  if (last < 2 || !id) return {};
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var r = 0; r < ids.length; r++) {
    if (String(ids[r][0]) === String(id)) {
      var purl = sh.getRange(r + 2, HEADERS.indexOf('photoUrls') + 1).getValue();
      return { row: r + 2, photoUrls: purl ? String(purl).split('  ') : [] };
    }
  }
  return {};
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
