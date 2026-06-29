/**
 * Group 5 — Carrier Perspective field survey: shared backend.
 *
 * ── SETUP (one person, ~3 min) ───────────────────────────────────────────
 * 1. Create a new Google Sheet (sheets.new). This is where all data lands.
 * 2. Extensions ▸ Apps Script. Delete any sample code, paste THIS whole file.
 * 3. Click Save (disk icon).
 * 4. Deploy ▸ New deployment ▸ (gear) Web app.
 *       Execute as  : Me
 *       Who has access : Anyone            <-- important
 *    Deploy ▸ Authorize access (allow the permissions it asks for).
 * 5. Copy the "Web app URL" (ends in /exec) and paste it into the page's
 *    ⚙️ Shared upload setup (or have it baked into index.html).
 *
 * ── UPDATING THIS SCRIPT LATER ───────────────────────────────────────────
 * After editing, you MUST publish a new version or the old code keeps running:
 *   Deploy ▸ Manage deployments ▸ (pencil) Edit ▸ Version: New version ▸ Deploy.
 * The web-app URL stays the same, so nothing else needs changing.
 *
 * Photos are saved to a Drive folder "G5 Survey Photos" and linked from the
 * sheet. The header row is created/repaired automatically.
 * ─────────────────────────────────────────────────────────────────────────
 */

var SHEET_NAME   = 'observations';
var PHOTO_FOLDER = 'G5 Survey Photos';

// Note: original columns 1-23 keep their positions; new fields are appended at
// the end so a header upgrade never shifts rows already written to the sheet.
var HEADERS = ['id','ts','date','observer','location','gps','arrive','depart','carrier',
  'vehicleType','powertrain','vehicleSize','packages','people','packageSize','equipment',
  'establishment','parking','distance','issues','challenges','notes','photoUrls',
  'test','accessZone','timeCritical','timingReason'];

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    return sh;
  }
  // Repair / upgrade the header row if it doesn't match (e.g. new 'test' column).
  var cur = sh.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  var same = cur.length === HEADERS.length && HEADERS.every(function (h, i) { return cur[i] === h; });
  if (!same) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function getFolder_() {
  var it = DriveApp.getFoldersByName(PHOTO_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_FOLDER);
}

function isTest_(v) { return String(v).toLowerCase() === 'true'; }

function durMin_(a, b) {
  var m = function (t) { var x = /^(\d{1,2}):(\d{2})$/.exec(String(t).trim()); return x ? (+x[1] * 60 + +x[2]) : null; };
  var A = m(a), B = m(b);
  if (A == null || B == null) return null;
  var d = B - A; if (d < 0) d += 1440;
  return d;
}

/** Live dashboard: per-member counts + aggregate stats (GET ...?action=summary).
 *  Test rows are excluded from every figure. */
function doGet(e) {
  var sh = getSheet_();
  var last = sh.getLastRow();
  var byMember = {}, total = 0, durSum = 0, durN = 0, elec = 0, comb = 0, chall = {};
  if (last > 1) {
    var vals = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
    var iObs = HEADERS.indexOf('observer'), iPow = HEADERS.indexOf('powertrain'),
        iArr = HEADERS.indexOf('arrive'),  iDep = HEADERS.indexOf('depart'),
        iCh  = HEADERS.indexOf('challenges'), iT = HEADERS.indexOf('test');
    vals.forEach(function (r) {
      if (isTest_(r[iT])) return;
      var n = String(r[iObs] || '').trim();
      if (n) { byMember[n] = (byMember[n] || 0) + 1; total++; }
      var d = durMin_(r[iArr], r[iDep]); if (d != null) { durSum += d; durN++; }
      var p = String(r[iPow] || '').toLowerCase();
      if (p === 'electric') elec++; else if (p === 'combustion') comb++;
      String(r[iCh] || '').split(';').forEach(function (c) { c = c.trim(); if (c) chall[c] = (chall[c] || 0) + 1; });
    });
  }
  var topChallenge = '', topN = 0;
  for (var k in chall) { if (chall[k] > topN) { topN = chall[k]; topChallenge = k; } }
  var stats = {
    avgDuration: durN ? Math.round(durSum / durN) : null,
    electric: elec, combustion: comb,
    pctElectric: (elec + comb) ? Math.round(elec / (elec + comb) * 100) : null,
    topChallenge: topChallenge, topChallengeN: topN
  };
  return json_({ ok: true, total: total, byMember: byMember, stats: stats,
                 url: SpreadsheetApp.getActiveSpreadsheet().getUrl() });
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

      var existing = findRowById_(sh, rec.id);
      if (existing.row) {
        return json_({ ok: true, duplicate: true, photoUrls: existing.photoUrls });
      }

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
        if (h === 'test')       return rec.test ? true : '';
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
