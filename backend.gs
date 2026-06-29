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
  var action = e && e.parameter && e.parameter.action;
  if (action === 'data') return dataRows_();
  if (action === 'photo') return photoB64_(e.parameter.id);
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
    if (body.action === 'report') return generateReport_();
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

/** Return all non-test observations as JSON objects (for the Insights view). */
function dataRows_() {
  var sh = getSheet_();
  var last = sh.getLastRow();
  var rows = [];
  if (last > 1) {
    var vals = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
    var iT = HEADERS.indexOf('test');
    vals.forEach(function (r) {
      if (isTest_(r[iT])) return;
      var o = {};
      HEADERS.forEach(function (h, i) { if (h !== 'test') o[h] = r[i]; });
      // Parse Drive file IDs out of the stored viewer links so each event's
      // photos can be fetched by id via ?action=photo (used for full reports).
      o.photoIds = String(o.photoUrls || '').split(/\s+/)
        .map(function (u) { var m = /\/d\/([-\w]+)/.exec(u); return m ? m[1] : ''; })
        .filter(String);
      rows.push(o);
    });
  }
  return json_({ ok: true, rows: rows });
}

/**
 * Generate a written AI report with Claude (Opus 4.8).
 *
 * Requires an Anthropic API key stored as a Script Property named
 * ANTHROPIC_API_KEY (Project Settings ▸ Script Properties ▸ Add property).
 * Get a key at https://console.anthropic.com (this is pay-per-use, separate
 * from any Claude.ai subscription — roughly a few cents per report). Without a
 * key this returns {ok:false, error:'no_key'} and the page falls back to the
 * free copy-the-prompt-into-claude.ai flow.
 */
function generateReport_() {
  var key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!key) return json_({ ok: false, error: 'no_key' });

  var rowsResp = JSON.parse(dataRows_().getContent());
  var rows = rowsResp.rows || [];
  if (!rows.length) return json_({ ok: false, error: 'no_data' });

  var prompt = buildReportPrompt_(rows);
  var payload = {
    model: 'claude-opus-4-8',
    max_tokens: 8000,
    thinking: { type: 'adaptive' },
    output_config: { effort: 'medium' },
    messages: [{ role: 'user', content: prompt }]
  };
  try {
    var resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    if (code !== 200) {
      return json_({ ok: false, error: 'Claude API ' + code + ': ' + resp.getContentText().slice(0, 300) });
    }
    var data = JSON.parse(resp.getContentText());
    var text = (data.content || [])
      .filter(function (b) { return b.type === 'text'; })
      .map(function (b) { return b.text; })
      .join('\n').trim();
    return json_({ ok: true, report: text || '(empty response)', model: data.model, n: rows.length });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function buildReportPrompt_(rows) {
  return 'You are an urban-freight researcher analysing a city-centre delivery survey of Antwerp '
    + '(the Meir / central shopping district), strictly from the CARRIER (delivery operator) perspective — Group 5.\n\n'
    + 'Below are ' + rows.length + ' field observations (one delivery stop each), collected around the morning peak.\n\n'
    + 'Write a clear, well-structured report (use headings and short paragraphs/bullets). Please:\n'
    + '1. Summarise recurring patterns in vehicle types, parking behaviour, walking distances, stop durations and powertrain (combustion vs electric).\n'
    + "2. Identify the carrier's main operational pain points (parking, waiting, congestion, access restrictions, long carries).\n"
    + '3. Quantify where possible (counts, averages, % of stops with each challenge).\n'
    + '4. Analyse delivery TIMING: using arrival times plus the fields "accessZone" (inside a pedestrian/restricted-access zone), '
    + '"timeCritical", and "timingReason", assess whether early/concentrated deliveries appear DRIVEN BY local access or time-window '
    + 'restrictions (forced) versus CHOSEN by carriers for operational efficiency. Report the share of stops in restricted zones and how it relates to timing.\n'
    + '5. Give evidence-based recommendations to improve carrier efficiency and cut emissions (consolidation, off-peak windows, micro-hubs, cargo bikes, loading-zone policy).\n'
    + '6. Clearly separate OBSERVATIONS from INTERPRETATIONS and from PROPOSED SOLUTIONS. Flag small sample size where relevant.\n\n'
    + 'DATA (JSON):\n' + JSON.stringify(rows);
}

/**
 * Return one photo as base64 (GET ...?action=photo&id=<driveFileId>), so a
 * report can correlate the image with its observation. SECURITY: only files
 * that live inside the G5 Survey Photos folder are served — the endpoint
 * cannot be used to read arbitrary files in the account's Drive.
 */
function photoB64_(id) {
  if (!id) return json_({ ok: false, error: 'no_id' });
  try {
    var file = DriveApp.getFileById(id);
    var inFolder = false, parents = file.getParents();
    while (parents.hasNext()) { if (parents.next().getName() === PHOTO_FOLDER) { inFolder = true; break; } }
    if (!inFolder) return json_({ ok: false, error: 'not_in_folder' });
    var blob = file.getBlob();
    return json_({ ok: true, id: id, name: file.getName(),
                   mime: blob.getContentType(), b64: Utilities.base64Encode(blob.getBytes()) });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
