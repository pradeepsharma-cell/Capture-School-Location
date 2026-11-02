/**
 * MSMS School Location Capture - Backend
 * Sheet columns: A Division | B Project office | C School Name | D Capture School Google Map Location
 */
var SHEET_GID = 1877270736; // aapki sheet ka gid

function out_(o, e) {
  var cb = e && e.parameter && e.parameter.callback;
  var txt = JSON.stringify(o);
  if (cb && /^[A-Za-z_][A-Za-z0-9_]{0,60}$/.test(cb)) {
    // JSONP: browser CORS ke bina data le sakta hai
    return ContentService.createTextOutput(cb + '(' + txt + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(txt).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var a = e && e.parameter && e.parameter.action;
  if (a === 'list') {
    try { return out_({ ok: true, data: getSchools() }, e); }
    catch (err) { return out_({ ok: false, error: String(err.message || err) }, e); }
  }
  if (a === 'save') {
    try {
      var p = e.parameter;
      return out_(saveLocation(Number(p.row), p.school, p.lat, p.lng, p.acc, p.link), e);
    } catch (err) { return out_({ ok: false, error: String(err.message || err) }, e); }
  }
  try {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('School Location Capture')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    // Index file na bani ho to bhi API (GitHub page) chalti rahe
    return HtmlService.createHtmlOutput('<p style="font-family:sans-serif">School Location API chal raha hai. ' +
      'Page kholne ke liye GitHub wala link use karein.</p>');
  }
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === SHEET_GID) return sheets[i];
  }
  return sheets[0];
}

/** Returns [{r: rowNumber, d: division, p: projectOffice, s: school, l: link}] */
function getSchools() {
  var sh = getSheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, 4).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var d = String(vals[i][0]).trim(), p = String(vals[i][1]).trim(), s = String(vals[i][2]).trim();
    if (!d || !s) continue;
    out.push({ r: i + 2, d: d, p: p, s: s, l: String(vals[i][3] || '').trim() });
  }
  return out;
}

/** Saves the Google Maps link for the given row (verifies school name to avoid wrong-row writes) */
function saveLocation(row, schoolName, lat, lng, accuracy, link) {
  var finalLink;
  if (link) {
    link = String(link).trim();
    var okHost = /^https:\/\/(maps\.app\.goo\.gl|goo\.gl\/maps|(www\.)?google\.[a-z.]+\/maps|maps\.google\.[a-z.]+)(\/|\?)/i;
    if (!okHost.test(link) || link.length > 500 || /\s/.test(link)) {
      throw new Error('Sirf Google Maps ka link daalein.');
    }
    finalLink = link;
  } else {
    lat = Number(lat); lng = Number(lng);
    if (!isFinite(lat) || !isFinite(lng) || Math.abs(lat) > 90 || Math.abs(lng) > 180) {
      throw new Error('Invalid location');
    }
    finalLink = 'https://www.google.com/maps?q=' + lat.toFixed(6) + ',' + lng.toFixed(6);
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = getSheet_();
    var current = String(sh.getRange(row, 3).getValue()).trim();
    if (current !== String(schoolName).trim()) {
      throw new Error('Sheet change ho gayi hai, page refresh karke dobara try karein.');
    }
    var cell = sh.getRange(row, 4);
    cell.setValue(finalLink);
    cell.setNote('Captured: ' + Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd-MMM-yyyy HH:mm') +
      ' | Accuracy: ' + Math.round(accuracy || 0) + ' m');
    SpreadsheetApp.flush();
    return { ok: true, link: finalLink };
  } finally {
    lock.releaseLock();
  }
}
