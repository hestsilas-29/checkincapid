/**
 * CAP Cadet Check-In backend (Google Apps Script, bound to a Google Sheet).
 *
 * Sheet tabs:
 *   Log       - master history: Date, Time, CAPID
 *   10-8-26   - one attendance sheet per meeting date (tab-safe form of 10/8/26)
 *
 * The date tabs are created automatically when the first check-in for a date
 * is recorded. Existing Log rows are migrated when setup() is run.
 */

const TZ = 'America/New_York';
const LOG_HEADERS = ['Date', 'Time', 'CAPID'];
const DATE_HEADERS = ['CAPID', 'Time'];

/* ---------- one-time setup: run this from the editor ---------- */
function cleanupLegacyNameFields_(ss) {
  const log = ss.getSheetByName('Log');
  if (log && log.getMaxColumns() >= 4 && String(log.getRange(1, 4).getValue()).toLowerCase() === 'name') {
    log.deleteColumn(4);
  }

  ss.getSheets().forEach(sh => {
    const n = sh.getName();
    if (!/^\d{1,2}-\d{1,2}-\d{2}$/.test(n)) return;
    if (sh.getMaxColumns() >= 3 && String(sh.getRange(2, 3).getValue()).toLowerCase() === 'name') {
      // Old date tabs used a merged A1:C1 date header. Break that merge before
      // removing the old Name column so setup is safe on existing sheets.
      sh.getRange('A1:C1').breakApart();
      sh.deleteColumn(3);
      sh.getRange('A1:B1').merge();
      sh.getRange('A1:B1').setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center');
      sh.setFrozenRows(2);
    }
  });
}
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  cleanupLegacyNameFields_(ss);
  const log = logSheet_();

  const oldRoster = ss.getSheetByName('Roster');
  if (oldRoster && ss.getSheets().length > 1) ss.deleteSheet(oldRoster);

  // Keep the workbook organized: Log first, dated attendance tabs next.
  moveSheet_(log, 1);

  // Backfill dated sheets from any existing Log data.
  const rows = log.getDataRange().getValues().slice(1);
  const byDate = {};
  rows.forEach(r => {
    const date = String(r[0] || '');
    const capid = String(r[2] || '');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !capid) return;
    (byDate[date] || (byDate[date] = [])).push([capid, String(r[1] || '')]);
  });
  Object.keys(byDate).sort().forEach(date => syncDateSheet_(date, byDate[date]));

  ensureTrigger_();
}

/* ---------- web app entry points ---------- */
function doGet() {
  return json_({ ok: true, service: 'cap-checkin' });
}

function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'Bad request' });
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    switch (req.action) {
      case 'checkin':    return json_(checkin_(req));
      case 'getConfig':  return json_(getConfig_(req));
      case 'saveConfig': return json_(saveConfig_(req));
      case 'sendTest':   return json_(sendTest_(req));
      default:           return json_({ ok: false, error: 'Unknown action' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- check-in ---------- */
function checkin_(req) {
  const key = props_().getProperty('STATION_KEY');
  if (!key) return { ok: false, error: 'Not set up yet. Open the config page first.' };
  if (String(req.key) !== key) return { ok: false, error: 'Wrong station code', badKey: true };

  const capid = String(req.capid || '').replace(/\D/g, '');
  if (capid.length < 4 || capid.length > 8) return { ok: false, error: 'Unreadable card, scan again' };

  let when = req.ts ? new Date(req.ts) : new Date();
  if (isNaN(when.getTime())) when = new Date();
  const today = fmt_(when, 'yyyy-MM-dd');
  const time = fmt_(when, 'HH:mm:ss');

  const log = logSheet_();
  const rows = log.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === today && String(rows[i][2]) === capid) {
      return {
        ok: true,
        duplicate: true,
        capid: capid,
        time: String(rows[i][1])
      };
    }
  }

  const r = log.getLastRow() + 1;
  log.getRange(r, 1, 1, 3).setNumberFormat('@').setValues([[today, time, capid]]);

  appendDateRow_(today, capid, time);
  return { ok: true, duplicate: false, capid: capid, time: time };
}

/* ---------- config ---------- */
function getConfig_(req) {
  const p = props_();
  if (!p.getProperty('ADMIN_PIN')) {
    return { ok: true, firstRun: true, email: '', time: '20:00', stationKey: '' };
  }
  if (!auth_(req.pin)) return { ok: false, error: 'Wrong admin PIN' };
  return {
    ok: true,
    email: p.getProperty('EMAIL') || '',
    time: p.getProperty('SEND_TIME') || '20:00',
    stationKey: p.getProperty('STATION_KEY') || '',
    lastSent: p.getProperty('LAST_SENT') || ''
  };
}

function saveConfig_(req) {
  const p = props_();
  const hasPin = !!p.getProperty('ADMIN_PIN');
  if (hasPin) {
    if (!auth_(req.pin)) return { ok: false, error: 'Wrong admin PIN' };
  } else if (!req.newPin || String(req.newPin).length < 4) {
    return { ok: false, error: 'Choose an admin PIN of at least 4 characters' };
  }

  const emails = String(req.email || '').split(/[,;\s]+/).filter(Boolean);
  if (!emails.length) return { ok: false, error: 'Enter at least one email address' };
  for (const em of emails) {
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(em)) return { ok: false, error: 'Invalid email: ' + em };
  }
  if (!/^([01]\d|2[0-3]):[0-5]\d$/.test(String(req.time || ''))) return { ok: false, error: 'Invalid time' };
  if (!req.stationKey || String(req.stationKey).length < 4) {
    return { ok: false, error: 'Station code must be at least 4 characters' };
  }

  if (req.newPin) p.setProperty('ADMIN_PIN', String(req.newPin));
  p.setProperty('EMAIL', emails.join(','));
  p.setProperty('SEND_TIME', String(req.time));
  p.setProperty('STATION_KEY', String(req.stationKey));
  ensureTrigger_();
  return { ok: true };
}

function sendTest_(req) {
  if (!auth_(req.pin)) return { ok: false, error: 'Wrong admin PIN' };
  const email = props_().getProperty('EMAIL');
  if (!email) return { ok: false, error: 'Save an email address first' };
  const today = fmt_(new Date(), 'yyyy-MM-dd');
  const rows = todayRows_(today);
  sendReport_(email, today, rows, true);
  return { ok: true, count: rows.length };
}

/* ---------- scheduled report ---------- */
function tick() {
  const p = props_();
  const email = p.getProperty('EMAIL');
  if (!email) return;

  const now = new Date();
  const today = fmt_(now, 'yyyy-MM-dd');
  if (p.getProperty('LAST_SENT') === today) return;

  const parts = (p.getProperty('SEND_TIME') || '20:00').split(':').map(Number);
  const nowMin = Number(fmt_(now, 'H')) * 60 + Number(fmt_(now, 'm'));
  if (nowMin < parts[0] * 60 + parts[1]) return;

  p.setProperty('LAST_SENT', today);
  const rows = todayRows_(today);
  if (!rows.length) return;
  sendReport_(email, today, rows, false);
}

function sendReport_(to, date, rows, isTest) {
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const trs = rows.map(r => '<tr><td>' + esc(r[1]) + '</td><td>' + esc(r[2]) + '</td></tr>').join('');
  const html =
    '<p>' + (isTest ? '<b>[TEST]</b> ' : '') + '<b>' + rows.length + '</b> checked in on ' + esc(date) + '.</p>' +
    '<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse">' +
    '<tr style="background:#eee"><th>Time</th><th>CAPID</th></tr>' + trs + '</table>' +
    '<p style="color:#666;font-size:12px">CSV copy attached.</p>';

  const csv = 'Date,Time,CAPID\n' + rows.map(r => [r[0], r[1], r[2]].join(',')).join('\n');

  MailApp.sendEmail({
    to: to,
    subject: (isTest ? '[TEST] ' : '') + 'CAP Attendance ' + date + ' (' + rows.length + ' checked in)',
    htmlBody: html,
    body: rows.length + ' checked in on ' + date + '. See attached CSV.',
    attachments: [Utilities.newBlob(csv, 'text/csv', 'attendance-' + date + '.csv')],
    name: 'CAP Check-In'
  });
}

/* ---------- helpers ---------- */
function props_() { return PropertiesService.getScriptProperties(); }

function auth_(pin) {
  const stored = props_().getProperty('ADMIN_PIN');
  return !!stored && String(pin) === stored;
}

function fmt_(d, pattern) { return Utilities.formatDate(d, TZ, pattern); }

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function logSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Log');
  if (!sh) {
    sh = ss.insertSheet('Log', 0);
    sh.getRange('A:C').setNumberFormat('@');
    sh.appendRow(LOG_HEADERS);
  }
  sh.setFrozenRows(1);
  sh.getRange('A:C').setNumberFormat('@');
  if (sh.getIndex() !== 1) moveSheet_(sh, 1);
  return sh;
}



function todayRows_(today) {
  return logSheet_().getDataRange().getValues().slice(1)
    .filter(r => String(r[0]) === today)
    .sort((a, b) => String(a[1]).localeCompare(String(b[1])));
}

/* ---------- dated attendance sheets ---------- */
function displayDate_(date) {
  const d = new Date(date + 'T12:00:00Z');
  return Utilities.formatDate(d, TZ, 'M/d/yy');
}

// Google Sheets tab names do not accept slash characters, so use a tab-safe
// version such as 10-8-26 while the sheet itself displays 10/8/26 as the title.
function dateSheetName_(date) {
  const d = new Date(date + 'T12:00:00Z');
  return Utilities.formatDate(d, TZ, 'M-d-yy');
}

function dateSheet_(date) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = dateSheetName_(date);
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange('A:C').setNumberFormat('@');
    sh.getRange('A1:B1').merge().setValue(displayDate_(date));
    sh.getRange('A2:B2').setValues([DATE_HEADERS]);
    sh.getRange('A1:B1').setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center');
    sh.getRange('A2:B2').setFontWeight('bold');
    sh.setFrozenRows(2);

    // Put all date tabs after Log.
    moveSheet_(sh, Math.max(2, ss.getSheets().length));
    sh.autoResizeColumns(1, 2);
  }
  return sh;
}

function appendDateRow_(date, capid, time) {
  const sh = dateSheet_(date);
  const existing = sh.getDataRange().getValues();
  for (let i = 2; i < existing.length; i++) {
    if (String(existing[i][0]) === capid) return;
  }
  const row = sh.getLastRow() + 1;
  sh.getRange(row, 1, 1, 2).setNumberFormat('@').setValues([[capid, time]]);
  sh.autoResizeColumns(1, 2);
}

function syncDateSheet_(date, rows) {
  const sh = dateSheet_(date);
  const existing = sh.getDataRange().getValues();
  const have = {};
  for (let i = 2; i < existing.length; i++) {
    if (existing[i][0]) have[String(existing[i][0])] = true;
  }
  const toAdd = rows.filter(r => r[0] && !have[String(r[0])]);
  if (toAdd.length) {
    sh.getRange(sh.getLastRow() + 1, 1, toAdd.length, 2).setNumberFormat('@').setValues(toAdd);
  }
  sh.autoResizeColumns(1, 2);
}

function ensureTrigger_() {
  const exists = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'tick');
  if (!exists) ScriptApp.newTrigger('tick').timeBased().everyMinutes(10).create();
}
