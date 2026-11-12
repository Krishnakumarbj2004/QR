/**
 * Visitor QR Attendance System
 * Google Sheets + Google Apps Script
 * Native-phone-camera workflow: each QR code contains a unique web-app URL.
 */

const CFG = {
  TIME_ZONE: 'Asia/Kolkata',
  SHEET: {
    VISITORS: 'Visitors',
    STAFF: 'Staff',
    ATTENDANCE: 'Attendance',
    DASHBOARD: 'Dashboard'
  },
  SESSION_TTL_SECONDS: 21600, // CacheService maximum: 6 hours
  LOCK_WAIT_MS: 20000,
  QR_VERSION: 'v1'
};

const HEADERS = {
  Visitors: [
    'Visitor ID', 'Name', 'QR Token (PRIVATE)', 'Token Hash (PRIVATE)',
    'QR URL (PRIVATE)', 'Status', 'Checked In At', 'Checked In By', 'Active'
  ],
  Staff: ['Staff ID', 'Staff Name', 'Access Code Hash (PRIVATE)', 'Active'],
  Attendance: ['Timestamp', 'Visitor ID', 'Name', 'Result', 'Staff Name', 'Details'],
  Dashboard: ['Visitor Attendance Dashboard', '']
};

// -----------------------------------------------------------------------------
// Spreadsheet menu and setup
// -----------------------------------------------------------------------------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Visitor QR System')
    .addItem('1. Initialize system', 'initializeSystem')
    .addItem('2. Add staff member', 'addStaffMember')
    .addItem('3. Set / reset admin print code', 'setAdminPrintCode')
    .addSeparator()
    .addItem('Show scanner web-app URL', 'showWebAppUrl')
    .addItem('Generate missing visitor QR codes', 'generateMissingQrCodes')
    .addItem('Refresh QR links after redeployment', 'refreshQrLinks')
    .addItem('Print all visitor QR codes', 'showQrPrintPage')
    .addToUi();
}

function initializeSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  ss.setSpreadsheetTimeZone(CFG.TIME_ZONE);

  Object.keys(HEADERS).forEach(name => ensureSheet_(ss, name, HEADERS[name]));
  configureDashboard_(ss.getSheetByName(CFG.SHEET.DASHBOARD));

  const visitors = ss.getSheetByName(CFG.SHEET.VISITORS);
  visitors.setFrozenRows(1);
  visitors.getRange('G:G').setNumberFormat('dd-mmm-yyyy hh:mm:ss');
  visitors.setColumnWidths(1, 2, 150);
  visitors.setColumnWidth(3, 260);
  visitors.setColumnWidth(4, 260);
  visitors.setColumnWidth(5, 420);
  visitors.setColumnWidth(7, 180);

  const attendance = ss.getSheetByName(CFG.SHEET.ATTENDANCE);
  attendance.setFrozenRows(1);
  attendance.getRange('A:A').setNumberFormat('dd-mmm-yyyy hh:mm:ss');
  attendance.autoResizeColumns(1, 6);

  const staff = ss.getSheetByName(CFG.SHEET.STAFF);
  staff.setFrozenRows(1);
  staff.autoResizeColumns(1, 4);

  SpreadsheetApp.getUi().alert(
    'System initialized.\n\n' +
    'Next: deploy the Apps Script as a web app, add staff, set the admin print code, ' +
    'paste your visitors into Visitors columns A:B, then generate QR codes.'
  );
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);

  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const isBlank = current.every(v => v === '' || v === null);
  if (isBlank) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#dbeafe');
  } else if (headers.some((h, i) => String(current[i]).trim() !== h)) {
    throw new Error(
      'The existing ' + name + ' sheet has a different layout. ' +
      'Use a new blank spreadsheet for this version, or rename the old sheet and run Initialize again.'
    );
  }
  sheet.setFrozenRows(1);
  return sheet;
}

function configureDashboard_(sheet) {
  sheet.clearContents();
  sheet.getRange(1, 1, 1, 2).setValues([['Visitor Attendance Dashboard', '']]);
  sheet.getRange('A1').setFontWeight('bold').setFontSize(16).setBackground('#dbeafe');
  sheet.getRange(3, 1, 7, 1).setValues([
    ['Total active visitors'],
    ['Checked in'],
    ['Pending'],
    ['Inactive visitors'],
    ['Successful entries'],
    ['Duplicate scans'],
    ['Invalid scans']
  ]);
  const formulas = [
    ['=COUNTIF(Visitors!I2:I,TRUE)'],
    ['=COUNTIF(Visitors!F2:F,"CHECKED IN")'],
    ['=COUNTIFS(Visitors!I2:I,TRUE,Visitors!F2:F,"<>CHECKED IN")'],
    ['=COUNTIF(Visitors!I2:I,FALSE)'],
    ['=COUNTIF(Attendance!D2:D,"VALID")'],
    ['=COUNTIF(Attendance!D2:D,"DUPLICATE")'],
    ['=COUNTIF(Attendance!D2:D,"INVALID")']
  ];
  sheet.getRange(3, 2, formulas.length, 1).setFormulas(formulas);
  sheet.getRange('A3:A9').setFontWeight('bold');
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 150);
}

function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('Run Visitor QR System → Initialize system first.');
  return SpreadsheetApp.openById(id);
}

// -----------------------------------------------------------------------------
// Utility functions
// -----------------------------------------------------------------------------

function sha256_(value) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(value),
    Utilities.Charset.UTF_8
  );
  return bytes.map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}

function randomHexToken_() {
  return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '').toLowerCase();
}

function randomStaffCode_() {
  // A UUIDv4 without hyphens provides a high-entropy 32-character staff code.
  return Utilities.getUuid().replace(/-/g, '').toLowerCase();
}

function isTrue_(value) {
  return value === true || String(value).toUpperCase() === 'TRUE';
}

function getWebAppUrl_() {
  const url = ScriptApp.getService().getUrl();
  if (!url) {
    throw new Error('Deploy this project as a Web app first (Deploy → New deployment → Web app).');
  }
  return url;
}

function showWebAppUrl() {
  SpreadsheetApp.getUi().alert('Scanner web-app URL:\n\n' + getWebAppUrl_());
}

// -----------------------------------------------------------------------------
// Staff management and staff sessions
// -----------------------------------------------------------------------------

function addStaffMember() {
  const ui = SpreadsheetApp.getUi();
  const nameResult = ui.prompt('Add staff member', 'Enter the staff member name:', ui.ButtonSet.OK_CANCEL);
  if (nameResult.getSelectedButton() !== ui.Button.OK) return;
  const name = nameResult.getResponseText().trim();
  if (!name) return ui.alert('Staff name cannot be empty.');

  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(CFG.SHEET.STAFF);
  const values = sheet.getDataRange().getValues();
  const staffId = 'STF' + String(Math.max(0, values.length - 1) + 1).padStart(3, '0');
  const accessCode = randomStaffCode_(); // shown once; only its hash is stored
  sheet.appendRow([staffId, name, sha256_(accessCode), true]);

  ui.alert(
    'Staff member created. Share this access code with ' + name + ':\n\n' +
    accessCode + '\n\n' +
    'This code is shown only once. Store it securely. To revoke access, set Active to FALSE in the Staff sheet.'
  );
}

function loginStaff(accessCode) {
  if (typeof accessCode !== 'string' || !/^[a-f0-9]{32}$/i.test(accessCode)) {
    return { ok: false, message: 'Access code is not valid.' };
  }

  const sheet = getSpreadsheet_().getSheetByName(CFG.SHEET.STAFF);
  const rows = sheet.getDataRange().getValues();
  const codeHash = sha256_(accessCode.toLowerCase());
  let staff = null;

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][2]) === codeHash && isTrue_(rows[i][3])) {
      staff = { staffId: String(rows[i][0]), staffName: String(rows[i][1]) };
      break;
    }
  }
  if (!staff) return { ok: false, message: 'Incorrect or inactive staff access code.' };

  const sessionToken = randomHexToken_();
  const cacheKey = 'STAFF_SESSION_' + sha256_(sessionToken);
  CacheService.getScriptCache().put(cacheKey, JSON.stringify(staff), CFG.SESSION_TTL_SECONDS);

  return { ok: true, sessionToken: sessionToken, staffName: staff.staffName };
}

function getStaffSession_(sessionToken) {
  if (typeof sessionToken !== 'string' || !/^[a-f0-9]{64}$/i.test(sessionToken)) return null;
  const cacheKey = 'STAFF_SESSION_' + sha256_(sessionToken.toLowerCase());
  const raw = CacheService.getScriptCache().get(cacheKey);
  if (!raw) return null;
  try {
    const session = JSON.parse(raw);
    // Re-check the staff record on every scan so setting Active=FALSE
    // revokes an existing session, not only future logins.
    const rows = getSpreadsheet_().getSheetByName(CFG.SHEET.STAFF).getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(session.staffId) && isTrue_(rows[i][3])) return session;
    }
    CacheService.getScriptCache().remove(cacheKey);
    return null;
  } catch (e) {
    return null;
  }
}

function logoutStaff(sessionToken) {
  if (typeof sessionToken === 'string' && /^[a-f0-9]{64}$/i.test(sessionToken)) {
    CacheService.getScriptCache().remove('STAFF_SESSION_' + sha256_(sessionToken.toLowerCase()));
  }
  return { ok: true };
}

// -----------------------------------------------------------------------------
// Admin print-code protection
// -----------------------------------------------------------------------------

function setAdminPrintCode() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Set administrator print code',
    'Enter a strong password with at least 16 characters. It protects access to the private QR token list.',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const code = response.getResponseText();
  if (typeof code !== 'string' || code.length < 16) return ui.alert('Use at least 16 characters.');

  PropertiesService.getScriptProperties().setProperty('ADMIN_PRINT_HASH', sha256_(code));
  ui.alert('Administrator print code saved. It is not stored in plain text.');
}

function verifyAdminPrintCode_(code) {
  const savedHash = PropertiesService.getScriptProperties().getProperty('ADMIN_PRINT_HASH');
  return Boolean(savedHash && typeof code === 'string' && sha256_(code) === savedHash);
}

// -----------------------------------------------------------------------------
// Visitor QR code generation and printing
// -----------------------------------------------------------------------------

function generateMissingQrCodes() {
  const sheet = getSpreadsheet_().getSheetByName(CFG.SHEET.VISITORS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('Paste visitor IDs in column A and names in column B first.');
    return;
  }

  const baseUrl = getWebAppUrl_();
  const rows = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  const ids = rows.filter(r => r[0] !== '').map(r => String(r[0]).trim());
  if (new Set(ids).size !== ids.length) {
    throw new Error('Duplicate Visitor IDs found. Each Visitor ID must be unique.');
  }

  let generated = 0;
  rows.forEach(row => {
    const visitorId = String(row[0] || '').trim();
    const name = String(row[1] || '').trim();
    if (!visitorId && !name) return;
    if (!visitorId || !name) throw new Error('Every visitor row must contain both Visitor ID and Name.');

    if (!row[2] || !row[3]) {
      const token = randomHexToken_();
      row[2] = token;
      row[3] = sha256_(token);
      generated++;
    }
    row[4] = baseUrl + '?t=' + encodeURIComponent(String(row[2]));
    if (!row[5]) row[5] = 'PENDING';
    if (row[8] === '' || row[8] === null) row[8] = true;
  });

  sheet.getRange(2, 1, rows.length, 9).setValues(rows);
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    generated + ' new QR token(s) generated.\n' +
    'QR links are ready. Use Visitor QR System → Print all visitor QR codes.'
  );
}

function refreshQrLinks() {
  const baseUrl = getWebAppUrl_();
  const sheet = getSpreadsheet_().getSheetByName(CFG.SHEET.VISITORS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return SpreadsheetApp.getUi().alert('No visitor records found.');

  const rows = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  let updated = 0;
  rows.forEach(row => {
    if (row[0] && row[2]) {
      row[4] = baseUrl + '?t=' + encodeURIComponent(String(row[2]));
      updated++;
    }
  });
  sheet.getRange(2, 1, rows.length, 9).setValues(rows);
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(updated + ' QR link(s) refreshed. Reprint cards if the deployment URL changed.');
}

function showQrPrintPage() {
  const html = HtmlService.createHtmlOutputFromFile('Print')
    .setWidth(1100)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Print visitor QR cards');
}

function getQrPrintData(adminCode) {
  if (!verifyAdminPrintCode_(adminCode)) throw new Error('Incorrect administrator print code.');
  const sheet = getSpreadsheet_().getSheetByName(CFG.SHEET.VISITORS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  return sheet.getRange(2, 1, lastRow - 1, 9).getValues()
    .filter(row => row[0] && row[1] && row[4] && isTrue_(row[8]))
    .map(row => ({ visitorId: String(row[0]), name: String(row[1]), qrUrl: String(row[4]) }));
}

// -----------------------------------------------------------------------------
// Web app and visitor check-in
// -----------------------------------------------------------------------------

function doGet(e) {
  // Laptop scanner is hosted outside Apps Script so the browser can access
  // the webcam. It calls these JSONP endpoints because Apps Script web apps
  // do not provide a configurable CORS response for a normal cross-origin fetch.
  if (e && e.parameter && e.parameter.api) return handleJsonpApi_(e);

  const rawToken = e && e.parameter ? String(e.parameter.t || '') : '';
  const token = /^[a-f0-9]{64}$/i.test(rawToken) ? rawToken.toLowerCase() : '';
  const template = HtmlService.createTemplateFromFile('Scanner');
  template.qrToken = token ? ('v1:' + token) : '';
  template.appUrl = ScriptApp.getService().getUrl() || '';
  return template.evaluate()
    .setTitle('Visitor QR Check-in')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** JSONP bridge for the separately hosted laptop webcam scanner. */
function handleJsonpApi_(e) {
  const p = e.parameter || {};
  const callback = String(p.callback || '');
  // Only allow a plain JavaScript identifier; never echo arbitrary code.
  if (!/^[A-Za-z_$][A-Za-z0-9_$]{0,63}$/.test(callback)) {
    return ContentService.createTextOutput('/* Invalid callback */')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  let result;
  try {
    switch (String(p.api)) {
      case 'login':
        result = loginStaff(String(p.code || ''));
        break;
      case 'checkin':
        result = checkInVisitor('v1:' + String(p.t || ''), String(p.s || ''));
        break;
      case 'logout':
        result = logoutStaff(String(p.s || ''));
        break;
      case 'ping':
        result = { ok: true, service: 'Visitor QR Attendance API' };
        break;
      default:
        result = { ok: false, status: 'ERROR', message: 'Unknown API action.' };
    }
  } catch (err) {
    console.error(err);
    result = { ok: false, status: 'ERROR', message: 'Server error. Please retry.' };
  }

  const json = JSON.stringify(result)
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e')
    .replace(/&/g, '\\u0026');
  return ContentService.createTextOutput(callback + '(' + json + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function checkInVisitor(qrToken, sessionToken) {
  const staff = getStaffSession_(sessionToken);
  if (!staff) return { status: 'UNAUTHORIZED', message: 'Staff session expired. Please sign in again.' };

  if (typeof qrToken !== 'string') {
    return { status: 'INVALID', message: 'QR code is invalid.' };
  }
  const match = qrToken.match(/^v1:([a-f0-9]{64})$/i);
  if (!match) {
    // The actual QR URL passes its token as the t= query parameter; the UI
    // sends that 64-character token with this version prefix.
    return { status: 'INVALID', message: 'QR code format is not recognized.' };
  }
  const tokenHash = sha256_(match[1].toLowerCase());

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(CFG.LOCK_WAIT_MS);
  } catch (err) {
    return { status: 'BUSY', message: 'The check-in system is busy. Please retry in a moment.' };
  }

  try {
    const ss = getSpreadsheet_();
    const visitors = ss.getSheetByName(CFG.SHEET.VISITORS);
    const attendance = ss.getSheetByName(CFG.SHEET.ATTENDANCE);
    const lastRow = visitors.getLastRow();
    const rows = lastRow >= 2 ? visitors.getRange(2, 1, lastRow - 1, 9).getValues() : [];
    let visitor = null;
    let visitorSheetRow = -1;

    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][3]) === tokenHash) {
        visitor = rows[i];
        visitorSheetRow = i + 2;
        break;
      }
    }

    const now = new Date();
    if (!visitor) {
      attendance.appendRow([now, '', '', 'INVALID', staff.staffName, 'Unregistered QR token']);
      SpreadsheetApp.flush();
      return { status: 'INVALID', message: 'This QR code does not belong to a registered visitor.' };
    }

    const visitorId = String(visitor[0]);
    const visitorName = String(visitor[1]);
    if (!isTrue_(visitor[8])) {
      attendance.appendRow([now, visitorId, visitorName, 'INVALID', staff.staffName, 'Visitor is inactive']);
      SpreadsheetApp.flush();
      return { status: 'INVALID', message: 'This visitor registration is inactive.', visitor: { visitorId, name: visitorName } };
    }

    const alreadyCheckedIn = String(visitor[5]).toUpperCase() === 'CHECKED IN' || visitor[6] instanceof Date;
    if (alreadyCheckedIn) {
      attendance.appendRow([now, visitorId, visitorName, 'DUPLICATE', staff.staffName, 'Already checked in']);
      SpreadsheetApp.flush();
      return {
        status: 'ALREADY_CHECKED_IN',
        message: 'This visitor has already checked in.',
        visitor: { visitorId: visitorId, name: visitorName },
        checkedInAt: visitor[6] instanceof Date ? visitor[6].toISOString() : ''
      };
    }

    // Update the visitor record and append the successful scan while holding
    // the same script-wide lock, so another staff phone cannot double-check-in.
    visitors.getRange(visitorSheetRow, 6, 1, 3).setValues([['CHECKED IN', now, staff.staffName]]);
    attendance.appendRow([now, visitorId, visitorName, 'VALID', staff.staffName, 'First successful check-in']);
    SpreadsheetApp.flush();

    return {
      status: 'VALID',
      message: 'Visitor verified. Entry recorded successfully.',
      visitor: { visitorId: visitorId, name: visitorName },
      checkedInAt: now.toISOString(),
      staffName: staff.staffName
    };
  } catch (err) {
    console.error(err);
    return { status: 'ERROR', message: 'Could not save the check-in. Please retry or contact the administrator.' };
  } finally {
    lock.releaseLock();
  }
}
