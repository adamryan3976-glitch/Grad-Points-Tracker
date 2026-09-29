/**
 * WPS Grad Point Tracker — data service (Google Apps Script)
 *
 * The app itself is hosted on GitHub Pages. This script, attached to the
 * tracker Google Sheet, receives requests from that page and reads/writes
 * the Sheet. Every request must include the staff passcode.
 *
 * Setup: paste into Extensions > Apps Script, run setup() once, set the
 * passcode from the "Grad Points" menu, then Deploy > Web app
 * (Execute as: Me, Who has access: Anyone). See README.md.
 */

const SHEETS = {
  STUDENTS: 'Students',
  CATEGORIES: 'Categories',
  ACTIVITIES: 'Activities',
  POINTS: 'Points',
  SUMMARY: 'Summary'
};

const HEADERS = {
  Students: ['Student ID', 'First Name', 'Last Name', 'Grade', 'Class', 'Active'],
  Categories: ['Category', 'Points Required'],
  Activities: ['Activity', 'Category', 'Points'],
  Points: ['Entry ID', 'Timestamp', 'Student ID', 'Student Name', 'Class',
           'Category', 'Activity', 'Points', 'Note', 'Entered By', 'Status']
};

const DEFAULT_CATEGORIES = [
  ['Arts', 4],
  ['Academics', 5],
  ['Athletics', 8],
  ['Service', 8]
];

// Suggestions only — teachers can type any activity.
const DEFAULT_ACTIVITIES = [
  ['Musical', 'Arts', 1],
  ['Choir', 'Arts', 1],
  ['Band', 'Arts', 1],
  ['Dance Crew', 'Arts', 1],
  ['Folk Dance', 'Arts', 1],
  ['Forest of Reading', 'Academics', 1],
  ['Tutoring - Term 1', 'Academics', 1],
  ['Tutoring - Term 2', 'Academics', 1],
  ['Honour Roll - Term 1', 'Academics', 1],
  ['Honour Roll - Term 2', 'Academics', 1],
  ['Ontario Skills Competition', 'Academics', 1],
  ['Folk Dance', 'Athletics', 1],
  ['School Team', 'Athletics', 1],
  ['Intramurals', 'Athletics', 1],
  ['Bus Kids', 'Service', 1],
  ['Welcome Committee', 'Service', 1],
  ['Gardening Club', 'Service', 1],
  ['Library Helpers', 'Service', 1],
  ['Score Keepers', 'Service', 1],
  ['Lunchroom Supervisors', 'Service', 1],
  ['Kindie Helpers', 'Service', 1],
  ['Teacher Helpers (frequent)', 'Service', 1],
  ['Office Helpers', 'Service', 1],
  ['Special Events Helpers', 'Service', 1]
];

// Only students in these grades appear in the app and can receive points.
const ELIGIBLE_GRADES = [7, 8];

const MAX_FAILED_LOGINS = 25;       // across all users
const LOCKOUT_SECONDS = 600;        // 10 minutes

/* ------------------------------------------------------------------ */
/* HTTP entry points (called by the GitHub Pages site)                 */
/* ------------------------------------------------------------------ */

const API = {
  getConfig: function () { return getConfig(); },
  getStudents: function () { return getStudents(); },
  addPoints: function (a) { return addPoints(a); },
  removeEntry: function (a) { return removeEntry(a.entryId, a.enteredBy); }
};

function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    checkPasscode_(req.passcode);
    const fn = API[req.action];
    if (!fn) throw new Error('Unknown request: ' + req.action);
    out = { ok: true, data: fn(req.args || {}) };
  } catch (err) {
    out = { ok: false, error: (err && err.message) || String(err), code: err && err.code };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, app: 'WPS Grad Point Tracker data service' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function checkPasscode_(code) {
  const expected = PropertiesService.getScriptProperties().getProperty('PASSCODE');
  if (!expected) throw authError_('The staff passcode has not been set yet. In the Sheet choose Grad Points > Set staff passcode.');
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('failedLogins') || 0);
  if (fails >= MAX_FAILED_LOGINS) throw authError_('Too many wrong passcodes. Please try again in 10 minutes.');
  if (str_(code) !== expected) {
    cache.put('failedLogins', String(fails + 1), LOCKOUT_SECONDS);
    throw authError_('That passcode is not correct.');
  }
}

function authError_(msg) {
  const e = new Error(msg);
  e.code = 'AUTH';
  return e;
}

/* ------------------------------------------------------------------ */
/* Sheet menu                                                          */
/* ------------------------------------------------------------------ */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Grad Points')
    .addItem('Set staff passcode', 'setPasscode')
    .addItem('Refresh Summary tab', 'refreshSummary')
    .addSeparator()
    .addItem('Set up / repair sheets (adds missing columns)', 'setup')
    .addToUi();
}

function setPasscode() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('Staff passcode',
    'Teachers enter this to use the app. Use at least 6 characters.\nChanging it signs everyone out.',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const code = str_(res.getResponseText());
  if (code.length < 6) { ui.alert('Please use at least 6 characters.'); return; }
  PropertiesService.getScriptProperties().setProperty('PASSCODE', code);
  CacheService.getScriptCache().remove('failedLogins');
  ui.alert('Passcode saved.');
}

/* ------------------------------------------------------------------ */
/* One-time setup                                                      */
/* ------------------------------------------------------------------ */

function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(HEADERS).forEach(function (name) {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    const h = HEADERS[name];
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, h.length).setValues([h]);
      styleHeader_(sh, h.length);
    } else {
      // Repair: add any missing columns (e.g. Grade) to the end of the header row.
      const have = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(str_);
      const missing = h.filter(function (x) { return have.indexOf(x) === -1; });
      if (missing.length) {
        sh.getRange(1, have.length + 1, 1, missing.length).setValues([missing]);
        styleHeader_(sh, have.length + missing.length);
      }
    }
  });

  ss.getSheetByName(SHEETS.STUDENTS).getRange('A:A').setNumberFormat('@');
  ss.getSheetByName(SHEETS.POINTS).getRange('B:B').setNumberFormat('yyyy-mm-dd hh:mm');
  ss.getSheetByName(SHEETS.POINTS).getRange('C:C').setNumberFormat('@');

  seedIfEmpty_(SHEETS.CATEGORIES, DEFAULT_CATEGORIES);
  seedIfEmpty_(SHEETS.ACTIVITIES, DEFAULT_ACTIVITIES);
  seedIfEmpty_(SHEETS.STUDENTS, [
    ['', 'Alex', 'Example', 8, 'Room 21', 'Yes'],
    ['', 'Jordan', 'Sample', 7, 'Room 18', 'Yes']
  ]);

  const blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(blank);

  ensureStudentIds_();
  refreshSummary();
}

function seedIfEmpty_(name, rows) {
  const sh = sheet_(name);
  if (sh.getLastRow() > 1 || !rows.length) return;
  sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function styleHeader_(sh, width) {
  sh.getRange(1, 1, 1, width)
    .setFontWeight('bold')
    .setBackground('#1d3557')
    .setFontColor('#ffffff');
  sh.setFrozenRows(1);
}

/* ------------------------------------------------------------------ */
/* Data functions                                                      */
/* ------------------------------------------------------------------ */

function getConfig() {
  const categories = readTable_(SHEETS.CATEGORIES).map(function (r) {
    return { name: str_(r['Category']), required: Number(r['Points Required']) || 0 };
  }).filter(function (c) { return c.name; });

  const activities = readTable_(SHEETS.ACTIVITIES).map(function (r) {
    return { name: str_(r['Activity']), category: str_(r['Category']), points: Number(r['Points']) || 1 };
  }).filter(function (a) { return a.name && a.category; });

  return { categories: categories, activities: activities, grades: ELIGIBLE_GRADES };
}

/** All active Grade 7-8 students (sorted by last name) and their point entries. */
function getStudents() {
  ensureStudentIds_();
  const students = readTable_(SHEETS.STUDENTS)
    .filter(function (s) { return isEligible_(s) && str_(s['Student ID']); })
    .map(function (s) {
      return {
        id: str_(s['Student ID']),
        first: str_(s['First Name']),
        last: str_(s['Last Name']),
        grade: gradeOf_(s['Grade']),
        cls: str_(s['Class'])
      };
    })
    .sort(function (a, b) {
      return a.last.localeCompare(b.last, 'en', { sensitivity: 'base' }) ||
             a.first.localeCompare(b.first, 'en', { sensitivity: 'base' });
    });

  const ids = {};
  students.forEach(function (s) { ids[s.id] = true; });

  const entries = readTable_(SHEETS.POINTS)
    .filter(function (p) { return ids[str_(p['Student ID'])] && !isRemoved_(p['Status']); })
    .map(function (p) {
      const t = p['Timestamp'];
      return {
        entryId: str_(p['Entry ID']),
        ts: t instanceof Date ? t.getTime() : (Date.parse(t) || 0),
        studentId: str_(p['Student ID']),
        category: str_(p['Category']),
        activity: str_(p['Activity']),
        points: Number(p['Points']) || 0,
        note: str_(p['Note']),
        by: str_(p['Entered By'])
      };
    });

  return { students: students, entries: entries };
}

function addPoints(payload) {
  payload = payload || {};
  const ids = unique_((payload.studentIds || []).map(str_).filter(Boolean));
  const category = str_(payload.category);
  const activity = str_(payload.activity).slice(0, 100);
  const note = str_(payload.note).slice(0, 500);
  const points = Number(payload.points);
  const by = str_(payload.enteredBy).slice(0, 60);

  if (!ids.length) throw new Error('Select at least one student.');
  const cats = readTable_(SHEETS.CATEGORIES).map(function (r) { return str_(r['Category']); });
  if (cats.indexOf(category) === -1) throw new Error('Unknown category: ' + category);
  if (!activity) throw new Error('Enter the activity the points are for.');
  if (!(points > 0 && points <= 20)) throw new Error('Points must be more than 0 and no more than 20.');
  if (!by) throw new Error('Your name is missing. Sign out and sign back in.');

  const studentMap = {};
  readTable_(SHEETS.STUDENTS).forEach(function (s) { studentMap[str_(s['Student ID'])] = s; });

  const now = new Date();
  const rows = ids.map(function (id) {
    const s = studentMap[id];
    if (!s) throw new Error('Student not found: ' + id);
    if (!isEligible_(s)) throw new Error((str_(s['First Name']) + ' ' + str_(s['Last Name'])).trim() +
      ' is not an active Grade 7 or 8 student.');
    return {
      'Entry ID': newId_(),
      'Timestamp': now,
      'Student ID': id,
      'Student Name': (str_(s['First Name']) + ' ' + str_(s['Last Name'])).trim(),
      'Class': str_(s['Class']),
      'Category': category,
      'Activity': activity,
      'Points': points,
      'Note': note,
      'Entered By': by,
      'Status': 'Active'
    };
  });

  withLock_(function () { appendObjects_(SHEETS.POINTS, rows); });
  safeRefreshSummary_();
  return getStudents();
}

function removeEntry(entryId, enteredBy) {
  entryId = str_(entryId);
  const who = str_(enteredBy).slice(0, 60) || 'teacher';
  withLock_(function () {
    const sh = sheet_(SHEETS.POINTS);
    const data = sh.getDataRange().getValues();
    const h = data[0].map(str_);
    const idCol = h.indexOf('Entry ID');
    const stCol = h.indexOf('Status');
    if (idCol < 0 || stCol < 0) throw new Error('The Points tab is missing its "Entry ID" or "Status" column.');
    for (let i = 1; i < data.length; i++) {
      if (str_(data[i][idCol]) === entryId) {
        const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
        sh.getRange(i + 1, stCol + 1).setValue('Removed by ' + who + ' ' + stamp);
        return;
      }
    }
    throw new Error('That entry was not found. It may already have been removed.');
  });
  safeRefreshSummary_();
  return getStudents();
}

/* ------------------------------------------------------------------ */
/* Summary tab (whole-school view for record keeping)                  */
/* ------------------------------------------------------------------ */

function refreshSummary() {
  const ss = SpreadsheetApp.getActive();
  const cats = readTable_(SHEETS.CATEGORIES).map(function (r) {
    return { name: str_(r['Category']), required: Number(r['Points Required']) || 0 };
  }).filter(function (c) { return c.name; });

  const totals = {};
  readTable_(SHEETS.POINTS).forEach(function (p) {
    if (isRemoved_(p['Status'])) return;
    const id = str_(p['Student ID']);
    const c = str_(p['Category']);
    totals[id] = totals[id] || {};
    totals[id][c] = (totals[id][c] || 0) + (Number(p['Points']) || 0);
  });

  const header = ['Student ID', 'Last Name', 'First Name', 'Grade', 'Class']
    .concat(cats.map(function (c) { return c.name + ' (' + c.required + ' needed)'; }))
    .concat(['Total Points', 'Grad Letter']);

  const rows = readTable_(SHEETS.STUDENTS)
    .filter(function (s) { return str_(s['Student ID']) && isEligible_(s); })
    .map(function (s) {
      const id = str_(s['Student ID']);
      const t = totals[id] || {};
      const vals = cats.map(function (c) { return t[c.name] || 0; });
      const total = vals.reduce(function (a, b) { return a + b; }, 0);
      const earned = cats.length > 0 && cats.every(function (c) { return (t[c.name] || 0) >= c.required; });
      return [id, str_(s['Last Name']), str_(s['First Name']), gradeOf_(s['Grade']), str_(s['Class'])]
        .concat(vals).concat([total, earned ? 'EARNED' : 'In progress']);
    })
    .sort(function (a, b) {
      return String(a[1]).localeCompare(String(b[1]), 'en', { sensitivity: 'base' }) ||
             String(a[2]).localeCompare(String(b[2]), 'en', { sensitivity: 'base' });
    });

  const sh = ss.getSheetByName(SHEETS.SUMMARY) || ss.insertSheet(SHEETS.SUMMARY);
  sh.clear();
  sh.getRange(1, 1, 1, header.length).setValues([header]);
  styleHeader_(sh, header.length);
  if (rows.length) {
    sh.getRange(2, 1, rows.length, 1).setNumberFormat('@');
    sh.getRange(2, 1, rows.length, header.length).setValues(rows);
  }
  sh.getRange(1, header.length + 2).setValue('Auto-generated — edits here are overwritten.')
    .setFontColor('#888888').setFontStyle('italic');
}

function safeRefreshSummary_() {
  try { refreshSummary(); } catch (e) { console.error('Summary refresh failed: ' + e); }
}

/* ------------------------------------------------------------------ */
/* Helpers                                                             */
/* ------------------------------------------------------------------ */

function sheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('The "' + name + '" tab is missing. In the Sheet, choose Grad Points > Set up / repair sheets.');
  return sh;
}

function readTable_(name) {
  const values = sheet_(name).getDataRange().getValues();
  const headers = (values.shift() || []).map(str_);
  return values
    .filter(function (r) { return r.some(function (c) { return str_(c) !== ''; }); })
    .map(function (r) {
      const o = {};
      headers.forEach(function (h, j) { if (h) o[h] = r[j]; });
      return o;
    });
}

function appendObjects_(name, objs) {
  if (!objs.length) return;
  const sh = sheet_(name);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(str_);
  const rows = objs.map(function (o) {
    return headers.map(function (h) { return Object.prototype.hasOwnProperty.call(o, h) ? o[h] : ''; });
  });
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
}

/** Gives any student row without an ID a new one (S0001, S0002, ...). */
function ensureStudentIds_() {
  const sh = sheet_(SHEETS.STUDENTS);
  if (sh.getLastRow() < 2) return;

  const check = sh.getDataRange().getValues();
  const h = check[0].map(str_);
  const idC = h.indexOf('Student ID'), fC = h.indexOf('First Name'), lC = h.indexOf('Last Name');
  if (idC < 0) return;
  const needsId = function (r) { return !str_(r[idC]) && (str_(r[fC]) || str_(r[lC])); };
  if (!check.slice(1).some(needsId)) return;

  withLock_(function () {
    const data = sh.getDataRange().getValues();
    let max = 0;
    data.slice(1).forEach(function (r) {
      const m = str_(r[idC]).match(/^S(\d+)$/);
      if (m) max = Math.max(max, Number(m[1]));
    });
    const col = data.slice(1).map(function (r) {
      if (needsId(r)) { max++; return ['S' + ('000' + max).slice(-4)]; }
      return [r[idC]];
    });
    sh.getRange(2, idC + 1, col.length, 1).setNumberFormat('@').setValues(col);
  });
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function newId_() { return Utilities.getUuid().split('-')[0].toUpperCase(); }
function str_(v) { return String(v == null ? '' : v).trim(); }
function unique_(arr) { return arr.filter(function (v, i) { return arr.indexOf(v) === i; }); }
function isRemoved_(v) { return str_(v).toLowerCase().indexOf('removed') === 0; }
/** Reads a grade like 8, "8", "Gr 8" or "Grade 8" as a number (blank/unknown = 0). */
function gradeOf_(v) {
  const m = str_(v).match(/\d+/);
  return m ? Number(m[0]) : 0;
}
function isEligible_(s) {
  return isActive_(s['Active']) && ELIGIBLE_GRADES.indexOf(gradeOf_(s['Grade'])) !== -1;
}
function isActive_(v) {
  const a = str_(v).toLowerCase();
  return !(a === 'no' || a === 'n' || a === 'false' || a === 'inactive');
}
