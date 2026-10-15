/**
 * MundyBall website forms -> Google Sheets
 *
 * Paste this whole file into the spreadsheet: Extensions -> Apps Script, replacing anything there.
 * Then: Deploy -> New deployment -> type "Web app" -> Execute as: Me -> Who has access: Anyone -> Deploy.
 * Copy the Web app URL (ends in /exec) and send it to Claude to connect the website.
 *
 * Each website form posts here with a "_form" value. The row goes to the tab named in TABS below.
 * Columns are matched by header text (case/punctuation-insensitive, with a few synonyms). A field with no
 * matching column gets a new column at the end, so nothing is dropped. Missing tabs are created.
 */

// ===== Settings: change the tab names on the right to match your spreadsheet =====
var TABS = {
  tryouts:  '2026 Tryout Registration',
  camp:     '2027 Camp Interest Form',
  contact:  'MundyBall Contact Us Page',
  coaching: 'Coaching Applicants'
};

// Form field -> exact column name, for tabs with their own established headers.
// Fields not listed here fall back to header matching with SYNONYMS below.
var FIELD_MAP = {
  tryouts: {
    'Camper First Name':              'First Name',
    'Camper Last Name':               'Last Name',
    'Camper Gender':                  'Gender',
    'Camper Grade':                   'Grade',
    'Camper Birthday':                'Birthday',
    'Camper School':                  'School',
    'Camper Experience':              'Experience',
    'Parent Name':                    'Parent/Guardian Name(s)',
    'Parent Phone Number':            'Parent Phone(s)',
    'Email':                          'Parent Email(s)',
    'Emergency Contact Name':         'Emergency Contact Name(s)',
    'Emergency Contact Phone Number': 'Emergency Phone(s)',
    'Allergies/Medical Notes':        'Allergies / Medical Notes',
    'Jersey Size':                    'Jersey Size',
    'Shorts Size':                    'Shorts Size',
    'T-Shirt Size':                   'Shirt Size'
  }
};

// Must match the token in the website code. Not a password; it just filters out random junk posts.
var TOKEN = 'mundyball-forms-v1';

// Header synonyms: a form field on the left can fill any column named on the right.
var SYNONYMS = {
  'camper': ['player', 'athlete', 'child', 'kid'],
  'birthday': ['date of birth', 'dob', 'birthdate'],
  'date of birth': ['birthday', 'dob', 'birthdate'],
  'gender': ['boys/girls', 'boys girls', 'sex'],
  'boys/girls': ['gender'],
  'phone number': ['phone', 'cell', 'mobile'],
  'phone': ['phone number', 'cell', 'mobile'],
  'email': ['email address', 'e-mail', 'parent email'],
  'experience': ['years of organized basketball', 'years experience'],
  'allergies/medical notes': ['medical notes', 'allergies', 'medical notes / allergies', 'allergies / medical notes'],
  'allergies / medical notes': ['medical notes', 'allergies', 'medical notes / allergies', 'allergies/medical notes']
};

var SKIP = { access_key: 1, botcheck: 1, subject: 1, from_name: 1, _form: 1, _token: 1, 'form-name': 1 };

// Light protection, since the token is visible in the page source:
// - each form must include these fields (a row without them is ignored)
// - every cell is capped at MAX_CELL characters, and a post with more than MAX_FIELDS fields is ignored
// - text starting with = + - @ is stored as plain text, never run as a spreadsheet formula
var REQUIRED = {
  tryouts:  ['Camper First Name', 'Camper Last Name', 'Email'],
  camp:     ['Camper First Name', 'Email'],
  contact:  ['Email', 'Message'],
  coaching: ['Applicant Name', 'Applicant Email']
};
var MAX_CELL = 2000;
var MAX_FIELDS = 60;

function doPost(e) {
  var p = (e && e.parameters) || {};
  var form = first_(p._form);
  if (first_(p._token) !== TOKEN || !TABS[form]) return reply_('ignored');
  if (first_(p.botcheck)) return reply_('ignored'); // honeypot filled in = bot
  if (Object.keys(p).length > MAX_FIELDS) return reply_('ignored');
  var missing = REQUIRED[form].some(function (k) { return !String(first_(p[k])).trim(); });
  if (missing) return reply_('ignored');

  // Build field -> value. Checkbox groups like "Grade: 4th" (value "on") become one "Grade" cell.
  var values = {}, order = [];
  Object.keys(p).forEach(function (key) {
    if (SKIP[key]) return;
    var vals = p[key].filter(function (v) { return String(v).trim() !== ''; });
    var m = key.match(/^([^:]+):\s*(.+)$/);
    if (m && vals.length && vals.every(function (v) { return /^(on|true|yes)$/i.test(v) || v === m[2]; })) {
      var group = m[1].trim();
      if (!(group in values)) { values[group] = []; order.push(group); }
      values[group].push(m[2].trim());
      return;
    }
    if (!(key in values)) order.push(key);
    values[key] = vals;
  });

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = sheet_(TABS[form]);
    var hRow = headerRow_(sheet);
    var headers = headers_(sheet, hRow);
    var row = new Array(headers.length).fill('');
    row[ensureColumn_(sheet, headers, 'Timestamp', hRow)] = new Date();
    if (row.length < headers.length) row.length = headers.length;
    order.forEach(function (field) {
      var target = (FIELD_MAP[form] || {})[field] || field;
      var col = findColumn_(headers, target);
      if (col < 0) col = ensureColumn_(sheet, headers, target, hRow);
      row[col] = safe_(values[field].join(', '));
    });
    for (var i = 0; i < headers.length; i++) if (row[i] === undefined) row[i] = '';
    var merged = form === 'tryouts' && mergeIntoHighlighted_(sheet, hRow, headers, row);
    if (!merged) {
      // Insert inside the player list (above its last row), so the table and any formulas that count
      // its range grow to include the new row; the sort below then puts it in its place.
      var count = dataRowCount_(sheet, hRow, headers.length);
      var at = count >= 2 ? hRow + count : hRow + count + 1;
      if (count >= 2) sheet.insertRowBefore(at); else sheet.insertRowAfter(at - 1);
      sheet.getRange(at, 1, 1, headers.length).setValues([row.slice(0, headers.length)]);
    }
    if (form === 'tryouts') sortTryouts_(sheet, hRow);
  } finally {
    lock.releaseLock();
  }
  return reply_('ok');
}

function doGet() { return reply_('MundyBall form endpoint is running.'); }

// ----- Tryout tab ordering -----
// Registered players: Boys first, then Girls, each by grade 3rd -> 8th (ties keep their current order).
// Rows whose First or Last Name cell is highlighted yellow (summer camp, not registered yet) or
// green (played winter feeder last season, not registered yet) stay at the bottom, in their current order.
// Rows move whole (values, colors, notes, formulas), and only rows that are out of place are moved.
// Also available from the sheet menu: MundyBall -> Sort tryout registrations.

function onOpen() {
  SpreadsheetApp.getUi().createMenu('MundyBall')
    .addItem('Sort tryout registrations', 'sortTryoutsNow')
    .addToUi();
}

function sortTryoutsNow() {
  var sheet = sheet_(TABS.tryouts);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { sortTryouts_(sheet, headerRow_(sheet)); } finally { lock.releaseLock(); }
}

function sortTryouts_(sheet, hRow) {
  var headers = headers_(sheet, hRow);
  var cFirst = findColumn_(headers, 'First Name'), cLast = findColumn_(headers, 'Last Name');
  var cGender = findColumn_(headers, 'Gender'), cGrade = findColumn_(headers, 'Grade');
  var first = hRow + 1, n = dataRowCount_(sheet, hRow, headers.length);
  if (n < 2 || cFirst < 0) return;

  var vals = sheet.getRange(first, 1, n, headers.length).getValues();
  var bgs = sheet.getRange(first, 1, n, headers.length).getBackgrounds();
  var main = [], highlighted = [], blank = [];
  vals.forEach(function (r, i) {
    var name = String(r[cFirst]).trim() + String(cLast >= 0 ? r[cLast] : '').trim();
    if (!name) blank.push(i);
    else if (isYellowOrGreen_(bgs[i][cFirst]) || (cLast >= 0 && isYellowOrGreen_(bgs[i][cLast]))) highlighted.push(i);
    else main.push(i);
  });
  main.sort(function (a, b) {
    return genderRank_(cGender >= 0 ? vals[a][cGender] : '') - genderRank_(cGender >= 0 ? vals[b][cGender] : '') ||
           gradeRank_(cGrade >= 0 ? vals[a][cGrade] : '') - gradeRank_(cGrade >= 0 ? vals[b][cGrade] : '') ||
           a - b;
  });
  var target = main.concat(highlighted, blank);

  // Move rows into place one at a time, skipping any already where they belong.
  var cur = vals.map(function (_, i) { return i; });
  for (var pos = 0; pos < target.length; pos++) {
    var from = cur.indexOf(target[pos]);
    if (from === pos) continue;
    sheet.moveRows(sheet.getRange(first + from, 1), first + pos);
    cur.splice(pos, 0, cur.splice(from, 1)[0]);
  }
}

function genderRank_(v) {
  var s = String(v).trim().toLowerCase();
  if (/^(boy|male|m$)/.test(s)) return 0;
  if (/^(girl|female|f$)/.test(s)) return 1;
  return 2;
}

function gradeRank_(v) {
  var m = String(v).match(/\d+/);
  return m ? Number(m[0]) : 99;   // "5th Grade", "5th", "5" -> 5; blank/unknown last
}

// 'yellow' or 'green' for a yellow/green fill (any shade); null for white, gray, other colors, or no fill.
function colorOf_(hex) {
  var m = /^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(String(hex || ''));
  if (!m) return null;
  var r = parseInt(m[1], 16) / 255, g = parseInt(m[2], 16) / 255, b = parseInt(m[3], 16) / 255;
  var max = Math.max(r, g, b), min = Math.min(r, g, b), d = max - min;
  if (max === 0 || d / max < 0.05) return null;            // white or gray: not a highlight (pale fills like #d9ead3 still count)
  var h;
  if (max === r) h = 60 * (((g - b) / d) % 6);
  else if (max === g) h = 60 * ((b - r) / d + 2);
  else h = 60 * ((r - g) / d + 4);
  if (h < 0) h += 360;
  if (h >= 40 && h < 80) return 'yellow';
  if (h >= 80 && h <= 170) return 'green';
  return null;
}
function isYellowOrGreen_(hex) { return colorOf_(hex) !== null; }

// ----- Merging a registration into a highlighted (not-yet-registered) row -----
// Match: same first + last name (ignoring case, spaces, punctuation) on exactly one yellow/green row.
// Each column: empty -> filled; same value (allowing format differences) -> kept; different -> "new / old".
// Then the yellow/green fill is cleared and Source Records Merged notes where the row came from.
// Two or more highlighted rows with the same name: not merged; the new row gets a Review Notes flag.
function mergeIntoHighlighted_(sheet, hRow, headers, row) {
  var cF = findColumn_(headers, 'First Name'), cL = findColumn_(headers, 'Last Name');
  if (cF < 0 || cL < 0) return false;
  var key = nameKey_(row[cF], row[cL]);
  var first = hRow + 1, n = dataRowCount_(sheet, hRow, headers.length);
  if (!key || n < 1) return false;
  var range = sheet.getRange(first, 1, n, headers.length);
  var vals = range.getValues(), bgs = range.getBackgrounds();
  var hits = [];
  vals.forEach(function (r, i) {
    var color = colorOf_(bgs[i][cF]) || colorOf_(bgs[i][cL]);
    if (color && nameKey_(r[cF], r[cL]) === key) hits.push({ i: i, color: color });
  });
  if (hits.length !== 1) {
    var cNote = findColumn_(headers, 'Review Notes');
    if (hits.length > 1 && cNote >= 0) row[cNote] = 'Name matches ' + hits.length + ' highlighted rows; not merged. Check for duplicates.';
    return false;
  }
  var hit = hits[0], old = vals[hit.i], rowNum = first + hit.i;
  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  for (var c = 0; c < headers.length; c++) {
    var nv = row[c];
    if (nv === '' || nv === null || nv === undefined) continue;
    var out;
    if (nv instanceof Date || old[c] === '' || old[c] === null) {
      out = nv;
    } else {
      var oldText = asText_(old[c], tz);
      var dup = oldText.split(' / ').some(function (part) { return sameValue_(headers[c], part, nv); });
      if (dup) continue;
      out = String(nv).replace(/^'/, '') + ' / ' + oldText;
      out = /^[=+\-@]/.test(out) ? "'" + out : out;
    }
    sheet.getRange(rowNum, c + 1).setValue(out);
  }
  var cSrc = findColumn_(headers, 'Source Records Merged');
  if (cSrc >= 0) {
    var label = (hit.color === 'yellow' ? 'Summer camp list' : 'Winter feeder last season') +
                ' + tryout form ' + Utilities.formatDate(new Date(), tz, 'MMM d, yyyy');
    var prev = asText_(old[cSrc], tz);
    sheet.getRange(rowNum, cSrc + 1).setValue(prev ? prev + '; ' + label : label);
  }
  sheet.getRange(rowNum, 1, 1, headers.length)
    .setBackgrounds([bgs[hit.i].map(function (b) { return colorOf_(b) ? null : b; })]);
  return true;
}

function nameKey_(firstName, lastName) {
  return (String(firstName) + '|' + String(lastName)).toLowerCase().replace(/[^a-z|]/g, '');
}

function asText_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  return String(v === null || v === undefined ? '' : v).trim();
}

// Same value for this column, allowing harmless format differences.
function sameValue_(header, a, b) {
  var h = norm_(header);
  a = String(a).replace(/^'/, '').trim(); b = String(b).replace(/^'/, '').trim();
  if (/name/.test(h)) return a.toLowerCase().replace(/[^a-z]/g, '') === b.toLowerCase().replace(/[^a-z]/g, '');
  if (/grade/.test(h)) return gradeRank_(a) === gradeRank_(b);
  if (/gender/.test(h)) return genderRank_(a) === genderRank_(b);
  if (/phone/.test(h)) return a.replace(/\D/g, '').slice(-10) === b.replace(/\D/g, '').slice(-10);
  if (/birthday|birth|dob/.test(h)) { var da = new Date(a), db = new Date(b); if (!isNaN(da) && !isNaN(db)) return da.toISOString().slice(0, 10) === db.toISOString().slice(0, 10); }
  return norm_(a) === norm_(b);
}

// ----- helpers -----
function first_(v) { return Array.isArray(v) ? v[0] : (v || ''); }
function safe_(v) {
  var s = String(v).slice(0, MAX_CELL);
  return /^[=+\-@]/.test(s) ? "'" + s : s;   // leading ' keeps it as text (e.g. phone "+1 847...")
}
function norm_(s) { return String(s).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim(); }
function reply_(msg) { return ContentService.createTextOutput(msg).setMimeType(ContentService.MimeType.TEXT); }

// Finds the tab ignoring capitals and extra spaces; creates it only if nothing close exists.
function sheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var exact = ss.getSheetByName(name);
  if (exact) return exact;
  var want = norm_(name);
  var match = ss.getSheets().filter(function (s) { return norm_(s.getName()) === want; })[0];
  return match || ss.insertSheet(name);
}

// The header row: the first of the top 30 rows that has both a "First Name" and a "Last Name" cell;
// otherwise the fullest of the top 10 rows (so a title row above the table is skipped).
function headerRow_(sheet) {
  var last = sheet.getLastColumn(), rows = Math.min(sheet.getLastRow(), 30);
  if (last === 0 || rows === 0) return 1;
  var grid = sheet.getRange(1, 1, rows, last).getValues();
  for (var i = 0; i < grid.length; i++) {
    var cells = grid[i].map(norm_);
    if (cells.indexOf('first name') >= 0 && cells.indexOf('last name') >= 0) return i + 1;
  }
  var best = 1, bestCount = 0;
  grid.slice(0, 10).forEach(function (r, i) {
    var n = r.filter(function (c) { return String(c).trim() !== ''; }).length;
    if (n > bestCount) { bestCount = n; best = i + 1; }
  });
  return best;
}

// Number of player rows under the header. The list ends at the first row that is completely empty,
// has no first or last name, or mentions one of the other tables on the tab (see OTHER_TABLES), so the
// small tables under the registration list are never moved, sorted, or merged into.
var OTHER_TABLES = ['tryout registration counts', 'coaching interest', 'coaches assignments'];

function dataRowCount_(sheet, hRow, width) {
  var n = sheet.getLastRow() - hRow;
  if (n < 1) return 0;
  var vals = sheet.getRange(hRow + 1, 1, n, width).getValues();
  var headers = headers_(sheet, hRow);
  var cF = findColumn_(headers, 'First Name'), cL = findColumn_(headers, 'Last Name');
  for (var i = 0; i < vals.length; i++) {
    var r = vals[i];
    var cells = r.map(norm_);
    if (cells.every(function (c) { return c === ''; })) return i;
    if (cF >= 0 && cells[cF] === '') return i;
    if (cL >= 0 && cells[cL] === '') return i;
    if (cells.some(function (c) { return OTHER_TABLES.some(function (t) { return c.indexOf(t) >= 0; }); })) return i;
    if (cF >= 0 && cells[cF] === 'first name') return i;      // another header row
  }
  return n;
}

function headers_(sheet, hRow) {
  var last = sheet.getLastColumn();
  if (last === 0) return [];
  return sheet.getRange(hRow || 1, 1, 1, last).getValues()[0].map(String);
}

function ensureColumn_(sheet, headers, name, hRow) {
  var col = findColumn_(headers, name);
  if (col >= 0) return col;
  headers.push(name);
  sheet.getRange(hRow || 1, headers.length).setValue(name).setFontWeight('bold');
  if (sheet.getFrozenRows() === 0) sheet.setFrozenRows(hRow || 1);
  return headers.length - 1;
}

// Exact (normalized) header match first; then variants: synonym swaps at the start/end/whole
// ("Camper First Name" -> "Player First Name"), and the name without a leading person word
// ("Camper Grade" -> "Grade", "Camper Birthday" -> "Birthday" -> "Date of Birth").
var PERSON_WORDS = ['camper', 'player', 'athlete', 'child'];

function findColumn_(headers, field) {
  var normHeaders = headers.map(norm_);
  var target = norm_(field);
  var bases = [target];
  PERSON_WORDS.forEach(function (w) {
    if (target.indexOf(w + ' ') === 0) bases.push(target.slice(w.length + 1));
  });
  var candidates = [];
  bases.forEach(function (b) {
    candidates.push(b);
    Object.keys(SYNONYMS).forEach(function (k) {
      var nk = norm_(k);
      SYNONYMS[k].forEach(function (alt) {
        var na = norm_(alt);
        if (b === nk) candidates.push(na);
        if (b.indexOf(nk + ' ') === 0) candidates.push(na + b.slice(nk.length));              // prefix swap
        if (b.slice(-nk.length - 1) === ' ' + nk) candidates.push(b.slice(0, -nk.length) + na); // suffix swap
      });
    });
  });
  for (var c = 0; c < candidates.length; c++) {
    var i = normHeaders.indexOf(candidates[c]);
    if (i >= 0) return i;
  }
  return -1;
}
