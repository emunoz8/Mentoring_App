/*** SignInServer.gs
 * Handles student sign-in sessions, known student directory, and suggestion helpers.
 * Relies on:
 *   - QueueServer.gs (queue_ensureSignInLogSheet_, STATUS)
 *   - FormServer.gs (_getSuggestRoster_, _norm_)
 *   - Utils.gs (ymd_, parseLooseDate_)
 */

const SIGNIN_KNOWN = { SHEET: 'known_students' };
const SIGNIN_SESSIONS = { SHEET: 'sign_in_sessions' };

function signinNorm_(value) {
  if (typeof _norm_ === 'function') return _norm_(value);
  return String(value || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim();
}

function signinEnsureKnownStudentsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SIGNIN_KNOWN.SHEET);
  if (!sh) sh = ss.insertSheet(SIGNIN_KNOWN.SHEET);

  if (sh.getLastRow() < 1 || sh.getLastColumn() < 1) {
    sh.clear();
    sh.getRange(1, 1, 1, 8).setValues([[
      'StudentID','FirstName','LastName','School','Email','Grade','CreatedAt','LastSignIn'
    ]]);
  } else {
    const lc0 = Math.max(1, sh.getLastColumn());
    const first = sh.getRange(1, 1, 1, lc0).getValues()[0];
    if (first.every(v => String(v || '').trim() === '')) {
      sh.getRange(1, 1, 1, 8).setValues([[
        'StudentID','FirstName','LastName','School','Email','Grade','CreatedAt','LastSignIn'
      ]]);
    }
  }

  const lc = Math.max(1, sh.getLastColumn());
  const header = sh.getRange(1, 1, 1, lc).getValues()[0].map(h => String(h || '').trim());
  const idx = new Map(header.map((h, i) => [signinNorm_(h), i]));

  const C = {
    StudentID : idx.get('studentid') ?? idx.get('id') ?? idx.get('cpsid') ?? idx.get('cpsidnumber'),
    FirstName : idx.get('firstname') ?? idx.get('first'),
    LastName  : idx.get('lastname')  ?? idx.get('last'),
    School    : idx.get('school')    ?? idx.get('site'),
    Email     : idx.get('email')     ?? idx.get('emailaddress'),
    Grade     : idx.get('grade')     ?? idx.get('currentgrade') ?? idx.get('currentgradelevel'),
    CreatedAt : idx.get('createdat') ?? idx.get('created'),
    LastSignIn: idx.get('lastsignin') ?? idx.get('lastsignedin'),
  };

  const needed = [
    ['StudentID', 'StudentID'],
    ['FirstName', 'FirstName'],
    ['LastName', 'LastName'],
    ['School', 'School'],
    ['Email', 'Email'],
    ['Grade', 'Grade'],
    ['CreatedAt', 'CreatedAt'],
    ['LastSignIn', 'LastSignIn'],
  ].filter(([key]) => C[key] == null);

  if (needed.length) {
    const start = sh.getLastColumn() + 1;
    sh.getRange(1, start, 1, needed.length).setValues([needed.map(x => x[1])]);
    const header2 = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(h => String(h || '').trim());
    const idx2 = new Map(header2.map((h, i) => [signinNorm_(h), i]));
    const resolve = (cur, ...keys) => C[cur] ?? keys.map(k => idx2.get(k)).find(x => x != null);
    C.StudentID  = resolve('StudentID','studentid','id','cpsid','cpsidnumber');
    C.FirstName  = resolve('FirstName','firstname','first');
    C.LastName   = resolve('LastName','lastname','last');
    C.School     = resolve('School','school','site');
    C.Email      = resolve('Email','email','emailaddress');
    C.Grade      = resolve('Grade','grade','currentgrade','currentgradelevel');
    C.CreatedAt  = resolve('CreatedAt','createdat','created');
    C.LastSignIn = resolve('LastSignIn','lastsignin','lastsignedin');
  }

  return { sh, C };
}

function signinEnsureSessionsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SIGNIN_SESSIONS.SHEET);
  if (!sh) sh = ss.insertSheet(SIGNIN_SESSIONS.SHEET);

  if (sh.getLastRow() < 1 || sh.getLastColumn() < 1) {
    sh.clear();
    sh.getRange(1, 1, 1, 9).setValues([[
      'SessionID','Label','Type','Date','IsActive','CreatedAt','ClosedAt','LastSignInAt','SignInCount'
    ]]);
  } else {
    const lc0 = Math.max(1, sh.getLastColumn());
    const first = sh.getRange(1, 1, 1, lc0).getValues()[0];
    if (first.every(v => String(v || '').trim() === '')) {
      sh.getRange(1, 1, 1, 9).setValues([[
        'SessionID','Label','Type','Date','IsActive','CreatedAt','ClosedAt','LastSignInAt','SignInCount'
      ]]);
    }
  }

  const lc = Math.max(1, sh.getLastColumn());
  const header = sh.getRange(1, 1, 1, lc).getValues()[0].map(h => String(h || '').trim());
  const idx = new Map(header.map((h, i) => [signinNorm_(h), i]));

  const C = {
    SessionID   : idx.get('sessionid') ?? idx.get('id'),
    Label       : idx.get('label') ?? idx.get('group') ?? idx.get('name'),
    Type        : idx.get('type') ?? idx.get('sessiontype'),
    Date        : idx.get('date') ?? idx.get('ymd'),
    IsActive    : idx.get('isactive') ?? idx.get('active'),
    CreatedAt   : idx.get('createdat') ?? idx.get('created'),
    ClosedAt    : idx.get('closedat') ?? idx.get('closed'),
    LastSignInAt: idx.get('lastsignin') ?? idx.get('lastsigninat') ?? idx.get('lastsignin_at'),
    SignInCount : idx.get('signincount') ?? idx.get('signins'),
  };

  const needed = [
    ['Type','Type'],
    ['SessionID','SessionID'],
    ['Label','Label'],
    ['Date','Date'],
    ['IsActive','IsActive'],
    ['CreatedAt','CreatedAt'],
    ['ClosedAt','ClosedAt'],
    ['LastSignInAt','LastSignInAt'],
    ['SignInCount','SignInCount']
  ].filter(([key]) => C[key] == null);

  if (needed.length) {
    const start = sh.getLastColumn() + 1;
    sh.getRange(1, start, 1, needed.length).setValues([needed.map(x => x[1])]);
    const header2 = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(h => String(h || '').trim());
    const idx2 = new Map(header2.map((h, i) => [signinNorm_(h), i]));
    const resolve = (cur, ...keys) => C[cur] ?? keys.map(k => idx2.get(k)).find(x => x != null);
    C.SessionID    = resolve('SessionID','sessionid','id');
    C.Label        = resolve('Label','label','group','name');
    C.Type         = resolve('Type','type','sessiontype');
    C.Date         = resolve('Date','date','ymd');
    C.IsActive     = resolve('IsActive','isactive','active');
    C.CreatedAt    = resolve('CreatedAt','createdat','created');
    C.ClosedAt     = resolve('ClosedAt','closedat','closed');
    C.LastSignInAt = resolve('LastSignInAt','lastsignin','lastsigninat','lastsignin_at');
    C.SignInCount  = resolve('SignInCount','signincount','signins');
  }

  return { sh, C };
}

function signinKnownCacheKey_() { return 'SIGNIN_KNOWN_V1'; }

function signinInvalidateKnownCache_() {
  try {
    CacheService.getScriptCache().remove(signinKnownCacheKey_());
  } catch (_) {}
}

function signinFetchKnownStudents_() {
  const cache = CacheService.getScriptCache();
  const key = signinKnownCacheKey_();
  const hit = cache.get(key);
  if (hit) {
    try { return JSON.parse(hit); }
    catch (_) { /* fall through */ }
  }

  const { sh, C } = signinEnsureKnownStudentsSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const width = Math.max(1, sh.getLastColumn());
  const data = sh.getRange(2, 1, lastRow - 1, width).getValues();
  const list = [];
  data.forEach(row => {
    const id = String(row[C.StudentID] || '').trim();
    if (!id) return;
    list.push({
      id,
      firstName : String(row[C.FirstName] || '').trim(),
      lastName  : String(row[C.LastName] || '').trim(),
      school    : String(row[C.School] || '').trim(),
      email     : String(row[C.Email] || '').trim(),
      grade     : String(row[C.Grade] || '').trim(),
      lastSignIn: row[C.LastSignIn] instanceof Date
        ? row[C.LastSignIn].toISOString()
        : String(row[C.LastSignIn] || '').trim()
    });
  });

  try { cache.put(key, JSON.stringify(list), 300); } catch (_) {}
  return list;
}

function signinUpsertKnownStudent(student, options = {}) {
  const opts = options || {};
  const hasTimestamp = Object.prototype.hasOwnProperty.call(opts, 'timestamp');
  const timestamp = hasTimestamp ? opts.timestamp : new Date();
  const updateLastSignIn = opts.updateLastSignIn !== false;
  const createdAtValue = Object.prototype.hasOwnProperty.call(opts, 'createdAt')
    ? opts.createdAt
    : new Date();

  const id = String(student?.studentId || student?.id || '').trim();
  if (!id) throw new Error('Student ID is required.');

  const firstName = String(student?.firstName || '').trim();
  const lastName  = String(student?.lastName || '').trim();
  const school    = String(student?.school || '').trim();
  const email     = String(student?.email || '').trim();
  const gradeRaw  = String(student?.grade || '').trim();
  const grade     = signinNormalizeGradeLabel_(gradeRaw);

  const { sh, C } = signinEnsureKnownStudentsSheet_();
  const width = Math.max(1, sh.getLastColumn());
  const lastRow = sh.getLastRow();

  if (lastRow >= 2) {
    const data = sh.getRange(2, 1, lastRow - 1, width).getValues();
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const rowId = String(row[C.StudentID] || '').trim();
      if (rowId !== id) continue;

      const next = row.slice();
      if (firstName) next[C.FirstName] = firstName;
      if (lastName)  next[C.LastName]  = lastName;
      if (school)    next[C.School]    = school;
      if (email)     next[C.Email]     = email;
      if (grade)     next[C.Grade]     = grade;
      if (C.LastSignIn != null && updateLastSignIn && hasTimestamp) next[C.LastSignIn] = timestamp;
      if (C.CreatedAt != null && !next[C.CreatedAt] && createdAtValue != null) next[C.CreatedAt] = createdAtValue;

      sh.getRange(i + 2, 1, 1, width).setValues([next]);
      signinInvalidateKnownCache_();
      return { created: false, rowIndex: i + 2 };
    }
  }

  const newRow = new Array(width).fill('');
  if (C.StudentID != null) newRow[C.StudentID] = id;
  if (C.FirstName != null) newRow[C.FirstName] = firstName;
  if (C.LastName != null) newRow[C.LastName] = lastName;
  if (C.School != null) newRow[C.School] = school;
  if (C.Email != null) newRow[C.Email] = email;
  if (C.Grade != null) newRow[C.Grade] = grade;
  if (C.CreatedAt != null) newRow[C.CreatedAt] = createdAtValue;
  if (C.LastSignIn != null && updateLastSignIn && hasTimestamp) newRow[C.LastSignIn] = timestamp;

  sh.appendRow(newRow);
  signinInvalidateKnownCache_();
  return { created: true, rowIndex: sh.getLastRow() };
}

function signinHydrateSessionRow_(C, row, rowIndex) {
  const tz = Session.getScriptTimeZone() || 'America/Chicago';
  const dateRaw = C.Date != null ? row[C.Date] : null;
  const isActiveCell = C.IsActive != null ? row[C.IsActive] : '';
  const typeRaw = C.Type != null ? String(row[C.Type] || '').trim().toLowerCase() : '';
  const type = typeRaw === 'group' ? 'group' : 'individual';
  return {
    rowIndex,
    id    : String(C.SessionID != null ? row[C.SessionID] : '').trim(),
    label : String(C.Label != null ? row[C.Label] : '').trim(),
    date  : ymd_(dateRaw, tz),
    type,
    isActive: String(isActiveCell || '').trim().toLowerCase() !== 'false' && isActiveCell !== false && isActiveCell !== 0,
    signInCount: Number(C.SignInCount != null ? row[C.SignInCount] : 0) || 0,
    lastSignInAt: (C.LastSignInAt != null && row[C.LastSignInAt] instanceof Date)
      ? row[C.LastSignInAt].toISOString()
      : String(C.LastSignInAt != null ? row[C.LastSignInAt] : '').trim()
  };
}

function signinFindSession_(sessionId) {
  const id = String(sessionId || '').trim();
  if (!id) return null;
  const { sh, C } = signinEnsureSessionsSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;

  const width = Math.max(1, sh.getLastColumn());
  const data = sh.getRange(2, 1, lastRow - 1, width).getValues();
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowId = String(C.SessionID != null ? row[C.SessionID] : '').trim();
    if (rowId === id) {
      return signinHydrateSessionRow_(C, row, i + 2);
    }
  }
  return null;
}

function listActiveSignInSessions(dateRaw) {
  const tz = Session.getScriptTimeZone() || 'America/Chicago';
  const targetDate = ymd_(parseLooseDate_(dateRaw) || new Date(), tz);
  const { sh, C } = signinEnsureSessionsSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok: true, sessions: [] };

  const width = Math.max(1, sh.getLastColumn());
  const data = sh.getRange(2, 1, lastRow - 1, width).getValues();
  const sessions = [];
  data.forEach((row, idx) => {
    const info = signinHydrateSessionRow_(C, row, idx + 2);
    if (!info.isActive) return;
    if (info.date !== targetDate) return;
    sessions.push(info);
  });
  sessions.sort((a, b) => a.label.localeCompare(b.label));
  return { ok: true, sessions };
}

function startSignInSession(labelRaw, dateRaw, typeRaw) {
  const label = String(labelRaw || '').trim();
  if (!label) return { ok: false, error: 'Session label is required.' };

  const tz = Session.getScriptTimeZone() || 'America/Chicago';
  const date = ymd_(parseLooseDate_(dateRaw) || new Date(), tz);
  const requestedType = String(typeRaw || '').trim().toLowerCase() === 'group' ? 'group' : 'individual';

  const lock = LockService.getDocumentLock();
  try {
    lock.waitLock(30000);

    const { sh, C } = signinEnsureSessionsSheet_();
    const lastRow = sh.getLastRow();
    const width = Math.max(1, sh.getLastColumn());
    const labelNorm = signinNorm_(label);

    if (lastRow >= 2) {
      const data = sh.getRange(2, 1, lastRow - 1, width).getValues();
      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const rowDate = ymd_(C.Date != null ? row[C.Date] : null, tz);
        const rowLabel = String(C.Label != null ? row[C.Label] : '').trim();
        const rowTypeRaw = C.Type != null ? String(row[C.Type] || '').trim().toLowerCase() : '';
        const rowType = rowTypeRaw === 'group' ? 'group' : 'individual';
        const activeCell = C.IsActive != null ? row[C.IsActive] : '';
        const isActive = String(activeCell || '').trim().toLowerCase() !== 'false' && activeCell !== false && activeCell !== 0;
        if (rowDate === date && signinNorm_(rowLabel) === labelNorm && rowType === requestedType) {
          if (!isActive) {
            if (C.IsActive != null) sh.getRange(i + 2, C.IsActive + 1).setValue(true);
            if (C.ClosedAt != null) sh.getRange(i + 2, C.ClosedAt + 1).setValue('');
          }
          return { ok: true, session: signinHydrateSessionRow_(C, row, i + 2) };
        }
      }
    }

    const id = Utilities.getUuid();
    const newRow = new Array(width).fill('');
    if (C.SessionID != null) newRow[C.SessionID] = id;
    if (C.Label != null) newRow[C.Label] = label;
    if (C.Type != null) newRow[C.Type] = requestedType;
    if (C.Date != null) newRow[C.Date] = date;
    if (C.IsActive != null) newRow[C.IsActive] = true;
    const now = new Date();
    if (C.CreatedAt != null) newRow[C.CreatedAt] = now;
    if (C.SignInCount != null) newRow[C.SignInCount] = 0;
    sh.appendRow(newRow);
    const rowIndex = sh.getLastRow();
    return {
      ok: true,
      session: {
        id,
        label,
        date,
        type: requestedType,
        rowIndex,
        isActive: true,
        signInCount: 0,
        lastSignInAt: ''
      }
    };
  } catch (err) {
    return { ok: false, error: err && err.message ? err.message : String(err) };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function signinNormalizeDate_(value) {
  if (value instanceof Date && !isNaN(value)) return value;
  if (typeof parseLooseDate_ === 'function') {
    const parsed = parseLooseDate_(value);
    if (parsed instanceof Date && !isNaN(parsed)) return parsed;
  }
  return null;
}

function signinSplitName_(full) {
  const str = String(full || '').trim();
  if (!str) return { firstName: '', lastName: '' };
  const parts = str.split(/\s+/);
  if (parts.length === 1) {
    return { firstName: parts[0], lastName: '' };
  }
  return {
    firstName: parts[0],
    lastName: parts.slice(1).join(' ')
  };
}

function signinNormalizeGradeLabel_(gradeRaw) {
  const input = String(gradeRaw || '').trim();
  if (!input) return '';
  const lower = input.toLowerCase();
  if (['freshman','freshmen','9','9th','9th grade','grade 9','year 1'].includes(lower)) return 'Freshman';
  if (['sophomore','sophomores','10','10th','10th grade','grade 10','year 2'].includes(lower)) return 'Sophomore';
  if (['junior','juniors','11','11th','11th grade','grade 11','year 3'].includes(lower)) return 'Junior';
  if (['senior','seniors','12','12th','12th grade','grade 12','year 4'].includes(lower)) return 'Senior';
  const num = Number(input.replace(/[^0-9.]/g, ''));
  if (!isNaN(num)) {
    if (num <= 9) return 'Freshman';
    if (num === 10) return 'Sophomore';
    if (num === 11) return 'Junior';
    if (num >= 12) return 'Senior';
  }
  return input.replace(/\b\w/g, c => c.toUpperCase());
}

function bootstrapKnownStudents() {
  const summary = {
    ok: true,
    added: 0,
    skippedExisting: 0,
    sources: {}
  };

  const SOURCE_PRIORITY = {
    form_2026: 2,
    sign_in_log: 1,
    attendance: 0
  };

  const { sh: knownSh, C: KC } = signinEnsureKnownStudentsSheet_();
  const knownWidth = Math.max(1, knownSh.getLastColumn());
  const existingRows = knownSh.getLastRow() >= 2
    ? knownSh.getRange(2, 1, knownSh.getLastRow() - 1, knownWidth).getValues()
    : [];
  const knownIds = new Set();
  existingRows.forEach(row => {
    const id = KC.StudentID != null ? String(row[KC.StudentID] || '').trim() : '';
    if (id) knownIds.add(id);
  });

  const aggregates = new Map();

  function ensureSourceBucket(source) {
    if (!summary.sources[source]) {
      summary.sources[source] = { considered: 0, merged: 0, skipped: 0 };
    }
    return summary.sources[source];
  }

  function mergeAggregate(source, info) {
    const bucket = ensureSourceBucket(source);
    bucket.considered += 1;
    const id = String(info?.studentId || info?.id || '').trim();
    if (!id) {
      bucket.skipped += 1;
      return;
    }
    let entry = aggregates.get(id);
    if (!entry) {
      entry = {
        studentId: id,
        firstName: '',
        lastName: '',
        school: '',
        email: '',
        grade: '',
        lastSignIn: null,
        _priority: {}
      };
      aggregates.set(id, entry);
    }
    function setField(field, value) {
      const val = String(value || '').trim();
      if (!val) return;
      const pr = SOURCE_PRIORITY[source] ?? 0;
      const current = entry._priority[field] ?? -Infinity;
      if (!entry[field] || pr >= current) {
        entry[field] = val;
        entry._priority[field] = pr;
      }
    }
    setField('firstName', info.firstName);
    setField('lastName', info.lastName);
    setField('school', info.school);
    setField('email', info.email);
    const normalizedGrade = signinNormalizeGradeLabel_(info.grade);
    setField('grade', normalizedGrade);
    const stamp = info.lastSignIn instanceof Date && !isNaN(info.lastSignIn)
      ? info.lastSignIn
      : null;
    if (stamp) {
      if (!entry.lastSignIn || entry.lastSignIn < stamp) {
        entry.lastSignIn = stamp;
      }
    }
    bucket.merged += 1;
  }

  function collectFormSheet(sheetName, source) {
    if (!sheetName || !CONFIG || !CONFIG.FORM || !CONFIG.FORM.COLS) return;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return;
    const values = sh.getDataRange().getValues();
    const C = CONFIG.FORM.COLS;
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const id = C.cpsIdNumber != null ? String(row[C.cpsIdNumber] || '').trim() : '';
      if (!id) {
        mergeAggregate(source, { studentId: '', firstName: '', lastName: '' });
        continue;
      }
      const first = C.firstName != null ? row[C.firstName] : '';
      const last = C.lastName != null ? row[C.lastName] : '';
      const school = C.school != null ? row[C.school] : '';
      const emailPrimary = C.emailAddress != null ? row[C.emailAddress] : '';
      const emailAlt = C.participantEmails != null ? row[C.participantEmails] : '';
      const gradeCurrent = C.currentGradeLevel != null ? signinNormalizeGradeLabel_(row[C.currentGradeLevel]) : '';
      const gradeAtIntake = C.gradeAtIntake != null ? signinNormalizeGradeLabel_(row[C.gradeAtIntake]) : '';
      mergeAggregate(source, {
        studentId: id,
        firstName: first,
        lastName: last,
        school,
        email: String(emailPrimary || emailAlt || '').split(/[;,]/)[0],
        grade: gradeCurrent || gradeAtIntake,
        lastSignIn: null
      });
    }
  }

  function collectAttendance() {
    if (!CONFIG || !CONFIG.ATTENDANCE || !CONFIG.ATTENDANCE.SHEET) return;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(CONFIG.ATTENDANCE.SHEET);
    if (!sh || sh.getLastRow() < 2) return;
    const lastCol = sh.getLastColumn();
    const values = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
    const A = CONFIG.ATTENDANCE.COLS || {};
    values.forEach(row => {
      const id = A.idNumber != null ? String(row[A.idNumber] || '').trim() : '';
      if (!id) {
        mergeAggregate('attendance', { studentId: '', firstName: '', lastName: '' });
        return;
      }
      const first = A.firstName != null ? row[A.firstName] : '';
      const last = A.lastName != null ? row[A.lastName] : '';
      const school = A.school != null ? row[A.school] : '';
      const grade = A.schoolYear != null ? signinNormalizeGradeLabel_(row[A.schoolYear]) : '';
      const ts = A.timestamp != null ? row[A.timestamp] : '';
      const lastSignIn = signinNormalizeDate_(ts);
      mergeAggregate('attendance', {
        studentId: id,
        firstName: first,
        lastName: last,
        school,
        grade,
        lastSignIn
      });
    });
  }

  function collectSignInLog() {
    const context = queue_ensureSignInLogSheet_();
    if (!context || !context.sh) return;
    const sh = context.sh;
    const C = context.C || {};
    if (sh.getLastRow() < 2) return;
    const width = Math.max(1, sh.getLastColumn());
    const values = sh.getRange(2, 1, sh.getLastRow() - 1, width).getValues();
    values.forEach(row => {
      const id = C.ID != null ? String(row[C.ID] || '').trim() : '';
      if (!id) {
        mergeAggregate('sign_in_log', { studentId: '', firstName: '', lastName: '' });
        return;
      }
      const nameRaw = C.Name != null ? String(row[C.Name] || '').trim() : '';
      const { firstName, lastName } = signinSplitName_(nameRaw);
      const school = C.School != null ? row[C.School] : '';
      const ts = C.Timestamp != null ? row[C.Timestamp] : '';
      const lastSignIn = signinNormalizeDate_(ts);
      mergeAggregate('sign_in_log', {
        studentId: id,
        firstName,
        lastName,
        school,
        lastSignIn
      });
    });
  }

  if (CONFIG && CONFIG.FORM) {
    collectFormSheet(CONFIG.FORM.SUBMISSIONS_SHEET, 'form_2026');
  }
  collectAttendance();
  collectSignInLog();

  const ids = Array.from(aggregates.keys()).sort();
  ids.forEach(id => {
    if (knownIds.has(id)) {
      summary.skippedExisting += 1;
      return;
    }
    const entry = aggregates.get(id);
    const cleanEntry = {
      studentId: entry.studentId,
      firstName: entry.firstName,
      lastName: entry.lastName,
      school: entry.school,
      email: entry.email,
      grade: entry.grade
    };
    const options = {};
    if (entry.lastSignIn instanceof Date && !isNaN(entry.lastSignIn)) {
      options.timestamp = entry.lastSignIn;
      options.updateLastSignIn = true;
    } else {
      options.updateLastSignIn = false;
    }
    const result = signinUpsertKnownStudent(cleanEntry, options);
    if (result && result.created) {
      summary.added += 1;
      knownIds.add(id);
    } else {
      summary.skippedExisting += 1;
    }
  });

  summary.totalKnown = knownIds.size;
  summary.collected = ids.length;
  return summary;
}

function endSignInSession(sessionId) {
  const info = signinFindSession_(sessionId);
  if (!info) return { ok: false, error: 'Session not found.' };
  if (!info.isActive) return { ok: true, session: info };

  const { sh, C } = signinEnsureSessionsSheet_();
  const now = new Date();
  if (C.IsActive != null) sh.getRange(info.rowIndex, C.IsActive + 1).setValue(false);
  if (C.ClosedAt != null) sh.getRange(info.rowIndex, C.ClosedAt + 1).setValue(now);
  const updated = Object.assign({}, info, { isActive: false, closedAt: now.toISOString() });
  return { ok: true, session: updated };
}

function lookupSignInById(idRaw) {
  const id = String(idRaw || '').trim();
  if (!id) return { ok: false, error: 'Student ID is required.' };

  const known = signinFetchKnownStudents_();
  const knownMatch = known.find(item => item.id === id);
  if (knownMatch) {
    return {
      ok: true,
      record: {
        id,
        firstName: knownMatch.firstName,
        lastName: knownMatch.lastName,
        school: knownMatch.school,
        email: knownMatch.email,
        grade: knownMatch.grade,
        isKnown: true
      }
    };
  }

  const rosterList = (typeof _getSuggestRoster_ === 'function') ? _getSuggestRoster_() : [];
  const rosterMatch = rosterList.find(item => String(item.id || '').trim() === id);
  if (rosterMatch) {
    return {
      ok: true,
      record: {
        id,
        firstName: String(rosterMatch.firstName || '').trim(),
        lastName: String(rosterMatch.lastName || '').trim(),
        school: String(rosterMatch.school || '').trim(),
        email: '',
        grade: String(rosterMatch.grade || '').trim(),
        isKnown: false
      }
    };
  }

  return { ok: false, error: `ID "${id}" not found.` };
}

function signInSuggestPeople(query, limit) {
  const raw = String(query || '').trim();
  if (!raw) return [];

  const norm = signinNorm_;
  const tokens = raw.split(/\s+/).map(norm).filter(Boolean);
  if (!tokens.length) return [];

  const known = signinFetchKnownStudents_();
  const rosterList = (typeof _getSuggestRoster_ === 'function') ? _getSuggestRoster_() : [];
  const combined = new Map();

  known.forEach(rec => {
    if (!rec.id) return;
    combined.set(rec.id, {
      id: rec.id,
      firstName: rec.firstName,
      lastName: rec.lastName,
      school: rec.school,
      email: rec.email,
      grade: signinNormalizeGradeLabel_(rec.grade),
      source: 'known'
    });
  });

  rosterList.forEach(item => {
    const id = String(item.id || '').trim();
    if (!id) return;
    if (combined.has(id)) {
      const tgt = combined.get(id);
      if (!tgt.firstName && item.firstName) tgt.firstName = item.firstName;
      if (!tgt.lastName && item.lastName) tgt.lastName = item.lastName;
      if (!tgt.school && item.school) tgt.school = item.school;
      if (!tgt.grade && item.grade) tgt.grade = signinNormalizeGradeLabel_(item.grade);
      return;
    }
    combined.set(id, {
      id,
      firstName: String(item.firstName || '').trim(),
      lastName: String(item.lastName || '').trim(),
      school: String(item.school || '').trim(),
      email: '',
      grade: signinNormalizeGradeLabel_(item.grade),
      source: 'roster'
    });
  });

  const entries = Array.from(combined.values()).map(entry => {
    const full = `${entry.firstName || ''} ${entry.lastName || ''}`.trim();
    const normFields = [
      norm(entry.id),
      norm(entry.firstName),
      norm(entry.lastName),
      norm(full),
      norm(entry.school),
      norm(entry.email),
      norm(entry.grade)
    ];
    let score = 0;
    tokens.forEach(tok => {
      normFields.forEach(field => {
        if (!tok || !field) return;
        if (field === tok) score += 6;
        else if (field.startsWith(tok)) score += 4;
        else if (field.includes(tok)) score += 2;
      });
    });
    const labelParts = [];
    if (full) labelParts.push(full);
    if (entry.id) labelParts.push(entry.id);
    const meta = [entry.school, entry.grade].map(x => String(x || '').trim()).filter(Boolean);
    const label = labelParts.length
      ? `${labelParts.join(' · ')}${meta.length ? ' (' + meta.join(' • ') + ')' : ''}`
      : entry.id;
    return Object.assign({}, entry, {
      fullName: full,
      label,
      score
    });
  }).filter(item => item.score > 0);

  entries.sort((a, b) => {
    if (b.score !== a.score) return b.score - a.score;
    return (a.fullName || '').localeCompare(b.fullName || '') || a.id.localeCompare(b.id);
  });

  const n = Math.max(1, Math.min(Number(limit) || 10, 20));
  return entries.slice(0, n).map(({ score, fullName, ...rest }) => rest);
}

function signinRecordGroupAttendance_(session, details, timestamp) {
  if (typeof CONFIG === 'undefined' || !CONFIG || !CONFIG.ATTENDANCE) {
    throw new Error('Attendance configuration unavailable.');
  }
  const sheetName = CONFIG.ATTENDANCE.SHEET;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Attendance sheet "' + sheetName + '" not found.');

  const width = Math.max(1, sh.getLastColumn());
  const row = new Array(width).fill('');
  const A = CONFIG.ATTENDANCE.COLS || {};
  if (A.timestamp != null) row[A.timestamp] = timestamp;
  if (A.firstName != null) row[A.firstName] = details.firstName || '';
  if (A.lastName != null) row[A.lastName] = details.lastName || '';
  if (A.schoolYear != null) row[A.schoolYear] = details.grade || '';
  if (A.school != null) row[A.school] = details.school || '';
  if (A.idNumber != null) row[A.idNumber] = details.studentId || '';
  if (A.group != null) row[A.group] = session.label || '';
  sh.appendRow(row);
  return sh.getLastRow();
}

function recordStudentSignIn(payload) {
  const sessionId = payload && payload.sessionId;
  const student = payload && payload.student;
  if (!sessionId) return { ok: false, error: 'Session ID is required.' };
  if (!student) return { ok: false, error: 'Student payload is required.' };

  const id = String(student.id || student.studentId || '').trim();
  if (!id) return { ok: false, error: 'Student ID is required.' };

  const lock = LockService.getDocumentLock();
  try {
    lock.waitLock(30000);

    const session = signinFindSession_(sessionId);
    if (!session || !session.isActive) {
      return { ok: false, error: 'Session is not active or not found.' };
    }
    const sessionType = session.type === 'group' ? 'group' : 'individual';

  const now = new Date();
  const details = {
    studentId: id,
    firstName: String(student.firstName || '').trim(),
    lastName : String(student.lastName || '').trim(),
    school   : String(student.school || '').trim(),
    email    : String(student.email || '').trim(),
    grade    : String(student.grade || '').trim()
  };
  if (details.grade) {
    details.grade = signinNormalizeGradeLabel_(details.grade);
  }

    const upsert = signinUpsertKnownStudent(details, { timestamp: now });
    const fullName = `${details.firstName || ''} ${details.lastName || ''}`.trim() || id;

    let rowIndex = null;
    let targetSheet = '';
    if (sessionType === 'group') {
      rowIndex = signinRecordGroupAttendance_(session, details, now);
      targetSheet = CONFIG && CONFIG.ATTENDANCE ? CONFIG.ATTENDANCE.SHEET : '';
    } else {
      const { sh: logSh, C: L } = queue_ensureSignInLogSheet_();
      const width = Math.max(1, logSh.getLastColumn());
      const row = new Array(width).fill('');
      if (L.Timestamp != null) row[L.Timestamp] = now;
      if (L.ID != null) row[L.ID] = id;
      if (L.Name != null) row[L.Name] = fullName;
      if (L.School != null) row[L.School] = details.school;
      if (L.Group != null) row[L.Group] = session.label;
      if (L.MentorRaw != null) row[L.MentorRaw] = '';
      if (L.Status != null) row[L.Status] = STATUS ? STATUS.PENDING : 'Pending';
      if (L.ClaimedBy != null) row[L.ClaimedBy] = '';
      if (L.ClaimedAt != null) row[L.ClaimedAt] = '';
      if (L.ProcessedAt != null) row[L.ProcessedAt] = '';
      if (L.ContactID != null) row[L.ContactID] = '';

      logSh.appendRow(row);
      rowIndex = logSh.getLastRow();
      targetSheet = SIGN_IN.SHEET;
    }

    const { sh: sessSh, C: SC } = signinEnsureSessionsSheet_();
    if (SC.LastSignInAt != null) {
      sessSh.getRange(session.rowIndex, SC.LastSignInAt + 1).setValue(now);
    }
    if (SC.SignInCount != null) {
      const cell = sessSh.getRange(session.rowIndex, SC.SignInCount + 1);
      const current = Number(cell.getValue()) || 0;
      cell.setValue(current + 1);
    }

    return {
      ok: true,
      session: { id: session.id, label: session.label, date: session.date, type: sessionType },
      student: details,
      createdKnown: upsert.created,
      rowIndex,
      targetSheet
    };
  } catch (err) {
    return { ok: false, error: err && err.message ? err.message : String(err) };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}
