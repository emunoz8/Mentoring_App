/*** SignInServer.gs
 * Handles student sign-in sessions, known student directory, and suggestion helpers.
 * Relies on:
 *   - QueueServer.gs (queue_ensureSignInLogSheet_, STATUS)
 *   - FormServer.gs (_getSuggestRoster_, _norm_)
 *   - Utils.gs (ymd_, parseLooseDate_)
 */

const SIGNIN_KNOWN = { SHEET: 'known_students' };
const SIGNIN_SESSIONS = { SHEET: 'sign_in_sessions' };
const SIGNIN_SUGGEST_CACHE = globalThis.SIGNIN_SUGGEST_CACHE || (globalThis.SIGNIN_SUGGEST_CACHE = { entries: null, expires: 0 });
const SIGNIN_SUGGEST_CACHE_TTL_MS = 5 * 60 * 1000; // 5 minutes

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
  SIGNIN_SUGGEST_CACHE.entries = null;
  SIGNIN_SUGGEST_CACHE.expires = 0;
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

function signinFetchSuggestIndex_() {
  const now = Date.now();
  if (SIGNIN_SUGGEST_CACHE.entries && SIGNIN_SUGGEST_CACHE.expires > now) {
    return SIGNIN_SUGGEST_CACHE.entries;
  }

  const combined = new Map();
  function ensureRecord(idRaw) {
    const id = String(idRaw || '').trim();
    if (!id) return null;
    let rec = combined.get(id);
    if (!rec) {
      rec = { id, firstName: '', lastName: '', school: '', email: '', grade: '', source: '' };
      combined.set(id, rec);
    }
    return rec;
  }
  function mergeInto(id, patch) {
    const rec = ensureRecord(id);
    if (!rec) return;
    if (patch.firstName && !rec.firstName) rec.firstName = patch.firstName;
    if (patch.lastName && !rec.lastName) rec.lastName = patch.lastName;
    if (patch.school && !rec.school) rec.school = patch.school;
    if (patch.email && !rec.email) rec.email = patch.email;
    if (patch.grade && !rec.grade) rec.grade = signinNormalizeGradeLabel_(patch.grade);
    if (patch.source) {
      // Prefer marking as known if either source is known
      if (rec.source !== 'known' || patch.source === 'known') rec.source = patch.source;
    }
  }

  const known = signinFetchKnownStudents_();
  known.forEach((rec) => {
    mergeInto(rec.id, {
      firstName: String(rec.firstName || '').trim(),
      lastName : String(rec.lastName || '').trim(),
      school   : String(rec.school || '').trim(),
      email    : String(rec.email || '').trim(),
      grade    : String(rec.grade || '').trim(),
      source   : 'known'
    });
  });

  const rosterList = (typeof _getSuggestRoster_ === 'function') ? _getSuggestRoster_() : [];
  rosterList.forEach((item) => {
    mergeInto(item && item.id, {
      firstName: String(item && item.firstName || '').trim(),
      lastName : String(item && item.lastName  || '').trim(),
      school   : String(item && item.school   || '').trim(),
      grade    : String(item && item.grade    || '').trim(),
      source   : 'roster'
    });
  });

  const entries = Array.from(combined.values()).map((rec) => {
    const first = String(rec.firstName || '').trim();
    const last  = String(rec.lastName  || '').trim();
    const full  = `${first} ${last}`.trim();
    const school = String(rec.school || '').trim();
    const grade  = signinNormalizeGradeLabel_(rec.grade);
    const email  = String(rec.email || '').trim();
    const labelParts = [];
    if (full) labelParts.push(full);
    if (rec.id) labelParts.push(rec.id);
    const metaParts = [school, grade].filter(Boolean);
    const label = labelParts.length
      ? `${labelParts.join(' · ')}${metaParts.length ? ' (' + metaParts.join(' • ') + ')' : ''}`
      : rec.id;
    const searchFields = [
      rec.id,
      first,
      last,
      full,
      school,
      email,
      grade
    ].map((part) => signinNorm_(part));
    return {
      id: rec.id,
      firstName: first,
      lastName: last,
      school,
      email,
      grade,
      source: rec.source || '',
      label: label || rec.id || '',
      _searchFields: searchFields,
      _fullName: full
    };
  });

  SIGNIN_SUGGEST_CACHE.entries = entries;
  SIGNIN_SUGGEST_CACHE.expires = now + SIGNIN_SUGGEST_CACHE_TTL_MS;
  return entries;
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

function signinFindSessionWithContext_(sessionId, sh, C, hintRow) {
  const id = String(sessionId || '').trim();
  if (!id) return null;
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;

  const width = Math.max(1, sh.getLastColumn());
  const hint = Number(hintRow);
  if (Number.isFinite(hint) && hint >= 2 && hint <= lastRow) {
    const rowValues = sh.getRange(hint, 1, 1, width).getValues();
    if (rowValues && rowValues.length) {
      const info = signinHydrateSessionRow_(C, rowValues[0], hint);
      if (info && info.id === id) return info;
    }
  }

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

function signinFindSession_(sessionId, hintRow) {
  const { sh, C } = signinEnsureSessionsSheet_();
  return signinFindSessionWithContext_(sessionId, sh, C, hintRow);
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
  const n = Math.max(1, Math.min(Number(limit) || 10, 20));
  const index = signinFetchSuggestIndex_();
  if (!Array.isArray(index) || !index.length) return [];

  const scored = [];
  index.forEach((entry) => {
    const fields = entry._searchFields || [];
    let score = 0;
    tokens.forEach((tok) => {
      fields.forEach((field) => {
        if (!tok || !field) return;
        if (field === tok) score += 6;
        else if (field.startsWith(tok)) score += 4;
        else if (field.indexOf(tok) !== -1) score += 2;
      });
    });
    if (score > 0) {
      scored.push({ entry, score });
    }
  });

  scored.sort((a, b) => {
    if (b.score !== a.score) return b.score - a.score;
    const nameA = a.entry._fullName || '';
    const nameB = b.entry._fullName || '';
    const cmp = nameA.localeCompare(nameB);
    if (cmp !== 0) return cmp;
    return (a.entry.id || '').localeCompare(b.entry.id || '');
  });

  return scored.slice(0, n).map(({ entry }) => ({
    id: entry.id,
    firstName: entry.firstName,
    lastName: entry.lastName,
    school: entry.school,
    email: entry.email,
    grade: entry.grade,
    source: entry.source,
    label: entry.label
  }));
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
  const mentorIdOpt = payload && payload.mentorId ? String(payload.mentorId).trim() : '';
  const student = payload && payload.student;
  const sessionRowHint = payload && payload.sessionRow;
  if (!sessionId) return { ok: false, error: 'Session ID is required.' };
  if (!student) return { ok: false, error: 'Student payload is required.' };

  const id = String(student.id || student.studentId || '').trim();
  if (!id) return { ok: false, error: 'Student ID is required.' };

  const now = new Date();
  const details = {
    studentId: id,
    firstName: String(student.firstName || '').trim(),
    lastName : String(student.lastName  || '').trim(),
    school   : String(student.school   || '').trim(),
    email    : String(student.email    || '').trim(),
    grade    : String(student.grade    || '').trim()
  };
  if (details.grade) {
    details.grade = signinNormalizeGradeLabel_(details.grade);
  }

  let sessionInfo;
  let sessionType;
  let rowIndex = null;
  let targetSheet = '';
  let createdKnown = null;
  let sessSh, SC;

  const lock = LockService.getDocumentLock();
  try {
    lock.waitLock(30000);

    const context = signinEnsureSessionsSheet_();
    sessSh = context.sh;
    SC = context.C;
    sessionInfo = signinFindSessionWithContext_(sessionId, sessSh, SC, sessionRowHint);
    if (!sessionInfo || !sessionInfo.isActive) {
      return { ok: false, error: 'Session is not active or not found.' };
    }
    sessionType = sessionInfo.type === 'group' ? 'group' : 'individual';

    if (sessionType === 'group') {
      rowIndex = signinRecordGroupAttendance_(sessionInfo, details, now);
      targetSheet = CONFIG && CONFIG.ATTENDANCE ? CONFIG.ATTENDANCE.SHEET : '';
    } else {
      const { sh: logSh, C: L } = queue_ensureSignInLogSheet_();
      const width = Math.max(1, logSh.getLastColumn());
      const row = new Array(width).fill('');
      if (L.Timestamp != null) row[L.Timestamp] = now;
      if (L.ID != null) row[L.ID] = id;
      if (L.Name != null) row[L.Name] = `${details.firstName || ''} ${details.lastName || ''}`.trim() || id;
      if (L.School != null) row[L.School] = details.school;
      if (L.Group != null) row[L.Group] = sessionInfo.label;
      if (L.MentorRaw != null) row[L.MentorRaw] = mentorIdOpt || '';
      if (L.Status != null) row[L.Status] = STATUS ? STATUS.PENDING : 'Pending';
      if (L.ClaimedBy != null) row[L.ClaimedBy] = '';
      if (L.ClaimedAt != null) row[L.ClaimedAt] = '';
      if (L.ProcessedAt != null) row[L.ProcessedAt] = '';
      if (L.ContactID != null) row[L.ContactID] = '';

      logSh.appendRow(row);
      rowIndex = logSh.getLastRow();
      targetSheet = SIGN_IN.SHEET;
    }

    if (SC.LastSignInAt != null) {
      sessSh.getRange(sessionInfo.rowIndex, SC.LastSignInAt + 1).setValue(now);
    }
    if (SC.SignInCount != null) {
      const cell = sessSh.getRange(sessionInfo.rowIndex, SC.SignInCount + 1);
      const current = Number(cell.getValue()) || 0;
      cell.setValue(current + 1);
    }
  } catch (err) {
    return { ok: false, error: err && err.message ? err.message : String(err) };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }

  try {
    signinEnqueueKnownStudentUpdate_(details, now, { updateLastSignIn: true });
  } catch (err) {
    console.error('Queueing known-student update failed; falling back to direct write.', err);
    try {
      const fallback = signinUpsertKnownStudent(details, { timestamp: now });
      createdKnown = !!(fallback && fallback.created);
    } catch (err2) {
      return { ok: false, error: err2 && err2.message ? err2.message : String(err2) };
    }
  }

  return {
    ok: true,
    session: sessionInfo ? { id: sessionInfo.id, label: sessionInfo.label, date: sessionInfo.date, type: sessionType } : null,
    student: details,
    createdKnown,
    rowIndex,
    targetSheet
  };
}

function recordStudentBatch(payloadRaw) {
  const payload = payloadRaw || {};
  const sessionId = String(payload.sessionId || '').trim();
  if (!sessionId) return { ok: false, error: 'Session ID is required.' };
  const studentsRaw = Array.isArray(payload.students) ? payload.students : [];
  if (!studentsRaw.length) return { ok: false, error: 'No students provided.' };

  const now = new Date();
  const results = studentsRaw.map((student) => ({
    studentId: String(student && (student.studentId || student.id) || '').trim(),
    firstName: String(student && student.firstName || '').trim(),
    lastName : String(student && student.lastName  || '').trim(),
    ok: false
  }));
  const lock = LockService.getDocumentLock();
  let sessionInfo;
  let targetSheet = '';
  try {
    lock.waitLock(30000);

    const { sh: sessSh, C: SC } = signinEnsureSessionsSheet_();
    sessionInfo = signinFindSessionWithContext_(sessionId, sessSh, SC, payload.sessionRow);
    if (!sessionInfo || !sessionInfo.isActive) {
      return { ok: false, error: 'Session is not active or not found.' };
    }
    if (sessionInfo.type !== 'group') {
      return {
        ok: true,
        results: studentsRaw.map((student) => {
          const res = recordStudentSignIn({ sessionId, sessionRow: sessionInfo.rowIndex, student });
          return {
            studentId: String(student && (student.studentId || student.id) || '').trim(),
            firstName: String(student && student.firstName || '').trim(),
            lastName : String(student && student.lastName  || '').trim(),
            ok: !!(res && res.ok),
            error: res && !res.ok ? res.error : undefined
          };
        })
      };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(CONFIG.ATTENDANCE.SHEET);
    if (!sh) throw new Error('Attendance sheet not found.');
    const width = Math.max(1, sh.getLastColumn());
    const A = CONFIG.ATTENDANCE.COLS || {};
    const lastRow = sh.getLastRow();

    let formulaTemplate = null;
    let formulaStartCol = null;
    let formulaSpan = 0;
    const formulaCols = [A.inDb, A.consent, A.pre, A.post].filter(c => c != null);
    if (formulaCols.length) {
      formulaStartCol = Math.min.apply(null, formulaCols);
      const formulaMaxCol = Math.max.apply(null, formulaCols);
      formulaSpan = formulaMaxCol - formulaStartCol + 1;
      if (lastRow >= 2) {
        const tmpl = sh.getRange(2, formulaStartCol + 1, 1, formulaSpan).getFormulas();
        if (tmpl && tmpl.length) {
          const rowTemplate = tmpl[0] || [];
          if (rowTemplate.some(f => f)) {
            formulaTemplate = rowTemplate;
          }
        }
      }
    }

    const rows = [];
    results.forEach((result, index) => {
      const student = studentsRaw[index] || {};
      if (!result.studentId) {
        result.error = 'Student ID is required.';
        return;
      }
      const details = {
        studentId: result.studentId,
        firstName: result.firstName,
        lastName : result.lastName,
        school   : String(student.school || '').trim(),
        email    : String(student.email || '').trim(),
        grade    : String(student.grade || '').trim()
      };
      if (details.grade) {
        details.grade = signinNormalizeGradeLabel_(details.grade);
      }
      const row = new Array(width).fill('');
      if (A.timestamp != null) row[A.timestamp] = now;
      if (A.firstName != null) row[A.firstName] = details.firstName || '';
      if (A.lastName != null) row[A.lastName] = details.lastName || '';
      if (A.schoolYear != null) row[A.schoolYear] = details.grade || '';
      if (A.school != null) row[A.school] = details.school || '';
      if (A.idNumber != null) row[A.idNumber] = details.studentId || '';
      if (A.group != null) row[A.group] = sessionInfo.label || '';
      if (formulaTemplate && formulaStartCol != null && formulaSpan > 0) {
        for (let offset = 0; offset < formulaSpan; offset++) {
          const formula = formulaTemplate[offset];
          if (formula) row[formulaStartCol + offset] = formula;
        }
      }
      rows.push({ row, details, index });
      result.ok = true;
      result.firstName = details.firstName;
      result.lastName = details.lastName;
    });

    if (rows.length) {
      sh.getRange(lastRow + 1, 1, rows.length, width).setValues(rows.map(entry => entry.row));
      if (SC.LastSignInAt != null) {
        sessSh.getRange(sessionInfo.rowIndex, SC.LastSignInAt + 1).setValue(now);
      }
      if (SC.SignInCount != null) {
        const cell = sessSh.getRange(sessionInfo.rowIndex, SC.SignInCount + 1);
        const current = Number(cell.getValue()) || 0;
        cell.setValue(current + rows.length);
      }
    }

    rows.forEach(({ details, index }) => {
      try {
        signinEnqueueKnownStudentUpdate_(details, now, { updateLastSignIn: true });
      } catch (err) {
        console.error('Queueing known-student update failed; falling back to direct write.', err);
        try {
          const fallback = signinUpsertKnownStudent(details, { timestamp: now });
          results[index].createdKnown = !!(fallback && fallback.created);
        } catch (err2) {
          results[index].ok = false;
          results[index].error = err2 && err2.message ? err2.message : String(err2);
        }
      }
    });

    targetSheet = CONFIG && CONFIG.ATTENDANCE ? CONFIG.ATTENDANCE.SHEET : '';
  } catch (err) {
    return { ok: false, error: err && err.message ? err.message : String(err) };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }

  return {
    ok: true,
    session: sessionInfo ? { id: sessionInfo.id, label: sessionInfo.label, date: sessionInfo.date, type: sessionInfo.type } : null,
    results,
    targetSheet
  };
}

function listKnownStudentsLite() {
  const index = signinFetchSuggestIndex_();
  if (!Array.isArray(index) || !index.length) return [];
  return index.map((entry) => ({
    id: entry.id,
    firstName: entry.firstName,
    lastName: entry.lastName,
    school: entry.school,
    email: entry.email,
    grade: entry.grade,
    source: entry.source || '',
    label: entry.label,
  }));
}
