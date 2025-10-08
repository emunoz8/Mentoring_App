/*** QueueServer.gs
 * Canonical queue APIs:
 *   - listQueue(dateStr)
 *   - claimRows(rowIndices, claimedBy?)
 *   - markProcessed(rowIndices, contactId)
 *   - getWebAppBaseUrl()
 *   - getNamesForIds(ids)
 * Internal helper:
 *   - queue_ensureSignInLogSheet_()
 * Depends on Utils.gs (ymd_, parseLooseDate_)
 */

const SIGN_IN = { SHEET: 'sign_in_log' };
const STATUS  = { PENDING: 'Pending', CLAIMED: 'Claimed', PROCESSED: 'Processed' };

// normalize header key
function _normKey_(s){ return String(s||'').toLowerCase().replace(/[^a-z0-9]+/g,''); }

/** Ensure sheet + tolerant header map (seed status columns if missing) */
function queue_ensureSignInLogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SIGN_IN.SHEET);
  if (!sh) sh = ss.insertSheet(SIGN_IN.SHEET);

  // Seed base columns if empty/blank
  if (sh.getLastRow() < 1 || sh.getLastColumn() < 1) {
    sh.clear();
    sh.getRange(1,1,1,10).setValues([[
      'Timestamp','ID number','First Name + Last Name','School','Mentor',
      'Status','ClaimedBy','ClaimedAt','ProcessedAt','ContactID'
    ]]);
  } else {
    const lc0 = Math.max(1, sh.getLastColumn());
    const first = sh.getRange(1,1,1,lc0).getValues()[0];
    if (first.every(v => String(v||'').trim()==='')) {
      sh.getRange(1,1,1,10).setValues([[
        'Timestamp','ID number','First Name + Last Name','School','Mentor',
        'Status','ClaimedBy','ClaimedAt','ProcessedAt','ContactID'
      ]]);
    }
  }

  const lc = Math.max(1, sh.getLastColumn());
  const header = sh.getRange(1,1,1,lc).getValues()[0].map(h => String(h||'').trim());
  const idx = new Map(header.map((h,i)=>[_normKey_(h), i]));

  const C = {
    Timestamp  : idx.get('timestamp') ?? idx.get('date') ?? idx.get('signindate'),
    ID         : idx.get('idnumber')  ?? idx.get('id') ?? idx.get('studentid') ?? idx.get('cpsid'),
    Name       : idx.get('firstnamelastname') ?? idx.get('name') ?? idx.get('fullname') ??
                 idx.get('studentname') ?? idx.get('firstlast') ?? idx.get('fullnamestudent'),
    School     : idx.get('school')    ?? idx.get('site'),
    MentorRaw  : idx.get('mentor')    ?? idx.get('mentorid') ?? idx.get('staff') ?? idx.get('advisor'),
    Status     : idx.get('status'),
    ClaimedBy  : idx.get('claimedby'),
    ClaimedAt  : idx.get('claimedat'),
    ProcessedAt: idx.get('processedat'),
    ContactID  : idx.get('contactid'),
  };

  // Add any missing admin columns to the RIGHT so we can write to them
  const needed = [
    ['Status','Status'],
    ['ClaimedBy','ClaimedBy'],
    ['ClaimedAt','ClaimedAt'],
    ['ProcessedAt','ProcessedAt'],
    ['ContactID','ContactID']
  ].filter(([k]) => C[k] == null);

  if (needed.length) {
    const start = sh.getLastColumn() + 1;
    sh.getRange(1, start, 1, needed.length).setValues([needed.map(x=>x[1])]);

    // Rebuild index/mapping
    const header2 = sh.getRange(1,1,1,Math.max(1, sh.getLastColumn())).getValues()[0].map(h => String(h||'').trim());
    const idx2 = new Map(header2.map((h,i)=>[_normKey_(h), i]));
    const resolve = (cur, ...keys) => C[cur] ?? keys.map(k=>idx2.get(k)).find(x=>x!=null);

    C.Status      = resolve('Status','status');
    C.ClaimedBy   = resolve('ClaimedBy','claimedby');
    C.ClaimedAt   = resolve('ClaimedAt','claimedat');
    C.ProcessedAt = resolve('ProcessedAt','processedat');
    C.ContactID   = resolve('ContactID','contactid');
  }

  return { sh, C };
}

/** Load queue for a given date (yyyy-MM-dd or anything parseLooseDate_ can handle) */
function listQueue(dateStr) {
  const tz = Session.getScriptTimeZone() || 'America/Chicago';
  const wantDay = ymd_(parseLooseDate_(dateStr) || new Date(), tz); // yyyy-MM-dd

  const { sh, C } = queue_ensureSignInLogSheet_();
  if (sh.getLastRow() < 2) return { items: [] };

  const width = Math.max(1, sh.getLastColumn());
  const vals = sh.getRange(2,1, sh.getLastRow()-1, width).getValues();

  const me = (Session.getActiveUser && Session.getActiveUser().getEmail && Session.getActiveUser().getEmail()) || '';

  const items = [];
  const idsNeedingName = new Set();

  for (let i=0; i<vals.length; i++) {
    const r = vals[i];

    const ts = (C.Timestamp != null) ? r[C.Timestamp] : null;
    const d  = parseLooseDate_(ts);
    const rowDay = d ? ymd_(d, tz) : '';
    if (rowDay !== wantDay) continue;

    const id     = String(C.ID     != null ? r[C.ID]     : '').trim();
    if (!id) continue;

    const name   = String(C.Name   != null ? r[C.Name]   : '').trim();
    const school = String(C.School != null ? r[C.School] : '').trim();

    const mentorRaw = String(C.MentorRaw != null ? r[C.MentorRaw] : '').trim();
    let mentorId = '', mentorName = '';
    if (mentorRaw) {
      if (/^[A-Za-z]{1,}\s+[A-Za-z].*/.test(mentorRaw)) mentorName = mentorRaw;
      else mentorId = mentorRaw.toUpperCase();
    }

    // --- existing lines ---
    // Normalize status
    const rawStatus = String(C.Status != null ? r[C.Status] : '').trim();
    const lowered = rawStatus.toLowerCase();

    // NEW: read ProcessedAt + ContactID
    const processedAtVal = (C.ProcessedAt != null) ? r[C.ProcessedAt] : null;
    const hasProcessed   = !!processedAtVal;
    const contactIdCell  = (C.ContactID != null) ? String(r[C.ContactID] || '').trim() : '';
    const hasContactId   = contactIdCell.length > 0;

    // Compute final status
    let status;
    if (hasProcessed) {
      status = STATUS.PROCESSED;
    } else if (hasContactId) {
      status = STATUS.CLAIMED;
    } else {
      status =
        lowered === 'claimed'   ? STATUS.CLAIMED :
        lowered === 'processed' ? STATUS.PROCESSED :
        STATUS.PENDING;
    }

    const claimedBy = String(C.ClaimedBy != null ? r[C.ClaimedBy] : '').trim();
    
    let displayName = name || '';
    if (!displayName && id) idsNeedingName.add(id);
    
    items.push({
      rowIndex: i + 2,
      id,
      displayName,
      school,
      mentorId,
      mentorName,
      status,
      claimedBy,
      contactId: contactIdCell, // <-- add this line
      mine: (!!me && claimedBy && claimedBy.toLowerCase() === me.toLowerCase())
    });




  }

  // Fill missing names from roster map once
  if (idsNeedingName.size) {
    let nameMap;
    try { nameMap = getNamesForIds_ ? getNamesForIds_(Array.from(idsNeedingName)) : new Map(); }
    catch (_) { nameMap = new Map(); }
    items.forEach(it => { if (!it.displayName) it.displayName = nameMap.get(it.id) || it.id; });
  } else {
    items.forEach(it => { if (!it.displayName) it.displayName = it.id; });
  }

  // convenience aliases for older UI bits
  items.forEach(it => {
    it.name   = it.displayName;
    it.mentor = it.mentorName || it.mentorId;
  });

  return { items };
}

/** Claim rows (by 1-based sheet row indices). Writes canonical `Claimed`. */
function claimRows(rowIndices, claimedByRaw) {
  const { sh, C } = queue_ensureSignInLogSheet_();
  const now = new Date();
  const who =
    String(claimedByRaw || '').trim() ||
    (Session.getEffectiveUser && Session.getEffectiveUser().getEmail && Session.getEffectiveUser().getEmail()) ||
    'unknown';

  const rows = (Array.isArray(rowIndices) ? rowIndices : [])
    .map(n => Number(n))
    .filter(n => Number.isFinite(n) && n >= 2);

  if (!rows.length) return { ok:false, error:'No rows provided.' };

  const failed = [];
  const lock = LockService.getDocumentLock();
  try {
    lock.waitLock(30000);
    rows.forEach(rn => {
      // If already processed, don't overwrite
      const cur = String(sh.getRange(rn, C.Status + 1).getValue() || '').trim().toLowerCase();
      if (cur === 'processed' || cur === STATUS.PROCESSED.toLowerCase()) { failed.push(rn); return; }

      sh.getRange(rn, C.Status + 1).setValue(STATUS.CLAIMED);

      // Optional: keep ClaimedBy/ClaimedAt updated if those columns exist
      if (C.ClaimedBy != null) sh.getRange(rn, C.ClaimedBy + 1).setValue(who);
      if (C.ClaimedAt != null) sh.getRange(rn, C.ClaimedAt + 1).setValue(now);
    });
  } catch (e) {
    return { ok:false, error: e && e.message ? e.message : String(e) };
  } finally {
    try { lock.releaseLock(); } catch(_) {}
  }

  return { ok:true, failed };
}

/** Mark rows processed after note save (also writes ProcessedAt + ContactID) */
function markProcessed(rowIndices, contactId) {
  const { sh, C } = queue_ensureSignInLogSheet_();
  const rows = (Array.isArray(rowIndices) ? rowIndices : []).map(n => Number(n)||0).filter(n=>n>=2);
  if (!rows.length) return { ok:false, error:'No rows to mark processed.' };

  const lock = LockService.getDocumentLock();
  const now = new Date();

  try {
    lock.waitLock(30000);
    rows.forEach(rn=>{
      if (rn < 2 || rn > sh.getLastRow()) return;

      // Status -> Processed
      sh.getRange(rn, C.Status + 1).setValue(STATUS.PROCESSED);

      // ProcessedAt -> now
      if (C.ProcessedAt != null) sh.getRange(rn, C.ProcessedAt + 1).setValue(now);

      // ContactID -> saved contact id
      if (C.ContactID != null) sh.getRange(rn, C.ContactID + 1).setValue(String(contactId || '').trim());
    });
    return { ok:true, rows: rows.length, contactId: contactId||'' };
  } catch (e) {
    return { ok:false, error: e && e.message ? e.message : String(e) };
  } finally {
    try { lock.releaseLock(); } catch(_) {}
  }
}

/** Used by client to build a reliable /exec URL for redirects */
function getWebAppBaseUrl() {
  try { return ScriptApp.getService().getUrl(); }
  catch (e) { return ''; }
}

/** Public wrapper so client can fetch names for a list of IDs (reuses the impl from IndividualContacts.gs if present) */
function getNamesForIds(ids) {
  const list = Array.isArray(ids) ? ids : [];
  let map;
  try { map = getNamesForIds_ ? getNamesForIds_(list) : new Map(); } catch (_) { map = new Map(); }
  return Array.from(map.entries()).map(([id, name]) => ({ id, name }));
}

/** Compatibility shim for older clients: getSignInsByDate(dateStr) */
function getSignInsByDate(dateStr) {
  const { items } = listQueue(dateStr);
  return items.map(it => ({
    id: it.id,
    name: it.displayName,
    school: it.school || '',
    mentor: it.mentorName || it.mentorId || '',
    rowIndex: it.rowIndex || null,
    status: it.status
  }));
}

/** Claim/attach note for a set of IDs on a given date.
 *  - dateStr: anything parseLooseDate_ can handle; normalized to yyyy-MM-dd
 *  - ids: array of student IDs (strings/numbers)
 *  - contactId: the saved Individual Note ContactID (string)
 *  - statusOpt: optional; defaults to STATUS.CLAIMED. If STATUS.PROCESSED, sets ProcessedAt=now.
 *
 * Returns: { ok:true, matched:<count>, rows:[rowIndex,...] } or { ok:false, error:"..." }
 */
/** Mark processed by student IDs for a given date (stamps ProcessedAt + ContactID). */
function markProcessedByIds(dateStr, ids, contactId, statusOpt) {
  const tz = Session.getScriptTimeZone() || 'America/Chicago';
  const wantDay = ymd_(parseLooseDate_(dateStr) || new Date(), tz); // yyyy-MM-dd
  const idSet = new Set((Array.isArray(ids) ? ids : [])
    .map(x => String(x||'').trim())
    .filter(Boolean)
    .map(s => s.toUpperCase()));

  const newStatus = String(statusOpt || STATUS.CLAIMED);

  if (!wantDay) return { ok:false, error:'Bad date.' };
  if (!idSet.size) return { ok:false, error:'No IDs provided.' };

  const { sh, C } = queue_ensureSignInLogSheet_();
  const lr = sh.getLastRow();
  if (lr < 2) return { ok:true, matched:0, rows:[] };

  const width = Math.max(1, sh.getLastColumn());
  const vals = sh.getRange(2, 1, lr-1, width).getValues();

  const now = new Date();
  const toUpdate = [];   // array of row indices (1-based in sheet)
  const newContactValues = new Map(); // rn -> string
  const newStatusValues  = new Map(); // rn -> string
  const newProcessedAt   = new Map(); // rn -> Date

  for (let i = 0; i < vals.length; i++) {
    const r = vals[i];
    const rn = i + 2; // sheet row

    // Date match
    const ts = (C.Timestamp != null) ? r[C.Timestamp] : null;
    const d  = parseLooseDate_(ts);
    const rowDay = d ? ymd_(d, tz) : '';
    if (rowDay !== wantDay) continue;

    // ID match
    const id = String(C.ID != null ? r[C.ID] : '').trim();
    if (!id || !idSet.has(id.toUpperCase())) continue;

    // ContactID append/unique
    let curContact = (C.ContactID != null) ? String(r[C.ContactID] || '').trim() : '';
    const parts = curContact ? curContact.split(/\s*,\s*/).filter(Boolean) : [];
    const has = parts.some(p => p === String(contactId || '').trim());
    const nextContact = (contactId && !has)
      ? (parts.concat([String(contactId)]).join(', '))
      : curContact;

    // Record updates
    toUpdate.push(rn);
    if (C.Status != null)      newStatusValues.set(rn, newStatus);
    if (C.ContactID != null)   newContactValues.set(rn, nextContact);

    // ✅ Always stamp ProcessedAt if a contactId was created
    if (C.ProcessedAt != null && contactId) {
      newProcessedAt.set(rn, now);
    }
  }

  if (!toUpdate.length) return { ok:true, matched:0, rows:[] };

  // Write updates
  const lock = LockService.getDocumentLock();
  try {
    lock.waitLock(30000);
    toUpdate.forEach(rn => {
      if (C.Status != null)      sh.getRange(rn, C.Status + 1).setValue(newStatusValues.get(rn) || newStatus);
      if (C.ContactID != null)   sh.getRange(rn, C.ContactID + 1).setValue(newContactValues.get(rn) || '');
      if (C.ProcessedAt != null && contactId) {
        sh.getRange(rn, C.ProcessedAt + 1).setValue(newProcessedAt.get(rn) || now);
      }
    });
    return { ok:true, matched: toUpdate.length, rows: toUpdate };
  } catch (e) {
    return { ok:false, error: e && e.message ? e.message : String(e) };
  } finally {
    try { lock.releaseLock(); } catch(_) {}
  }
}





