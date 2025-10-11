/*** GroupContactSessions.gs
 * Creates/updates group note sessions with unique ContactID
 * Schema:
 *   group_contact_sessions
 *     - ContactID (UUID)
 *     - Date
 *     - Group
 *     - Topic
 *     - Summary
 *     - DurationMinutes
 *     - CreatedAt
 *     - EditedAt
 */

// ------------------ Ensure sheet exists ------------------
function ensureGroupContactSessionsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('group_contact_sessions');
  if (!sh) sh = ss.insertSheet('group_contact_sessions');

  // Seed headers
  if (sh.getLastRow() < 1 || sh.getLastColumn() < 1) {
    sh.clear();
    sh.getRange(1,1,1,8).setValues([[
      'ContactID','Date','Group','Topic','Summary','DurationMinutes','CreatedAt','EditedAt'
    ]]);
  }

  // Build header map (normalized keys)
  const header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const norm = s => String(s||'').toLowerCase().replace(/[^a-z0-9]+/g,'');
  const idxMap = new Map(header.map((h,i)=>[norm(h), i]));
  const C = {
    ContactID: idxMap.get('contactid'),
    Date: idxMap.get('date'),
    Group: idxMap.get('group'),
    Topic: idxMap.get('topic'),
    Summary: idxMap.get('summary'),
    Duration: idxMap.get('durationminutes'),
    CreatedAt: idxMap.get('createdat'),
    EditedAt: idxMap.get('editedat')
  };
  return { sh, C };
}

function _groupSessionCacheKey_(dateYMD, groupName) {
  return `GROUP_SESSION_V1_${dateYMD}_${String(groupName || '').toUpperCase()}`;
}

// ------------------ Core API ------------------
/**
 * Create or update a group session (returns a ContactID)
 * @param {string} dateStr
 * @param {string} groupName
 * @param {string} topic
 * @param {string} summary
 * @param {number} durationMinutes
 * @returns {Object} { ok, contactId, created, updated }
 */
function createOrUpdateGroupContactSession_(dateStr, groupName, topic, summary, durationMinutes) {
  const tz = Session.getScriptTimeZone() || 'America/Chicago';
  const date = ymd_(dateStr, tz);
  const group = String(groupName || '').trim();
  if (!date) return { ok:false, error:'Missing/invalid date.' };
  if (!group) return { ok:false, error:'Missing group.' };

  const { sh, C } = ensureGroupContactSessionsSheet_();
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const cacheKey = _groupSessionCacheKey_(date, group);
  const cached = _scriptCacheGet_(cacheKey);

  let rowIdx = null;
  let contactId = null;
  if (cached && typeof cached.row === 'number') {
    const candidate = cached.row;
    if (candidate >= 2 && candidate <= lastRow) {
      const rowValues = sh.getRange(candidate, 1, 1, lastCol).getValues()[0];
      const rowDate = ymd_(rowValues[C.Date], tz) || rowValues[C.Date];
      const rowGroup = String(rowValues[C.Group] || '').trim();
      if (rowDate === date && rowGroup === group) {
        rowIdx = candidate;
        contactId = String(rowValues[C.ContactID] || '').trim() || null;
      }
    }
  }

  if (rowIdx == null && lastRow >= 2) {
    const data = sh.getRange(2,1,lastRow-1,lastCol).getValues();
    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      const d = ymd_(r[C.Date], tz);
      const g = String(r[C.Group] || '').trim();
      if (d === date && g === group) { rowIdx = i + 2; contactId = String(r[C.ContactID] || '').trim() || null; break; }
    }
  }

  const now = new Date();
  if (rowIdx) {
    if (!contactId) {
      contactId = Utilities.getUuid();
      sh.getRange(rowIdx, C.ContactID + 1).setValue(contactId);
    }
    sh.getRange(rowIdx, C.Topic + 1).setValue(topic || '');
    sh.getRange(rowIdx, C.Summary + 1).setValue(summary || '');
    sh.getRange(rowIdx, C.Duration + 1).setValue(Number(durationMinutes) || 0);
    sh.getRange(rowIdx, C.EditedAt + 1).setValue(now);
    _scriptCachePut_(cacheKey, { row: rowIdx, contactId }, 300);
    return { ok:true, contactId, updated:true };
  } else {
    const contactId = Utilities.getUuid();
    sh.appendRow([contactId, date, group, topic || '', summary || '', Number(durationMinutes) || 0, now, now]);
    const newRowIdx = sh.getLastRow();
    _scriptCachePut_(cacheKey, { row: newRowIdx, contactId }, 300);
    return { ok:true, contactId, created:true };
  }
}


/**
 * Get the ContactID for a given date/group (if exists)
 */
function getContactIdForGroup_(dateStr, groupName) {
  const tz = Session.getScriptTimeZone() || 'America/Chicago';
  const date = ymd_(dateStr, tz);
  const group = String(groupName || '').trim();
  if (!date || !group) return null;

  const { sh, C } = ensureGroupContactSessionsSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;

  const cacheKey = _groupSessionCacheKey_(date, group);
  const cached = _scriptCacheGet_(cacheKey);
  if (cached && cached.contactId) return cached.contactId || null;

  const vals = sh.getRange(2,1,lastRow-1,sh.getLastColumn()).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    const r = vals[i];
    const d = ymd_(r[C.Date], tz);
    const g = String(r[C.Group] || '').trim();
    if (d === date && g === group) {
      const contactId = String(r[C.ContactID] || '').trim() || null;
      if (contactId) {
        _scriptCachePut_(cacheKey, { row: i + 2, contactId }, 300);
      }
      return contactId;
    }
  }
  return null;
}
