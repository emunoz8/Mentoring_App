
/* ---------------- Notes.gs---------------- */
const GROUP_PREFILL_CACHE = globalThis.GROUP_PREFILL_CACHE || (globalThis.GROUP_PREFILL_CACHE = {});
function ensureGroupNotesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('group_notes');
  if (!sh) sh = ss.insertSheet('group_notes');

  // Seed headers if empty/blank
  if (sh.getLastRow() < 1 || sh.getLastColumn() < 1) {
    sh.clear();
    sh.getRange(1,1,1,9).setValues([[
      'Date','Group','Topic','Summary','DurationMinutes','ID','FirstName','LastName','LastEdited'
    ]]);
  } else {
    const lc0 = Math.max(1, sh.getLastColumn());
    const firstRow = sh.getRange(1,1,1,lc0).getValues()[0];
    const allBlank = firstRow.every(v => String(v||'').trim()==='');
    if (allBlank) {
      sh.getRange(1,1,1,9).setValues([[
        'Date','Group','Topic','Summary','DurationMinutes','ID','FirstName','LastName','LastEdited'
      ]]);
    }
  }

  // Build tolerant header index
  const lc = Math.max(1, sh.getLastColumn());
  const header = sh.getRange(1,1,1,lc).getValues()[0].map(h => String(h||'').trim());
  const norm = s => s.toLowerCase().replace(/[^a-z0-9]+/g, '');
  const idxMap = new Map(header.map((h,i)=>[norm(h), i]));

  const idx = {
    Date:     idxMap.get('date'),
    Group:    idxMap.get('group'),
    Topic:    idxMap.get('topic') ?? idxMap.get('subject'),
    Summary:  idxMap.get('summary') ?? idxMap.get('note') ?? idxMap.get('notes') ?? idxMap.get('description'),
    Duration: idxMap.get('durationminutes') ?? idxMap.get('duration') ?? idxMap.get('minutes') ?? idxMap.get('mins'),
    ID:       idxMap.get('id') ?? idxMap.get('studentid') ?? idxMap.get('cpsid') ?? idxMap.get('participantid'),
    First:    idxMap.get('firstname') ?? idxMap.get('first') ?? idxMap.get('fname') ?? idxMap.get('givenname'),
    Last:     idxMap.get('lastname') ?? idxMap.get('last') ?? idxMap.get('lname') ?? idxMap.get('surname') ?? idxMap.get('familyname'),
    Edited:   idxMap.get('lastedited') ?? idxMap.get('updated') ?? idxMap.get('modified') ?? idxMap.get('lastupdate'),
  };

  // Add any missing columns to the right
  const want = [
    ['Date','Date'],
    ['Group','Group'],
    ['Topic','Topic'],
    ['Summary','Summary'],
    ['Duration','DurationMinutes'],
    ['ID','ID'],
    ['First','FirstName'],
    ['Last','LastName'],
    ['Edited','LastEdited'],
  ];
  const missing = want.filter(([k]) => idx[k] == null);
  if (missing.length) {
    const startCol = sh.getLastColumn() + 1;
    sh.getRange(1, startCol, 1, missing.length).setValues([missing.map(x => x[1])]);

    const header2 = sh.getRange(1,1,1,Math.max(1, sh.getLastColumn())).getValues()[0].map(h => String(h||'').trim());
    const idxMap2 = new Map(header2.map((h,i)=>[norm(h), i]));
    idx.Date     = idx.Date     ?? idxMap2.get('date');
    idx.Group    = idx.Group    ?? idxMap2.get('group');
    idx.Topic    = idx.Topic    ?? idxMap2.get('topic');
    idx.Summary  = idx.Summary  ?? (idxMap2.get('summary') ?? idxMap2.get('note') ?? idxMap2.get('notes') ?? idxMap2.get('description'));
    idx.Duration = idx.Duration ?? (idxMap2.get('durationminutes') ?? idxMap2.get('duration') ?? idxMap2.get('minutes') ?? idxMap2.get('mins'));
    idx.ID       = idx.ID       ?? (idxMap2.get('id') ?? idxMap2.get('studentid') ?? idxMap2.get('cpsid') ?? idxMap2.get('participantid'));
    idx.First    = idx.First    ?? (idxMap2.get('firstname') ?? idxMap2.get('first') ?? idxMap2.get('fname') ?? idxMap2.get('givenname'));
    idx.Last     = idx.Last     ?? (idxMap2.get('lastname') ?? idxMap2.get('last') ?? idxMap2.get('lname') ?? idxMap2.get('surname') ?? idxMap2.get('familyname'));
    idx.Edited   = idx.Edited   ?? (idxMap2.get('lastedited') ?? idxMap2.get('updated') ?? idxMap2.get('modified') ?? idxMap2.get('lastupdate'));
  }

  return { sh, C: idx };
}

/**
 * Save or update a Group Note session (one row per ContactID)
 * Replaces the old per-participant logic.
 */
function saveGroupNote(dateStr, groupName, topic, summary, durationMinutes, participants) {
  const tz = Session.getScriptTimeZone() || 'America/Chicago';
  const date = ymd_(dateStr, tz);
  const group = String(groupName || '').trim();
  if (!date) return { ok:false, error:'Missing/invalid date.' };
  if (!group) return { ok:false, error:'Missing group.' };

  // 1️⃣ Ensure the group_contact_sessions sheet exists
  const { sh, C } = ensureGroupContactSessionsSheet_();

  // 2️⃣ Check if session exists for this date + group
  const lastRow = sh.getLastRow();
  const lastCol = Math.max(1, sh.getLastColumn());
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
      if (d === date && g === group) {
        rowIdx = i + 2;
        contactId = String(r[C.ContactID] || '').trim() || null;
        break;
      }
    }
  }

  const now = new Date();
  const topicVal = String(topic || '').trim();
  const summaryVal = String(summary || '').trim();
  const durVal = Number(durationMinutes) || 0;

  // 3️⃣ If existing session found → update
  if (rowIdx) {
    if (!contactId) {
      contactId = Utilities.getUuid();
      sh.getRange(rowIdx, C.ContactID + 1).setValue(contactId);
    }
    sh.getRange(rowIdx, C.Topic + 1).setValue(topicVal);
    sh.getRange(rowIdx, C.Summary + 1).setValue(summaryVal);
    sh.getRange(rowIdx, C.Duration + 1).setValue(durVal);
    sh.getRange(rowIdx, C.EditedAt + 1).setValue(now);
    _scriptCachePut_(cacheKey, { row: rowIdx, contactId }, 300);
    return { ok:true, contactId, updated:true };
  }

  // 4️⃣ Otherwise, create new session
  contactId = Utilities.getUuid();
  const newRow = [
    contactId,
    date,
    group,
    topicVal,
    summaryVal,
    durVal,
    now,
    now
  ];
  sh.appendRow(newRow);
  const newRowIdx = sh.getLastRow();
  _scriptCachePut_(cacheKey, { row: newRowIdx, contactId }, 300);

  // 5️⃣ Optionally link participants (future feature)
  // if (Array.isArray(participants) && participants.length) {
  //   saveGroupParticipants(contactId, participants);
  // }

  return { ok:true, contactId, created:true };
}


function getLatestGroupContactSession(dateStr, groupName) {
  const tz = Session.getScriptTimeZone() || 'America/Chicago';
  const d = ymd_(dateStr, tz);
  const g = String(groupName||'').trim();

  const { sh, C } = ensureGroupContactSessionsSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok:true, note:null };

  const vals = sh.getRange(2,1,lastRow-1,sh.getLastColumn()).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    const r = vals[i];
    const rd = ymd_(r[C.Date], tz);
    const rg = String(r[C.Group] || '').trim();
    if (rd === d && rg === g) {
      return {
        ok:true,
        note:{
          topic:String(r[C.Topic]||'').trim(),
          summary:String(r[C.Summary]||'').trim(),
          duration:Number(r[C.Duration]||0)
        }
      };
    }
  }
  return { ok:true, note:null };
}


function saveFullGroupNote(dateStr, groupName, topic, summary, durationMinutes, participants, mentors) {
  const logTag = `[GroupNote ${dateStr}|${groupName}]`;
  const t0 = Date.now();
  const session = saveGroupNote(dateStr, groupName, topic, summary, durationMinutes);
  const t1 = Date.now();
  console.log(`${logTag} saveGroupNote ${t1 - t0}ms (ok=${session.ok})`);
  if (!session.ok) return session;
  const contactId = session.contactId;

  const partRes = saveGroupParticipants(contactId, participants);
  const t2 = Date.now();
  console.log(`${logTag} saveGroupParticipants ${t2 - t1}ms (added=${partRes.added || 0})`);

  const mentRes = saveGroupMentors(contactId, mentors);
  const t3 = Date.now();
  console.log(`${logTag} saveGroupMentors ${t3 - t2}ms (added=${mentRes.added || 0})`);

  try {
    if (typeof _scriptCacheRemove_ === 'function') {
      const tz = Session.getScriptTimeZone() || 'America/Chicago';
      const dateKey = ymd_(dateStr, tz);
      if (dateKey) {
        _scriptCacheRemove_(`GROUP_PREFILL_V1_${dateKey}`);
        delete GROUP_PREFILL_CACHE[dateKey];
      }
    }
  } catch (_) {}

  console.log(`${logTag} total ${t3 - t0}ms`);

  return {
    ok: session.ok && partRes.ok && mentRes.ok,
    contactId,
    created: session.created,
    participantsAdded: partRes.added || 0,
    mentorsAdded: mentRes.added || 0
  };
}

/**
 * Bulk-prefill helper for Group Notes UI.
 * Returns the latest note + mentor chips for each requested group on a date.
 *
 * @param {string} dateStr
 * @param {Array<string>} groupNames
 * @return {{ ok: boolean, groups: Object<string,{note: Object|null, mentors: Array, contactId: string|null}> }}
 */
function getGroupPrefill(dateStr, groupNames) {
  const tz = Session.getScriptTimeZone() || 'America/Chicago';
  const targetGroups = Array.isArray(groupNames) ? groupNames : [];
  if (!targetGroups.length) return { ok: true, groups: {} };

  const normalizedMap = new Map();
  const result = {};
  targetGroups.forEach(raw => {
    const g = String(raw || '').trim();
    if (!g) return;
    const key = g.toUpperCase();
    if (!normalizedMap.has(key)) {
      normalizedMap.set(key, g);
      result[g] = { note: null, mentors: [], contactId: null };
    }
  });
  if (!normalizedMap.size) return { ok: true, groups: result };

  const dateYMD = ymd_(dateStr, tz);
  if (!dateYMD) return { ok: true, groups: result };

  const cacheKey = 'GROUP_PREFILL_V1_' + dateYMD;
  let fullMap = GROUP_PREFILL_CACHE[dateYMD];
  if (!fullMap) {
    const cached = _scriptCacheGet_(cacheKey);
    if (cached && cached.groups) {
      fullMap = cached.groups;
    }
  }

  if (!fullMap) {
    fullMap = {};
    const contactKeyByGroup = new Map();

    const { sh: sessSh, C: Sess } = ensureGroupContactSessionsSheet_();
    const lastRow = sessSh.getLastRow();
    if (lastRow >= 2) {
      const data = sessSh.getRange(2, 1, lastRow - 1, sessSh.getLastColumn()).getValues();
      data.forEach(r => {
        const rowDate = ymd_(r[Sess.Date], tz) || r[Sess.Date];
        if (rowDate !== dateYMD) return;
        const groupName = String(r[Sess.Group] || '').trim();
        if (!groupName) return;
        if (fullMap[groupName]) return; // keep first match
        const contactId = String(r[Sess.ContactID] || '').trim();
        contactKeyByGroup.set(groupName, contactId ? contactId.toUpperCase() : '');
        fullMap[groupName] = {
          note: {
            topic: String(r[Sess.Topic] || '').trim(),
            summary: String(r[Sess.Summary] || '').trim(),
            duration: Number(r[Sess.Duration] || 0)
          },
          mentors: [],
          contactId: contactId || null
        };
      });
    }

    const { sh: mentorSh, C: MentC } = ensureGroupContactMentorsSheet_();
    const mentorRowCount = mentorSh.getLastRow();
    if (mentorRowCount >= 2 && MentC.ContactID != null && MentC.MentorID != null && MentC.Name != null) {
      const mentorValues = mentorSh.getRange(2, 1, mentorRowCount - 1, mentorSh.getLastColumn()).getValues();
      const mentorsByContact = new Map();
      mentorValues.forEach(r => {
        const cid = String(r[MentC.ContactID] || '').trim().toUpperCase();
        if (!cid) return;
        const id = String(r[MentC.MentorID] || '').trim().toUpperCase();
        if (!id) return;
        const name = String(r[MentC.Name] || '').trim() || id;
        if (!mentorsByContact.has(cid)) mentorsByContact.set(cid, []);
        mentorsByContact.get(cid).push({ id, name });
      });

      Object.keys(fullMap).forEach(groupName => {
        const cidKey = contactKeyByGroup.get(groupName) || '';
        if (cidKey && mentorsByContact.has(cidKey)) {
          fullMap[groupName].mentors = mentorsByContact.get(cidKey).slice();
        }
      });
    }

    _scriptCachePut_(cacheKey, { groups: fullMap }, 300);
    GROUP_PREFILL_CACHE[dateYMD] = fullMap;
  }

  Object.entries(result).forEach(([groupName, entry]) => {
    const source = fullMap[groupName];
    if (source) {
      entry.note = source.note;
      entry.mentors = Array.isArray(source.mentors) ? source.mentors.slice() : [];
      entry.contactId = source.contactId || null;
    }
  });

  return { ok: true, groups: result };
}
