
/* ---------------- Notes.gs---------------- */
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
  const data = lastRow >= 2 ? sh.getRange(2,1,lastRow-1,lastCol).getValues() : [];

  let rowIdx = null;
  let contactId = null;
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
  const session = saveGroupNote(dateStr, groupName, topic, summary, durationMinutes);
  if (!session.ok) return session;
  const contactId = session.contactId;

  const partRes = saveGroupParticipants(contactId, participants);
  const mentRes = saveGroupMentors(contactId, mentors);

  return {
    ok: session.ok && partRes.ok && mentRes.ok,
    contactId,
    created: session.created,
    participantsAdded: partRes.added || 0,
    mentorsAdded: mentRes.added || 0
  };
}



