/*** IndividualContacts.gs
 * Sheets used:
 *   - individual_contact_sessions
 *   - individual_contact_participants
 * Depends on helpers from Utils.gs:
 *   - ymd_(v, tz)
 *   - toMDYString_(date)
 *
 * Note: saveIndividualContactSession(dateStr, people, payload, queueRowIndices?)
 *       The optional 4th arg can be an array of 1-based row indices from sign_in_log
 *       and will be passed to markProcessed(...) if available.
 */

// ---------- Sessions sheet ----------

// ---- Lightweight script cache helpers ----
function _scriptCacheGet_(k){
  try {
    const hit = CacheService.getScriptCache().get(k);
    return hit ? JSON.parse(hit) : null;
  } catch(_) { return null; }
}
function _scriptCachePut_(k, v, sec){
  try {
    CacheService.getScriptCache().put(k, JSON.stringify(v), sec);
  } catch(_) {}
}

function ensureContactSessionsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('individual_contact_sessions');
  if (!sh) sh = ss.insertSheet('individual_contact_sessions');

  // Seed EXACTLY 13 columns (fixed)
  if (sh.getLastRow() < 1 || sh.getLastColumn() < 1) {
    sh.clear();
    sh.getRange(1,1,1,13).setValues([[
      'ContactID','Date','DurationMinutes',
      'ContactWith','TypeOfContact','Topic','Success',
      'Notes','Referrals','Location','MentorID',
      'CreatedAt','EditedAt'
    ]]);
  }

  const lc = Math.max(1, sh.getLastColumn());
  const header = sh.getRange(1,1,1,lc).getValues()[0].map(h => String(h||'').trim());
  const norm = s => s.toLowerCase().replace(/[^a-z0-9]+/g,'');
  const idx = new Map(header.map((h,i)=>[norm(h), i]));

  const C = {
    ContactID : idx.get('contactid') ?? 0,
    Date      : idx.get('date'),
    Duration  : idx.get('durationminutes') ?? idx.get('minutes') ?? idx.get('duration'),
    With      : idx.get('contactwith') ?? idx.get('with'),
    Type      : idx.get('typeofcontact') ?? idx.get('channel') ?? idx.get('type'),
    Topic     : idx.get('topic') ?? idx.get('topicprimary'),
    Success   : idx.get('success') ?? idx.get('outcome'),
    Notes     : idx.get('notes') ?? idx.get('summary') ?? idx.get('description'),
    Referrals : idx.get('referrals') ?? idx.get('referralsmade'),
    Location  : idx.get('location') ?? idx.get('place'),
    MentorID  : idx.get('mentorid') ?? idx.get('mentor'),
    CreatedAt : idx.get('createdat') ?? idx.get('created'),
    EditedAt  : idx.get('editedat') ?? idx.get('lastedited') ?? idx.get('updated'),
  };

  // Add any missing columns to the right
  const want = [
    ['ContactID','ContactID'], ['Date','Date'], ['Duration','DurationMinutes'],
    ['With','ContactWith'], ['Type','TypeOfContact'], ['Topic','Topic'], ['Success','Success'],
    ['Notes','Notes'], ['Referrals','Referrals'], ['Location','Location'], ['MentorID','MentorID'],
    ['CreatedAt','CreatedAt'], ['EditedAt','EditedAt']
  ].filter(([k]) => C[k] == null);

  if (want.length) {
    const startCol = sh.getLastColumn() + 1;
    sh.getRange(1,startCol,1,want.length).setValues([want.map(x=>x[1])]);

    const header2 = sh.getRange(1,1,1,Math.max(1, sh.getLastColumn())).getValues()[0].map(h => String(h||'').trim());
    const idx2 = new Map(header2.map((h,i)=>[norm(h), i]));
    const resolve = (k, ...keys) => C[k] ?? (keys.map(x=>idx2.get(x)).find(x=>x!=null));
    C.ContactID = resolve('ContactID','contactid');
    C.Date      = resolve('Date','date');
    C.Duration  = resolve('Duration','durationminutes','minutes','duration');
    C.With      = resolve('With','contactwith','with');
    C.Type      = resolve('Type','typeofcontact','channel','type');
    C.Topic     = resolve('Topic','topic','topicprimary');
    C.Success   = resolve('Success','success','outcome');
    C.Notes     = resolve('Notes','notes','summary','description');
    C.Referrals = resolve('Referrals','referrals','referralsmade');
    C.Location  = resolve('Location','location','place');
    C.MentorID  = resolve('MentorID','mentorid','mentor');
    C.CreatedAt = resolve('CreatedAt','createdat','created');
    C.EditedAt  = resolve('EditedAt','editedat','lastedited','updated');
  }

  return { sh, C };
}

// ---------- Participants/link sheet ----------
function ensureContactParticipantsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('individual_contact_participants');
  if (!sh) sh = ss.insertSheet('individual_contact_participants');

  if (sh.getLastRow() < 1 || sh.getLastColumn() < 1) {
    sh.clear();
    sh.getRange(1,1,1,4).setValues([[
      'ContactID','StudentID','NotesStudent','CreatedAt'
    ]]);
  }

  const lc = Math.max(1, sh.getLastColumn());
  const header = sh.getRange(1,1,1,lc).getValues()[0].map(h => String(h||'').trim());
  const norm = s => s.toLowerCase().replace(/[^a-z0-9]+/g,'');
  const idx = new Map(header.map((h,i)=>[norm(h), i]));

  const C = {
    ContactID   : idx.get('contactid'),
    StudentID   : idx.get('studentid') ?? idx.get('id') ?? idx.get('cpsid') ?? idx.get('participantid'),
    NotesStudent: idx.get('notesstudent') ?? idx.get('studentnotes'),
    CreatedAt   : idx.get('createdat') ?? idx.get('created'),
  };

  // Add any missing columns to the right
  const want = [
    ['ContactID','ContactID'], ['StudentID','StudentID'],
    ['NotesStudent','NotesStudent'], ['CreatedAt','CreatedAt']
  ].filter(([k]) => C[k] == null);

  if (want.length) {
    const startCol = sh.getLastColumn() + 1;
    sh.getRange(1,startCol,1,want.length).setValues([want.map(x=>x[1])]);

    const header2 = sh.getRange(1,1,1,Math.max(1, sh.getLastColumn())).getValues()[0].map(h => String(h||'').trim());
    const idx2 = new Map(header2.map((h,i)=>[norm(h), i]));
    const resolve = (k, ...keys) => C[k] ?? (keys.map(x=>idx2.get(x)).find(x=>x!=null));
    C.ContactID    = resolve('ContactID','contactid');
    C.StudentID    = resolve('StudentID','studentid','id','cpsid','participantid');
    C.NotesStudent = resolve('NotesStudent','notesstudent','studentnotes');
    C.CreatedAt    = resolve('CreatedAt','createdat','created');
  }

  return { sh, C };
}

// ---------- Save one session + links (uses Mentors.gs for mentor list; QueueServer.gs for markProcessed) ----------
/**
 * Saves an Individual Contact session and its participant links.
 * If queueRowIndices are provided (or present as people[*].rowIndex), will call markProcessed()
 * to update sign_in_log (Status/ProcessedAt/ContactID).
 *
 * @param {string|Date} dateStr
 * @param {Array<{id:string,rowIndex?:number}>} people
 * @param {Object} payload
 * @param {Array<number>=} queueRowIndices  (optional) sign_in_log row numbers (1-based)
 * @return {Object} { ok, contactId, participantsSaved, processed? }
 */
function saveIndividualContactSession(dateStr, people, payload, queueRowIndices) {
  const tz = Session.getScriptTimeZone() || 'America/Chicago';
  const dateYMD = ymd_(dateStr, tz) || Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  // normalize unique student IDs
  const uniq = new Map();
  (Array.isArray(people) ? people : []).forEach(p=>{
    const id = String(p && p.id || '').trim();
    if (id && !uniq.has(id)) uniq.set(id, true);
  });
  if (!uniq.size) return { ok:false, error:'Select at least one student.' };

  const {
    contactWith, typeOfContact, topic, success, notes, referrals, location,
    durationMinutes, mentorId
  } = (payload || {});

  const { sh: sessSh, C: S } = ensureContactSessionsSheet_();
  const { sh: linkSh, C: L } = ensureContactParticipantsSheet_();

  const widthSess = Math.max(1, sessSh.getLastColumn());
  const widthLink = Math.max(1, linkSh.getLastColumn());

  const now = new Date();
  const contactId = Utilities.getUuid();

  const sessionRow = new Array(widthSess).fill('');
  sessionRow[S.ContactID] = contactId;
  sessionRow[S.Date]      = dateYMD;
  sessionRow[S.Duration]  = Number(durationMinutes)||0;
  sessionRow[S.With]      = String(contactWith||'').trim();
  sessionRow[S.Type]      = String(typeOfContact||'').trim();
  sessionRow[S.Topic]     = String(topic||'').trim();
  sessionRow[S.Success]   = String(success||'').trim();
  sessionRow[S.Notes]     = String(notes||'').trim();
  sessionRow[S.Referrals] = String(referrals||'').trim();
  sessionRow[S.Location]  = String(location||'').trim();
  sessionRow[S.MentorID]  = String(mentorId||'').trim().toUpperCase();
  sessionRow[S.CreatedAt] = now;
  sessionRow[S.EditedAt]  = now;

  const linkRows = Array.from(uniq.keys()).map(id=>{
    const row = new Array(widthLink).fill('');
    row[L.ContactID] = contactId;
    row[L.StudentID] = id;
    row[L.NotesStudent] = '';
    row[L.CreatedAt] = now;
    return row;
  });

  // Collect queue rows to mark processed (either explicit arg or inferred from people[*].rowIndex)
  let rowsToMark = [];
  if (Array.isArray(queueRowIndices) && queueRowIndices.length) {
    rowsToMark = queueRowIndices.map(n => Number(n)||0).filter(n => n >= 2);
  } else {
    rowsToMark = (Array.isArray(people) ? people : [])
      .map(p => Number(p && p.rowIndex))
      .filter(n => Number.isFinite(n) && n >= 2);
  }
  // de-dupe
  rowsToMark = Array.from(new Set(rowsToMark));

  const lock = LockService.getDocumentLock();
  try {
    lock.waitLock(30000);

    // Write session + links
    sessSh.getRange(sessSh.getLastRow()+1, 1, 1, widthSess).setValues([sessionRow]);
    if (linkRows.length) {
      linkSh.getRange(linkSh.getLastRow()+1, 1, linkRows.length, widthLink).setValues(linkRows);
    }

    // Optionally mark sign_in_log rows processed
    let processedInfo = null;
    if (rowsToMark.length && typeof markProcessed === 'function') {
      try {
        processedInfo = markProcessed(rowsToMark, contactId);
      } catch (e) {
        processedInfo = { ok:false, error: e && e.message ? e.message : String(e) };
      }
    }

    return {
      ok: true,
      contactId,
      participantsSaved: linkRows.length,
      processed: processedInfo || null
    };

  } catch (e) {
    return { ok:false, error: e && e.message ? e.message : String(e) };
  } finally {
    try { lock.releaseLock(); } catch(_) {}
  }
}

// ---------- Recent contacts per ID ----------
// ---- Faster recent contacts (uses cached snapshot + cached mentors + cached roster map) ----
function listRecentContactsForIds(ids, perId){
  const target = new Set((Array.isArray(ids)?ids:[]).map(x=>String(x||'').trim()).filter(Boolean));
  if (!target.size) return {};
  const n = Math.max(1, Math.min(Number(perId)||5, 50));
  const tz = Session.getScriptTimeZone() || 'America/Chicago';

  const { C: S } = ensureContactSessionsSheet_();
  const { C: L } = ensureContactParticipantsSheet_();
  const { sessVals, linkVals } = _getRecentData_();

  const mentorNameById = new Map();
  try { (getMentors(true)||[]).forEach(m => mentorNameById.set(String(m.id||'').toUpperCase(), m.name||m.id)); } catch(_){}

  const byContact = new Map();
  sessVals.forEach(r=>{
    const cid = String(r[S.ContactID]||'').trim(); if(!cid) return;
    const mid = String(r[S.MentorID]||'').trim().toUpperCase();
    byContact.set(cid, {
      dateYMD: ymd_(r[S.Date], tz) || r[S.Date],
      duration: Number(r[S.Duration]||0),
      contactWith: String(r[S.With]||''),
      typeOfContact: String(r[S.Type]||''),
      topic: String(r[S.Topic]||''),
      success: String(r[S.Success]||''),
      notes: String(r[S.Notes]||''),
      referrals: String(r[S.Referrals]||''),
      location: String(r[S.Location]||''),
      edited: r[S.EditedAt] instanceof Date ? toMDYString_(r[S.EditedAt]) : String(r[S.EditedAt]||''),
      mentorId: mid,
      mentorName: mid ? (mentorNameById.get(mid) || mid) : ''
    });
  });

  const nameMap = _getRosterIdNameMap_();

  const out = {}; target.forEach(id=>out[id]=[]);
  linkVals.forEach(r=>{
    const id = String(r[L.StudentID]||'').trim();
    if(!target.has(id)) return;
    const cid = String(r[L.ContactID]||'').trim();
    const s = byContact.get(cid); if(!s) return;
    out[id].push(Object.assign({ contactId: cid, displayName: nameMap.get(id)||'' }, s));
  });

  Object.keys(out).forEach(id=>{
    out[id].sort((a,b)=> String(b.dateYMD||'').localeCompare(String(a.dateYMD||'')));
    out[id] = out[id].slice(0,n);
  });
  return out;
}


// ---------- Names (roster lookup) ----------
function getNamesForIds(ids) {
  const list = Array.isArray(ids) ? ids : [];
  const map = getNamesForIds_(list);
  return Array.from(map.entries()).map(([id, name]) => ({ id, name }));
}

// ---- Cached roster map: ID -> Display Name ----
function _getRosterIdNameMap_() {
  const key = 'ROSTER_ID_NAME_V2';
  const hit = _scriptCacheGet_(key);
  if (hit) return new Map(hit);  // cached as array of [id, name]

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rosterTabs = ['2026','2025'];
  const norm = s => String(s||'').toLowerCase().replace(/[^a-z0-9]+/g,'');

  const pairs = [];
  for (const tab of rosterTabs) {
    const sh = ss.getSheetByName(tab);
    if (!sh || sh.getLastRow() < 2) continue;

    const lc = sh.getLastColumn();
    const header = sh.getRange(1,1,1,lc).getValues()[0].map(h=>String(h||'').trim());
    const idx = new Map(header.map((h,i)=>[norm(h), i]));
    const C = {
      ID   : idx.get('cpsidnumber') ?? idx.get('cpsid') ?? idx.get('id') ?? idx.get('studentid'),
      Full : idx.get('firstnamelastname') ?? idx.get('fullname') ?? idx.get('name') ?? idx.get('studentname'),
      First: idx.get('firstname') ?? idx.get('first') ?? idx.get('fname') ?? idx.get('givenname'),
      Last : idx.get('lastname')  ?? idx.get('last')  ?? idx.get('lname') ?? idx.get('surname') ?? idx.get('familyname'),
    };
    if (C.ID == null) continue;

    const vals = sh.getRange(2,1, sh.getLastRow()-1, lc).getValues();
    vals.forEach(r=>{
      const id = String(r[C.ID]||'').trim(); if (!id) return;
      let nm = C.Full!=null ? String(r[C.Full]||'').trim() : '';
      if (!nm && (C.First!=null || C.Last!=null)) {
        const f = C.First!=null ? String(r[C.First]||'').trim() : '';
        const l = C.Last !=null ? String(r[C.Last ]||'').trim() : '';
        nm = (f||l) ? (f+' '+l).trim() : '';
      }
      pairs.push([id, nm || id]);
    });
  }

  const map = new Map();
  pairs.forEach(([id,name])=>{ if(!map.has(id)) map.set(id, name); });
  _scriptCachePut_(key, Array.from(map.entries()), 600); // 10 minutes
  return map;
}

// ---- Name lookup using cached roster map ----
function getNamesForIds_(ids){
  const need = (ids||[]).map(x=>String(x||'').trim()).filter(Boolean);
  if (!need.length) return new Map();
  const all = _getRosterIdNameMap_();
  const out = new Map();
  need.forEach(id => out.set(id, all.get(id) || id));
  return out;
}

// ---- Cache snapshot of sessions/links for 60s ----
function _getRecentData_() {
  const cache = CacheService.getScriptCache();
  const { sh: sessSh } = ensureContactSessionsSheet_();
  const { sh: linkSh } = ensureContactParticipantsSheet_();

  const metaNow = JSON.stringify({ s:sessSh.getLastRow(), l:linkSh.getLastRow() });
  const metaHit = cache.get('RECENT_META_V1');
  const dataHit = cache.get('RECENT_DATA_V1');

  if (metaHit === metaNow && dataHit) return JSON.parse(dataHit);

  const sessVals = sessSh.getRange(2,1, Math.max(0, sessSh.getLastRow()-1), sessSh.getLastColumn()).getValues();
  const linkVals = linkSh.getRange(2,1, Math.max(0, linkSh.getLastRow()-1), linkSh.getLastColumn()).getValues();

  cache.put('RECENT_META_V1', metaNow, 60);
  cache.put('RECENT_DATA_V1', JSON.stringify({ sessVals, linkVals }), 60);
  return { sessVals, linkVals };
}

// ---- One-shot bootstrap for Individual Notes ----
function bootstrapIndividualNotes(){
  const tz = Session.getScriptTimeZone() || 'America/Chicago';
  return {
    mentors: getMentors(true),
    today: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd')
  };
}

// ---------- Light presence check ----------
function checkIdStatus(id) {
  const m = getNamesForIds_([id]);
  return { ok:true, exists: m.has(String(id||'').trim()) };
}
