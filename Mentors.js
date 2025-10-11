/* ---------------- Mentors.gs ---------------- */

// ---- Mentors (cached) ----
function getMentors(activeOnly) {
  const cacheKey = 'MENTORS_LIST_V2_' + (activeOnly ? 'A' : 'ALL');
  const hit = _scriptCacheGet_(cacheKey);
  if (hit) return hit;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('mentors');
  if (!sh) return [];
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];

  const headers = vals[0].map(h => String(h || '').trim());
  const norm = s => s.toLowerCase().replace(/[^a-z0-9]+/g, '');
  const idxMap = new Map(headers.map((h,i)=>[norm(h), i]));

  const cID     = idxMap.get('mentorid') ?? idxMap.get('id') ?? idxMap.get('employeeid') ?? idxMap.get('staffid');
  const cFirst  = idxMap.get('firstname') ?? idxMap.get('first') ?? idxMap.get('fname') ?? idxMap.get('givenname');
  const cLast   = idxMap.get('lastname')  ?? idxMap.get('last')  ?? idxMap.get('lname') ?? idxMap.get('surname') ?? idxMap.get('familyname');
  const cActive = idxMap.get('active')    ?? idxMap.get('isactive') ?? idxMap.get('enabled') ?? idxMap.get('status');

  if (cID == null) return [];

  const toBool = v => {
    if (v === true || v === false) return v;
    const s = String(v == null ? '' : v).trim().toLowerCase();
    return ['true','t','yes','y','1','✓'].includes(s);
  };

  const list = vals.slice(1).map(r => {
    const id = String((cID!=null ? r[cID] : '') || '').trim();
    const first = String((cFirst!=null ? r[cFirst] : '') || '').trim();
    const last  = String((cLast !=null ? r[cLast ] : '') || '').trim();
    const active = cActive != null ? toBool(r[cActive]) : true;
    return {
      id,
      first,
      last,
      name: (first || last) ? (first + ' ' + last).trim() : id,
      active
    };
  })
  .filter(m => m.id)
  .filter(m => activeOnly ? m.active : true)
  .sort((a,b) => (a.last||'').localeCompare(b.last||'') || (a.first||'').localeCompare(b.first||''));

  _scriptCachePut_(cacheKey, list, 600); // 10 minutes
  return list;
}

function getMentorNameMap_(activeOnly) {
  const cacheKey = 'MENTOR_NAME_MAP_V1_' + (activeOnly ? 'A' : 'ALL');
  const hit = _scriptCacheGet_(cacheKey);
  if (hit) return new Map(hit);
  const list = getMentors(activeOnly);
  const pairs = list.map(m => [String(m.id || '').trim().toUpperCase(), m.name || m.id]);
  _scriptCachePut_(cacheKey, pairs, 600);
  return new Map(pairs);
}


// Read mentors for date+group
function getGroupMentors(dateStr, groupName) {
  const tz = Session.getScriptTimeZone() || 'America/Chicago';
  const d = ymd_(dateStr, tz);
  const g = String(groupName||'').trim();
  if (!d || !g) return [];

  // Prefer the contact-based storage (group_contact_mentors) if a session exists.
  const contactId = getContactIdForGroup_(dateStr, groupName);
  const contactKey = String(contactId || '').trim().toUpperCase();
  if (!contactKey) return [];

  const { sh, C } = ensureGroupContactMentorsSheet_();
  if (C.ContactID == null || C.MentorID == null || C.Name == null) return [];

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const data = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  return data
    .map(r => ({
      contact: String(r[C.ContactID] || '').trim().toUpperCase(),
      id: String(r[C.MentorID] || '').trim().toUpperCase(),
      name: String(r[C.Name] || '').trim()
    }))
    .filter(r => r.contact === contactKey && r.id)
    .map(r => ({ id: r.id, name: r.name }));
}
