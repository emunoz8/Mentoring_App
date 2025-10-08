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


// Mentors sheet for date+group
function ensureGroupNoteMentorsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('group_note_mentors');
  if (!sh) sh = ss.insertSheet('group_note_mentors');

  if (sh.getLastRow() < 1 || sh.getLastColumn() < 1) {
    sh.clear();
    sh.getRange(1,1,1,3).setValues([['Date','Group','MentorID']]);
  } else {
    const lc = Math.max(1, sh.getLastColumn());
    const firstRow = sh.getRange(1,1,1,lc).getValues()[0];
    const allBlank = firstRow.every(v => String(v||'').trim()==='');
    if (allBlank) {
      sh.getRange(1,1,1,3).setValues([['Date','Group','MentorID']]);
    }
  }
  return sh;
}

//** De-dupe ONLY the rows for a specific (date, group) */
function dedupeGroupMentorsFor(dateStr, groupName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('group_note_mentors');
  if (!sh) return { ok:true, removed:0 };

  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return { ok:true, removed:0 };

  const header = vals[0];
  const d = String(dateStr || '').trim();
  const g = String(groupName || '').trim();

  // Partition rows
  const keepOther = [header];
  const block = []; // rows for this (d,g)
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    if (String(r[0]).trim() === d && String(r[1]).trim() === g) {
      block.push(r);
    } else {
      keepOther.push(r);
    }
  }

  if (!block.length) return { ok:true, removed:0 };

  // Build unique set by MentorID (col C = index 2)
  const seen = new Set();
  const uniqBlock = [];
  for (const r of block) {
    const id = String(r[2]).trim().toUpperCase();
    if (!id || seen.has(id)) continue;
    seen.add(id);
    uniqBlock.push([d, g, id]);
  }

  // Rewrite: header + other + unique block
  sh.clearContents();
  const out = keepOther.concat(uniqBlock);
  sh.getRange(1,1,out.length,3).setValues(out);

  const removed = block.length - uniqBlock.length;
  return { ok:true, removed };
}

/** Global, one-time de-dupe of the whole sheet by (Date|Group|MentorID) */
function dedupeAllGroupMentors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('group_note_mentors');
  if (!sh) throw new Error('group_note_mentors not found');

  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return { ok:true, removed:0 };

  const header = vals[0];
  const map = new Map(); // key = Date|Group|MentorID
  for (let i = vals.length - 1; i >= 1; i--) { // keep last occurrence
    const r = vals[i];
    const key = [String(r[0]).trim(), String(r[1]).trim(), String(r[2]).trim().toUpperCase()].join('|');
    if (!map.has(key)) map.set(key, [String(r[0]).trim(), String(r[1]).trim(), String(r[2]).trim().toUpperCase()]);
  }

  const uniq = Array.from(map.values())
    .sort((a,b)=> a[0].localeCompare(b[0]) || a[1].localeCompare(b[1]) || a[2].localeCompare(b[2]));

  sh.clearContents();
  sh.getRange(1,1,1,3).setValues([header]);
  if (uniq.length) sh.getRange(2,1,uniq.length,3).setValues(uniq);

  const removed = (vals.length - 1) - uniq.length;
  return { ok:true, removed };
}

/** Save mentors uniquely for a given (date, group).
 * - Normalizes date (yyyy-MM-dd) and MentorID (UPPERCASE)
 * - Unions existing IDs for (date, group) with the incoming list
 * - Rewrites that block once (no duplicates)
 */
function saveGroupMentors(dateStr, groupName, mentors) {
  const tz = Session.getScriptTimeZone() || 'America/Chicago';
  const dYMD = ymd_(dateStr, tz) || Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const gKey = String(groupName || '').trim();
  if (!gKey) return { ok:false, error:'Missing group.' };

  // Incoming → normalized, unique (uppercase)
  const incoming = new Set(
    (Array.isArray(mentors) ? mentors : [])
      .map(m => String(m && m.id || '').trim())
      .filter(Boolean)
      .map(id => id.toUpperCase())
  );

  const lock = LockService.getDocumentLock();
  try {
    lock.waitLock(30000);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName('group_note_mentors');
    if (!sh) {
      sh = ss.insertSheet('group_note_mentors');
      sh.getRange(1,1,1,3).setValues([['Date','Group','MentorID']]);
    }

    const vals = sh.getDataRange().getValues();
    const header = (vals && vals.length) ? vals[0] : ['Date','Group','MentorID'];

    // Partition into: rows for this (date, group) and others
    const others = [header];
    const existingForKey = new Set(); // existing IDs for (dYMD, gKey)

    for (let i = 1; i < (vals ? vals.length : 0); i++) {
      const r = vals[i];
      const rowYMD = ymd_(r[0], tz);
      const rowGroup = String(r[1] == null ? '' : r[1]).trim();
      const rowId = String(r[2] == null ? '' : r[2]).trim().toUpperCase();

      if (rowYMD === dYMD && rowGroup === gKey) {
        if (rowId) existingForKey.add(rowId);
      } else {
        others.push(r);
      }
    }

    // Union existing + incoming
    const finalSet = new Set([...existingForKey, ...incoming]);
    const newRows = Array.from(finalSet).map(id => [dYMD, gKey, id]);

    // Rewrite once
    sh.clearContents();
    const out = others.concat(newRows);
    if (out.length) sh.getRange(1,1,out.length,3).setValues(out);

    return {
      ok: true,
      mentorsSaved: finalSet.size,
      addedNew: Math.max(0, finalSet.size - existingForKey.size)
    };
  } catch (e) {
    return { ok:false, error: e && e.message ? e.message : String(e) };
  } finally {
    try { lock.releaseLock(); } catch(_) {}
  }
}

// Read mentors for date+group
function getGroupMentors(dateStr, groupName) {
  const tz = Session.getScriptTimeZone() || 'America/Chicago';
  const d = ymd_(dateStr, tz);
  const g = String(groupName||'').trim();
  if (!d || !g) return [];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('group_note_mentors');
  if (!sh) return [];

  const v = sh.getDataRange().getValues();
  if (v.length < 2) return [];
  return v.slice(1)
    .filter(r => ymd_(r[0], tz)===d && String(r[1]||'').trim()===g)
    .map(r => ({ id: String(r[2]||'').trim() }));
}
