/*** GroupContactMentors.gs
 * Links mentors to a Group Contact session by ContactID.
 */

function ensureGroupContactMentorsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('group_contact_mentors');
  if (!sh) sh = ss.insertSheet('group_contact_mentors');

  if (sh.getLastRow() < 1 || sh.getLastColumn() < 1) {
    sh.clear();
    sh.getRange(1,1,1,5).setValues([[
      'ContactID','MentorID','Name','CreatedAt','EditedAt'
    ]]);
  }

  const header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const norm = s => String(s||'').toLowerCase().replace(/[^a-z0-9]+/g,'');
  const idx = new Map(header.map((h,i)=>[norm(h),i]));
  const C = {
    ContactID: idx.get('contactid'),
    MentorID:  idx.get('mentorid'),
    Name:      idx.get('name'),
    CreatedAt: idx.get('createdat'),
    EditedAt:  idx.get('editedat')
  };
  return { sh, C };
}

function saveGroupMentors(contactId, mentors) {
  if (!contactId) return { ok:false, error:'Missing contactId' };

  const { sh, C } = ensureGroupContactMentorsSheet_();
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();

  // Normalize & de-dupe by MentorID
  const now = new Date();
  const seen = new Set();
  const rows = (Array.isArray(mentors) ? mentors : [])
    .map(m => {
      const id = String(m?.id || '').trim();
      const name = String(m?.name || '').trim() || id;
      if (!id) return null;
      if (seen.has(id)) return null;
      seen.add(id);
      return [contactId, id, name, now, now];
    })
    .filter(Boolean);

  const data = (lastRow >= 2) ? sh.getRange(2,1,lastRow-1,lastCol).getValues() : [];
  const keep = data.filter(r => String(r[C.ContactID] || '').trim() !== contactId);
  const finalRows = keep.concat(rows);
  const changed = rows.length || keep.length !== data.length;

  if (changed) {
    _clearDataRows_(sh);
    if (finalRows.length) {
      sh.getRange(2, 1, finalRows.length, lastCol).setValues(finalRows);
    }
    if (typeof _scriptCacheRemove_ === 'function') {
      _scriptCacheRemove_('RECENT_META_V2');
      _scriptCacheRemove_('RECENT_DATA_V2');
    }
  }

  return { ok:true, contactId, added: rows.length, totalForContact: rows.length };
}
