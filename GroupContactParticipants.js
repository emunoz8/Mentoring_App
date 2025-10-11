/*** GroupContactParticipants.gs
 * Links students to a Group Contact session by ContactID.
 */

function ensureGroupContactParticipantsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('group_contact_participants');
  if (!sh) sh = ss.insertSheet('group_contact_participants');

  if (sh.getLastRow() < 1 || sh.getLastColumn() < 1) {
    sh.clear();
    sh.getRange(1,1,1,6).setValues([[
      'ContactID','StudentID','FirstName','LastName','CreatedAt','EditedAt'
    ]]);
  }

  const header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const norm = s => String(s||'').toLowerCase().replace(/[^a-z0-9]+/g,'');
  const idx = new Map(header.map((h,i)=>[norm(h),i]));
  const C = {
    ContactID: idx.get('contactid'),
    StudentID: idx.get('studentid'),
    First:     idx.get('firstname'),
    Last:      idx.get('lastname'),
    CreatedAt: idx.get('createdat'),
    EditedAt:  idx.get('editedat')
  };
  return { sh, C };
}

function saveGroupParticipants(contactId, participants) {
  if (!contactId) return { ok:false, error:'Missing contactId' };

  const { sh, C } = ensureGroupContactParticipantsSheet_();
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();

  // Normalize & de-dupe incoming participants by id
  const now = new Date();
  const seen = new Set();
  const rows = (Array.isArray(participants) ? participants : [])
    .map(p => {
      const id = String(p?.id || '').trim();
      const name = String(p?.name || '').trim();
      if (!id) return null; // skip blanks
      if (seen.has(id)) return null; // de-dupe
      seen.add(id);

      // split name -> first/last (best-effort)
      let first = '', last = '';
      if (name) {
        const parts = name.split(/\s+/);
        last = parts.slice(-1)[0] || '';
        first = parts.slice(0, -1).join(' ') || (name && !last ? name : '');
      }
      return [contactId, id, first, last, now, now];
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
