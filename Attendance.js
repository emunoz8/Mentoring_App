/*** Attendance.gs ***/
/* Uses CONFIG from Code.gs:
   CONFIG.ATTENDANCE.SHEET
   CONFIG.ATTENDANCE.COLS
*/


// ---------- API: read roster by date ----------
function getRosterByDate(dateStr) {
  var tz = Session.getScriptTimeZone() || 'America/Chicago';
  var target = ymd_(dateStr, tz) || Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CONFIG.ATTENDANCE.SHEET);
  if (!sh) throw new Error('Attendance sheet "' + CONFIG.ATTENDANCE.SHEET + '" not found');

  var A = CONFIG.ATTENDANCE.COLS;
  var vals = sh.getDataRange().getValues();
  if (vals.length < 2) return {};

  var roster = {};
  var seen = new Set(); // de-dupe by id|group|date

  for (var i = 1; i < vals.length; i++) {
    var r = vals[i];

    var rowDate = ymd_(r[A.timestamp], tz);
    if (rowDate !== target) continue;

    var group = String(r[A.group] == null ? '' : r[A.group]).trim();
    var id    = String(r[A.idNumber] == null ? '' : r[A.idNumber]).trim();
    if (!group || !id) continue;

    var key = id + '|' + group + '|' + rowDate;
    if (seen.has(key)) continue;
    seen.add(key);

    var first = String(r[A.firstName] == null ? '' : r[A.firstName]).trim();
    var last  = String(r[A.lastName]  == null ? '' : r[A.lastName]).trim();
    var name  = (first || last) ? (first + ' ' + last).trim() : id;

    var school     = String(r[A.school]     == null ? '' : r[A.school]).trim();
    var schoolYear = String(r[A.schoolYear] == null ? '' : r[A.schoolYear]).trim();

    var inDb    = (A.inDb    != null) ? toBool_(r[A.inDb])    : null;
    var consent = (A.consent != null) ? toBool_(r[A.consent]) : null;
    var pre     = (A.pre     != null) ? toBool_(r[A.pre])     : null;
    var post    = (A.post    != null) ? toBool_(r[A.post])    : null;
    var shirt   = (A.shirtSize != null) ? String(r[A.shirtSize] || '').trim() : '';

    (roster[group] = roster[group] || []).push({
      id: id,
      name: name,
      school: school,
      grade: schoolYear,
      inDb: inDb,
      consent: consent,
      pre: pre,
      post: post,
      shirt: shirt
    });
  }

  var result = {};
  Object.keys(roster).forEach(function(g){
    if (roster[g].length) result[g] = roster[g];
  });
  return result;
}

// ---------- API: move member to a new group ----------
function moveMember(idRaw, fromGroupRaw, toGroupRaw, dateStr) {
  var id = (idRaw == null ? '' : String(idRaw)).trim();
  var fromGroup = (fromGroupRaw == null ? '' : String(fromGroupRaw)).trim();
  var toGroup = (toGroupRaw == null ? '' : String(toGroupRaw)).trim();
  if (!id || !fromGroup || !toGroup) return { ok: false, error: 'Missing id/from/to.' };
  if (fromGroup === toGroup) return { ok: true, noop: true };

  var tz = Session.getScriptTimeZone() || 'America/Chicago';
  var target = ymd_(dateStr, tz) || Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CONFIG.ATTENDANCE.SHEET);
  if (!sh) return { ok: false, error: 'Attendance sheet "' + CONFIG.ATTENDANCE.SHEET + '" not found' };

  var A = CONFIG.ATTENDANCE.COLS;
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2) return { ok: false, error: 'No attendance rows.' };

  var lock = LockService.getDocumentLock();
  try {
    lock.waitLock(30000);
    var vals = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

    for (var i = 0; i < vals.length; i++) {
      var r = vals[i];

      var rowDate = ymd_(r[A.timestamp], tz);
      if (rowDate !== target) continue;

      var rowId = String(r[A.idNumber] == null ? '' : r[A.idNumber]).trim();
      if (rowId !== id) continue;

      var rowGroup = String(r[A.group] == null ? '' : r[A.group]).trim();
      if (rowGroup !== fromGroup) continue;

      var sheetRow = i + 2;
      sh.getRange(sheetRow, A.group + 1).setValue(toGroup);

      return { ok: true, row: sheetRow, id: id, fromGroup: fromGroup, toGroup: toGroup, date: target };
    }

    return { ok: false, error: 'Matching row not found for that date/id/group.' };
  } catch (e) {
    return { ok: false, error: e && e.message ? e.message : String(e) };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

// ---------- Update ID on a single attendance row ----------
function updateMemberId(originalIdRaw, groupRaw, dateStr, newIdRaw) {
  var id0 = String(originalIdRaw || '').trim();
  var grp = String(groupRaw || '').trim();
  var id1 = String(newIdRaw || '').trim();
  if (!id0 || !grp || !id1) return { ok: false, error: 'Missing id/group.' };

  var tz = Session.getScriptTimeZone() || 'America/Chicago';
  var target = ymd_(dateStr, tz) || Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CONFIG.ATTENDANCE.SHEET);
  if (!sh) return { ok: false, error: 'Attendance sheet "' + CONFIG.ATTENDANCE.SHEET + '" not found' };

  var A = CONFIG.ATTENDANCE.COLS;
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2) return { ok: false, error: 'No attendance rows.' };

  var lock = LockService.getDocumentLock();
  try {
    lock.waitLock(30000);

    var vals = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
    for (var i = 0; i < vals.length; i++) {
      var r = vals[i];

      var rowDate = ymd_(r[A.timestamp], tz);
      if (rowDate !== target) continue;

      var rowId = String(r[A.idNumber] == null ? '' : r[A.idNumber]).trim();
      var rowGroup = String(r[A.group] == null ? '' : r[A.group]).trim();
      if (rowId !== id0 || rowGroup !== grp) continue;

      var sheetRow = i + 2;
      sh.getRange(sheetRow, A.idNumber + 1).setValue(id1);

      return { ok: true, row: sheetRow, oldId: id0, newId: id1, group: grp, date: target };
    }

    return { ok: false, error: 'Matching row not found for that date/id/group.' };
  } catch (e) {
    return { ok: false, error: e && e.message ? e.message : String(e) };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}
