/*** Utils.gs ***/

/** Convert truthy/falsy-ish values to real booleans */
function toBool_(v) {
  const s = (v == null ? '' : String(v)).trim().toLowerCase();
  if (['true','yes','y','1','✓','done','complete','completed'].includes(s)) return true;
  if (['false','no','n','0'].includes(s)) return false;
  return null;
}

/** Normalize to yyyy-MM-dd string */
function ymd_(v, tz) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, tz || (Session.getScriptTimeZone() || 'America/Chicago'), 'yyyy-MM-dd');
  }
  const s = String(v || '').trim();
  return s ? s.slice(0,10) : '';
}

/** Convert Date → MM/DD/YYYY string */
function toMDYString_(d) {
  if (!(d instanceof Date) || isNaN(d)) return '';
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  const yyyy = d.getFullYear();
  return `${mm}/${dd}/${yyyy}`;
}

/** Parse loose date into Date at midnight */
function parseLooseDate_(v) {
  if (v instanceof Date && !isNaN(v)) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  if (typeof v === 'number' && isFinite(v)) {
    const epoch = new Date(Date.UTC(1899, 11, 30));
    return new Date(epoch.getTime() + v * 24 * 60 * 60 * 1000);
  }
  const s = String(v || '').trim();
  if (!s) return null;

  const ymd = /^(\d{4})-(\d{2})-(\d{2})$/;
  const mdy = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/;
  if (ymd.test(s)) {
    const [y, m, d] = s.split('-').map(Number);
    return new Date(y, m - 1, d);
  }
  if (mdy.test(s)) {
    const [m, d, y] = s.split('/').map(Number);
    return new Date(y, m - 1, d);
  }
  const guess = new Date(s);
  return isNaN(guess) ? null : new Date(guess.getFullYear(), guess.getMonth(), guess.getDate());
}


/**
 * Clear all data rows (below header) from a sheet.
 * Safely skips if sheet has only the header row.
 */
function _clearDataRows_(sh) {
  if (!sh) return;
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow <= 1) return; // nothing to clear
  sh.getRange(2, 1, lastRow - 1, lastCol).clearContent();
}

