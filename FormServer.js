/*** FormServer.gs ***/
/* Uses CONFIG.FORM from Code.gs */

function _norm_(s){return String(s||'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim();}

function _hydrateRosterRecords_(pack){
  if (!Array.isArray(pack)) return [];
  return pack.map(item => {
    const arr = Array.isArray(item) ? item : [];
    const id = String(arr[0] || '').trim();
    const firstName = String(arr[1] || '').trim();
    const lastName  = String(arr[2] || '').trim();
    const fullRaw   = String(arr[3] || '').trim();
    const school    = String(arr[4] || '').trim();
    const grade     = String(arr[5] || '').trim();
    const tab       = String(arr[6] || '').trim();
    const rowIndex  = arr[7] ?? null;

    const full = fullRaw || `${firstName} ${lastName}`.trim();
    const labelParts = [];
    if (full) labelParts.push(full);
    if (id) labelParts.push(id);
    const metaParts = [school, grade].filter(Boolean);
    const label = labelParts.length
      ? `${labelParts.join(' · ')}${metaParts.length ? ' (' + metaParts.join(' • ') + ')' : ''}`
      : id;

    return {
      id,
      firstName,
      lastName,
      school,
      grade,
      tab,
      rowIndex,
      label,
      _n: {
        id: _norm_(id),
        first: _norm_(firstName),
        last: _norm_(lastName),
        full: _norm_(full),
        school: _norm_(school),
        grade: _norm_(grade)
      }
    };
  });
}

function _dataRows_(){
  const cache = CacheService.getScriptCache();
  const hit = cache.get('FORM_DATA_ROWS_V1');
  if (hit) return JSON.parse(hit);
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.FORM.DATA_SHEET);
  if (!sh) throw new Error(`Missing tab: ${CONFIG.FORM.DATA_SHEET}`);
  const rows = sh.getDataRange().getValues().slice(1);
  cache.put('FORM_DATA_ROWS_V1', JSON.stringify(rows), 300);
  return rows;
}

function _rowToRecord_(r){
  const C = CONFIG.FORM.COLS;
  const record = {};
  Object.keys(C).forEach(key => {
    record[key] = r[C[key]];
  });
  return record;
}

function _getSuggestRoster_(){
  const cache = CacheService.getScriptCache();
  const key = 'FORM_SUGGEST_ROSTER_V2';
  const hit = cache.get(key);
  if (hit) {
    try { return _hydrateRosterRecords_(JSON.parse(hit)); }
    catch (_) { /* fall through */ }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const priorityTabs = ['2026', CONFIG.FORM.DATA_SHEET]; // prefer current year, then legacy/intake
  const additionalTabs = ['2024','2023']; // include prior cohorts for broader lookup
  const seen = new Set();
  const rows = [];
  const packed = [];

  const allTabs = [...priorityTabs, ...additionalTabs];

  allTabs.forEach(tab => {
    const sh = ss.getSheetByName(tab);
    if (!sh || sh.getLastRow() < 2) return;

    const lc = sh.getLastColumn();
    const header = sh.getRange(1,1,1,lc).getValues()[0].map(h => String(h||'').trim());
    const norm = s => String(s||'').toLowerCase().replace(/[^a-z0-9]+/g,'');
    const idx = new Map(header.map((h,i)=>[norm(h), i]));

    const col = {
      id    : idx.get('cpsidnumber') ?? idx.get('cpsid') ?? idx.get('id') ?? idx.get('studentid'),
      first : idx.get('firstname') ?? idx.get('first') ?? idx.get('fname') ?? idx.get('givenname'),
      last  : idx.get('lastname')  ?? idx.get('last')  ?? idx.get('lname') ?? idx.get('surname') ?? idx.get('familyname'),
      full  : idx.get('fullname') ?? idx.get('studentname') ?? idx.get('firstnamelastname') ?? idx.get('name'),
      school: idx.get('school'),
      grade : idx.get('currentgradelevel') ?? idx.get('grade') ?? idx.get('gradelvl') ?? idx.get('gradelevel'),
    };
    if (col.id == null) return;

    const vals = sh.getRange(2,1, sh.getLastRow()-1, lc).getValues();
    vals.forEach((r, i) => {
      const id = String(r[col.id] || '').trim();
      if (!id || seen.has(id)) return;
      seen.add(id);
      const first = col.first != null ? String(r[col.first] || '').trim() : '';
      const last  = col.last  != null ? String(r[col.last ] || '').trim() : '';
      const fullFromCols = (first || last) ? (first + ' ' + last).trim() : '';
      const full = col.full != null ? (String(r[col.full] || '').trim() || fullFromCols) : fullFromCols;
      const school = col.school != null ? String(r[col.school] || '').trim() : '';
      const grade  = col.grade  != null ? String(r[col.grade ] || '').trim() : '';
      const packedRow = [
        id,
        first,
        last,
        full,
        school,
        grade,
        tab,
        tab === CONFIG.FORM.DATA_SHEET ? (i + 2) : null
      ];
      packed.push(packedRow);
      const hydrated = _hydrateRosterRecords_([packedRow])[0];
      if (hydrated) rows.push(hydrated);
    });
  });

  try {
    cache.put(key, JSON.stringify(packed), 300);
  } catch (_) {
    // If cache exceeds size limits, safely ignore (we still return computed rows)
  }
  return rows;
}

/** Suggest by ID or name */
function suggestPeople(query, limit){
  const q = _norm_(query); if (!q) return [];
  const n = Math.max(1, Math.min(Number(limit)||10, 20));
  const items = _getSuggestRoster_();

  function score(p){
    const fields=[p._n.id,p._n.first,p._n.last,p._n.full,p._n.school,p._n.grade];
    return q.split(/\s+/).filter(Boolean).reduce((s,t)=>{
      fields.forEach(h=>{
        if(!t||!h) return;
        if(h===t) s+=5; else if(h.startsWith(t)) s+=4; else if(h.includes(t)) s+=2;
      }); return s;
    },0);
  }

  return items.map(p=>({p,s:score(p)})).filter(x=>x.s>0)
    .sort((a,b)=>b.s-a.s||a.p.lastName.localeCompare(b.p.lastName))
    .slice(0,n)
    .map(({p})=>({
      label:p.label,
      value:p.id || p.label,
      firstName:p.firstName,
      lastName:p.lastName,
      id:p.id,
      school:p.school,
      grade:p.grade,
      rowIndex:p.rowIndex
    }));
}



/** Exact lookup by CPS ID; block if already in SUBMISSIONS */
function lookupById(idRaw){
  const id = String(idRaw??'').trim(); if(!id) return {ok:false,error:'No CPS ID Number provided.'};
  const ss = SpreadsheetApp.getActive(); const C = CONFIG.FORM.COLS;

  const sub = ss.getSheetByName(CONFIG.FORM.SUBMISSIONS_SHEET);
  if (sub){
    const vals = sub.getDataRange().getValues().slice(1);
    for (const r of vals){ if (String(r[C.cpsIdNumber]??'').trim()===id) return {ok:false,error:'This person is already in 2026.'}; }
  }

  const rows=_dataRows_();
  for (const r of rows){ if (String(r[C.cpsIdNumber]??'').trim()===id) return {ok:true,record:_rowToRecord_(r)}; }
  return {ok:false,error:`CPS ID Number "${id}" not found.`};
}

/** Submit to SUBMISSIONS_SHEET (2026). Computes intake date & age. */
function submitForm(payload){
  if(!payload) return {ok:false,error:'No payload.'};
  if(!payload.cpsIdNumber) return {ok:false,error:'CPS ID Number is required.'};
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.FORM.SUBMISSIONS_SHEET);
  if(!sh) throw new Error(`Missing tab: ${CONFIG.FORM.SUBMISSIONS_SHEET}`);
  const C = CONFIG.FORM.COLS;

  function parseLooseDate_(v){
    if (v instanceof Date && !isNaN(v)) return new Date(v.getFullYear(),v.getMonth(),v.getDate());
    if (typeof v==='number' && isFinite(v)){ const epoch=new Date(Date.UTC(1899,11,30)); return new Date(epoch.getTime()+v*86400000); }
    const s=String(v||'').trim(); if(!s) return null;
    const ymd=/^(\d{4})-(\d{2})-(\d{2})$/, mdy=/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/;
    if(ymd.test(s)){ const [y,m,d]=s.split('-').map(Number); return new Date(y,m-1,d); }
    if(mdy.test(s)){ const [m,d,y]=s.split('/').map(Number); return new Date(y,m-1,d); }
    const g=new Date(s); return isNaN(g)?null:new Date(g.getFullYear(),g.getMonth(),g.getDate());
  }
  function toMDY(d){ if(!(d instanceof Date)||isNaN(d))return''; return `${String(d.getMonth()+1).padStart(2,'0')}/${String(d.getDate()).padStart(2,'0')}/${d.getFullYear()}`; }
  function yearsBetween_(b,r){ if(!b||!r) return ''; let age=r.getFullYear()-b.getFullYear(); const m=r.getMonth()-b.getMonth(); if(m<0||(m===0&&r.getDate()<b.getDate())) age--; return age; }

  const today = new Date();
  const birthObj = parseLooseDate_(payload.birthDate);
  const age = birthObj ? yearsBetween_(birthObj, today) : '';
  const row = [];
  row[C.timestamp] = new Date();
  row[C.emailAddress]=payload.emailAddress||'';
  row[C.lastName]=payload.lastName||'';
  row[C.firstName]=payload.firstName||'';
  row[C.intakeDate]=toMDY(today);
  row[C.participantStatus]=payload.participantStatus||'';
  row[C.joinedProgramYear]=payload.joinedProgramYear||'';
  row[C.birthDate]= birthObj?toMDY(birthObj):'';
  row[C.gender]=payload.gender||'';
  row[C.address]=payload.address||'';
  row[C.zipCode]=payload.zipCode||'';
  row[C.participantPhone]=payload.participantPhone||'';
  row[C.parentPhone]=payload.parentPhone||'';
  row[C.participantEmails]=payload.participantEmails||'';
  row[C.race]=payload.race||'';
  row[C.spanishOnly]=payload.spanishOnly||'';
  row[C.ageAtIntake]=(age===''?'':age);
  row[C.gradeAtIntake]=payload.gradeAtIntake||'';
  row[C.currentGradeLevel]=payload.currentGradeLevel||'';
  row[C.school]=payload.school||'';
  row[C.cpsIdNumber]=payload.cpsIdNumber||'';
  row[C.familyType]=payload.familyType||'';
  row[C.householdSize]=payload.householdSize||'';
  row[C.siblingsCount]=payload.siblingsCount||'';
  row[C.grandparentsInHouse]=payload.grandparentsInHouse||'';
  row[C.housingStatus]=payload.housingStatus||'';
  row[C.incomeSource]=payload.incomeSource||'';
  row[C.yearlyIncome]=payload.yearlyIncome||'';
  row[C.publicAssistance]=payload.publicAssistance||'';
  row[C.healthInsurance]=payload.healthInsurance||'';
  row[C.everWorked]=payload.everWorked||'';
  row[C.workingNow]=payload.workingNow||'';
  row[C.hasIEP]=payload.hasIEP||'';
  row[C.has504]=payload.has504||'';
  row[C.medicalIssues]=payload.medicalIssues||'';
  row[C.relationshipStatus]=payload.relationshipStatus||'';
  row[C.grades]=payload.grades||'';
  row[C.attendance]=payload.attendance||'';
  row[C.punctuality]=payload.punctuality||'';
  row[C.involvementTeachers]=payload.involvementTeachers||'';
  row[C.involvementStaff]=payload.involvementStaff||'';
  row[C.extracurricular]=payload.extracurricular||'';
  row[C.comments]=payload.comments||'';

  sh.appendRow(row.map(v=>v===undefined?'':v));
  try {
    const cache = CacheService.getScriptCache();
    cache.remove('FORM_SUGGEST_ROSTER_V2');
    cache.remove('FORM_DATA_ROWS_V1');
  } catch (_) {}
  return {ok:true,message:'Submission saved to 2026.'};
}
