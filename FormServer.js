/*** FormServer.gs ***/
/* Uses CONFIG.FORM from Code.gs */

function _norm_(s){return String(s||'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim();}

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
  return {
    timestamp:r[C.timestamp], emailAddress:r[C.emailAddress],
    lastName:r[C.lastName], firstName:r[C.firstName],
    intakeDate:r[C.intakeDate], participantStatus:r[C.participantStatus],
    joinedProgramYear:r[C.joinedProgramYear], birthDate:r[C.birthDate],
    gender:r[C.gender], address:r[C.address], zipCode:r[C.zipCode],
    participantPhone:r[C.participantPhone], parentPhone:r[C.parentPhone],
    participantEmails:r[C.participantEmails], race:r[C.race],
    spanishOnly:r[C.spanishOnly], ageAtIntake:r[C.ageAtIntake],
    gradeAtIntake:r[C.gradeAtIntake], currentGradeLevel:r[C.currentGradeLevel],
    school:r[C.school], cpsIdNumber:r[C.cpsIdNumber], familyType:r[C.familyType],
    householdSize:r[C.householdSize], siblingsCount:r[C.siblingsCount],
    grandparentsInHouse:r[C.grandparentsInHouse], housingStatus:r[C.housingStatus],
    incomeSource:r[C.incomeSource], yearlyIncome:r[C.yearlyIncome],
    publicAssistance:r[C.publicAssistance], healthInsurance:r[C.healthInsurance],
    everWorked:r[C.everWorked], workingNow:r[C.workingNow],
    hasIEP:r[C.hasIEP], has504:r[C.has504], medicalIssues:r[C.medicalIssues],
    relationshipStatus:r[C.relationshipStatus], grades:r[C.grades],
    attendance:r[C.attendance], punctuality:r[C.punctuality],
    involvementTeachers:r[C.involvementTeachers], involvementStaff:r[C.involvementStaff],
    extracurricular:r[C.extracurricular], comments:r[C.comments],
  };
}

/** Suggest by ID or name */
function suggestPeople(query, limit){
  const q = _norm_(query); if (!q) return [];
  const n = Math.max(1, Math.min(Number(limit)||10, 20));
  const C = CONFIG.FORM.COLS;
  const items = _dataRows_().map((r,i)=>{
    const first=r[C.firstName]||'', last=r[C.lastName]||'';
    const id=r[C.cpsIdNumber]||'', school=r[C.school]||'', grade=r[C.currentGradeLevel]||'';
    const full = `${first} ${last}`.trim();
    return {rowIndex:i+2, firstName:first, lastName:last, id, school, grade,
            _n:{id:_norm_(id), first:_norm_(first), last:_norm_(last), full:_norm_(full), school:_norm_(school), grade:_norm_(grade)},
            label:`${full}${id?' - '+id:''} (${school} - ${grade})`};
  });

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
    .slice(0,n).map(({p})=>({label:p.label,value:p.id||p.label,firstName:p.firstName,lastName:p.lastName,
                              id:p.id,school:p.school,grade:p.grade,rowIndex:p.rowIndex}));
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
  return {ok:true,message:'Submission saved to 2026.'};
}
