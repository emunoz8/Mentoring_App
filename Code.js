/*** Code.gs ***/
/** One shared config, namespaced to avoid collisions */
const FORM_REQUIRED_KEYS = ['firstName','lastName','cpsIdNumber','intakeDate'];
const FORM_FIELDS = [
  { key:'emailAddress',        label:'Email Address' },
  { key:'lastName',            label:'Last Name' },
  { key:'firstName',           label:'First Name' },
  { key:'intakeDate',          label:'Intake Date', type:'date', readonly:true },
  { key:'participantStatus',   label:'Participant Status' },
  { key:'joinedProgramYear',   label:"Joined What's Up in Program Year" },
  { key:'birthDate',           label:'Birth Date', type:'date' },
  {
    key:'gender', label:'Gender', type:'select',
    options:['Male','Female','Nonbinary/Genderqueer','Transgender','Cisgender','Prefer not to say']
  },
  { key:'address',             label:'Address' },
  { key:'zipCode',             label:'Zip Code' },
  { key:'participantPhone',    label:'Participant Telephone Number' },
  { key:'parentPhone',         label:'Parent Telephone Number' },
  { key:'participantEmails',   label:'Participant Email(s)' },
  { key:'race',                label:'Race' },
  { key:'spanishOnly',         label:'Spanish-Language Only' },
  { key:'ageAtIntake',         label:'Age at Intake [number]', type:'number', readonly:true },
  { key:'gradeAtIntake',       label:'Grade at Intake' },
  { key:'currentGradeLevel',   label:'Current Grade Level' },
  { key:'school',              label:'School' },
  { key:'cpsIdNumber',         label:'CPS ID Number' },
  {
    key:'familyType', label:'Family Type', type:'select',
    options:['Single Parent/Female','Single Parent/Male','Two-parent Household','Independent Youth','Relative','Guardian','Foster/DCFS']
  },
  { key:'householdSize',       label:'Household Size [number]', type:'number' },
  { key:'siblingsCount',       label:'How Many Siblings? [number]', type:'number' },
  { key:'grandparentsInHouse', label:'How Many Grandparents in House? [number]', type:'number' },
  {
    key:'housingStatus', label:'Housing Status', type:'select',
    options:['Rent','Own','Homeless/Shelter','In temporary Housing','Other']
  },
  {
    key:'incomeSource',  label:'Income Source', type:'multiselect',
    options:[
      'Employment-One Parent/Guardian','Employment-Both Parents/Guardians','Employment-Participant',
      'Pension','TANF','Social Security','Unemployment','Other(including SSDI, Child Support, VA benefits)','SSI'
    ]
  },
  { key:'yearlyIncome',        label:'Yearly Income [if known]' },
  {
    key:'publicAssistance', label:'Public Assistance', type:'multiselect',
    options:['SNAP','TANF','Medicaid','SSI']
  },
  { key:'healthInsurance',     label:'Health Insurance' },
  {
    key:'everWorked', label:'Have You Ever Worked? (Select all that apply)', type:'select',
    options:['I have never worked','I have had summer jobs','I have worked part-time during the school year','I have worked full-time during the school year']
  },
  {
    key:'workingNow', label:'Are You Working Now?', type:'select',
    options:['Working full-time','Working part-time','Not working, looking for work','Not working, not looking for work']
  },
  { key:'hasIEP',              label:'Do You Have an IEP?' },
  { key:'has504',              label:'Any Medical Issues or 504 Plan?' },
  { key:'medicalIssues',       label:'If Yes, What Are Your Medical Issues?', type:'textarea' },
  {
    key:'relationshipStatus', label:'Relationship Status', type:'select',
    options:['Currently dating','Not currently dating, but looking','Not dating, not looking','Other']
  },
  {
    key:'grades', label:'Are Your Grades:', type:'select',
    options:["Mostly A's","Mostly B's","Mostly C's","Mostly Below C's"]
  },
  {
    key:'attendance', label:'How Would You Describe, Your School Attendance?', type:'select',
    options:['Almost Never Miss a Day','Occasionally Miss a Day','Frequently Miss Days']
  },
  {
    key:'punctuality', label:'How Would You Describe Getting to School on Time?', type:'select',
    options:['Always on Time','Occasionally Late','Sometimes Late','Always Late']
  },
  {
    key:'involvementTeachers', label:'Involvement with Teachers (1-5) (1 being very comfortable)', type:'select',
    options:['1','2','3','4','5']
  },
  {
    key:'involvementStaff', label:'Involvement with School Staff (1-5) (1 being very comfortable)', type:'select',
    options:['1','2','3','4','5']
  },
  { key:'extracurricular',     label:'Extracurricular Activities' },
  { key:'comments',            label:'Comments / Concerns', type:'textarea' },
];

const CONFIG = {
  ATTENDANCE: {
    SHEET: 'attendance',
    COLS: {
      timestamp : 0,
      firstName : 1,
      lastName  : 2,
      contact   : 3,
      schoolYear: 4,
      school    : 5,
      idNumber  : 6,
      group     : 7,
      inDb: 8, consent: 9, pre: 10, post: 11, shirtSize: 12
    }
  },
  FORM: {
    DATA_SHEET: '2025',
    SUBMISSIONS_SHEET: '2026',
    FIELDS: FORM_FIELDS,
    REQUIRED_KEYS: FORM_REQUIRED_KEYS,
    COLS: (() => {
      const cols = { timestamp: 0 };
      FORM_FIELDS.forEach((field, idx) => {
        cols[field.key] = idx + 1;
      });
      return Object.freeze(cols);
    })()
  }
};

/** Serve pages:
 *  - Default: Ui.html
 *  - ?page=Form to open the Form.html page
 */
function doGet(e) {
  var page = (e && e.parameter && e.parameter.page)
    ? String(e.parameter.page)
    : 'GroupNotes'; // change to 'Queue' if you want Queue as default

  // Only allow known pages; fallback if typo/unknown
  var allowed = new Set(['GroupNotes', 'Form', 'IndividualNotes', 'Queue', 'Login', 'LoginGroup', 'LoginIndividual', 'SessionAdmin']);
  if (!allowed.has(page)) page = 'GroupNotes';

  var titleMap = {
    GroupNotes: 'Group Notes',
    Form: 'Intake Form',
    IndividualNotes: 'Individual Notes',
    Queue: 'Sign-In Queue',
    Login: 'Program Sign-In',
    LoginGroup: 'Group Sign-In',
    LoginIndividual: 'Individual Sign-In',
    SessionAdmin: 'Session & Kiosk Links'
  };

  var tpl = HtmlService.createTemplateFromFile(page);
  return tpl.evaluate()
    .setTitle(titleMap[page] || 'App')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/** Allow including partial .html files (Style, scripts, etc.) */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** Returns the base URL of the deployed web app (works for /dev and /exec). */
function getWebAppBaseUrl() {
  try {
    return ScriptApp.getService().getUrl();
  } catch (e) {
    return '';
  }
}
