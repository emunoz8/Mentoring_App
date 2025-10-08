/*** Code.gs ***/
/** One shared config, namespaced to avoid collisions */
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
    COLS: {
      timestamp: 0,
      emailAddress: 1,
      lastName: 2,
      firstName: 3,
      intakeDate: 4,
      participantStatus: 5,
      joinedProgramYear: 6,
      birthDate: 7,
      gender: 8,
      address: 9,
      zipCode: 10,
      participantPhone: 11,
      parentPhone: 12,
      participantEmails: 13,
      race: 14,
      spanishOnly: 15,
      ageAtIntake: 16,
      gradeAtIntake: 17,
      currentGradeLevel: 18,
      school: 19,
      cpsIdNumber: 20, // lookup key
      familyType: 21,
      householdSize: 22,
      siblingsCount: 23,
      grandparentsInHouse: 24,
      housingStatus: 25,
      incomeSource: 26,
      yearlyIncome: 27,
      publicAssistance: 28,
      healthInsurance: 29,
      everWorked: 30,
      workingNow: 31,
      hasIEP: 32,
      has504: 33,
      medicalIssues: 34,
      relationshipStatus: 35,
      grades: 36,
      attendance: 37,
      punctuality: 38,
      involvementTeachers: 39,
      involvementStaff: 40,
      extracurricular: 41,
      comments: 42,
    }
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
  var allowed = new Set(['GroupNotes', 'Form', 'IndividualNotes', 'Queue']);
  if (!allowed.has(page)) page = 'GroupNotes';

  var titleMap = {
    GroupNotes: 'Group Notes',
    Form: 'Intake Form',
    IndividualNotes: 'Individual Notes',
    Queue: 'Sign-In Queue'
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

/** Tiny connectivity check (debug) */
function ping() { return 'pong'; }

/** Returns the base URL of the deployed web app (works for /dev and /exec). */
function getWebAppBaseUrl() {
  try {
    return ScriptApp.getService().getUrl();
  } catch (e) {
    return '';
  }
}

