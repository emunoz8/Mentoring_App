const REPORTS_SOURCE_SHEET = "reports";
const REPORTS_MAIL_SHEET   = "School Mailing Lists";
const REPORTS_MAILING_INPUT_SHEET = "ReportMailingList";
const REPORTS_MENU_MAX_SCHOOLS = 4;
const REPORTS_MENU_SLOT_PREFIX = "REPORTS_MENU_SLOT_";
const REPORTS_FOLDER_ID = "";           // Optional: set to a Drive folder ID to store PDFs
const REPORTS_FOLDER_NAME = "Reports";  // Fallback folder name if ID not set

function generateAllSchoolAttendancePdfs(options) {
  let onlySchools = options && Array.isArray(options.onlySchools)
    ? new Set(options.onlySchools.map(reportsNormalizeSchool_).filter(Boolean))
    : null;
  const forceRecipientsKey = options && options.forceRecipientsKey
    ? reportsNormalizeSchool_(options.forceRecipientsKey)
    : null;
  const skipEmail = !!(options && options.skipEmail);

  // If forcing one recipient group (e.g., OFY), send all schools to them.
  if (forceRecipientsKey) {
    onlySchools = null;
  }

  const SOURCE_SHEET_NAME = REPORTS_SOURCE_SHEET;

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SOURCE_SHEET_NAME);
  if (!sheet) throw new Error("Sheet not found: " + SOURCE_SHEET_NAME);

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    Logger.log("No data in reports.");
    return;
  }

  // Data rows, skip header
  const rows = values.slice(1);

  // Column indices (0-based)
  // 0: CPS ID Number (not used here)
  // 1: First Name
  // 2: Last Name
  // 3: School
  // 4: Current Grade Level
  // 5: Individual Attendance
  // 6: Group Attendance
  const IDX_FIRST  = 1;
  const IDX_LAST   = 2;
  const IDX_SCHOOL = 3;
  const IDX_GRADE  = 4;
  const IDX_INDIV  = 5;
  const IDX_GROUP  = 6;

  // -----------------------------
  // Group rows: school -> grade -> rows[]
  // -----------------------------
  const bySchool = Object.create(null); // key -> { label, grades }

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const schoolRaw = row[IDX_SCHOOL];
    const school = String(schoolRaw || "").trim();
    const schoolKey = reportsNormalizeSchool_(schoolRaw);
    if (!schoolKey) continue;

    if (onlySchools && !onlySchools.has(schoolKey)) continue;

    const gradeKey = row[IDX_GRADE] || "Unknown";

    if (!bySchool[schoolKey]) bySchool[schoolKey] = { label: school || "Unknown", grades: Object.create(null) };
    if (!bySchool[schoolKey].grades[gradeKey]) bySchool[schoolKey].grades[gradeKey] = [];
    bySchool[schoolKey].grades[gradeKey].push(row);
  }

  const schoolEntries = Object.values(bySchool);
  if (!schoolEntries.length) {
    Logger.log("No matching schools found for filter.");
    return;
  }

  Logger.log(`Generating PDFs for ${schoolEntries.length} school(s).`);

  // -----------------------------
  // Build recipients map: school -> [emails...]
  // -----------------------------
  const recipientsBySchool = getReportMailingListRecipients_();

  // Optional: target folder
  const folder = reportsResolveTargetFolder_();

  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");

  function gradeLabel(grade) {
    const n = Number(grade);
    if (n === 9)  return "9th Grade (Freshman)";
    if (n === 10) return "10th Grade (Sophomore)";
    if (n === 11) return "11th Grade (Junior)";
    if (n === 12) return "12th Grade (Senior)";
    return "Grade " + (grade || "Unknown");
  }

  function cmp(a, b) {
    if (a < b) return -1;
    if (a > b) return 1;
    return 0;
  }

  // -----------------------------
  // One PDF (and email) per school
  // -----------------------------
  schoolEntries.forEach((entry, schoolIndex) => {
    const gradesMap = entry.grades;
    const schoolName = entry.label;
    const schoolKey = reportsNormalizeSchool_(schoolName);
    Logger.log(`Processing school ${schoolIndex + 1}/${schoolEntries.length}: ${schoolName}`);

    const docName = `Attendance Report - ${schoolName} - ${todayStr}`;
    const pdfFileName = `${docName}.pdf`;
    let pdfFile = reportsFindExistingPdf_(pdfFileName, folder);
    let pdfBlob = null;

    if (pdfFile) {
      pdfBlob = pdfFile.getBlob();
      Logger.log(`Reusing existing PDF for ${schoolName}: ${pdfFile.getUrl()}`);
    } else {
      const doc = DocumentApp.create(docName);
      const body = doc.getBody();

      // First page: title + school info
      body.appendParagraph("Student Attendance Report")
          .setHeading(DocumentApp.ParagraphHeading.HEADING1);

      body.appendParagraph(`School: ${schoolName}`)
          .setHeading(DocumentApp.ParagraphHeading.HEADING2);

      body.appendParagraph(`Generated on: ${todayStr}`)
          .setHeading(DocumentApp.ParagraphHeading.NORMAL);

      body.appendParagraph(""); // spacer

      // Sorted grades (numeric where possible)
      const gradeKeys = Object.keys(gradesMap).sort((a, b) => {
        const ga = Number(a) || 0;
        const gb = Number(b) || 0;
        return ga - gb;
      });

      let isFirstGrade = true;

      gradeKeys.forEach(gKey => {
        const rowsForGrade = gradesMap[gKey];
        if (!rowsForGrade || !rowsForGrade.length) return;

        // Start each grade on a new page (except the first one)
        if (!isFirstGrade) {
          body.appendPageBreak();
        }
        isFirstGrade = false;

        // Grade heading at top of page
        body.appendParagraph(gradeLabel(gKey))
            .setHeading(DocumentApp.ParagraphHeading.HEADING2);

        // Sort by last name, then first name
        rowsForGrade.sort((a, b) => {
          const lastA  = (a[IDX_LAST]  || "").toString().toLowerCase();
          const lastB  = (b[IDX_LAST]  || "").toString().toLowerCase();
          const lastCmp = cmp(lastA, lastB);
          if (lastCmp !== 0) return lastCmp;

          const firstA = (a[IDX_FIRST] || "").toString().toLowerCase();
          const firstB = (b[IDX_FIRST] || "").toString().toLowerCase();
          return cmp(firstA, firstB);
        });

        Logger.log(`  Grade ${gKey}: ${rowsForGrade.length} students`);

        // Build table data for this grade
        const tableData = new Array(rowsForGrade.length + 1);
        tableData[0] = [
          "Student",
          "Individual Attendance (dates)",
          "Group Attendance (dates)"
        ];

        for (let i = 0; i < rowsForGrade.length; i++) {
          const r = rowsForGrade[i];
          const fullName = `${r[IDX_FIRST]} ${r[IDX_LAST]}`;
          const indiv = r[IDX_INDIV] || "";
          const group = r[IDX_GROUP] || "";
          tableData[i + 1] = [fullName, indiv, group];
        }

        const table = body.appendTable(tableData);
        table.setBorderWidth(0.5);

        // Bold header row
        const headerRow = table.getRow(0);
        const numCells = headerRow.getNumCells();
        for (let c = 0; c < numCells; c++) {
          headerRow.getCell(c).editAsText().setBold(true);
        }
      });

      doc.saveAndClose();

      // Convert to PDF
      const docFile = DriveApp.getFileById(doc.getId());
      pdfBlob = docFile.getAs("application/pdf").setName(pdfFileName);

      // Save PDF
      pdfFile = folder ? folder.createFile(pdfBlob) : DriveApp.createFile(pdfBlob);

      // Remove temp Doc
      docFile.setTrashed(true);

      Logger.log(`Created PDF for ${schoolName}: ${pdfFile.getUrl()}`);
    }

    // -----------------------------
    // Email this PDF to that school's recipients
    // -----------------------------
    if (!skipEmail) {
      const recipients = forceRecipientsKey
        ? (recipientsBySchool[forceRecipientsKey] || [])
        : (recipientsBySchool[schoolKey] || recipientsBySchool[schoolName] || []);
      if (recipients.length) {
        const subject = `Attendance Report - ${schoolName} - ${todayStr}`;
        const bodyText =
          `Hello,\n\n` +
          `Attached is the latest attendance report for ${schoolName}, grouped by grade level.\n\n` +
          `Generated on ${todayStr}.\n\n` +
          `Best,\n` +
          `Options for Youth Team`;

        MailApp.sendEmail({
          to: "",                  // leave "to" blank; use BCC only
          bcc: recipients.join(","), // BCC all recipients
          subject: subject,
          body: bodyText,
          attachments: [pdfBlob]
        });

        const targetLabel = forceRecipientsKey
          ? `forced group "${forceRecipientsKey}"`
          : `school "${schoolName}" (key "${schoolKey}")`;
        Logger.log(`Emailed report for ${schoolName} to ${targetLabel}: ${recipients.join(", ")}`);
      } else {
        const targetLabel = forceRecipientsKey
          ? `forced group "${forceRecipientsKey}"`
          : `school "${schoolName}" (key "${schoolKey}")`;
        Logger.log(`No email recipients found for ${targetLabel}.`);
      }
    }
  });

  Logger.log("Finished generating and emailing selected school PDFs.");
}

/**
 * UI helper: ensures the "School Mailing Lists" sheet lists every school
 * present on the reports tab, keeping any existing email/name entries, and
 * adds mailto links per school (and leaves blank if no emails).
 */
function generateSchoolMailingListSheet() {
  const ss = SpreadsheetApp.getActive();
  const source = ss.getSheetByName(REPORTS_SOURCE_SHEET);
  if (!source) throw new Error(`Sheet not found: ${REPORTS_SOURCE_SHEET}`);

  const sourceValues = source.getDataRange().getValues();
  if (sourceValues.length < 2) throw new Error("Reports sheet has no data rows.");

  const norm = reportsNormalize_;
  const srcHeader = sourceValues[0].map(norm);
  const idxSchool = srcHeader.indexOf("school");
  if (idxSchool === -1) throw new Error("Reports sheet is missing a School column.");

  const schools = new Set();
  for (let r = 1; r < sourceValues.length; r++) {
    const school = String(sourceValues[r][idxSchool] || "").trim();
    if (school) schools.add(school);
  }

  let sh = ss.getSheetByName(REPORTS_MAIL_SHEET);
  if (!sh) sh = ss.insertSheet(REPORTS_MAIL_SHEET);

  // Capture existing entries so we don't wipe manual emails/names.
  const existing = new Map(); // school -> { email, name }
  if (sh.getLastRow() >= 2) {
    const width = Math.max(1, sh.getLastColumn());
    const values = sh.getRange(1, 1, sh.getLastRow(), width).getValues();
    const header = values[0].map(norm);
    const idxMailSchool = header.indexOf("school");
    const idxMailEmail  = header.indexOf("email");
    const idxMailName   = header.indexOf("name");

    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      const school = idxMailSchool !== -1 ? String(row[idxMailSchool] || "").trim() : "";
      if (!school) continue;
      const email = idxMailEmail !== -1 ? String(row[idxMailEmail] || "").trim() : "";
      const name  = idxMailName  !== -1 ? String(row[idxMailName]  || "").trim() : "";
      existing.set(school, { email, name });
      schools.add(school); // include any manual schools not on reports tab
    }
  }

  const sortedSchools = Array.from(schools).sort((a, b) => a.localeCompare(b));

  // Build output rows: [School, Email, Name, Email Link Text]
  const rows = sortedSchools.map(school => {
    const prev = existing.get(school) || {};
    const emailStr = prev.email || "";
    const nameStr = prev.name || "";
    const mailto = reportsBuildMailtoUrl_(emailStr, `School update - ${school}`);
    const linkText = mailto ? `Email ${school}` : "No emails found";
    return { school, emailStr, nameStr, mailto, linkText };
  });

  // Rewrite sheet
  sh.clear();
  sh.getRange(1, 1, 1, 4).setValues([["School", "Email", "Name", "Email Link"]]);
  if (rows.length) {
    sh.getRange(2, 1, rows.length, 4).setValues(
      rows.map(r => [r.school, r.emailStr, r.nameStr, r.linkText])
    );

    // Apply mailto links in column D
    rows.forEach((r, idx) => {
      if (!r.mailto) return;
      const cell = sh.getRange(2 + idx, 4);
      const rich = SpreadsheetApp.newRichTextValue()
        .setText(r.linkText)
        .setLinkUrl(r.mailto)
        .build();
      cell.setRichTextValue(rich);
    });
  }

  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, 4).setFontWeight("bold");
  sh.autoResizeColumns(1, 3);
  sh.getRange(1, 1, Math.max(2, sh.getLastRow()), 3)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  Logger.log(`Updated "${REPORTS_MAIL_SHEET}" with ${rows.length} school(s).`);
  return { schools: rows.length };
}

/**
 * Reads the "ReportMailingList" sheet and returns:
 *   { [schoolName]: [ "email1@...", "email2@..." ] }
 */
function getSchoolRecipientsMap_() {
  return getReportMailingListRecipients_();
}

function getReportMailingListRecipients_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(REPORTS_MAILING_INPUT_SHEET);
  const recipientsBySchool = Object.create(null); // key (normalized) -> [emails]

  if (!sh) {
    Logger.log(`Mailing list sheet "${REPORTS_MAILING_INPUT_SHEET}" not found; emails will be skipped.`);
    return recipientsBySchool;
  }

  const values = sh.getDataRange().getValues();
  if (values.length < 2) {
    Logger.log(`Mailing list sheet "${REPORTS_MAILING_INPUT_SHEET}" has no data.`);
    return recipientsBySchool;
  }

  const header = values[0].map(reportsNormalize_);
  const idxSchool = header.indexOf("school");
  const idxEmail = header.indexOf("email") !== -1 ? header.indexOf("email") : header.indexOf("emailaddress");

  if (idxSchool === -1 || idxEmail === -1) {
    Logger.log(`Mailing list sheet "${REPORTS_MAILING_INPUT_SHEET}" is missing School or Email column.`);
    return recipientsBySchool;
  }

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const schoolRaw = row[idxSchool];
    const school = String(schoolRaw || "").trim();
    const schoolKey = reportsNormalizeSchool_(schoolRaw);
    const emails = reportsSplitEmails_(row[idxEmail]);
    if (!schoolKey || !emails.length) continue;

    if (!recipientsBySchool[schoolKey]) recipientsBySchool[schoolKey] = [];

    emails.forEach(email => {
      const key = email.toLowerCase();
      if (!recipientsBySchool[schoolKey].some(e => e.toLowerCase() === key)) {
        recipientsBySchool[schoolKey].push(email);
      }
    });

    // Also store by raw trimmed label to cover menu selections that use display names
    if (school) {
      if (!recipientsBySchool[school]) recipientsBySchool[school] = recipientsBySchool[schoolKey];
    }
  }

  Logger.log(`Loaded recipients for ${Object.keys(recipientsBySchool).length} school(s) from "${REPORTS_MAILING_INPUT_SHEET}".`);
  return recipientsBySchool;
}

function sendReportToSchool(schoolNameRaw) {
  const schoolName = String(schoolNameRaw || "").trim();
  if (!schoolName) throw new Error("School name is required.");
  return generateAllSchoolAttendancePdfs({ onlySchools: [schoolName] });
}

function sendReportsToEveryone() {
  return generateAllSchoolAttendancePdfs({ onlySchools: null });
}

function sendReportsToOfy() {
  return generateAllSchoolAttendancePdfs({ forceRecipientsKey: "OFY" });
}

function reportsFetchSchoolsForMenu_() {
  const ss = SpreadsheetApp.getActive();
  const schools = new Map(); // key -> display label

  // Prefer mailing list sheet
  const ml = ss.getSheetByName(REPORTS_MAILING_INPUT_SHEET);
  if (ml && ml.getLastRow() >= 2) {
    const values = ml.getRange(1, 1, ml.getLastRow(), Math.max(1, ml.getLastColumn())).getValues();
    const header = values[0].map(reportsNormalize_);
    const idxSchool = header.indexOf("school");
    if (idxSchool !== -1) {
      for (let r = 1; r < values.length; r++) {
        const school = String(values[r][idxSchool] || "").trim();
        const key = reportsNormalizeSchool_(school);
        if (key && !schools.has(key)) schools.set(key, school);
      }
    }
  }

  // Fallback: schools from reports tab
  if (!schools.size) {
    const source = ss.getSheetByName(REPORTS_SOURCE_SHEET);
    if (source && source.getLastRow() >= 2) {
      const vals = source.getDataRange().getValues();
      const header = vals[0].map(reportsNormalize_);
      const idxSchool = header.indexOf("school");
      if (idxSchool !== -1) {
        for (let r = 1; r < vals.length; r++) {
          const school = String(vals[r][idxSchool] || "").trim();
          const key = reportsNormalizeSchool_(school);
          if (key && !schools.has(key)) schools.set(key, school);
        }
      }
    }
  }

  return Array.from(schools.values()).sort((a, b) => a.localeCompare(b));
}

function reportsBuildMenuActions_() {
  const schools = reportsFetchSchoolsForMenu_();
  // Skip OFY from per-school menu; OFY gets its own "all schools" option
  const filtered = schools.filter(s => reportsNormalizeSchool_(s) !== "ofy");
  const limited = filtered.slice(0, REPORTS_MENU_MAX_SCHOOLS);

  // Store slot -> school mapping so menu handlers can be static functions.
  reportsStoreMenuSlots_(limited);

  const actions = limited.map((school, idx) => ({
    label: `Send report to ${school}`,
    handler: `reportsMenuSendSlot${idx + 1}`
  }));

  // Always include OFY (sends all schools to OFY recipients)
  actions.push({ label: "Send report to OFY (all schools)", handler: "reportsMenuSendOFY" });

  actions.push({ label: "Send reports to everyone", handler: "reportsSendAllReports_" });
  return actions;
}

function reportsSendAllReports_() {
  return sendReportsToEveryone();
}

function reportsMenuSendOFY() {
  return sendReportsToOfy();
}

function reportsMenuSendSlot_(slot) {
  const key = `${REPORTS_MENU_SLOT_PREFIX}${slot}`;
  const props = PropertiesService.getUserProperties();
  const school = props.getProperty(key);
  if (!school) throw new Error(`No school configured for menu slot ${slot}. Please reopen the sheet to refresh the menu.`);
  return sendReportToSchool(school);
}

function reportsMenuSendSlot1() { return reportsMenuSendSlot_(1); }
function reportsMenuSendSlot2() { return reportsMenuSendSlot_(2); }
function reportsMenuSendSlot3() { return reportsMenuSendSlot_(3); }
function reportsMenuSendSlot4() { return reportsMenuSendSlot_(4); }

function reportsStoreMenuSlots_(schools) {
  const props = PropertiesService.getUserProperties();
  const limit = REPORTS_MENU_MAX_SCHOOLS;
  for (let i = 1; i <= limit; i++) {
    const school = schools[i - 1] || "";
    if (school) {
      props.setProperty(`${REPORTS_MENU_SLOT_PREFIX}${i}`, school);
    } else {
      props.deleteProperty(`${REPORTS_MENU_SLOT_PREFIX}${i}`);
    }
  }
}

// ----------------- Helpers -----------------
function reportsNormalize_(value) {
  return String(value || "").toLowerCase().replace(/[^a-z0-9]+/g, "");
}

function reportsNormalizeSchool_(value) {
  return String(value || "").trim().toLowerCase();
}

function reportsSplitEmails_(value) {
  if (!value) return [];
  return String(value)
    .replace(/\s*\n\s*/g, ",")
    .split(/[,;]+/)
    .map(e => e.trim())
    .filter(Boolean);
}

function reportsBuildMailtoUrl_(emailsRaw, subject) {
  const emails = Array.from(new Set(reportsSplitEmails_(emailsRaw).map(e => e.toLowerCase())));
  if (!emails.length) return "";
  const bcc = encodeURIComponent(emails.join(","));
  const subj = encodeURIComponent(subject || "School update");
  return `mailto:?bcc=${bcc}&subject=${subj}`;
}

function reportsResolveTargetFolder_() {
  // Prefer explicit folder ID
  if (REPORTS_FOLDER_ID) {
    try {
      return DriveApp.getFolderById(REPORTS_FOLDER_ID);
    } catch (err) {
      Logger.log(`Could not open folder by ID "${REPORTS_FOLDER_ID}": ${err}`);
    }
  }

  // Try to reuse an existing folder with the configured name
  if (REPORTS_FOLDER_NAME) {
    const it = DriveApp.getFoldersByName(REPORTS_FOLDER_NAME);
    if (it && it.hasNext()) return it.next();
    try {
      return DriveApp.createFolder(REPORTS_FOLDER_NAME);
    } catch (err) {
      Logger.log(`Could not create folder "${REPORTS_FOLDER_NAME}": ${err}`);
    }
  }

  // Fallback: My Drive
  return null;
}

function reportsFindExistingPdf_(fileName, folder) {
  if (!fileName) return null;
  if (folder) {
    const it = folder.getFilesByName(fileName);
    if (it && it.hasNext()) return it.next();
  }
  const it2 = DriveApp.getFilesByName(fileName);
  return (it2 && it2.hasNext()) ? it2.next() : null;
}
