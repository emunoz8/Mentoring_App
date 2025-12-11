function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Text Tools")
    .addItem("Capitalize Selected Cells", "capitalizeSelection")
    .addToUi();

  const reportsMenu = ui.createMenu("Reports");
  reportsMenu.addItem("Build School Mailing Lists", "generateSchoolMailingListSheet");

  if (typeof reportsBuildMenuActions_ === 'function') {
    const actions = reportsBuildMenuActions_();
    actions.forEach(action => {
      reportsMenu.addItem(action.label, action.handler);
    });
  } else {
    reportsMenu.addItem("Send reports to everyone", "reportsSendAllReports_");
  }

  reportsMenu.addToUi();
}

function capitalizeSelection() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  const values = range.getValues();

  const newValues = values.map(row =>
    row.map(cell =>
      (typeof cell === "string")
        ? toTitleCase(cell)
        : cell
    )
  );

  range.setValues(newValues);
}

// Helper function: Title Case
function toTitleCase(str) {
  return str
    .toLowerCase()
    .replace(/\b\w/g, c => c.toUpperCase());
}
