function checkAllBirthdays() {
  const sheet = getSheetByName("Master");
  if (!sheet) return;

  const headers = getHeaders(sheet);
  const targetColumns = ["First Name", "Last Name", "School", "Current Grade Level", "Birth Date"];
  const columnIndexes = getColumnIndexes(headers, targetColumns);

  if (!columnIndexes || !columnIndexes.includes(headers.indexOf("Birth Date") + 1)) {
    Logger.log("Ensure the 'Birth Date' column exists in the header.");
    return;
  }

  const data = getColumnData(sheet, columnIndexes);
  const allBirthdays = getAllBirthdays(data, targetColumns, "Birth Date");

  addBirthdaysToCalendar(allBirthdays);  // Add all birthdays to the calendar
}

function getSheetByName(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`Sheet '${sheetName}' not found.`);
  }
  return sheet;
}

function getHeaders(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function getColumnIndexes(headers, targetColumns) {
  return targetColumns.map(colName => {
    const index = headers.findIndex(header => header === colName);
    if (index === -1) {
      Logger.log(`Column '${colName}' not found.`);
    }
    return index + 1; // Convert 0-based to 1-based index
  }).filter(index => index > 0); // Exclude missing columns
}

function getColumnData(sheet, columnIndexes) {
  const lastRow = sheet.getLastRow();
  return columnIndexes.map(colIndex =>
    sheet.getRange(2, colIndex, lastRow - 1).getValues().flat()
  );
}

function getAllBirthdays(data, targetColumns, dateColumnName) {
  const results = [];
  const dateColIndex = targetColumns.indexOf(dateColumnName);
  if (dateColIndex === -1) {
    Logger.log(`Date column '${dateColumnName}' not found in target columns.`);
    return [];
  }

  // Loop through each row and gather birthday information
  for (let rowIndex = 0; rowIndex < data[dateColIndex].length; rowIndex++) {
    const cellDate = data[dateColIndex][rowIndex];
    if (cellDate instanceof Date) {
      let rowData = {};
      targetColumns.forEach((colName, colIndex) => {
        rowData[colName] = data[colIndex][rowIndex] || ""; // Handle empty cells
      });
      results.push(rowData);
    }
  }
  return results;
}
function addBirthdaysToCalendar(birthdays) {
  let calendar = CalendarApp.getCalendarsByName("OFY Student Birthdays")[0];
  
  // If the "Birthdays" calendar does not exist, create a new one
  if (!calendar) {
    Logger.log("No calendar named 'OFY Student Birthdays' found.");
    // Uncomment to create a new calendar if needed
    calendar = CalendarApp.createCalendar("OFY Student Birthdays");
    Logger.log("Created new calendar 'OFY Student Birthdays'.");
  }

  // Get the current year
  const currentYear = new Date().getFullYear();

  // Add each birthday as an all-day event for this year and the next five years
  birthdays.forEach(birthday => {
    const firstName = birthday["First Name"];
    const lastName = birthday["Last Name"];
    const fullName = firstName + " " + lastName;  // Combining First and Last Name
    const birthDate = new Date(birthday["Birth Date"]);
    let gradeLevel = birthday["Current Grade Level"]; // Capture the grade level

    // For each year from current year to the next 5 years
    for (let yearOffset = 0; yearOffset <= 5; yearOffset++) {
      // Set the birth year to the current year + offset (this year and the next 5 years)
      const eventDate = new Date(birthDate);
      eventDate.setFullYear(currentYear + yearOffset);

      // Make sure the event's date is correct, even if the month/day is invalid for leap years, etc.
      if (eventDate.getDate() !== birthDate.getDate()) {
        eventDate.setMonth(birthDate.getMonth() + 1); // Adjust if the date changes due to leap year
      }

      // Increment gradeLevel after each year, and if greater than 12, set to "Graduate"
      gradeLevel = (gradeLevel < 12) ? gradeLevel + 1 : "Graduate";

      const school = birthday["School"];
      const eventDescription = `School: ${school}\nGrade Level: ${gradeLevel}`;

      // Check if an event already exists on the same date
      const existingEvents = calendar.getEventsForDay(eventDate);
      const eventExists = existingEvents.some(event => event.getTitle() === fullName + "'s Birthday");

      if (!eventExists) {
        // Create an all-day event for the birthday if it doesn't already exist
        calendar.createAllDayEvent(fullName + "'s Birthday", eventDate, {description: eventDescription});
        Logger.log(`Added ${fullName}'s birthday for ${eventDate.toDateString()} to the calendar.`);
      } else {
        Logger.log(`Birthday event for ${fullName} already exists on ${eventDate.toDateString()}. Skipping.`);
      }
    }
  });
}


