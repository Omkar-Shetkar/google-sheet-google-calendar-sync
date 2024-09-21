function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Sync Data with Calendar")
    .addItem("Import Calendar Events", "exportCalendarEventsToSheet")
    .addToUi();
}

function exportCalendarEventsToSheet() {
  const calendarId = "your-calendar-id"; // Replace with your calendar ID
  const startDate = new Date("2024-01-01"); // Set your desired start date
  const endDate = new Date("2024-12-31"); // Set your desired end date

  const calendar = CalendarApp.getCalendarById(calendarId);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const events = calendar.getEvents(startDate, endDate);
  const data = [];

  deleteDataRowsExceptHeader(); // Clear the sheet to populate with the latest event details

  if (events.length > 0) {
    for (let i = 0; i < events.length; i++) {
      const event = events[i];
      const eventID = event.getId();
      const eventTitle = event.getTitle();
      const startTime = event.getStartTime();
      const endTime = event.getEndTime();
      const description = event.getDescription();
      const color = event.getColor();

      data.push([eventID, eventTitle, startTime, endTime, description, color]);
    }
    // Clear previous data and set new values
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  } else {
    Logger.log("No events exist for the specified range");
  }
}

function deleteDataRowsExceptHeader() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); // Get the active sheet
  var lastRow = sheet.getLastRow(); // Get the last row number with content

  if (lastRow > 1) {
    // Check if there are rows below the header
    sheet.deleteRows(2, lastRow - 1); // Delete all rows from row 2 to the last row
  }
}
