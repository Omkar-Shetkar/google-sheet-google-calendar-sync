const CALENDAR_ID = ""; // Replace with your calendar ID
const CALENDAR = CalendarApp.getCalendarById(CALENDAR_ID);

const SHEET = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

const COLUMN = {
  ID: 0,
  TITLE: 1,
  DATE: 2,
  AMOUNT: 3,
};

const DATE_RANGE_START = new Date("2024-01-01");
const DATE_RANGE_END = new Date("2024-12-31");

function updateSheet() {
  var data = SHEET.getDataRange().getValues();

  const events = CALENDAR.getEvents(DATE_RANGE_START, DATE_RANGE_END);

  const eventMap = new Map(); // Cache of events

  events.forEach((event) => {
    eventMap.set(event.getId(), event);
  });

  eventMap.forEach((entry) =>
    Logger.log(entry.getId + " " + entry.getDescription())
  );

  const deleteRows = [];

  for (var i = 1; i < data.length; i++) {
    // Loop through rows by skipping header row. Rows and Columns start at index 0.
    var eventId = data[i][COLUMN.ID];
    var event = eventMap.get(eventId);
    if (!event) {
      // Event doesn't exist in calendar, delete the row
      deleteRows.push(i + 1);
      continue;
    }
    SHEET.getRange(i + 1, COLUMN.AMOUNT + 1).setValue(event.getDescription()); // Row and Column for range start at index 1.
    eventMap.delete(eventId); // Done with update, remove the event from cache
  }

  deleteRows.forEach((row) => SHEET.deleteRow(row));

  // Append remaining events in the cache to sheet
  var nextRow = SHEET.getLastRow() + 1;
  var startColumn = 1;
  var remainingRows = eventMap.size;
  var totalColumns = data[0].length;

  const newEvents = [];

  if (eventMap.size != 0) {
    eventMap.forEach((entry) => {
      const event = entry;
      const eventID = event.getId();
      const eventTitle = event.getTitle();
      const startTime = event.getStartTime();
      const description = event.getDescription();

      newEvents.push([eventID, eventTitle, startTime, description]);
    });

    SHEET.getRange(nextRow, startColumn, remainingRows, totalColumns).setValues(
      newEvents
    );
  }
}

function importCalendarEvents() {
  const events = CALENDAR.getEvents(DATE_RANGE_START, DATE_RANGE_END);
  const data = [];

  deleteDataRowsExceptHeader(); // Clear the sheet to populate with the latest event details

  SHEET.getRowGroup;
  if (events.length > 0) {
    for (let i = 0; i < events.length; i++) {
      const event = events[i];
      const eventID = event.getId();
      const eventTitle = event.getTitle();
      const startTime = event.getStartTime();
      const description = event.getDescription();

      data.push([eventID, eventTitle, startTime, description]);
    }
    // Clear previous data and set new values
    SHEET.getRange(2, 1, data.length, data[0].length).setValues(data);
  } else {
    Logger.log("No events exist for the specified range");
  }
}

function deleteDataRowsExceptHeader() {
  var lastRow = SHEET.getLastRow(); // Get the last row number with content

  if (lastRow > 1) {
    // Check if there are rows below the header
    SHEET.deleteRows(2, lastRow - 1); // Delete all rows from row 2 to the last row
  }
}
