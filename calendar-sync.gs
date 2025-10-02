/**
 * @OnlyCurrentDoc
 * This comment limits the script's authorization scope to the current spreadsheet,
 * enhancing security by ensuring the script cannot access other files.
 */

/**
 * Creates a custom menu in the Google Sheet UI when the spreadsheet is opened.
 * This provides an easy way for the user to run the script.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Calendar Sync')
    .addItem('Add Plan to Calendar', 'addEventsToCalendar')
    .addToUi();
}

/**
 * A mapping from the numeric color IDs provided in the sheet to the
 * corresponding color names available in the CalendarApp.EventColor enum.
 */
const COLOR_MAP = {
  '1': CalendarApp.EventColor.MAUVE,      // Lavender
  '2': CalendarApp.EventColor.PALE_GREEN, // Sage
  '3': CalendarApp.EventColor.MAUVE,      // Grape
  '4': CalendarApp.EventColor.PALE_RED,   // Flamingo
  '5': CalendarApp.EventColor.YELLOW,     // Banana
  '6': CalendarApp.EventColor.ORANGE,     // Tangerine
  '7': CalendarApp.EventColor.CYAN,       // Peacock
  '8': CalendarApp.EventColor.GRAY,       // Graphite
  '9': CalendarApp.EventColor.BLUE,       // Blueberry
  '10': CalendarApp.EventColor.GREEN,     // Basil
  '11': CalendarApp.EventColor.RED,       // Tomato
};

/**
 * Calculates the upcoming date for a given day of the week.
 * @param {string} dayName The name of the day (e.g., "Saturday").
 * @returns {Date} A Date object representing the next occurrence of that day.
 */
function getNextDateForDay(dayName) {
  const dayOfWeekMap = {
    'sunday': 0, 'monday': 1, 'tuesday': 2, 'wednesday': 3,
    'thursday': 4, 'friday': 5, 'saturday': 6
  };
  const targetDay = dayOfWeekMap[dayName.toLowerCase()];

  if (targetDay === undefined) {
    return null; // Invalid day name
  }

  const today = new Date();
  const currentDay = today.getDay();
  let dayDifference = targetDay - currentDay;

  if (dayDifference < 0) {
    dayDifference += 7;
  }

  const targetDate = new Date();
  targetDate.setDate(today.getDate() + dayDifference);
  return targetDate;
}


/**
 * Main function to read events from the sheet and add them to Google Calendar.
 */
function addEventsToCalendar() {
  const ui = SpreadsheetApp.getUi();
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getDisplayValues();

    const calendar = CalendarApp.getDefaultCalendar();
    let eventsAddedCount = 0;

    for (const row of data) {
      const [day, title, startTimeStr, endTimeStr, colorId, reminderMinutes] = row;

      if (!title) {
        continue;
      }
      Logger.log(`Processing: ${title}`);

      const eventDate = getNextDateForDay(day);
      if (!eventDate) {
        Logger.log(`Skipping row due to invalid day name: ${day}`);
        continue;
      }

      const [startHours, startMins] = startTimeStr.split(':').map(Number);
      const [endHours, endMins] = endTimeStr.split(':').map(Number);

      const startDateTime = new Date(eventDate);
      startDateTime.setHours(startHours, startMins, 0, 0);

      const endDateTime = new Date(eventDate);
      endDateTime.setHours(endHours, endMins, 0, 0);
      
      if (isNaN(startDateTime.getTime()) || isNaN(endDateTime.getTime())) {
          Logger.log(`Skipping row due to invalid time value for title: ${title}`);
          continue;
      }

      const recurrence = CalendarApp.newRecurrence().addWeeklyRule();

      const eventSeries = calendar.createEventSeries(title, startDateTime, endDateTime, recurrence);
      eventsAddedCount++;
      
      if (colorId && COLOR_MAP[colorId]) {
        eventSeries.setColor(COLOR_MAP[colorId]);
      }

      eventSeries.removeAllReminders();
      const reminder = parseInt(reminderMinutes, 10);
      if (!isNaN(reminder) && reminder > 0) {
        eventSeries.addPopupReminder(reminder);
      }
    }

    ui.alert('Success!', `Successfully added ${eventsAddedCount} events to your calendar.`, ui.ButtonSet.OK);

  } catch (e) {
    Logger.log(e);
    ui.alert('Error', 'An error occurred while adding events. Please check the logs for more details (Extensions > Apps Script > Executions).', ui.ButtonSet.OK);
  }
}
