const calendarId = ScriptProperties.getProperty('calendarID')
const calendar = CalendarApp.getCalendarById(calendarId);

function loadAddOn() {
  syncEmails()
}

// Check if the event is already there so we don't have duplicates if it reads a previously read email. 
function checkIfExists(date, startTime) {
  let events = calendar.getEventsForDay(new Date(date))
  let exists = false;
  if (events) {
    for (var i = 0; i < events.length; i++) {
      if (events[i].getStartTime().toISOString() === startTime.toISOString() || events[i].getTitle === "Scheduled Shift") {
        exists = true
        break
      }
    }
  }
  return exists
}

function syncEmails() {
  var searchQuery = "subject:Your New Schedule Has Been Published In Teamworx!";
  var threads = GmailApp.search(searchQuery);

  var message = threads[0].getMessages()[0].getPlainBody();

  // Matches MM/DD/YYYY, HH:MM AM|PM - HH:MM AM|PM
  var dateTimePattern = /(\d{1,2}\/\d{1,2}\/\d{4}, \d{1,2}:\d{2} (?:AM|PM) - \d{1,2}:\d{2} (?:AM|PM))/g;
  var dateTimes = message.match(dateTimePattern);

  if (dateTimes) {
    for (var i = 0; i < dateTimes.length; i++) {
      var dateTime = dateTimes[i];
      dateTime = dateTime.replace(/,/g, '');
      Logger.log("Found an event: " + dateTime);

      var day = dateTime.match(/(\d{1,2}\/\d{1,2}\/\d{4})/g)

      var dateTimeParts = dateTime.split(' - ');
      var startDateTime = new Date(dateTimeParts[0]);
      var endDateTime = new Date(`${day} ${dateTimeParts[1]}`);

      if (checkIfExists(day, startDateTime)) {
        Logger.log('It looks like this event already exists! Skipping...')
      } else {
        var eventTitle = "Scheduled Shift";
        var event = calendar.createEvent(eventTitle, startDateTime, endDateTime);
        Logger.log("Event created: " + event.getTitle() + " (" + event.getStartTime() + " - " + event.getEndTime() + ")");
      }

    }
  }
}

