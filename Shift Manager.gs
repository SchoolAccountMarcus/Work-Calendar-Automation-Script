//Personal Calendar Vars
var personalCalendarId = "mquantwenty22@gmail.com";

//Work Calendar Vars
var workCalendarId = "bvpc5u0ramrjmcv64bi6947q6k@group.calendar.google.com";
var shiftTimeCellRange = "A9:E15";
var workLocation = "Smashburger";

//Time Vars
var miliDay = 1000 * 60 * 60 * 24;
var weekdayTargetOffset = 4;

//Main Function
function shiftManager()
{

  //Clear events
  clearWeekEvents(workCalendarId, "Work");

  //Update events
  updateWorkShifts();

  //Update Conflicts Text
  returnTextMessage = "Conflicts Next Week:\n\n-------------------------------" + eventsToText(personalCalendarId, 7, false) + "-------------------------------";

  //Update Discord Text
  var disReturnTextMessage = eventsToText(workCalendarId, 0, true);

  //Update Spreadsheet
  SpreadsheetApp.openById("1_8-1yfn7sUtYccWq4QdS3LeeCURpYSK-azoQFv5efKo").getSheetByName("Schedule Manager").getRange("B16").setValue(returnTextMessage);
  SpreadsheetApp.openById("1_8-1yfn7sUtYccWq4QdS3LeeCURpYSK-azoQFv5efKo").getSheetByName("Schedule Manager").getRange("D16").setValue(disReturnTextMessage);

}

//Generate weeks schedule
function eventsToText(calendarID, dayOffset, disFormat)
{
  Logger.log("START: eventsToText");

  //Define the calendar and weekdays array
  var cal = CalendarApp.getCalendarById(calendarID);
  var weekdays = getWeekdays(dayOffset, 1, weekdayTargetOffset);

  //Define time zone
  var timeZone = "America/New_York";

  //Define return message
  var returnTextMessage = "";

  //Iterate through days in the weekday array defined above
  for(var day = 0; day < weekdays.length; day++)
  {
    //Define current day vars
    var currentDay = new Date(weekdays[day].getTime());
    Logger.log("Day: " + currentDay);
    var events = cal.getEventsForDay(currentDay);

    //Add header
    if(disFormat) returnTextMessage += Utilities.formatDate(currentDay, timeZone, "\n> **EEEE** - MMMM dd\n> ```");
    else returnTextMessage += Utilities.formatDate(currentDay, timeZone, "\nEEEE - MMMM dd\n");

    //Add each event for the day
    var returnDayMessage = "";
    var isDayFree = true;
    for(var dayIndex = 0; dayIndex < events.length; dayIndex++)
    {
      //Define the free day variable
      isDayFree = false;

      //Get the event variable
      var event = events[dayIndex];

      //Get the event title, start time and end time
      var title = event.getTitle();
      var start = Utilities.formatDate(event.getStartTime(), timeZone, "hh:mm aa");
      var end = Utilities.formatDate(event.getEndTime(), timeZone, "hh:mm aa");

      //Append the message based on discord or normal message
      if(disFormat) returnDayMessage += "\n> " + start + " - " + end + "  | " + title;
      else returnDayMessage += title + ": " + start + " to " + end + "\n";
    }

    //If day is free then add that for the text
    if(isDayFree && !disFormat)
    {
      returnDayMessage += " - Free\n";
    }

    //Close code block
    if(disFormat && returnDayMessage == "") returnTextMessage += " ```";
    else if(disFormat) returnTextMessage += returnDayMessage + "```";
    else returnTextMessage += returnDayMessage;
  }

  //Return the message
  Logger.log(returnTextMessage);
  return returnTextMessage;
} 

//Clear week events
function clearWeekEvents(calendarId, eventTitle)
{
  Logger.log("START: clearWeekEvents");
  Logger.log("Deleting events titled: " + eventTitle);

  //Define vars
  var calendar = CalendarApp.getCalendarById(calendarId);
  var weekdays = getWeekdays(0, 1, weekdayTargetOffset);

  //Main deleting loop
  for(var i = 0; i < weekdays.length; i++)
  {
    var currentDay = weekdays[i];
    var events = calendar.getEventsForDay(currentDay);
    Logger.log("Current Day: " + currentDay);
    // Logger.log(events.length + " for the day");
    
    //Delete events for day if there are multiple
    for(var k = 0; k < events.length; k++)
    {
      if(events[k].getTitle() == eventTitle)
      {
        Logger.log("Deleting event " + events[k].getTitle() + "(" + events[k].getId() + ")");
        events[k].deleteEvent();
      }
    }
  }

  Logger.log("END: clearWeekEvents");
}

//Main work shifts function
function updateWorkShifts() 
{
  Logger.log("START: updateWorkShifts");

  //Define spreadsheet and get calendar id
  var spreadsheet = SpreadsheetApp.getActiveSheet();

  //Get info from spreadsheet
  var workCal = CalendarApp.getCalendarById(workCalendarId);
  var signups = spreadsheet.getRange(shiftTimeCellRange).getValues();

  //Iterate through days and create events
  for(var x = 0; x < signups.length; x++)
  {

    var shift = signups[x];

    var startTimeShift1 = shift[1];
    var endTimeShift1 = shift[2];

    var startTimeShift2 = shift[3];
    var endTimeShift2 = shift[4];

    createFormatEvent(workCal, workLocation, startTimeShift1, endTimeShift1);
    createFormatEvent(workCal, workLocation, startTimeShift2, endTimeShift2);

  }

  Logger.log("END: updateWorkShifts");
}

//Work shift support function
function createFormatEvent(eventCal, workLocation, startTime, endTime)
{
  Logger.log("START: createFormatEvent");
  
  //Check to see if the shift is "FREE"
  if(startTime != "FREE" && endTime != "FREE")
    {
      //Define start and end date
      var startDate = new Date(startTime);
      var endDate = new Date(endTime);

      Logger.log(startTime);
      Logger.log(endTime);

      //Create event
      var event = eventCal.createEvent("Work", startDate, endDate, {location: workLocation});

      Logger.log("Event ID: " + event.getId());
    }
    else
    {
      Logger.log("DAY HAS NO SHIFT");
    }

    Logger.log("END: createFormatEvent");
}

//Get weekdays
function getWeekdays(dayOffset, weekCount, weekCorrectionShift)
{
  //Define Date Vars
  var today = new Date();
  var dayOfWeek = new Date(today.getTime() + (weekCorrectionShift * miliDay)).getDay();
  var firstDay = new Date(today.getTime() + (weekCorrectionShift * miliDay) - (dayOfWeek * miliDay));

  //Create return array
  var weekdays = new Array(7 * weekCount);

  for(var i = 0; i < weekdays.length; i++)
  {
    weekdays[i] = new Date(firstDay.getTime() + ((i + dayOffset) * miliDay));
  }

  //Return
  return weekdays;
}