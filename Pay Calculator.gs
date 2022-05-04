//Work Calendar Vars
var workCalendarId = "bvpc5u0ramrjmcv64bi6947q6k@group.calendar.google.com";

//Spreadsheet Vars
var paySheet = SpreadsheetApp.openById("1_8-1yfn7sUtYccWq4QdS3LeeCURpYSK-azoQFv5efKo").getSheetByName("Pay Tracker");

//Time Vars
var firstPayday = new Date('April 9, 2021 18:00:00');

//Main Function
function payCalculator() 
{
  var sheetPayDay = new Date(paySheet.getRange("A2").getValue());
  sheetPayDay = standardizeHours(sheetPayDay);
  var currentPayDay = getCurrentPayday();

  Logger.log(currentPayDay);
  Logger.log(addDays(sheetPayDay, 14));

  var nextPayDay = addDays(sheetPayDay, 14);

  //Update sheet
  if(currentPayDay.getTime() >= addDays(sheetPayDay, 14).getTime())
  {
    Logger.log("Updating Sheet");

    nextPayDay = new Date(sheetPayDay.getTime() + miliDay * 14);

    //Shift cells down
    shiftCellsDown(paySheet, "A", 2);
    shiftCellsDown(paySheet, "B", 2);
    shiftCellsDown(paySheet, "C", 2);
    shiftCellsDown(paySheet, "D", 2);
    shiftCellsDown(paySheet, "E", 2);

    //Set date only if it needs updating
    paySheet.getRange("A2").setValue(nextPayDay);
  }
  else
  {
    Logger.log("Sheet up to date, skipping...");

    nextPayDay = new Date(sheetPayDay.getTime());
  }

  var payPeriod = getPayPeriod(nextPayDay, 2, 1);

  Logger.log(nextPayDay);
  Logger.log(payPeriod);

  //Set updated values regardless if row needed updating
  paySheet.getRange("B2").setValue(payPeriod[0]);
  paySheet.getRange("C2").setValue(payPeriod[payPeriod.length - 1]);
  paySheet.getRange("D2").setValue(paySheet.getRange("M2").getValue());
  paySheet.getRange("E2").setValue(getHours(workCalendarId, "Work", payPeriod));
}

function shiftCellsDown(spreadsheet, column, startingRow)
{
  var previousValue = "";

  var row = startingRow;
  while(true)
  {
    var currentCell = spreadsheet.getRange(column + row);
    var currentValue = currentCell.getValue();

    if(!currentCell.isBlank())
    {
      currentCell.setValue(previousValue);

      previousValue = currentValue;
      row++;
    }
    else
    {
      currentCell.setValue(previousValue);
      // spreadsheet.getRange(column + startingRow).setValue("");
      break;
    }

  }
}

//Get hours in week
function getHours(calendarId, eventTitle, payPeriodDays)
{
  //Return Var
  var totalDuration = 0;

  //Define vars
  var calendar = CalendarApp.getCalendarById(calendarId);
  var days = payPeriodDays;

  //Main loop
  for(var i = 0; i < days.length; i++)
  {
    var currentDay = days[i];
    var events = calendar.getEventsForDay(currentDay);
    
    //Delete events for day if there are multiple
    for(var k = 0; k < events.length; k++)
    {
      if(events[k].getTitle() == eventTitle)
      {

        var duration = (events[k].getEndTime() - events[k].getStartTime()) / 3600000;
        totalDuration += duration;
        
      }
    }
  }

  return totalDuration;
}

//Get weekdays
function getPayPeriod(paydayDate, weekCount, weekCorrectionShift)
{
  //Define Date Vars
  var today = new Date(paydayDate.getTime() - 19 * miliDay);
  var dayOfWeek = new Date(today.getTime() + (weekCorrectionShift * miliDay)).getDay();
  var firstDay = new Date(today.getTime() + (weekCorrectionShift * miliDay) - (dayOfWeek * miliDay));

  //Create return array
  var weekdays = new Array(7 * weekCount);

  for(var i = 0; i < weekdays.length; i++)
  {
    weekdays[i] = new Date(firstDay.getTime() + (i * miliDay));
  }

  //Return
  return weekdays;

  // //First day of new pay period
  // var firstDay = new Date(currentPayDay.getTime() - 19 * miliDay)

  // //Create return array
  // var days = new Array(14);

  // for(var i = 0; i < days.length; i++)
  // {
  //   days[i] = new Date(firstDay.getTime() + (i * miliDay));
  // }

  // //Return
  // return days;
}

//Get get the day of the next payday
function getCurrentPayday()
{
  //Var for while loop
  var previousPayday = firstPayday;
  var nextPayday;
  var today = new Date();

  //Find current paydate
  while(true)
  {
    nextPayday = new Date(previousPayday.getTime() + 14 * miliDay);

    if(nextPayday.getTime() > today.getTime())
    {
      break;
    }
    else
    {
      previousPayday = nextPayday;
    }
  }

  return nextPayday;
}

function addDays(inputDate, additionalDays)
{
  //Time Vars
  var miliDay = 1000 * 60 * 60 * 24;

  var outputDate = new Date(inputDate.getTime() + additionalDays * miliDay);
  outputDate = standardizeHours(outputDate);

  return outputDate;
}

function standardizeHours(inputDate)
{
  inputDate.setHours(17, 0, 0, 0);
  return inputDate;
}