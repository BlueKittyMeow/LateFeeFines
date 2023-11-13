function calculateFees() {
  // Constants for office hours and late fee rate
  const OFFICE_HOURS = {
    'MONDAY': {start: 9, end: 19}, // represents 9:00 AM to 7:00 PM
    'TUESDAY': {start: 9, end: 19},
    'WEDNESDAY': {start: 9, end: 19},
    'THURSDAY': {start: 9, end: 19},
    'FRIDAY': {start: 9, end: 17}, // represents 9:00 AM to 5:00 PM
    // Closed days
    'SATURDAY': {start: null, end: null},
    'SUNDAY': {start: null, end: null}
  };
  const LATE_FEE_PER_HOUR = 0.50;

  // Get the active spreadsheet and the sheet you are working on
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Retrieve the due date, due time components, return date, return time components, and number of items from the spreadsheet
  var dueDateTime = new Date(sheet.getRange("A3").getValue());
  Logger.log('Due Date: ' + dueDateTime);

  // Construct the 12-hour time format from the individual components
  var dueTime = constructTime(sheet.getRange("B3").getValue(), sheet.getRange("C3").getValue(), sheet.getRange("D3").getValue());
  Logger.log('Due Time (12-hour format): ' + dueTime);

  var dueHoursMinutes = convertTo24HourFormat(dueTime).split(':');
  dueDateTime.setHours(dueHoursMinutes[0], dueHoursMinutes[1]); // set the hours and minutes for the due date
  Logger.log('Due Date and Time (24-hour format): ' + dueDateTime);

  var returnDateTime = new Date(sheet.getRange("E3").getValue());
  Logger.log('Return Date: ' + returnDateTime);

  // Construct the 12-hour time format from the individual components
  var returnTime = constructTime(sheet.getRange("F3").getValue(), sheet.getRange("G3").getValue(), sheet.getRange("H3").getValue());
  Logger.log('Return Time (12-hour format): ' + returnTime);

  var returnHoursMinutes = convertTo24HourFormat(returnTime).split(':');
  returnDateTime.setHours(returnHoursMinutes[0], returnHoursMinutes[1]); // set the hours and minutes for the return date
  Logger.log('Return Date and Time (24-hour format): ' + returnDateTime);

  var numberOfItems = sheet.getRange("I3").getValue();
  Logger.log('Number of Items: ' + numberOfItems);

  // Calculate the effective office hours late
  var effectiveHoursLate = calculateEffectiveHoursLate(dueDateTime, returnDateTime, OFFICE_HOURS);
  Logger.log('Effective Hours Late: ' + effectiveHoursLate);

  // Calculate the late fees
  var lateFees = effectiveHoursLate * LATE_FEE_PER_HOUR * numberOfItems;
  Logger.log('Late Fees: ' + lateFees);

  // Output the results to the spreadsheet
  sheet.getRange("J3").setValue(effectiveHoursLate); // Effective Office Hours Late
  sheet.getRange("K3").setValue(lateFees); // Late Fees
}

// This function constructs a time string in 12-hour format from individual components
function constructTime(hour, minute, ampm) {
  return hour + ':' + minute + ' ' + ampm;
}

// This function converts a time string from 12-hour format to 24-hour format
function convertTo24HourFormat(time12h) {
  const [time, modifier] = time12h.split(' ');
  let [hours, minutes] = time.split(':');

  if (hours === '12') {
    hours = '00';
  }

  if (modifier === 'PM') {
    hours = parseInt(hours, 10) + 12;
  }

  return `${hours}:${minutes}`;
}

function calculateEffectiveHoursLate(dueDateTime, returnDateTime, officeHours) {
  // This function calculates the effective hours late based on office hours
  // The function handles different days of the week, office hours, and excludes weekends

  // Check if the due date and return date are the same
  if (dueDateTime.toDateString() === returnDateTime.toDateString()) {
    // Same day, so we calculate the hours based on the time difference today
    var hoursLate = (returnDateTime - dueDateTime) / 36e5; // milliseconds to hours
    return Math.ceil(hoursLate); // rounding up to the nearest hour
  }

  var effectiveHoursLate = 0;
  var currentDateTime = new Date(dueDateTime); // cloning the due date

  // Flag to check if it's the first day (due date)
  var isFirstDay = true;

  Logger.log('Calculating effective hours late between: ' + dueDateTime + ' and ' + returnDateTime);

  while (currentDateTime < returnDateTime) {
    var dayOfWeek = ['SUNDAY', 'MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY', 'SATURDAY'][currentDateTime.getDay()];
    Logger.log('Current Day: ' + dayOfWeek + ', Date and Time: ' + currentDateTime);

    if (officeHours[dayOfWeek].start != null) { // if the office is open that day
      var startOfDay = new Date(currentDateTime);
      startOfDay.setHours(officeHours[dayOfWeek].start, 0, 0, 0); // start of office hours

      var endOfDay = new Date(currentDateTime);
      endOfDay.setHours(officeHours[dayOfWeek].end, 0, 0, 0); // end of office hours

      Logger.log('Office Hours: ' + startOfDay + ' to ' + endOfDay);

      if (currentDateTime < endOfDay) { // if current time is before office closing time
        if (isFirstDay) {
          // If it's the first day, we calculate hours based on due time and office closing time
          var hoursLateFirstDay = (endOfDay - currentDateTime) / 36e5; // milliseconds to hours
          effectiveHoursLate += hoursLateFirstDay;
          Logger.log('First day, hours late added: ' + hoursLateFirstDay);
          isFirstDay = false; // reset the flag
        } else if (returnDateTime < endOfDay) {
          // if the item is returned before the end of the workday
          // Corrected the start time for calculation to the office start time instead of midnight
          effectiveHoursLate += (returnDateTime - startOfDay) / 36e5; // milliseconds to hours
          Logger.log('Returned same day, hours late added: ' + effectiveHoursLate);
        } else {
          // here we need to add only the total office hours for a full day
          var totalOfficeHoursToday = (endOfDay - startOfDay) / 36e5; // milliseconds to hours
          effectiveHoursLate += totalOfficeHoursToday;
          Logger.log('Returned another day, total office hours today: ' + totalOfficeHoursToday);
        }
      }
    }

    currentDateTime.setDate(currentDateTime.getDate() + 1);
    currentDateTime.setHours(0, 0, 0, 0); // reset to start of the day
  }

  effectiveHoursLate = Math.ceil(effectiveHoursLate); // rounding up to the nearest hour
  Logger.log('Total effective hours late: ' + effectiveHoursLate);
  return effectiveHoursLate;
}

function resetForm() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Clear the content of cells; this doesn't remove formatting or data validation
  sheet.getRange('A3').clearContent();
  sheet.getRange('E3:H3').clearContent();
  sheet.getRange('J3:K3').clearContent();

  // Set default values
  sheet.getRange('B3').setValue(1);
  sheet.getRange('C3').setValue('00');
  sheet.getRange('D3').setValue('PM');
  sheet.getRange('I3').setValue(1);
}
