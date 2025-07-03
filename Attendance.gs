function validateAndUpdateNames() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var attendanceSheet = ss.getSheetByName("Attendance");
  var updatesStudentsSheet = ss.getSheetByName("Updates-Students");

  // Create "Attendance" sheet if it doesn't exist
  if (!attendanceSheet) {
    attendanceSheet = ss.insertSheet("Attendance");
    attendanceSheet.getRange("A2").setValue("Name");
    attendanceSheet.getRange("B2").setValue("E-Mail");
    attendanceSheet.getRange("B1").setValue("Total Present");
  }

  // Get the names and emails from "Attendance"
  var attendanceLastRow = attendanceSheet.getLastRow();
  var attendanceNamesRange = attendanceSheet.getRange('A3:A' + attendanceLastRow);
  var attendanceNames = attendanceNamesRange.getValues().flat();
  var attendanceEmailsRange = attendanceSheet.getRange('B3:B' + attendanceLastRow);
  var attendanceEmails = attendanceEmailsRange.getValues().flat();

  // Get the names and emails from "Updates-Students"
  var updatesStudentsLastRow = updatesStudentsSheet.getLastRow();
  var updatesStudentsNamesRange = updatesStudentsSheet.getRange('A2:A' + updatesStudentsLastRow);
  var updatesStudentsEMailRange = updatesStudentsSheet.getRange('C2:C' + updatesStudentsLastRow);
  var updatesStudentsNames = updatesStudentsNamesRange.getValues().flat();
  var updatesStudentsEMail = updatesStudentsEMailRange.getValues().flat();

  // Ensure data integrity by removing blank entries
  var updatesData = [];
  for (var i = 0; i < updatesStudentsNames.length; i++) {
    if (updatesStudentsNames[i] && updatesStudentsEMail[i]) {
      updatesData.push([updatesStudentsNames[i], updatesStudentsEMail[i]]);
    }
  }

  // Compare the names and emails
  var dataToUpdate = updatesData.map(function(row) {
    return row;
  });

  var namesMatch = attendanceNames.every(function(name, index) {
    return name === (updatesData[index] ? updatesData[index][0] : "");
  });
  var emailsMatch = attendanceEmails.every(function(email, index) {
    return email === (updatesData[index] ? updatesData[index][1] : "");
  });

  // If names or emails do not match, copy from "Updates-Students" to "Attendance"
  if (!namesMatch || !emailsMatch) {
    attendanceSheet.getRange('A3:B' + attendanceLastRow).clearContent(); // Clear the existing names and emails in "Attendance" from A3 and B3 downwards
    attendanceSheet.getRange(3, 1, dataToUpdate.length, 2).setValues(dataToUpdate); // Update with new names and emails
  }

  attendanceSheet.getRange("A2").setValue("Name");
  attendanceSheet.getRange("B2").setValue("E-Mail");
  attendanceSheet.getRange("B1").setValue("Total Present");
}

function isThereClassToday(){
  var today = new Date();
  var day = today.getDay(); // getDay() returns 0 for Sunday, 1 for Monday, ..., 6 for Saturday

  // Check if the day is a weekday (1 to 5)
  if (day >= 1 && day <= 5) {
    // Perform the task only on weekdays
    putNewDate();
  }
}

function putNewDate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var attendanceSheet = ss.getSheetByName("Attendance");
  
  var today = new Date();
  var formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), "MM/dd/yyyy"); // Format the date
  
  var lastColumn = attendanceSheet.getLastColumn() + 1; // Determine the next empty column
  var headerCell = attendanceSheet.getRange(2, lastColumn); // Header cell for the day number
  headerCell.setValue(formattedDate + " Morning"); // Set the header value
  
  var headerCellNext = attendanceSheet.getRange(2, lastColumn + 1);
  headerCellNext.setValue(formattedDate + " Noon"); // Set the next header value
  
  // Add data validation and conditional formatting for the new columns
  addDropdownAndFormatting(attendanceSheet, lastColumn);
  addDropdownAndFormatting(attendanceSheet, lastColumn + 1);
}

function addDropdownAndFormatting(sheet, column) {
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(3, column, lastRow - 2); // Adjust the range to exclude headers
  
  // Add data validation
  var rule = SpreadsheetApp.newDataValidation().requireValueInList([, 'Present', 'Absent'], true).build();
  range.setDataValidation(rule);

  // Set default value "Absent"  
  range.setValue('Absent');

  // Apply conditional formatting
  var conditionalFormatRules = sheet.getConditionalFormatRules();
  
  // Rule for "Present" (Green)
  var presentRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Present')
    .setBackground('#A2D9A1') // Green
    .setRanges([range])
    .build();
  
  // Rule for "Absent" (Red)
  var absentRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Absent')
    .setBackground('#D2686E') // Red
    .setRanges([range])
    .build();
  
  // Add new rules to the existing ones
  conditionalFormatRules.push(presentRule);
  conditionalFormatRules.push(absentRule);
  
  // Apply the conditional formatting rules to the sheet
  sheet.setConditionalFormatRules(conditionalFormatRules);
}

function checkMorningAttendance(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var attendanceSheet = ss.getSheetByName("Attendance");

  var checkColumn = attendanceSheet.getLastColumn() - 1; // Corrected to use getLastColumn() as a function
  updatePresentCount(attendanceSheet, checkColumn);
}

function checkNoonAttendance() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var attendanceSheet = ss.getSheetByName("Attendance");

  var checkColumn = attendanceSheet.getLastColumn(); // Corrected to use getLastColumn() as a function
  updatePresentCount(attendanceSheet, checkColumn);
}

function updatePresentCount(sheet, column) {
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(3, column, lastRow - 2); // Adjust the range to exclude headers
  var values = range.getValues();
  
  var presentCount = values.flat().filter(function(value) {
    return value === 'Present';
  }).length;
  
  // Set the present count in the top row of the column
  sheet.getRange(1, column).setValue(presentCount);
}

function sendAttendanceMail(email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var attendanceSheet = ss.getSheetByName("Attendance");
  if (!attendanceSheet) {
    Logger.log("Attendance sheet not found.");
    return;
  }

  // Find the email in column B
  var emailRange = attendanceSheet.getRange('B:B');
  var emailValues = emailRange.getValues().flat();
  var rowIndex = emailValues.indexOf(email) + 1; // +1 because getValues() returns a 0-based index
  
  if (rowIndex === 0) {
    Logger.log("Email not found.");
    return;
  }

  // Iterate through the cells in that row from the 3rd column to the last column
  var lastColumn = attendanceSheet.getLastColumn();
  var rowRange = attendanceSheet.getRange(rowIndex, 3, 1, lastColumn - 2); // From column 3 to the last column
  var rowValues = rowRange.getValues()[0]; // Get the first (and only) row from the range

  // Count the number of "Present" entries
  var presentCount = rowValues.filter(function(value) {
    return value === 'Present';
  }).length;

  // Count the number of "Absent" entries
  var absentCount = rowValues.filter(function(value) {
    return value === 'Absent';
  }).length;
  
  // Calculate the percentage of "Present" entries
  var presentPercentage = (presentCount / (absentCount + presentCount)) * 100;

  // Send an email with the attendance percentage
  var subject = "Your Attendance Report";
  var body = "Hello " + attendanceSheet.getRange(rowIndex, 1).getValue() + ",\n\nYour attendance percentage is " + presentPercentage.toFixed(2) + "%.\n\nBest regards,\nTeam AI Vicharana Shala";

  // CC recipients
  var ccRecipients = ["cc1@example.com", "cc2@example.com"]; // Add CC recipients here
  
  var options = {
    cc: ccRecipients.join(',') // Join the CC recipients into a comma-separated string
  };
  
  MailApp.sendEmail(email, subject, body, options);
}

function testMailSending() {
  var testEmail = "2023eeb1212@iitrpr.ac.in"; // Replace with a test email
  sendAttendanceMail(testEmail);
}


// Function to process all students and send emails with their attendance percentage
function processAllStudentsAndSendEmails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var updatesStudentsSheet = ss.getSheetByName('Updates-Students');

  // Get all emails from Updates-Mentors
  var emailsRange = updatesStudentsSheet.getRange('C2:C' + updatesStudentsSheet.getLastRow());
  var emails = emailsRange.getValues().flat();

  // Iterate through all emails and process each student
  for (var i = 0; i < emails.length; i++) {
    var email = emails[i];
    if (email) {
      sendAttendanceMail(email);
    }
  }
}

function lockSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var protection = sheet.protect();
  protection.setWarningOnly(false); // Disable warning-only mode to enforce protection
  protection.removeEditors(protection.getEditors());
  Logger.log('Sheet locked at 9:05 AM');
}

function unlockSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Attendance");
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  for (var i = 0; i < protections.length; i++) {
    var protection = protections[i];
    protection.remove();
  }
  Logger.log('Sheet unlocked at 8:20 AM');
}

function createTriggers() {  
  // Create new triggers
  ScriptApp.newTrigger('unlockSheet')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .nearMinute(20)
    .create();

  ScriptApp.newTrigger('lockSheet')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .nearMinute(5)
    .create();

  ScriptApp.newTrigger('checkMorningAttendance')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .nearMinute(15)
    .create();

  ScriptApp.newTrigger('unlockSheet')
    .timeBased()
    .everyDays(1)
    .atHour(12)
    .nearMinute(5)
    .create();

  ScriptApp.newTrigger('lockSheet')
    .timeBased()
    .everyDays(1)
    .atHour(12)
    .nearMinute(55)
    .create();

    ScriptApp.newTrigger('checkNoonAttendance')
      .timeBased()
      .everyDays(1)
      .atHour(1)
      .nearMinute(15)
      .create();

    ScriptApp.newTrigger("processAllStudentsAndSendEmails")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY) // Choose the desired day of the week
    .atHour(10) // Set the desired time (24-hour format)
    .create();
}
