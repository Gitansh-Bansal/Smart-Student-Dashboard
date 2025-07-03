function validateAndUpdateNamesLogger() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var regLogSheet = ss.getSheetByName("Reg-Log");
    var updatesMentorsSheet = ss.getSheetByName("Updates-Mentors");
    if (!regLogSheet) {
      regLogSheet = ss.insertSheet("Reg-Log");
      regLogSheet.getRange("A3").setValue("Reg No.");
      regLogSheet.getRange("A1").setValue("Average");
      regLogSheet.getRange("A2").setValue("Top 10 Average");
    }
    
    // Get the names from "Reg-Log" and "Updates-Mentors"
    var regLogNamesRange = regLogSheet.getRange('A4:A' + regLogSheet.getLastRow());
    var regLogNames = regLogNamesRange.getValues().flat();
    var updatesMentorsNamesRange = updatesMentorsSheet.getRange('B2:B' + updatesMentorsSheet.getLastRow());
    var updatesMentorsNames = updatesMentorsNamesRange.getValues().flat();
  
    // Compare the names
    var namesMatch = regLogNames.every(function(name, index) {
      return name === updatesMentorsNames[index];
    });
  
    // If names do not match, copy from "Updates-Mentors" to "Reg-Log"
    if (!namesMatch) {
      regLogSheet.getRange('A4:A').clearContent(); // Clear the existing names in "Reg-Log" from A3 downwards
      regLogSheet.getRange(4, 1, updatesMentorsNames.length, 1).setValues(updatesMentorsNames.map(function(name) {
        return [name];
      }));
    }
}
  
function logDailyYesCountsLogger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var updatesMentorsSheet = ss.getSheetByName("Updates-Mentors");
  var regLogSheet = ss.getSheetByName("Reg-Log");
    
  var today = new Date();
  var dayNumber = getDayNumberLogger(today);
  var lastColumn = regLogSheet.getLastColumn() + 1; // Determine the next empty column
  var headerCell = regLogSheet.getRange(3, lastColumn); // Header cell for the day number
  headerCell.setValue("Day " + dayNumber); // Set the header value
    
  var lastRow = updatesMentorsSheet.getRange('C:C').getValues().filter(String).length;  // Get the last row in "Updates-Mentors"
    
  for (var i = 2; i <= lastRow; i++) { // Start from the 2nd row
    var rowRange = updatesMentorsSheet.getRange(i, 6, 1, updatesMentorsSheet.getLastColumn() - 5); // Get range from column F onwards
    var rowValues = rowRange.getValues()[0]; // Get the row values
    var yesCount = rowValues.filter(value => value === 'Yes').length; // Count the number of 'Yes'
      
    // Write the count in the "Reg-Log" sheet
    if(yesCount){
      regLogSheet.getRange(i + 2, lastColumn).setValue(yesCount);
    } else {
      regLogSheet.getRange(i + 2, lastColumn).setValue(0.5);
    }
  }
}  
  
// Function to get the day number of the year
function getDayNumberLogger(d) {
  
  // ATTENTION !!
  var startDay = 141; // set this according to the day the program started
  
  var start = new Date(d.getFullYear(), 0, 0);
  var diff = (d - start) + ((start.getTimezoneOffset() - d.getTimezoneOffset()) * 60 * 1000);
  var oneDay = 1000 * 60 * 60 * 24;
  var day = Math.floor(diff / oneDay);
  return day - startDay;
}
  
function averageFinderLogger() {
  
  // NOTE: The function is such that it will find the average of the last column populated. 
  // Expect reuslts accordingly
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var regLogSheet = ss.getSheetByName("Reg-Log");

  // Check if "Reg-Log" sheet exists, create it if it doesn't
  if (!regLogSheet) {
    regLogSheet = ss.insertSheet("Reg-Log");
    regLogSheet.getRange("A1").setValue("Name");
  }
  
  // Get today's date
  var today = new Date();
  var formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");
  
  // Find the last column with value starting from B3
  var lastRow = regLogSheet.getLastRow();
  var lastColumn = regLogSheet.getLastColumn();
  var range = regLogSheet.getRange(3, 2, 1, lastColumn - 1);
  var values = range.getValues()[0];
  var lastPopulatedColumn = values.reduce((lastCol, value, index) => {
    if (value !== "") {
      lastCol = index + 2; // Adjust index to column number
    }
    return lastCol;
  }, 0);
  
  if (lastPopulatedColumn === 0) {
    Logger.log("No data found in row 3 onwards.");
    return;
  }
  
  // Calculate average of marks for the day
  var marksColumn = regLogSheet.getRange(4, lastPopulatedColumn, lastRow - 3, 1);
  var marksValues = marksColumn.getValues().flat().filter(value => !isNaN(value));
  var marksSum = marksValues.reduce((sum, mark) => sum + mark, 0);
  var marksAverage = marksSum / marksValues.length;
  
  // Calculate average of top 10 marks
  var sortedMarks = marksValues.sort((a, b) => b - a);
  var topTenMarks = sortedMarks.slice(0, 10);
  var topTenSum = topTenMarks.reduce((sum, mark) => sum + mark, 0);
  var topTenAverage = topTenSum / topTenMarks.length;
  
  // Write the averages to the sheet
  regLogSheet.getRange(1, lastPopulatedColumn).setValue(marksAverage);
  regLogSheet.getRange(2, lastPopulatedColumn).setValue(topTenAverage);
}

// Function to send email with chart attached
function sendEmailWithChartLogger(email, chartBlob, ccEmail) {
  var subject = 'Your Daily Performance Chart';
  var body = 'Hello,\n\nAttached is your daily performance chart.\n\nBest regards,\nTeam AI Vicharana Shala';
  MailApp.sendEmail({
    to: email,
    cc: ccEmail,
    subject: subject,
    body: body,
    attachments: [chartBlob]
  });
}

// Function to get user row by regNo
function getUserRowByRegNoLogger(regNo, sheet) {
  var data = sheet.getRange('A:A').getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === regNo) {
      return i + 1; // Adding 1 because Sheets is 1-indexed
    }
  }
  return -1; // User not found
}

// Function to get user details by regNo from Dashboard-Students
function getUserDetailsByRegNoLogger(regNo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboardSheet = ss.getSheetByName('Dashboard-Mentors');
  var data = dashboardSheet.getRange('B:B').getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === regNo) {
      var name = dashboardSheet.getRange(i + 1, 1).getValue();
      var rank = i + 1; // Rank based on the order in the sheet
      var badge = dashboardSheet.getRange(i + 1, 5).getValue(); // Assuming Badge is in column E
      return { name: name, rank: rank, badge: badge };
    }
  }
  return { name: '', rank: '', badge: '' };
}

// Function to create temporary graph and return the chart as a Blob
function createTempGraphLogger(regNo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var regLogSheet = ss.getSheetByName('Reg-Log');
  var tempGraphSheet = ss.getSheetByName('Temp-Graph') || ss.insertSheet('Temp-Graph');
  
  // Clear the Temp-Graph sheet before adding new data
  tempGraphSheet.clear();
  // Clear the graph made before making the new graph
  var charts = tempGraphSheet.getCharts();
  if (charts.length > 0) {
    // Assuming you want to delete the first chart in the sheet
    var chart = charts[0]; // Get the first chart
    tempGraphSheet.removeChart(chart); // Remove the chart from the sheet
    Logger.log('Chart deleted successfully.');
  } else {
    Logger.log('No charts found in the sheet.');
  }
  
  // Find the corresponding row in Reg-Log using the user's regNo
  var userRow = getUserRowByRegNoLogger(regNo, regLogSheet);
  
  if (userRow === -1) {
    // User not found, exit function or handle error
    Logger.log('User not found in Reg-Log');
    return;
  }
  
  // Get the user data row from Reg-Log
  var userDataRange = regLogSheet.getRange(userRow, 1, 1, regLogSheet.getLastColumn());
  var userData = userDataRange.getValues()[0];
  
  // Get the data from the first and second rows
  var row1Data = regLogSheet.getRange(1, 2, 1, regLogSheet.getLastColumn() - 1).getValues()[0];
  var row2Data = regLogSheet.getRange(2, 2, 1, regLogSheet.getLastColumn() - 1).getValues()[0];
  
  // Retrieve the user's name, rank, and badge from Dashboard-Students using the regNo
  var userDetails = getUserDetailsByRegNoLogger(regNo);
  var userName = userDetails.name;
  var userRank = userDetails.rank - 1;
  var userBadge = userDetails.badge;
  
  if (userName === '') {
    Logger.log('Reg No not found in Dashboard-Students');
    return;
  }
  
  // Set the Reg No in B1 and the user's name in B2
  tempGraphSheet.getRange(1, 2).setValue(regNo);
  tempGraphSheet.getRange(2, 2).setValue(userName);
  
  // Set "Day" in A column and enumerate
  for (var i = 0; i < userData.length - 1; i++) {
    tempGraphSheet.getRange(i + 3, 1).setValue(i + 1);
  }
  
  // Copy the user data to column B
  tempGraphSheet.getRange(3, 2, userData.length - 1, 1).setValues(userData.slice(1).map(item => [item]));
  
  // Copy row 1 data to column C
  tempGraphSheet.getRange(3, 3, row1Data.length, 1).setValues(row1Data.map(item => [item]));
  
  // Copy row 2 data to column D
  tempGraphSheet.getRange(3, 4, row2Data.length, 1).setValues(row2Data.map(item => [item]));
  tempGraphSheet.getRange("A2").setValue("Day");
  tempGraphSheet.getRange("B2").setValue("Your Performance");
  tempGraphSheet.getRange("C2").setValue("Average Performance");
  tempGraphSheet.getRange("D2").setValue("Top 10 Average Performance");
  
  // Create the chart
  var dataRange = tempGraphSheet.getRange('A2:D');
  var chartBuilder = tempGraphSheet.newChart()
  .setChartType(Charts.ChartType.LINE)
  .addRange(dataRange)
  .setPosition(5, 5, 0, 0)
  .setOption('title', "Daily Performance Analysis - " + userName + " (Rank: " + userRank + ", Badge: " + userBadge + ")")
  .setOption('hAxis', {
    title: 'Day',
    slantedText: true,
    slantedTextAngle: 45
  })
  .setOption('vAxis', {
    title: 'Performance Value'
  })
  .setOption('curveType', 'function') // Setting curve style to smooth
  .setOption('series', {
    0: { color: 'blue', labelInLegend: tempGraphSheet.getRange('B2').getValue() },
    1: { color: 'green', labelInLegend: tempGraphSheet.getRange('C2').getValue() },
    2: { color: 'red', labelInLegend: tempGraphSheet.getRange('D2').getValue() }
  })
  .setOption('legend', { position: 'top' })
  .setOption('annotations', {
    textStyle: {
      fontSize: 12,
      color: 'black'
    }
  })
  .setOption('colors', ['#3366CC', '#DC3912', '#FF9900'])
  .setOption('lineWidth', 3)
  .setOption('pointSize', 5)
  .setOption('animation', {
    duration: 1000,
    easing: 'out'
  })
  .setOption('backgroundColor', {
    fill: '#f1f8e9'
  });

var chart = chartBuilder.build();
tempGraphSheet.insertChart(chart);

  // Convert chart to Blob
  var chartBlob = chart.getAs('image/png').setName(userName + '-Performance-Chart.png');
  return chartBlob;
}

function findEmailLogger(regNo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var updatesStudentsSheet = ss.getSheetByName('Updates-Students');
  
  // Get all values in column B (registration numbers) and column C (emails)
  var regNos = updatesStudentsSheet.getRange('B2:B' + updatesStudentsSheet.getLastRow()).getValues();
  var emails = updatesStudentsSheet.getRange('C2:C' + updatesStudentsSheet.getLastRow()).getValues();

  // Iterate through regNos to find the matching regNo and return the corresponding email
  for (var i = 0; i < regNos.length; i++) {
    if (regNos[i][0] === regNo) {
      return emails[i][0]; // Return the email in the same row
    }
  }

  // Return a message if regNo is not found
  return 'Reg No not found';
}

// Function to process all students and send emails with their charts
function processAllStudentsAndSendEmailsLogger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var updatesMentorsSheet = ss.getSheetByName('Updates-Mentors');

  // Get all emails from Updates-Mentors
  var regNoRange = updatesMentorsSheet.getRange('B2:B' + updatesMentorsSheet.getLastRow());
  var regNos = regNoRange.getValues().flat();
  
  var ccEmail = "sudarshan@iitrpr.ac.in"; // Add the email address you want to CC
  
  // Iterate through all emails and process each student
  for (var i = 0; i < regNos.length; i++) {
    var regNo = regNos[i];
    if (regNo) {
      var chartBlob = createTempGraphLogger(regNo);
      var email = findEmailLogger(regNo);
      Logger.log(email);
      if (chartBlob) {
        sendEmailWithChartLogger(email, chartBlob, ccEmail);
      }
    }
  }
}





function mainLogger() {
  try {
    Logger.log("Starting validateAndUpdateNames...");
    validateAndUpdateNamesLogger();
    Logger.log("validateAndUpdateNames completed.");
 
    Logger.log("Starting logDailyYesCounts...");
    logDailyYesCountsLogger();
    Logger.log("logDailyYesCounts completed.");
  
    Logger.log("Starting averageFinder...");
    averageFinderLogger();
    Logger.log("averageFinder completed.");
  
    Logger.log("Starting processAllStudentsAndSendEmails...");
    processAllStudentsAndSendEmailsLogger();
    Logger.log("processAllStudentsAndSendEmails completed.");
      
  } catch (e) {
    Logger.log("Error in main function: " + e.toString());
  }
}

// Create time-driven trigger to run the main function daily at 8 AM
function createTimeDrivenTrigger() {
  ScriptApp.newTrigger('mainLogger')
    .timeBased()
    .atHour(10)
    .everyDays(1)
    .create();
}

// Optional: Function to delete all existing triggers (useful if you need to reset the trigger)
function deleteTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}
