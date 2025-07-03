function removeAllUsersExceptOwnerAndClearProtections() {
  // Get the active spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // Get the email address of the owner
  var ownerEmail = spreadsheet.getOwner().getEmail();
  
  // Remove all editors except the owner
  var editors = spreadsheet.getEditors();
  for (var i = 0; i < editors.length; i++) {
    var editorEmail = editors[i].getEmail();
    if (editorEmail !== ownerEmail) {
      spreadsheet.removeEditor(editorEmail);
    }
  }
  
  // Remove all viewers except the owner
  var viewers = spreadsheet.getViewers();
  for (var j = 0; j < viewers.length; j++) {
    var viewerEmail = viewers[j].getEmail();
    if (viewerEmail !== ownerEmail) {
      spreadsheet.removeViewer(viewerEmail);
    }
  }

  // Remove all protections across all sheets
  var sheets = spreadsheet.getSheets();
  for (var k = 0; k < sheets.length; k++) {
    var sheet = sheets[k];
    var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    for (var l = 0; l < protections.length; l++) {
      var protection = protections[l];
      protection.remove();
    }
    var rangeProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    for (var m = 0; m < rangeProtections.length; m++) {
      var rangeProtection = rangeProtections[m];
      rangeProtection.remove();
    }
  }

  Logger.log('All users except the owner have been removed and all protections cleared.');
}


function grantAdminAccess() {
  //Tested by Sudarshan. Works perfectly as it should. Ensure that you have included your email in the admin sheet.
  // Get the active spreadsheet and the "Admins" sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var adminsSheet = ss.getSheetByName("Admins");
  
  // Get the range of email addresses in the first column
  var emailRange = adminsSheet.getRange("A:A");
  var emailValues = emailRange.getValues();
  
  // Get all sheets in the spreadsheet
  var sheets = ss.getSheets();
  
  // Loop through each email address
  for (var i = 0; i < emailValues.length; i++) {
    var email = emailValues[i][0];
    
    // Check if the email is not empty
    if (email) {
      // Loop through each sheet
      for (var j = 0; j < sheets.length; j++) {
        var sheet = sheets[j];
        
        // Protect the sheet
        var protection = sheet.protect();
        protection.removeEditors(protection.getEditors()); // Remove existing editors
        
        // Grant edit access to the email for the current sheet
        protection.addEditor(email);
        
        // Optionally, you can unprotect the sheet for the current user
        protection.setUnprotectedRanges([sheet.getRange('A1')]);
      }
    }
  }
}



function removeUserFromSheet() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Enter the email address of the user to remove:', ui.ButtonSet.OK_CANCEL);
  
  // Process the user's response.
  if (response.getSelectedButton() == ui.Button.OK) {
    var email = response.getResponseText();
    
    if (email) {
      // Get the active spreadsheet
      var sheet = SpreadsheetApp.getActiveSpreadsheet();
      
      // Check if the email is a valid user of the sheet
      var currentEditors = sheet.getEditors();
      var userFound = false;
      for (var i = 0; i < currentEditors.length; i++) {
        if (currentEditors[i].getEmail() == email) {
          userFound = true;
          break;
        }
      }

      if (userFound) {
        // Unshare the spreadsheet with the user
        sheet.removeEditor(email);
        ui.alert('User ' + email + ' has been removed from the spreadsheet.');
      } else {
        ui.alert('User ' + email + ' is not an editor of the spreadsheet.');
      }
    } else {
      ui.alert('No email address entered.');
    }
  } else {
    ui.alert('Action canceled.');
  }
}



function onEdit(e) {
  logEvent('Edit', e);
}

// Custom function to log events
function logEvent(action, e) {
  // Open the "History" sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("History");
  
  // If the sheet doesn't exist, create it and add headers
  if (sheet == null) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("History");
    sheet.appendRow(['Date', 'Action', 'Sheet', 'Range', 'Value', 'User']); // Add headers
  }
  
  // Prepare the new row to add to the log
  var date = new Date();
  var editedSheet = e.source.getSheetName(); // Get the name of the sheet where the event happened
  var range = e.range.getA1Notation();
  var value = e.value || "N/A";
  var user = e.user ? e.user.getEmail() : "N/A";
  
  // Insert the new log entry at the second row (keeping headers at the first row)
  sheet.insertRowAfter(1);
  sheet.getRange('A2:F2').setValues([[date, action, editedSheet, range, value, user]]);
}

// Trigger to capture the row insert or delete event
function createSpreadsheetOpenTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('onChange')
      .forSpreadsheet(ss)
      .onChange()
      .create();
}

// Actual function that listens to changes for insert or delete
function onChange(e) {
  if(e.changeType === "INSERT_ROW") {
    logEvent("INSERT_ROW", e);
  } else if (e.changeType === "REMOVE_ROW") {
    logEvent("REMOVE_ROW", e);
  }
}

function sortAndCopyData_LogSheet() {
  copyPasteValue()
  // Access the 'Log' sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Log');
  
  // Get the data from the first 4 columns
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(2, 1, lastRow, 4).getValues();
  
  // Sort the data in descending order
  data.sort(function(a, b) {
    for (var i = 0; i < a.length; i++) {
      if (a[i] < b[i]) return 1;
      if (a[i] > b[i]) return -1;
    }
    return 0;
  });
  
  // Write the sorted data starting from F21
  sheet.getRange(2, 9, lastRow, 4).setValues(data);
}


function copyPasteValue() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Access the 'Progress' sheet and get the value from D18
  var progressSheet = ss.getSheetByName('Progress');
  var valueToCopy = progressSheet.getRange('D18').getValue();
  
  // Access the 'Log' sheet
  var logSheet = ss.getSheetByName('Log');
  
  // Find the last filled row in the second column
  var lastRow = logSheet.getLastRow();
  var lastValue = logSheet.getRange(lastRow, 2).getValue();
  
  // Check if the value to copy is different from the last value in the Log sheet
  if (valueToCopy !== lastValue) {
    // Add the current date and time to the first column
    logSheet.getRange(lastRow + 1, 1).setValue(new Date());
    
    // Paste the value to the second column
    logSheet.getRange(lastRow + 1, 2).setValue(valueToCopy);
    
    // Compute the new value for the third column based on the value just set in the second column
    var valueToLeft = logSheet.getRange(lastRow + 1, 2).getValue();
    var thirdColumnValue = -5579 + valueToLeft;
    
    // Set the new value in the third column
    logSheet.getRange(lastRow + 1, 3).setValue(thirdColumnValue);
  }
}




/**
 * Sets row-specific permissions on the 'Updates-Students' sheet.
 * Each student will have edit permission only for their respective row.
 * The owner of the spreadsheet retains edit rights for all rows.
 */
function grantRowPermissions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  const numRows = data.length;

  for (let i = 1; i < numRows; i++) { // Start from 1 to skip header row
    const email = data[i][2]; // Assuming email addresses are in column C (index 2)
    const rowRange = sheet.getRange(i + 1, 1, 1, sheet.getLastColumn()); // Range for the entire row

    // Create a new temporary sheet to copy the row data
    const tempSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(`R ${i+1}`);
    rowRange.copyTo(tempSheet.getRange(1, 1));

    // Share the temporary sheet with the specific email
    const protection = tempSheet.protect().setWarningOnly(true);
    //protection.removeEditors(protection.getEditors()); // Remove any default editors
    protection.addEditor(email); // Add the specific email as an editor

    // Optionally, you might want to hide the temporary sheet from the main UI
    tempSheet.hideSheet();
  }
}

  
function setRowPermissions() {
  // Get the active Google Spreadsheet instance.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Access the sheet named 'Updates-Students'.
  var sheet = ss.getSheetByName('Updates-Students');
  
  // Get the data range of the sheet that contains actual data.
  var range = sheet.getDataRange();
  
  // Retrieve all values present in the data range.
  var values = range.getValues();
  
  // Get the email of the spreadsheet owner.
  var ownerEmail = ss.getOwner().getEmail();
  
  // Log the owner email for debugging.
  Logger.log('Spreadsheet owner email: ' + ownerEmail);

  // Start iterating from the second row assuming the first row contains headers.
  for (var i = 1; i < values.length; i++) {
    // Fetch the email address from column C for the current row.
    var email = values[i][2];
    
    // Log the email for debugging.
    Logger.log('Processing row ' + (i + 1) + ' for email: ' + email);
    
    // Check if the email is valid.
    if (email && email.includes('@')) {
      try {
        // Determine the range of cells corresponding to the current student's data.
        var studentRange = sheet.getRange(i + 1, 1, 1, values[0].length);
        
        // Protect this range and add a description for easy identification.
        var studentProtection = studentRange.protect().setDescription('Protection for ' + email);
        
        // Remove all editors from the current protection to start afresh.
        studentProtection.removeEditors(studentProtection.getEditors());
        
        // Add the student as an editor to their own row.
        studentProtection.addEditor(email);
        
        // Ensure the spreadsheet's owner retains editing rights.
        if (!studentProtection.getEditors().includes(ownerEmail)) {
          studentProtection.addEditor(ownerEmail);
        }

        // Log success
        Logger.log('Row ' + (i + 1) + ' successfully protected and editor added: ' + email);
      } catch (e) {
        Logger.log('Error processing row ' + (i + 1) + ': ' + e.toString());
      }
    } else {
      Logger.log('Invalid email in row ' + (i + 1) + ': ' + email);
    }
  }
}





/**
 * Removes all range protections from the 'Updates-Students' sheet of the active spreadsheet.
 * 
 * @example
 * removeAllProtections();
 *
 * Note: Make sure you have the necessary permissions to modify protections before running this function.
 */
function removeAllProtections() {
  
  // Get the active spreadsheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Updates-Students');
  
  // Retrieve all range protections from the 'Updates-Students' sheet
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  
  // Loop through each protection
  for (var j = 0; j < protections.length; j++) {
    
    // Remove the current protection
    protections[j].remove();
  }
}




/**
 * This function color-codes rows in the "Ranking" sheet of the active spreadsheet. 
 * Rows 1-10 are colored light blue, 11-20 light orange, 21-30 light green, 
 * and 31-40 light pink. More colors can be added as needed.
 */
function colorCodeRows() {
  // Get the active spreadsheet object
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get the "Ranking" sheet from the active spreadsheet
  var sheet = ss.getSheetByName("Ranking");
  
  // Get the range of data present in the sheet
  var range = sheet.getDataRange();
  
  // Determine the number of rows in the data range
  var numRows = range.getNumRows();
  
  // Define an array of colors to be used for coloring rows
  var colorsArray = [
    "#E6F2FF", // Light blue for 1-10
    "#FFF3E6", // Light orange for 11-20
    "#F2FFE6", // Light green for 21-30
    "#FFE6F9", // Light pink for 31-40
    // Add more colors as needed
  ];
  
  // Variable to track the current index in the colors array
  var currentColorIndex = 0;
  
  // Loop through rows, starting from the second row and moving in increments of 10
  for (var i = 2; i <= numRows; i += 10) { 
    // Get a range of 10 rows and 4 columns for coloring
    var colorRange = sheet.getRange(i, 1, 10, 4); // Assuming 4 columns
    
    // Set the background color for the current range
    colorRange.setBackground(colorsArray[currentColorIndex]);
    
    // Increment the color index or reset it if it goes beyond the array's length
    currentColorIndex = (currentColorIndex + 1) % colorsArray.length;
  }
}

/**
 * Function to update the progress sheet with the average values of each mentor from the "Updates-Students" sheet.
 * It calculates the average progress value for each mentor and then updates the "Progress" sheet.
 */
function updateProgressSheet() {
  
  // Get a handle to the active Google spreadsheet.
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get a reference to the sheet named "Updates-Students".
  var updatesStudentsSheet = ss.getSheetByName("Updates-Students");
  
  // Retrieve all the values from this sheet.
  var data = updatesStudentsSheet.getDataRange().getValues();

  // Initialize dictionaries to accumulate the values for each mentor.
  var mentorData = {};
  
  // Initialize dictionaries to count the number of entries for each mentor.
  var mentorCounts = {};

  // Loop through all the rows of the "Updates-Students" sheet starting from the second row.
  for (var i = 1; i < data.length; i++) {
    
    // Extract the mentor's name from column D.
    var mentorName = data[i][3];
    
    // Extract the value from column E and convert it to a float directly since '%' sign has been removed.
    var value = parseFloat(data[i][4]);

    // If this mentor is already in our dictionary, update the cumulative value and increment the count.
    if (mentorData[mentorName]) {
      mentorData[mentorName] += value;
      mentorCounts[mentorName]++;
    } 
    // If this mentor is not yet in our dictionary, initialize them with the current value and count of 1.
    else {
      mentorData[mentorName] = value;
      mentorCounts[mentorName] = 1;
    }
  }

  // Get a reference to the sheet named "Progress".
  var progressSheet = ss.getSheetByName("Progress");
  
  // Retrieve all the values from the "Progress" sheet.
  var progressData = progressSheet.getDataRange().getValues();

  // Loop through all the rows of the "Progress" sheet starting from the second row.
  for (var j = 1; j < progressData.length; j++) {
    
    // Extract the mentor's name from column A.
    var mentorInProgress = progressData[j][0];

    // If this mentor is in our mentorData dictionary, then we have data to update for them.
    if (mentorData[mentorInProgress]) {
      
      // Calculate the average value for this mentor.
      var averageValue = mentorData[mentorInProgress] / mentorCounts[mentorInProgress];
      
      // Update column E in the "Progress" sheet with the average value.
      progressSheet.getRange(j + 1, 5).setValue(averageValue.toFixed(2) + "%");
    }
  }
  // Get a reference to the sheet named "Updates-Mentors".
  var updatesMentorsSheet = ss.getSheetByName("Updates-Mentors");
  
  // Retrieve all the values from this sheet.
  var data = updatesMentorsSheet.getDataRange().getValues();

  // Initialize dictionaries to accumulate the values for each mentor.
  var mentorData = {};
  
  // Initialize dictionaries to count the number of entries for each mentor.
  var mentorCounts = {};

  // Loop through all the rows of the "Updates-Mentors" sheet starting from the second row.
  for (var i = 1; i < data.length; i++) {
    
    // Extract the mentor's name from column D.
    var mentorName = data[i][3];
    
    // Extract the value from column E and convert it to a float directly since '%' sign has been removed.
    var value = parseFloat(data[i][4]);

    // If this mentor is already in our dictionary, update the cumulative value and increment the count.
    if (mentorData[mentorName]) {
      mentorData[mentorName] += value;
      mentorCounts[mentorName]++;
    } 
    // If this mentor is not yet in our dictionary, initialize them with the current value and count of 1.
    else {
      mentorData[mentorName] = value;
      mentorCounts[mentorName] = 1;
    }
  }

  // Get a reference to the sheet named "Progress".
  var progressSheet = ss.getSheetByName("Progress");
  
  // Retrieve all the values from the "Progress" sheet.
  var progressData = progressSheet.getDataRange().getValues();

  // Loop through all the rows of the "Progress" sheet starting from the second row.
  for (var j = 1; j < progressData.length; j++) {
    
    // Extract the mentor's name from column A.
    var mentorInProgress = progressData[j][0];

    // If this mentor is in our mentorData dictionary, then we have data to update for them.
    if (mentorData[mentorInProgress]) {
      
      // Calculate the average value for this mentor.
      var averageValue = mentorData[mentorInProgress] / mentorCounts[mentorInProgress];
      
      // Update column F in the "Progress" sheet with the average value.
      progressSheet.getRange(j + 1, 6).setValue(averageValue.toFixed(2) + "%");
    }
  }
}
function setDropdownsInRange() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var startRow = 2; // Starting row of your range
  var startCol = 6; // Starting column of your range
  var numRows = 100; // Number of rows in your range
  var numCols = 60; // Number of columns in your range
  var option = "No"; // The option you want to set

  var range = sheet.getRange(startRow, startCol, numRows, numCols);
  var values = range.getValues();

  for (var i = 0; i < numRows; i++) {
    for (var j = 0; j < numCols; j++) {
      if (values[i][j] != option) {
        values[i][j] = option;
      }
    }
  }

  range.setValues(values);
}

function setDropdowns() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getRange("G2:G100"); // Adjust this to the range where your dropdowns are.
  var option = "No"; // Replace this with the option you want to set.

  var values = range.getValues();
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] != option) {
      values[i][0] = option;
    }
  }
  range.setValues(values);
}
function change() {

}
function addDropdownToColumn() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getRange("D2:D50"); // Change "A1:A" to your column range
  var values = ["Ashutosh"]; // Change these to your dropdown values
  
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values)
    .setAllowInvalid(false)
    .build();
  
  range.setDataValidation(rule);
}

function setAllColumnsToValue() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getRange("F2:BM32"); // Change "A1:C10" to your rectangular range
  var value = "No"; // Change this to your desired value
  
  range.setValue(value);
}

function setPermissions(){
  //REMOVE ALL PROTECTIONS AND EDITORS
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();

  // Iterate through all sheets to remove protections
  sheets.forEach(function(sheet) {
    var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    protections.forEach(function(protection) {
      if (protection.canEdit()) {
        protection.remove();
      }
    });

    var sheetProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    sheetProtections.forEach(function(protection) {
      if (protection.canEdit()) {
        protection.remove();
      }
    });
  });

  // Ensure only the owner has access to the sheet
  var ownerEmail = spreadsheet.getOwner().getEmail();

  // Remove all other editors
  var editors = spreadsheet.getEditors();
  editors.forEach(function(editor) {
    if (editor.getEmail() !== ownerEmail && editor.getEmail() !== '2023mcb1294@iitrpr.ac.in') {
      spreadsheet.removeEditor(editor);
    }
  });

  // Remove all viewers
  var viewers = spreadsheet.getViewers();
  viewers.forEach(function(viewer) {
    if (viewer.getEmail() !== ownerEmail && viewer.getEmail() !== '2023mcb1294@iitrpr.ac.in') {
      spreadsheet.removeViewer(viewer);
    }
  });

    
  // Notify the user

  //SpreadsheetApp.getUi().alert('All protections removed and only the owner has access to the sheet.');

  // ADD EDIT ACCESS TO ADMINS AND STUDENTS\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  var additionalEditorsSheet = spreadsheet.getSheetByName('Admins');
  var additionalEditorEmails = additionalEditorsSheet.getRange("A1:A").getValues();
  var additionalEditors = additionalEditorEmails.map(function(row) { return row[0]; }).filter(function(email) { return email && email.includes('@'); });

  // Grant edit access to additional editors
  additionalEditors.forEach(function(email) {
    try {
      spreadsheet.addEditor(email);
      Logger.log("Added additional editor: " + email);
    } catch (e) {
      Logger.log("Failed to add additional editor: " + email);
    }
  });

  Logger.log("step-0");
  var activeSheet = spreadsheet.getSheetByName('Updates-Students');  
  var emailAddresses = activeSheet.getRange("C2:C").getValues().flat();
  Logger.log("step-1");
  Logger.log(emailAddresses);

  // Loop through the email addresses and grant edit access
  Logger.log("step-2");
  emailAddresses.forEach(function(email) {
    if (email) {
      try {
        spreadsheet.addEditor(email);
        Logger.log("Added editor: " + email);
      } catch (e) {
        Logger.log("Failed to add editor: " + email);
      }
    }
  });

  // GRANT SPECIAL ACCESS TO INDIVIDUAL STUDENTS AND ADMINS\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  var sheet = spreadsheet.getSheetByName('Updates-Students');
  var values = sheet.getDataRange().getValues();

  // Protect the first row and add additional editors
  var firstRowRange = sheet.getRange(1, 1, 1, values[0].length);
  var firstRowProtection = firstRowRange.protect().setDescription('Protection for first row');
  
  // Remove all existing editors
  firstRowProtection.removeEditors(firstRowProtection.getEditors());

  // Add additional editors to the first row
  additionalEditors.forEach(function(email) {
    try {
      firstRowProtection.addEditor(email);
      Logger.log("Added first row editor: " + email);
    } catch (e) {
      Logger.log("Failed to add first row editor: " + email);
    }
  });

  // Ensure the spreadsheet owner retains edit access to the first row
  if (!firstRowProtection.getEditors().includes(ownerEmail)) {
    firstRowProtection.addEditor(ownerEmail);
  }

  // Protect entire columns A to E for admins
  var adminProtectionRange = sheet.getRange(2, 1, sheet.getMaxRows() - 1, 5);
  var adminProtection = adminProtectionRange.protect().setDescription('Admin protection for columns A to E');
  adminProtection.removeEditors(adminProtection.getEditors());

  additionalEditors.forEach(function(email) {
    try {
      adminProtection.addEditor(email);
    } catch (e) {
      Logger.log("Failed to add admin editor: " + email);
    }
  });

  if (!adminProtection.getEditors().includes(ownerEmail)) {
    adminProtection.addEditor(ownerEmail);
  }

  // Loop through each row starting from the second row
  for (var i = 1; i < values.length; i++) {
    var email = values[i][2]; // Assuming email is in column C
    if (!email || !email.includes('@')) {
      continue; // Skip rows without a valid email
    }

    Logger.log('Processing row ' + (i + 1) + ' for email: ' + email);

    // Grant edit access to the student for columns F onward
    var studentRange = sheet.getRange(i + 1, 6, 1, values[0].length - 5); // Columns F onwards
    try {
      var studentProtection = studentRange.protect().setDescription('Editable columns for ' + email);
      studentProtection.removeEditors(studentProtection.getEditors());

      studentProtection.addEditor(email);

      if (!studentProtection.getEditors().includes(ownerEmail)) {
        studentProtection.addEditor(ownerEmail);
      }

      additionalEditors.forEach(function(additionalEditor) {
        if (!studentProtection.getEditors().includes(additionalEditor)) {
          studentProtection.addEditor(additionalEditor);
        }
      });

      Logger.log('Editable columns F onward for row ' + (i + 1) + ' successfully protected and editors set.');
    } catch (e) {
      Logger.log('Error processing editable columns for row ' + (i + 1) + ': ' + e.toString());
    }
  }

  // Flush changes to apply all updates
  SpreadsheetApp.flush();

  ///////////////////////////////////////////////////////////////////////////////////////////////////////

  //MENTOR UPDATE SHEET
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("Restricting Mentor Updates Sheet....")
  // Get the sheet where additional editor email addresses are stored
  var additionalEditorsSheet = spreadsheet.getSheetByName('Admins'); // Assuming the sheet is named 'AdditionalEditors'
  var additionalEditorEmails = additionalEditorsSheet.getRange("G1:G").getValues(); // Assuming emails are listed in column A starting from row 2
  var additionalEditors = additionalEditorEmails.map(function(row) { return row[0]; }).filter(function(email) { return email && email.includes('@'); });

  // Grant edit access to additional editors
  additionalEditors.forEach(function(email) {
    try {
      spreadsheet.addEditor(email);
      Logger.log("Added Editor: " + email);
    } catch (e) {
      Logger.log("Failed to add Editor: " + email);
    }
  });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetToRestrict = ss.getSheetByName('Updates-Mentors'); // Replace with the name of the sheet you want to restrict
  
  // Retrieve additional editor emails from the 'AdditionalEditors' sheet
  var editorsSheet = ss.getSheetByName('Admins'); // Assuming the sheet is named 'AdditionalEditors'
  var editorEmails = editorsSheet.getRange("G1:G").getValues(); // Assuming emails are listed in column A starting from row 2
  var additionalEditors = editorEmails.map(function(row) { return row[0]; }).filter(function(email) { return email && email.includes('@'); });

  // Protect the specified sheet
  var protection = sheetToRestrict.protect().setDescription('Restricted to AdditionalEditors');

  // Remove all existing editors to start fresh
  protection.removeEditors(protection.getEditors());

  // Grant edit access only to the additional editors
  additionalEditors.forEach(function(email) {
    try {
      protection.addEditor(email);
      Logger.log("Added Mentor: " + email);
    } catch (e) {
      Logger.log("Failed to add mentor: " + email + " - " + e.toString());
    }
  });

  // Ensure the spreadsheet owner retains edit access
  var ownerEmail = ss.getOwner().getEmail();
  if (!protection.getEditors().includes(ownerEmail)) {
    protection.addEditor(ownerEmail);
  }

  Logger.log('Sheet ' + sheetToRestrict.getName() + ' is now restricted to TAs.');

  //Attendance Sheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("Restricting Mentor Updates Sheet....")
  // Get the sheet where additional editor email addresses are stored
  var additionalEditorsSheet = spreadsheet.getSheetByName('Admins'); // Assuming the sheet is named 'AdditionalEditors'
  var additionalEditorEmails = additionalEditorsSheet.getRange("E1:E").getValues(); // Assuming emails are listed in column A starting from row 2
  var additionalEditors = additionalEditorEmails.map(function(row) { return row[0]; }).filter(function(email) { return email && email.includes('@'); });

  // Grant edit access to additional editors
  additionalEditors.forEach(function(email) {
    try {
      spreadsheet.addEditor(email);
      Logger.log("Added Editor: " + email);
    } catch (e) {
      Logger.log("Failed to add Editor: " + email);
    }
  });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetToRestrict = ss.getSheetByName('Attendance'); // Replace with the name of the sheet you want to restrict
  
  // Retrieve additional editor emails from the 'AdditionalEditors' sheet
  var editorsSheet = ss.getSheetByName('Admins'); // Assuming the sheet is named 'AdditionalEditors'
  var editorEmails = editorsSheet.getRange("E1:E").getValues(); // Assuming emails are listed in column A starting from row 2
  var additionalEditors = editorEmails.map(function(row) { return row[0]; }).filter(function(email) { return email && email.includes('@'); });

  // Protect the specified sheet
  var protection = sheetToRestrict.protect().setDescription('Restricted to AdditionalEditors');

  // Remove all existing editors to start fresh
  protection.removeEditors(protection.getEditors());

  // Grant edit access only to the additional editors
  additionalEditors.forEach(function(email) {
    try {
      protection.addEditor(email);
      Logger.log("Added Attendance Admin: " + email);
    } catch (e) {
      Logger.log("Failed to add attendance admin: " + email + " - " + e.toString());
    }
  });

  // Ensure the spreadsheet owner retains edit access
  var ownerEmail = ss.getOwner().getEmail();
  if (!protection.getEditors().includes(ownerEmail)) {
    protection.addEditor(ownerEmail);
  }

  Logger.log('Sheet ' + sheetToRestrict.getName() + ' is now restricted to Anuraag and Ved.');

  /////////////////////////////////////////////////////////////////////////////////////////////

  //DASHBOARDS
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("Restricting Dashboards....")
  // Get the sheet where additional editor email addresses are stored
  var additionalEditorsSheet = spreadsheet.getSheetByName('Admins'); // Assuming the sheet is named 'AdditionalEditors'
  var additionalEditorEmails = additionalEditorsSheet.getRange("C1:C").getValues(); // Assuming emails are listed in column A starting from row 2
  var additionalEditors = additionalEditorEmails.map(function(row) { return row[0]; }).filter(function(email) { return email && email.includes('@'); });

  // Grant edit access to additional editors
  additionalEditors.forEach(function(email) {
    try {
      spreadsheet.addEditor(email);
      Logger.log("Added editor: " + email);
    } catch (e) {
      Logger.log("Failed to add editor: " + email);
    }
  });


  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetsToRestrict = ['Dashboard-Students', 'Dashboard-Mentors', 'Reg-Log', 'Analytics', 'Peer-Instruction', 'Log', 'Progress', 'History']; // Replace with the names of the sheets you want to restrict
  
  // Retrieve additional editor emails from the 'AdditionalEditors' sheet
  var editorsSheet = ss.getSheetByName('Admins'); // Assuming the sheet is named 'AdditionalEditors'
  var editorEmails = editorsSheet.getRange("C1:C").getValues(); // Assuming emails are listed in column A starting from row 2
  var additionalEditors = editorEmails.map(function(row) { return row[0]; }).filter(function(email) { return email && email.includes('@'); });

  // Get the owner email
  var ownerEmail = ss.getOwner().getEmail();

  // Loop through each sheet to restrict
  sheetsToRestrict.forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      // Protect the specified sheet
      var protection = sheet.protect().setDescription('Restricted to AdditionalEditors');

      // Remove all existing editors to start fresh
      protection.removeEditors(protection.getEditors());

      // Grant edit access only to the additional editors
      additionalEditors.forEach(function(email) {
        try {
          protection.addEditor(email);
          Logger.log("Added dashboard editor: " + email);
        } catch (e) {
          Logger.log("Failed to add dashboard editor: " + email + " - " + e.toString());
        }
      });

      // Ensure the spreadsheet owner retains edit access
      if (!protection.getEditors().includes(ownerEmail)) {
        protection.addEditor(ownerEmail);
      }

      Logger.log('Sheet ' + sheetName + ' is now restricted.');
    } else {
      Logger.log('Sheet ' + sheetName + ' not found.');
    }
  });

  //ADMINS SHEET
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("Restricting Admins Sheet....")
  // Get the sheet where additional editor email addresses are stored
  var additionalEditorsSheet = spreadsheet.getSheetByName('Admins'); // Assuming the sheet is named 'AdditionalEditors'
  var additionalEditorEmails = additionalEditorsSheet.getRange("D1:D").getValues(); // Assuming emails are listed in column A starting from row 2
  var additionalEditors = additionalEditorEmails.map(function(row) { return row[0]; }).filter(function(email) { return email && email.includes('@'); });

  // Grant edit access to additional editors
  additionalEditors.forEach(function(email) {
    try {
      spreadsheet.addEditor(email);
      Logger.log("Added editor: " + email);
    } catch (e) {
      Logger.log("Failed to add editor: " + email);
    }
  });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetToRestrict = ss.getSheetByName('Admins'); // Replace with the name of the sheet you want to restrict
  
  // Retrieve additional editor emails from the 'AdditionalEditors' sheet
  var editorsSheet = ss.getSheetByName('Admins'); // Assuming the sheet is named 'AdditionalEditors'
  var editorEmails = editorsSheet.getRange("C1:C").getValues(); // Assuming emails are listed in column A starting from row 2
  var additionalEditors = editorEmails.map(function(row) { return row[0]; }).filter(function(email) { return email && email.includes('@'); });

  // Protect the specified sheet
  var protection = sheetToRestrict.protect().setDescription('Restricted to AdditionalEditors');

  // Remove all existing editors to start fresh
  protection.removeEditors(protection.getEditors());

  // Grant edit access only to the additional editors
  additionalEditors.forEach(function(email) {
    try {
      protection.addEditor(email);
      Logger.log("Added sheet editor: "+email);
    } catch (e) {
      Logger.log("Failed to add editor: " + email + " - " + e.toString());
    }
  });

  // Ensure the spreadsheet owner retains edit access
  var ownerEmail = ss.getOwner().getEmail();
  if (!protection.getEditors().includes(ownerEmail)) {
    protection.addEditor(ownerEmail);
  }

  Logger.log('Sheet ' + sheetToRestrict.getName() + ' is now restricted.');

  //adding viewers
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Admins'); // Change 'Sheet1' to the name of your sheet
  var emailRange = sheet.getRange('J1:J'); 
  var emails = emailRange.getValues();
  
  var uniqueEmails = [...new Set(emails.flat())].filter(String); // Flatten the array and remove duplicates

  var file = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());

  uniqueEmails.forEach(function(email) {
    file.addViewer(email);
  });
}

function removeDataValidation() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Updates-Mentors");
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  //var range = sheet.getRange(1, 66, 50, lastColumn - 5); 
  var range  = sheet.getRange('MU2:PM50');
  range.setDataValidation(null);
  range.clearContent();
  range.clearFormat();
  Logger.log('Data validation rules removed for range: ' + range.getA1Notation());
}

function addMentorDropdown(n=34) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Updates-Mentors');
  s="D2:D"+n.toString()
  Logger.log(s)
  var range = sheet.getRange(s); 
  var values = ["Ashutosh","Gitansh","Jatin","Khushi","Jaskirat","Nishit","Sreehitha","Vaishnavi"]; // Change these to your dropdown values
  
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values)
    .setAllowInvalid(false)
    .build();
  
  range.setDataValidation(rule);
}

function getActiveSheetName() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var sheetName = sheet.getName();
    Logger.log(sheetName);
}

function setActiveSheetByName(sheetName) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Updates-Students');
    SpreadsheetApp.setActiveSheet(sheet);
}


function giveViewAccess() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Admins'); // Change 'Sheet1' to the name of your sheet
  var emailRange = sheet.getRange('J1:J'); // Change 'A1:A' to the range where your emails are listed
  var emails = emailRange.getValues();
  
  var uniqueEmails = [...new Set(emails.flat())].filter(String); // Flatten the array and remove duplicates

  var file = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());

  uniqueEmails.forEach(function(email) {
    file.addViewer(email);
  });
}

function giveStudentPermissions(x1=42,x2=44){
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getSheetByName('Updates-Students');
  //make additional editors list
  var admin_sheet = ss.getSheetByName('Admins');
  var editorEmails = admin_sheet.getRange("A1:A").getValues(); // Assuming emails are listed in column A starting from row 2
  var additionalEditors = editorEmails.map(function(row) { return row[0]; }).filter(function(email) { return email && email.includes('@'); });

  for (var i=x1;i < (x2+1);i++) {
    //Logger.log(i)
    var i_s = i.toString()
    var currcell = 'C'+ i_s;
    var range_beg = 'F'+ i_s;
    var email_ = sheet.getRange(currcell).getValue();
    Logger.log(email_);
    ss.addEditor(email_);
    var range= sheet.getRange(range_beg+':'+i_s);
    //Logger.log(range.getValues());
    var protection = range.protect().setDescription('Protected range for ' + email_);
    protection.removeEditors(protection.getEditors());
    additionalEditors.forEach(function(email) {
      try {
        protection.addEditor(email);
        Logger.log("Added Admin: " + email);
      } catch (e) {
        Logger.log("Failed admin: " + email + " - " + e.toString());
      }
    });
    protection.addEditor(email_);
  }
}
