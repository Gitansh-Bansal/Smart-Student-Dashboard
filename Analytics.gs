//Charts & Data Visualisation//

//Part 1//
function createBadgePieChart() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = spreadsheet.getSheetByName("Dashboard-Mentors");
  var targetSheet = spreadsheet.getSheetByName("Analytics");

  if (!sourceSheet || !targetSheet) {
    Logger.log("One or both sheets do not exist.");
    return;
  }

  var lastRow = sourceSheet.getLastRow();
  var badgeColumnIndex = 5; // Column E is index 5

  // Get the data range for badges
  var badgeRange = sourceSheet.getRange(2, badgeColumnIndex, lastRow - 1, 1); // Start from row 2 to exclude header
  var badgeValues = badgeRange.getValues();

  // Initialize badge counts object
  var badgeCounts = {
    'Emerald': 0,
    'Diamond': 0,
    'Copper': 0,
    'Bronze': 0,
    'Aluminium': 0
  };

  // Count the occurrences of each badge
  badgeValues.forEach(function(row) {
    var badgeName = row[0].trim();
    if (badgeName in badgeCounts) {
      badgeCounts[badgeName]++;
    }
  });

  // Prepare data for pie chart
  // var chartData = [['Badge', 'Count']];
  // for (var badge in badgeCounts) {
  //   if (badgeCounts[badge] > 0) {
  //     chartData.push([badge, badgeCounts[badge]]);
  //   }
  // }

  // Write the chart data to the "Analyt" sheet
  var dataRange = sourceSheet.getRange("L1:M");
  //dataRange.clearContent();
  //dataRange.setValues(chartData);

  // Create a pie chart
  var chartBuilder = targetSheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(sourceSheet.getRange("L1:M")) // Range for chart data
    .setOption('title', 'Badge Distribution')
    .setOption('width', 840)
    .setOption('height', 480)
    .setPosition(1, 1, 5, 5); // Position the chart in the sheet

  // Add inference note
  var inferenceNote = "This pie chart vividly displays the allocation of badges earned by students, reflecting their performance and achievements.";
  chartBuilder.setOption('subtitle', inferenceNote);
  
  // Build and insert the chart
  var chart = chartBuilder.build();
  targetSheet.insertChart(chart);
}

//Part 2//
function createChart_1() {
  // Open the spreadsheet by its ID
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Select the sheet from which to create the chart
  var sheet = ss.getSheetByName("Progress");

  // Select the destination sheet by its name
  var ds = ss.getSheetByName("Analytics");

  // Get the last row with data in columns A and E
  var lastRowA = sheet.getRange("A:A").getValues().filter(String).length;
  var lastRowE = sheet.getRange("E:E").getValues().filter(String).length;
  var lastRowF = sheet.getRange("F:F").getValues().filter(String).length;

  // Determine the maximum of these rows to get the upper bound for the range
  var lastRow = Math.max(lastRowA, lastRowE, lastRowF);

  // Define the range for the chart data dynamically
  var range1 = sheet.getRange(2, 1, lastRow); // Range for column A
  var range2 = sheet.getRange(2, 5, lastRow); // Range for column E
  var range3 = sheet.getRange(2, 6, lastRow); // Range for column F

  // Create the chart
  var chart = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(range1)
      .addRange(range2)
      .addRange(range3)
      .setPosition(25, 1, 5, 5) // Set the position where the chart will appear in the sheet
      .setOption('title', 'Percentage Progress of the Students vs Percentage Evaluated by the Mentors') // Set the chart title
      .setOption('subtitle', 'This graph highlights the percentage of questions attempted by students alongside the percentage of questions evaluated by mentors.')
      .setOption('hAxis.title', 'Name of the Mentors') // Set the x-axis title
      .setOption('vAxis.title', 'Percentage') // Set the y-axis title
      .setOption('series', {
        0: {labelInLegend: 'Percentage Progress'}, // Label for the first series
        1: {labelInLegend: 'Percentage Evaluated'}, // Label for the second series
        // 2: {labelInLegend: 'Percentage Evaluated'}  // Label for the third series
      })
      .setOption('width', 900)
      .setOption('height', 545)
      .build();

  // Insert the chart into the destination sheet
  ds.insertChart(chart);
}

//Part 3//
function createStudentColumnChart() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = spreadsheet.getSheetByName("Updates-Mentors");
  var targetSheet = spreadsheet.getSheetByName("Analytics");

  var data = sourceSheet.getRange(1, 5, sourceSheet.getMaxRows(), 1).getValues();
  var LastRow = 1;
  for (var i = data.length - 1; i >= 0; i--) {
    if (data[i][0] !== "") {
      LastRow = i + 1;
      break;
    }
  }

  // Define the range for the data starting from F217 to the end
  var startRow = LastRow;
  var startColumn = 5; // Column F
  var numRows = 1;
  var numColumns = sourceSheet.getLastColumn() - startColumn + 1;
  var dataRange = sourceSheet.getRange(startRow, startColumn, numRows, numColumns);

  // Create a new chart
  var chart = targetSheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(dataRange)
      .setPosition(25, 1, 5, 5)
      .setOption('title', 'Questions Attempted by No. of Students')
      .setOption('subtitle', ' This graph illustrates the number of students who attempted each question. From this visual representation, we can easily identify the most and least attempted questions, providing valuable insights into student engagement and question popularity.')
      .setOption('hAxis', {title: 'Which Question has how many attempts'})
      .setOption('vAxis', {title: 'No. of Students Attempted'})
      .setOption('height', 700) // Set the height of the chart (in pixels)
      .setOption('width', 1670) // Set the width of the chart (in pixels)
      .build();

  //  // Add inference note
  // var inferenceNote = " This graph illustrates the number of students who attempted each question. From this visual representation, we can easily identify the most and least attempted questions, providing valuable insights into student engagement and question popularity.";
  // chart.setOption('subtitle', inferenceNote);

  // Insert the chart into the target sheet
  targetSheet.insertChart(chart);
}

//Part 4//
function createStudentHistogramChart() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var updatesStudentsSheet = ss.getSheetByName("Updates-Students");
  var analytSheet = ss.getSheetByName('Analytics');

  // Determine the last row with data in column E
  var lastRow = updatesStudentsSheet.getLastRow();
  var dataRange = updatesStudentsSheet.getRange('E1:E' + lastRow);

  // Create a new chart builder
  var chartBuilder = analytSheet.newChart();

  // Set chart type to histogram chart
  chartBuilder
    .setChartType(Charts.ChartType.HISTOGRAM)
    .addRange(dataRange)
    .setOption('title', 'Distribution of Scores')
    .setOption('subtitle', 'This graph illustrates the distribution of students across different ranges of marks, offering a clear insight into student performance and score segmentation.')
    .setOption('hAxis', {title: 'Percentage Completion'} )
    .setOption('vAxis',{title: 'Number of Students'})
    .setPosition(1, 10, 5, 5)
    .setOption('height', 610)
    .setOption('width', 1080)
    .setOption('legend', { position: 'none' })
    .setOption('histogram', { bucketSize: 10}); // Set bucket size

  // Insert the chart into the "Analyt" sheet
  var chart = chartBuilder.build();
  analytSheet.insertChart(chart);
}

function all_student_graph() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get data from Dashboard-Mentors
  var sheetMentors = ss.getSheetByName("Dashboard-Mentors");
  var lastRowMentors = sheetMentors.getLastRow();
  var mentorsData = sheetMentors.getRange(2, 1, lastRowMentors - 1, 4).getValues(); // Get all data from columns A to D
  
  // Get data from Dashboard-Students
  var sheetStudents = ss.getSheetByName("Dashboard-Students");
  var lastRowStudents = sheetStudents.getLastRow();
  var studentsData = sheetStudents.getRange(2, 1, lastRowStudents - 1, 4).getValues(); // Get all data from columns A to D

  // Create a new sheet for the compiled data
  var finalSheet = ss.getSheetByName("Analytics");

  // Set headers for the compiled data
  finalSheet.getRange(1, 36, 1, 3).setValues([["Student Name", "Percentage Attempted", "Percentage Completed"]]);

  // Compile data
  var compiledData = [];
  var studentMap = {};

  // Add mentors data to map
  for (var i = 0; i < mentorsData.length; i++) {
    var name = mentorsData[i][0];
    if (!studentMap[name]) {
      studentMap[name] = {};
    }
    studentMap[name].completed = mentorsData[i][3];
  }

  // Add students data to map
  for (var j = 0; j < studentsData.length; j++) {
    var studentName = studentsData[j][0];
    if (!studentMap[studentName]) {
      studentMap[studentName] = {};
    }
    studentMap[studentName].attempted = studentsData[j][3];
  }

  // Prepare data for the final sheet
  for (var student in studentMap) {
    compiledData.push([
      student,
      studentMap[student].attempted || 0,
      studentMap[student].completed || 0
    ]);
  }

  // Write compiled data to the final sheet
  finalSheet.getRange(1, 36, compiledData.length, 3).setValues(compiledData);

  // Create the chart using the compiled data
  var range = finalSheet.getRange(1, 36, compiledData.length, 3);
  var chart = finalSheet.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(range)
      .setPosition(65, 1, 0, 0) // Set the position where the chart will appear in the sheet
      .setOption('title', 'Student Attempted vs Mentor Evaluated') // Set the chart title
      .setOption('subtitle', 'This graph highlights the percentage of questions attempted by students and the percentage of questions evaluated by mentors')
      .setOption('hAxis.title', 'Percentage') // Set the x-axis title
      .setOption('vAxis.title', 'Name of the Student') // Set the y-axis title
      .setOption('series', {
        0: {labelInLegend: 'Percentage Attempted'}, // Label for the first series (Attempted)
        1: {labelInLegend: 'Percentage Evaluated'}, // Label for the second series (Completed)
      })
      .setOption('width', 1250)
      .setOption('height', 3000) // Adjusted height for better readability
      .build();

  // Insert the chart into the final sheet
  finalSheet.insertChart(chart);
}
