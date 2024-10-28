// 1. Extract application data from the SEF form spreadsheet
// 2. Write the SEF minute
// 3. Write the SEF tracker

// Manual change is required
var docName = "Fall 2024";

var scriptProperties = PropertiesService.getScriptProperties();

var sefFormId = scriptProperties.getProperty('SEF_FORM_ID');
var sefMinutesTemplateId = scriptProperties.getProperty('SEF_MINUTES_TEMPLATE_ID');
var sefTrackerSheetId = scriptProperties.getProperty('SEF_TRACKER_SHEET_ID');
var sefCommentSheetId = scriptProperties.getProperty('SEF_COMMENT_SHEET_ID');
var sefFolderUrl = scriptProperties.getProperty('SEF_FOLDER_URL');

var sefForm = FormApp.openById(sefFormId);
var sefFormSheetId = sefForm.getDestinationId();
var sefFormSheet = SpreadsheetApp.openById(sefFormSheetId);

var sef = []

var orgName_idx = 2; // Column C
var applicationLink_idx = 5; // Column F
var requested_idx = 6; // Column G
var totalAmount_idx = 7; // Column H
var lifespan_idx = 8; // Column I
var numStudents_idx = 9; // Column J
var projectTitle_idx = 10; // Column K
var approved_idx = 11; // Column L

function closeTheFormSEF() {
  extractDataSEF();
  createCommentSheetSEF();
  //createMeetingMinutesSEF();
  //createBudgetTrackerSEF();
}

// Extract application data from the form spreadsheet
function extractDataSEF() {
  // Open the SEF form spreadsheet
  var data = sefFormSheet.getActiveSheet().getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    
    var extractedData = {
      organizationName: row[orgName_idx],
      requested: row[requested_idx],
      totalAmount: row[totalAmount_idx],
      numStudents: row[numStudents_idx],
      lifespan: row[lifespan_idx],
      applicationLink: row[applicationLink_idx],
      projectTitle: row[projectTitle_idx]
    };

    // Push extracted data into the array
    sef.push(extractedData);
  }

  // Sort each array by organization name
  sef.sort(function(a, b) {
    return a.organizationName.localeCompare(b.organizationName);
  });

  Logger.log(sef)

  Logger.log("Data are successfully extracted.")
}



// Create a meeting minutes
function createMeetingMinutesSEF() {
  // Duplicate the sample meeting minutes and rename it
  // If the array of objects, applicationData, is not empty, search through the copied document until you find "2.0 Applications"
  // If it is found, call a function, InsertPDTable(applicationData[i])
  var sefFolderId = scriptProperties.getProperty('SEF_FOLDER_ID');

  var newDocumentName = docName + ' SEF Sub-Committee Meeting Minutes';
  var copiedDocument = DriveApp.getFileById(sefMinutesTemplateId).makeCopy(newDocumentName, DriveApp.getFolderById(sefFolderId));
  var minutesDoc = DocumentApp.openById(copiedDocument.getId());
  var minutesBody = minutesDoc.getBody();
  var tables = minutesBody.getTables();

  Logger.log("Meeting minutes created successfully.");

  insertTablesSEF(tables[2], minutesBody);
  Logger.log("SEF added to meeting minutes.");
}

// A helper function for createMeetingMinutesSEF()
function insertTablesSEF(table, minutesBody) {
  var namePrefix = "2."
  var lastInsertedTable = table;
  var lastInsertedIndex = minutesBody.getChildIndex(lastInsertedTable);

  // Formatter for currency unit
  var formatter = new Intl.NumberFormat('en-US', {
      style: 'currency',
      currency: 'USD',
      minimumFractionDigits: 2
  });

  for (var i = 0; i < sef.length; i++) {
    var data = sef[i];

    // Insert the organization name text above the table
    var text = namePrefix + (i + 1) + " " + data.organizationName;
    if ((i > 0 && data.organizationName == sef[i - 1].organizationName) || (i < sef.length - 1 && data.organizationName == sef[i + 1].organizationName)) {
      text += " - " + data.projectTitle;
    }
    var paragraph = minutesBody.insertParagraph(lastInsertedIndex, text);

    paragraph.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    paragraph.setFontFamily("Times New Roman");
    paragraph.setFontSize(15);
    paragraph.setSpacingAfter(10);

    // Insert the copied table
    lastInsertedTable = minutesBody.insertTable(lastInsertedIndex + 1, table.copy());
    lastInsertedIndex = minutesBody.getChildIndex(lastInsertedTable);

    // Write lifespan and number of students
    var cell = lastInsertedTable.getRow(1).getCell(0);
    var cellText = cell.getText();

    targetText = "Number of students:";
    if (cellText.indexOf(targetText) !== -1) {
      var updatedText = cellText.replace(targetText, targetText + ' ' + data.numStudents);
      cell.setText(updatedText);
    }

    var cellText = cell.getText();

    targetText = "Lifespan:";
    if (cellText.indexOf(targetText) !== -1) {
      var updatedText = cellText.replace(targetText, targetText + ' ' + data.lifespan);
      cell.setText(updatedText);
    }

    // Update the table cells with the application data
    lastInsertedTable.getRow(1).getCell(2).setText(formatter.format(data.totalAmount));
    lastInsertedTable.getRow(2).getCell(2).setText(formatter.format(data.requested));

    // Add a few line breaks between tables
    minutesBody.insertParagraph(++lastInsertedIndex, '\n');
  }

  // Remove the original table
  table.removeFromParent();
}

function createCommentSheetSEF() {
  var commentSheet = SpreadsheetApp.openById(sefCommentSheetId);
  commentSheet = commentSheet.getSheets()[0];
  commentSheet.getRange('C2').setValue(sefFolderUrl);

  // Insert sef data
  var range = commentSheet.getRange(5, 1, sef.length, 2);
  var colorRange = range.getSheet().getRange(range.getRow(), 1, range.getNumRows(), 1);
  colorRange.setBackground('#d9d2e9');

  var pairs = pickDistinctMembers(sef.length);
  Logger.log(pairs);

  for (let i = 0; i < sef.length; i++) {
    commentSheet.getRange(i + 5, 1).setValue(sef[i].organizationName);
    commentSheet.getRange(i + 5, 2).setValue(sef[i].applicationLink);
    commentSheet.getRange(i + 5, 3).setValue(members[pairs[i][0]]);
    commentSheet.getRange(i + 5, 5).setValue(members[pairs[i][1]]);
  }
}

function createBudgetTrackerSEF() {
  var sheet = formSheet.getActiveSheet();
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var row = data[i];

    var extractedData = {
      organizationName: row[orgName_idx],
      requested: row[requested_idx],
      totalAmount: row[totalAmount_idx],
      numStudents: row[numStudents_idx],
      lifespan: row[lifespan_idx],
      approved: row[approved_idx]
    };

    sef.push(extractedData);
  }

  Logger.log(sef);

  // Open the tracker spreadsheet
  var budgetTracker = SpreadsheetApp.openById(sefTrackerSheetId);
  var sheet = budgetTracker.getSheetByName(docName);

  if (sheet) {
    Logger.log('The sheet "' + docName + '" exists.');
  }
  else {
    Logger.log("The sheet doesn't exist.");
  }

  var organizationName = "";
  var columnBValues = [];
  var rowIndex = 0;
  var startIndex = 0;

  // If organization name is in the spreadsheet, get to that row line and add a new row, new value and change the range of sum
  for (var i = 0; i < sef.length; i++) {
    organizationName = sef[i].organizationName;
    columnBValues = sheet.getRange("B:B").getValues().flat();
    rowIndex = columnBValues.indexOf(organizationName);

    // If org name is in the spreadsheet, append the requested value
    if (rowIndex !== -1) {
      var endIndex = columnBValues.indexOf(organizationName + " Total");

      sheet.insertRowBefore(endIndex);

      startIndex = rowIndex + 2
      rowIndex = endIndex;
    }

    // If org name is new, find the correct position
    else {
      rowIndex = 4;

      // Insert a new section (5 columns)
      for (var j = 0; j < 5; j++) {
        sheet.insertRowBefore(rowIndex);
      }
      
      // Insert data
      sheet.getRange(rowIndex + 1, 2).setValue(organizationName).setFontWeight('bold');
      sheet.getRange(rowIndex + 4, 2).setValue(organizationName + " Total");

      // Apply background color and border
      for (var col = 2; col <= 10; col++) {
        var cell = sheet.getRange(rowIndex + 4, col);
        cell.setBackground("#fff2cc") // Light yellow
          .setBorder(true, null, null, null, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      }

      rowIndex += 2;
      startIndex = rowIndex;
    }

    if (sef[i].approved == "TABLED") {
      sheet.getRange(rowIndex, 3)
        .setValue(sef[i].approved)
        .setFontWeight("bold")
        .setFontColor("red")
        .setHorizontalAlignment("right");
    }
    else {
      sheet.getRange(rowIndex, 3).setValue(sef[i].approved);
    }
    sheet.getRange(rowIndex, 4).setValue(sef[i].requested); // D
    sheet.getRange(rowIndex, 5).setValue(today); // E
    sheet.getRange(rowIndex, 9).setValue(sef[i].numStudents); // I
    sheet.getRange(rowIndex, 11).setValue(sef[i].lifespan); // M

    sheet.getRange(rowIndex + 2, 3).setFormula(`=SUM(C${startIndex}:C${rowIndex})`);
    sheet.getRange(rowIndex + 2, 4).setFormula(`=SUM(D${startIndex}:D${rowIndex})`);
    sheet.getRange(rowIndex + 2, 9).setFormula(`=SUM(D${startIndex}:D${rowIndex})`);
  }

  setSumFormulaForTotalBudget(sheet);
}

// Helper function of insertBudgetTracker(sef, applicationType) 
function setSumFormulaForTotalBudgetSEF(sheet) {
  var data = sheet.getRange("B:B").getValues();

  var rowIndices = [];
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().includes("Total")) { // Check if any cell in column B contains "Total"
      rowIndices.push(i + 1);
    }
  }
  
  if (rowIndices.length === 0) {
    Logger.log('No rows with "Total" found.');
    return;
  }
  
  // Create a formula string to sum values in column C at these row indices
  var columns = ['C', 'D'];

  columns.forEach(function(col) {
    var lastRow = sheet.getLastRow();

    var cellReferences = rowIndices.map(function(rowIndex) {
      return col + rowIndex;
    }).join(',');
    
    var formula = '=SUM(' + cellReferences + ')';
    
    var targetCell = sheet.getRange(col + lastRow);
    targetCell.setFormula(formula);
  });
}
