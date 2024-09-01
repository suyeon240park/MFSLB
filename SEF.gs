// 1. Extract application data from the SEF form spreadsheet
// 2. Write the SEF minute
// 3. Write the SEF tracker

// Manual change is required
var docName = "Fall 2024";

var scriptProperties = PropertiesService.getScriptProperties();

var formId = scriptProperties.getProperty('SEF_FORM_ID');
var minutesTemplateId = scriptProperties.getProperty('SEF_MINUTES_TEMPLATE_ID');
var trackerSheetId = scriptProperties.getProperty('SEF_TRACKER_SHEET_ID');

var form = FormApp.openById(formId);
var formSheetId = form.getDestinationId();
var formSheet = SpreadsheetApp.openById(formSheetId);

var sef = []

var orgName_idx = 2; // Column C
var requested_idx = 6; // Column G
var totalAmount_idx = 7; // Column H
var lifespan_idx = 8; // Column I
var numStudents_idx = 9; // Column J
var approved_idx = 10;

function closeTheForm() {
  //extractData();
  //createMeetingMinutes();
  createBudgetTracker();
}

// Extract application data from the form spreadsheet
function extractData() {
  // Open the SEF form spreadsheet
  var data = formSheet.getActiveSheet().getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    
    var extractedData = {
      organizationName: row[orgName_idx],
      requested: row[requested_idx],
      totalAmount: row[totalAmount_idx],
      numStudents: row[numStudents_idx],
      lifespan: row[lifespan_idx]
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
function createMeetingMinutes() {
  // Duplicate the sample meeting minutes and rename it
  // If the array of objects, applicationData, is not empty, search through the copied document until you find "2.0 Applications"
  // If it is found, call a function, InsertPDTable(applicationData[i])
  var sefFolderId = scriptProperties.getProperty('SEF_FOLDER_ID');

  var newDocumentName = docName + ' SEF Sub-Committee Meeting Minutes';
  var copiedDocument = DriveApp.getFileById(minutesTemplateId).makeCopy(newDocumentName, DriveApp.getFolderById(sefFolderId));
  var minutesDoc = DocumentApp.openById(copiedDocument.getId());
  var minutesBody = minutesDoc.getBody();
  var tables = minutesBody.getTables();

  Logger.log("Meeting minutes created successfully.");

  insertTables(tables[2], minutesBody);
  Logger.log("SEF added to meeting minutes.");
}

// A helper function for createMeetingMinutes()
function insertTables(table, minutesBody) {
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
    var paragraph = minutesBody.insertParagraph(lastInsertedIndex + 1, namePrefix + (i + 1) + " " + data.organizationName);
    paragraph.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    paragraph.setFontFamily("Times New Roman");
    paragraph.setFontSize(15);
    paragraph.setSpacingAfter(10);

    // Insert the copied table
    lastInsertedTable = minutesBody.insertTable(lastInsertedIndex + 2, table.copy());
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



function createBudgetTracker() {
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
  var budgetTracker = SpreadsheetApp.openById(trackerSheetId);
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
function setSumFormulaForTotalBudget(sheet) {
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
