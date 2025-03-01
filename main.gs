/*
TODO: don't use the RNG, just randomize order then apply ppls names
Status: not started

TODO: inspect why 6 projects weren't assigned ppl to review in the comment sheet
Status: not started

TODO: implement SEF option in unified form
Status: not started, assigned to Kenneth

TODO: add functionality to get conf date/time/place on budget tracker
Status: Started by Kenneth.

TODO: add SEF to the finance wizard.
Status: not started, assigned to Kenneth

Suyeon TO-DO:
2. Fix the random name assignment algorithm
*/

// Intro about this program
// link: 
// The application form is prone to frequent changes.
// Update these variables in Script Properties to reflect any changes in the future.
var scriptProperties = PropertiesService.getScriptProperties();

var formId = scriptProperties.getProperty('FORM_ID');
var internalDriveId = scriptProperties.getProperty('INTERNAL_DRIVE_ID');
var parentFolderId = scriptProperties.getProperty('PARENT_FOLDER_ID');
var minutesTemplateId = scriptProperties.getProperty('MINUTES_TEMPLATE_ID');
var commentSheetId = scriptProperties.getProperty('COMMENT_SHEET_ID');
var trackerSheetId = scriptProperties.getProperty('TRACKER_SHEET_ID');

var form = FormApp.openById(formId);
var formSheetId = form.getDestinationId();
var formSheet = SpreadsheetApp.openById(formSheetId);
var parentFolder = DriveApp.getFolderById(parentFolderId);

// Can we build this online?
const applicantType_idx = 2; // Column C
const clubOrgName_idx = 4; // Column E
const pdOrgName_idx = 5; // Column F
const associatedEntityName = 7; // Column G
const applicantName_idx = 7; // Column H
const appLink_idx = 8; // Column I
const suppLink_idx = 9; // Column J
const requested_idx = 10; // Column K
const appType_idx = 11; // Column L
const prev_idx = 12; // Column M
const projectTitle_idx = 13; // Column N
const totalAmount_idx = 14; // Column O
const conferenceName_idx = 15; // Column P
const conferenceLocation_idx = 16// Column Q
const conferenceTime_idx = 17// Column R
const approved_idx = 18; // Column S


// Date
var today = new Date();
var thisMonth = today.toLocaleString('default', { month: 'long' });
var thisYear = today.getFullYear();

// Arrays with different application types
var projectDirectorships = [];
var specialProjects = [];
var conferenceFunding = [];


// Open the form and make new monthly folders
function openForm() {
  // Delete all the responses in the form (files in the drive and data in the spreadsheet remain intact)
  form.deleteAllResponses();

  // Archive the current linked sheet
  var responseSheet = formSheet.getSheets()[0];
  responseSheet.setName(thisMonth + " " + thisYear);

  // Unlink and relink the spreadsheet to create another linked sheet
  form.removeDestination();
  form.setDestination(FormApp.DestinationType.SPREADSHEET, formSheetId);

  Logger.log("The application data has been saved to the new monthly sheet.")

  // Make a new monthly folder in the drive
  var newFolder = parentFolder.createFolder(thisMonth + " " + thisYear + " Meeting");
  PropertiesService.getScriptProperties().setProperty('monthFolderId', newFolder.getId());
  PropertiesService.getScriptProperties().setProperty('monthFolderUrl', newFolder.getUrl());

  // Create subfolders in the new monthly folder
  newFolder.createFolder('Project Directorships');
  newFolder.createFolder('Special Projects');
  newFolder.createFolder('Conference Funding');

  Logger.log("Monthly folder created.");

  // Open the form
  form.setAcceptingResponses(true);
  Logger.log("The unified form is now ready to work.");
}

// When the user submits the form, categorize applications and allocate files to corresponding folders.
function onFormSubmit(e) {
  if (!e) { // TODO: implement debouncing
    Logger.log("Event object is undefined");
  }

  // Retrieve the month folder ID from the Properties Service
  var monthFolderId = PropertiesService.getScriptProperties().getProperty('monthFolderId');
  if (!monthFolderId) {
    Logger.log("monthFolder doesn't exist");
    return;
  }

  var monthFolder = DriveApp.getFolderById(monthFolderId);
  Logger.log("Opened a folder with id " + monthFolderId);

  // Get the form response
  var formResponse = e.response;
  var itemResponses = formResponse.getItemResponses();

  var targetFolder = null;
  var response = '';

  // Iterate over the form responses to determine the target folder
  for (var i = 0; i < itemResponses.length; i++) {
    var itemResponse = itemResponses[i];

    // Check if the item is "Application Type" question
    if (itemResponse.getItem().getTitle() == "Application Type") { // CHANGE THE TITLE
      response = itemResponse.getResponse();
      Logger.log(response);
      var folders = monthFolder.getFoldersByName(response);
      if (folders.hasNext()) {
        targetFolder = folders.next();
      }
    }
  }

  // If no target folder found, log an error and exit
  if (!targetFolder) {
    Logger.log("Target folder doesn't exist.")
    return;
  }

  // Process file uploads and move them to the target folder
  for (var i = 0; i < itemResponses.length; i++) {
    var itemResponse = itemResponses[i];

    // Check if the response is a file upload type
    if (itemResponse.getItem().getType() === FormApp.ItemType.FILE_UPLOAD) {
      var fileUploads = itemResponse.getResponse();

      // Process each uploaded file
      for (var j = 0; j < fileUploads.length; j++) {
        var fileId = fileUploads[j];
        var file = DriveApp.getFileById(fileId);
        file.moveTo(targetFolder); // Move the file to the target folder
        Logger.log("Moved");
      }
    }
  }
}

// Close the form and write meeting documents
// Pre-meeting: extract data, comment sheet, and meeting minutes
// Post-meeting: budget tracker
function closeTheForm() {
  form.setAcceptingResponses(false);
  //Logger.log("The form is closed.")

  // Make a comment sheet and meeting minutes
  extractData();
  createCommentSheet();
  createMeetingMinutes();

  Logger.log("All done. Congrats for finishing another month. Make Finance Secretary Life Better.");
}

// Extract application data from the form spreadsheet
function extractData() {
  // Open the unified form spreadsheet
  var data = formSheet.getActiveSheet().getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    var applicationType = data[i][appType_idx];
    var applicantType = data[i][applicantType_idx];

    if (applicantType == "Individual") {
      var organizationName = data[i][applicantName_idx];
    }
    else if (applicantType == "Project directorship") {
      var organizationName = data[i][pdOrgName_idx];
    }
    else if (applicantType == "Student club") {
      var organizationName = data[i][clubOrgName_idx];
    }
    else if (applicantType == "Associated entity") {
      var organizationName = data[i][associatedEntityName];
    }
    else {
      Logger.log("Option not on the list");
      continue;
    }
    
    var extractedData = {
      organizationName: organizationName,
      applicationType: applicationType,
      applicationLink: data[i][appLink_idx],
      suppFileLink: data[i][suppLink_idx],
      requested: data[i][requested_idx]
    };

    // Push extracted data into the corresponding array based on the application type
    if (applicationType === "Project Directorships") {
      extractedData.previous = data[i][prev_idx];
      projectDirectorships.push(extractedData);
    }
    else if (applicationType === "Special Projects") {
      specialProjects.push(extractedData);
    }
    else if (applicationType === "Conference Funding") {
      extractedData.totalAmount = data[i][totalAmount_idx];
      conferenceFunding.push(extractedData);
    }
  }

  // Sort each array by organization name
  projectDirectorships.sort(function(a, b) {
    return a.organizationName.localeCompare(b.organizationName);
  });

  specialProjects.sort(function(a, b) {
    return a.organizationName.localeCompare(b.organizationName);
  });

  conferenceFunding.sort(function(a, b) {
    return a.organizationName.localeCompare(b.organizationName);
  });

  Logger.log(projectDirectorships)
  Logger.log(specialProjects)
  Logger.log(conferenceFunding)

  Logger.log("Data are successfully extracted.")
}

// Create a comment sheet for Fincomm members to review the applications
function createCommentSheet() {
  // Duplicate the comment sheet sample and rename it
  var commentSheet = SpreadsheetApp.openById(commentSheetId);
  var sheetToDuplicate = commentSheet.getSheetByName("Sample");
  
  if (!sheetToDuplicate) {
    throw new Error("Sheet named 'Sample' not found.");
  }
  commentSheet = sheetToDuplicate.copyTo(commentSheet);
  commentSheet.setName(thisMonth + " " + thisYear);

  var monthFolderUrl = PropertiesService.getScriptProperties().getProperty('monthFolderUrl');
  commentSheet.getRange('C2').setValue(monthFolderUrl);

  // Insert project directorships data
  var rangePD = commentSheet.getRange(5, 1, projectDirectorships.length, 4);
  setValuesAndColor(rangePD, projectDirectorships, '#d9ead3'); // Light green 3

  // Insert special projects data
  var startRowSP = 5 + projectDirectorships.length;
  var rangeSP = commentSheet.getRange(startRowSP, 1, specialProjects.length, 4);
  setValuesAndColor(rangeSP, specialProjects, '#fce5cd'); // Light orange 3

  // Insert conference funding data
  var startRowCF = startRowSP + specialProjects.length;
  var rangeCF = commentSheet.getRange(startRowCF, 1, conferenceFunding.length, 4);
  setValuesAndColor(rangeCF, conferenceFunding, '#cfe2f3'); // Light cornflower blue 3

  var numApplications = projectDirectorships.length + specialProjects.length + conferenceFunding.length;
  Logger.log(numApplications);

  const pairs = pickDistinctMembers(numApplications);
  Logger.log(pairs);

  for (let i = 0; i < numApplications; i++) {
    commentSheet.getRange(i + 5, 5).setValue(members[pairs[i][0]]);
    commentSheet.getRange(i + 5, 7).setValue(members[pairs[i][1]]);
  }
}

function pickDistinctMembers(appLength) {
  const pairs = [];
  const pairCounts = new Map(); // Map to track the counts of each pair

  // Generate all possible distinct pairs
  for (let i = 0; i < numMembers; i++) {
    for (let j = i + 1; j < numMembers; j++) {
      pairCounts.set(`${i},${j}`, 0); // Initialize count for each pair
    }
  }

  for (let i = 0; i < appLength; i++) {
    // Convert pairCounts to an array and sort based on counts
    const sortedPairs = Array.from(pairCounts.entries())
      .sort((a, b) => a[1] - b[1]); // Sort pairs by their counts (ascending)

    // Pick a random pair from the least selected pairs
    const leastSelectedPairs = sortedPairs.filter(pair => pair[1] === sortedPairs[0][1]);
    const randomPairIndex = Math.floor(Math.random() * leastSelectedPairs.length);
    const selectedPair = leastSelectedPairs[randomPairIndex][0];

    // Increment the count for the selected pair
    pairCounts.set(selectedPair, pairCounts.get(selectedPair) + 1);
    
    // Store the selected pair in the results
    pairs.push(selectedPair.split(',').map(num => parseInt(num, 10)));
  }
  return pairs;
}

// A helper function for createCommentSheet()
function setValuesAndColor(range, data, color) {
  var values = data.map(function(item) {
    return [item.organizationName, item.applicationType, item.applicationLink, item.suppFileLink];
  });
  range.setValues(values);
  
  // Set the background color for column A
  var colorRange = range.getSheet().getRange(range.getRow(), 1, range.getNumRows(), 1);
  colorRange.setBackground(color);
}

// Create a meeting minutes
function createMeetingMinutes() {
  // Duplicate the sample meeting minutes and rename it
  // If the array of objects, applicationData, is not empty, search through the copied document until you find "3.0 Budget Submission"
  // If it is found, call a function, InsertPDTable(applicationData[i]) until applicationType in applicationData is not "Project Directorships" anymore
  // Then, starting from the point you were working on, search for ""

  var monthFolderId = PropertiesService.getScriptProperties().getProperty('monthFolderId');
  var newDocumentName = thisMonth + ' ' + thisYear + ' Finance Committee Agenda';
  var copiedDocument = DriveApp.getFileById(minutesTemplateId).makeCopy(newDocumentName, DriveApp.getFolderById(monthFolderId));
  var minutesDoc = DocumentApp.openById(copiedDocument.getId());
  var minutesBody = minutesDoc.getBody();
  var tables = minutesBody.getTables();

  Logger.log("Meeting minutes created successfully.");

  insertTables(tables[2], minutesBody, projectDirectorships, "3.");
  Logger.log("Project Directorships added to meeting minutes.");

  insertTables(tables[3], minutesBody, specialProjects, "4.");
  Logger.log("Special Projects added to meeting minutes.");

  insertTables(tables[4], minutesBody, conferenceFunding, "5.");
  Logger.log("Conference Fundings added to meeting minutes.");

  Logger.log("All meeting minutes operations completed.");
}

// A helper function for createMeetingMinutes()
function insertTables(table, minutesBody, appData, namePrefix) {
  var lastInsertedTable = table;
  var lastInsertedIndex = minutesBody.getChildIndex(lastInsertedTable);

  // Formatter for currency unit
  var formatter = new Intl.NumberFormat('en-US', {
      style: 'currency',
      currency: 'USD',
      minimumFractionDigits: 2
  });

  for (var i = 0; i < appData.length; i++) {
    var data = appData[i];

    // Insert the organization name text above the table
    var paragraph = minutesBody.insertParagraph(lastInsertedIndex + 1, namePrefix + (i + 1) + " " + data.organizationName);
    paragraph.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    paragraph.setFontFamily("Times New Roman");
    paragraph.setFontSize(15);
    paragraph.setSpacingAfter(10);

    // Insert the copied table
    lastInsertedTable = minutesBody.insertTable(lastInsertedIndex + 2, table.copy());
    lastInsertedIndex = minutesBody.getChildIndex(lastInsertedTable);

    // Update the table cells with the application data
    var firstValue, secondValue = "";
    if (data.applicationType == "Project Directorships") {
      firstValue = data.requested;
      secondValue = data.previous;
    }
    else if (data.applicationType == "Special Projects") {
      firstValue = data.requested;
    }
    else if (data.applicationType == "Conference Funding") {
      firstValue = data.totalAmount;
      secondValue = data.requested;
    }
    lastInsertedTable.getRow(1).getCell(2).setText(formatter.format(firstValue));
    lastInsertedTable.getRow(2).getCell(2).setText(formatter.format(secondValue));

    // Add a few line breaks between tables
    minutesBody.insertParagraph(++lastInsertedIndex, '\n');
  }

  // Remove the original table
  table.removeFromParent();
}


// Create a budget tracker after holding a meeting
// Remember: write approved amounts (currently in Column S) in the spreadsheet connected to the form
function createBudgetTracker() {
  var data = formSheet.getSheets()[0].getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var organizationName = row[pdOrgName_idx] || row[clubOrgName_idx] || row[applicantName_idx];
    var eventName = row[projectTitle_idx] || row[conferenceName_idx];
    var applicationType = row[appType_idx];

    var extractedData = {
      organizationName: organizationName,
      requested: row[requested_idx],
      approved: row[approved_idx]
    };

    // Push extracted data into the corresponding array based on the application type
    switch (applicationType) {
      case "Project Directorships":
        projectDirectorships.push(extractedData);
        break;
      case "Special Projects":
        extractedData["eventName"] = eventName;
        specialProjects.push(extractedData);
        break;
      case "Conference Funding":
        extractedData["eventName"] = eventName;
        extractedData["location"] = row[conferenceLocation_idx];
        extractedData["eventDate"] = row[conferenceTime_idx];
        conferenceFunding.push(extractedData);
        break;
      default:
        Logger.log("Option not on the list");
        continue;
    }
  }

  Logger.log(projectDirectorships)
  Logger.log(specialProjects)
  Logger.log(conferenceFunding)

  insertBudgetTracker(projectDirectorships, "Project Directorships", 8);
  insertBudgetTracker(specialProjects, "Special Projects", 9);
  insertBudgetTracker(conferenceFunding, "Conference Funding", 12);
}

// Actual recording of budget tracker
function insertBudgetTracker(appData, applicationType, numCol) {
  var approvedDate = "11/30/2024"
  var budgetTracker = SpreadsheetApp.openById(trackerSheetId);
  var sheet = budgetTracker.getSheetByName(applicationType);

  // Set everything in the sheet to Times New Roman
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setFontFamily("Times New Roman");

  var organizationName = "";
  var columnBValues = [];
  var rowIndex = 0;
  var startIndex = 0;

  // If organization name is in the spreadsheet, get to that row line and add a new row, new value and change the range of sum
  for (var i = 0; i < appData.length; i++) {
    organizationName = appData[i].organizationName;
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
      
      // Insert organization data
      sheet.getRange(rowIndex + 1, 2)
        .setValue(organizationName)
        .setFontWeight('bold')
      sheet.getRange(rowIndex + 4, 2)
        .setValue(organizationName + " Total")

      // Apply background color and border
      for (var col = 2; col <= numCol; col++) {
        var cell = sheet.getRange(rowIndex + 4, col);
        cell.setBackground("#fff2cc") // Light yellow
          .setBorder(true, null, null, null, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      }

      rowIndex += 2;
      startIndex = rowIndex;
    }

    // Set approved value with currency format
    if (appData[i].approved == "TABLED") {
      sheet.getRange(rowIndex, 3)
        .setValue(appData[i].approved)
        .setFontWeight("bold")
        .setFontColor("red")
        .setHorizontalAlignment("right")
    } else {
      sheet.getRange(rowIndex, 3)
        .setValue(appData[i].approved)
        .setNumberFormat("$#,##0.00")
    }

    // Set requested value with currency format
    sheet.getRange(rowIndex, 4)
      .setValue(appData[i].requested)
      .setNumberFormat("$#,##0.00") // D

    // Set approved date
    sheet.getRange(rowIndex, 5).setValue(approvedDate) // E

    // Set event name
    sheet.getRange(rowIndex, 8).setValue(appData[i].eventName) // H

    // Set conference location and time
    if (appData[i].location || appData[i].eventDate) {
      sheet.getRange(rowIndex, 11).setValue(appData[i].location)
      sheet.getRange(rowIndex, 12).setValue(appData[i].eventDate)
    }

    /*
    // FLAG KENNETH - Insert conference date/time here, and find appData prerequisites to include.
    if (applicationType == "Conference Funding"){
      sheet.getRange(rowIndex, 11).setValue(appdata[i].kargs.location) // K
      sheet.getRange(rowIndex, 12).setValue(appdata[i].kargs.conf_time) // L
    }
    */

    sheet.getRange(rowIndex + 2, 3).setFormula(`=SUM(C${startIndex}:C${rowIndex})`);
    sheet.getRange(rowIndex + 2, 4).setFormula(`=SUM(D${startIndex}:D${rowIndex})`);
  }

  setSumFormulaForTotalBudget(sheet);
}

// Helper function of insertBudgetTracker(appData, applicationType) 
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
