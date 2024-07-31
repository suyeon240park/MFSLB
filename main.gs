// Intro about this project
// Video recorded

var formUrl = "https://docs.google.com/forms/d/1PDQjAXXF9TTKFwtHjVHsNzIuH6D3eXn-mqFmetab88c/edit";
var form = FormApp.openByUrl(formUrl);
var formSheetId = form.getDestinationId();
var formSheet = SpreadsheetApp.openById(formSheetId);
var internalDriveId = "1QIwzrrSNl6WqIjX55BY55fQ6_Vw9Gtl5";
var parentFolderId = "1ktU2mzhB1rGbVyxGIIDq4V2CblYHaZyZ";
var parentFolder = DriveApp.getFolderById(parentFolderId);
var minutesTemplateId = "1TRJGuduquuOtu8wpV-AmK_UUjHuggnXvFDUBFKO9gQk";
var commentSheetId = "1OiW0FKIjcui6qY9n0FgIcbf1yZbioI6FGY54AgXktdU";
var trackerSheetId = "1PtkOfr48Turfb8sQZBC45pl2Wxq6RTh451xARNK4i4A";

var today = new Date();
var thisMonth = today.toLocaleString('default', { month: 'long' });
var thisYear = today.getFullYear();
var members = ["Suyeon Park", "Kenneth Sulimro", "Edlyn Li", "Beatriz Correa de Mello", "Kelvin Lo", "Zayneb Hussain", "Sean Huang"];
var numMembers = members.length;

var projectDirectorships = [];
var specialProjects = [];
var conferenceFunding = [];

function openForm() {
  // Delete all the responses in the form (files in the drive and data in the spreadsheet remain intact)
  form.deleteAllResponses();

  // Rename the linked sheet
  var responseSheet = formSheet.getSheets()[0];
  responseSheet.setName(thisMonth + " " + thisYear);

  // Unlink and relink the spreadsheet to create another linked sheet
  form.removeDestination();
  form.setDestination(FormApp.DestinationType.SPREADSHEET, formSheetId);

  Logger.log("The application data has been saved to the new monthly sheet.")

  // Clear all rows from row 2 to the last row
  var lastRow = linkedSheet.getLastRow();
  if (lastRow > 1) {
    linkedSheet.getRange(2, 1, lastRow - 1, linkedSheet.getLastColumn()).clearContent();
  }

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


function closeTheForm() {
  form.setAcceptingResponses(false);
  Logger.log("The form is closed.")

  // Make a comment sheet and meeting minutes
  extractData();
  //createCommentSheet();
  createMeetingMinutes();
  //createBudgetTracker();

  Logger.log("All done. Congrats for finishing another month. Make Finance Secretary Life Better.");
}


function extractData() {
  // Open the unified form spreadsheet
  var data = formSheet.getActiveSheet().getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var applicationType = row[6]; // Column G

    if (row[9] == "Individual") {
      var organizationName = row[2]; // Column C
    }
    else if (row[9] == "Project directorship") {
      var organizationName = row[12]; // Column M
    }
    else if (row[9] == "Student club") {
      var organizationName = row[10]; // Column K
    }
    else {
      Logger.log("Option not on the list");
      continue;
    }
    
    var extractedData = {
      organizationName: organizationName,
      applicationType: applicationType,
      applicationLink: row[3], // Column D
      suppFileLink: row[4], // Column E
      requested: row[5], // Column F
    };

    // Push extracted data into the corresponding array based on the application type
    if (applicationType === "Project Directorships") {
      extractedData.previous = row[7];
      projectDirectorships.push(extractedData);
    }
    else if (applicationType === "Special Projects") {
      specialProjects.push(extractedData);
    }
    else if (applicationType === "Conference Funding") {
      extractedData.totalAmount = row[8];
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

  var numApplications = startRowCF; // Will be changed later for other applications
  Logger.log(numApplications);

  // Randomly list FC members, make sure each member is called an equal number of times
  var count = new Array(numMembers).fill(0);

  var randomNum1 = Math.floor(Math.random() * numMembers);
  var randomNum2 = Math.floor(Math.random() * numMembers);

  for (var i = 0; i < numApplications - 4; i++) {
    while (randomNum1 == randomNum2 || count[randomNum2] > Math.max(...count)) {
      randomNum2 = Math.floor(Math.random() * numMembers);
    }
    count[randomNum1] += 1;
    count[randomNum2] += 1;

    commentSheet.getRange(5 + i, 5).setValue(members[randomNum1]);
    commentSheet.getRange(5 + i, 7).setValue(members[randomNum2]);

    randomNum1 = Math.floor(Math.random() * numMembers);
    randomNum2 = Math.floor(Math.random() * numMembers);
  }
}

function setValuesAndColor(range, data, color) {
  var values = data.map(function(item) {
    return [item.organizationName, item.applicationType, item.applicationLink, item.suppFileLink];
  });
  range.setValues(values);
  
  // Set the background color for column A
  var colorRange = range.getSheet().getRange(range.getRow(), 1, range.getNumRows(), 1);
  colorRange.setBackground(color);
}


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



function createBudgetTracker() {
  var data = formSheet.getActiveSheet().getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var applicationType = row[6]; // Column G

    if (row[9] == "Individual") {
      var organizationName = row[2]; // Column C
    }
    else if (row[9] == "Project directorship") {
      var organizationName = row[12]; // Column M
    }
    else if (row[9] == "Student club") {
      var organizationName = row[10]; // Column K
    }
    else {
      Logger.log("Option not on the list");
      continue;
    }
    
    var extractedData = {
      organizationName: organizationName,
      applicationType: applicationType,
      requested: row[5], // Column F
      approved: row[24] // Column Z
    };

    // Push extracted data into the corresponding array based on the application type
    if (applicationType === "Project Directorships") {
      projectDirectorships.push(extractedData);
    }
    else if (applicationType === "Special Projects") {
      extractedData.projectTitle = row[13];
      specialProjects.push(extractedData);
    }
    else if (applicationType === "Conference Funding") {
      extractedData.confName = row[14];
      conferenceFunding.push(extractedData);
    }
    else {
      Logger.log("Option not on the list");
      continue;
    }
  }

  insertBudgetTracker(projectDirectorships);
  //insertBudgetTracker(specialProjects);
  //insertBudgetTracker(conferenceFunding);
}

function insertBudgetTracker(appData) {
  var budgetTracker = SpreadsheetApp.openById(trackerSheetId);
  var applicationType = appData[0].applicationType;
  var sheet = budgetTracker.getSheetByName(applicationType);

  // If organization name is in the spreadsheet, get to that row line and add a new row, new value and change the range of sum
  var columnBValues = sheet.getRange("B:B").getValues();
  for (var i = 0; i < appData.length; i++) {
    var organizationName = appData[i].organizationName;
    var rowIndex = columnBValues.indexOf(organizationName);

    // If org name is in the spreadsheet, append the requested value
    if (rowIndex !== -1) {
      var endIndex = columnBValues.indexOf(organizationName + " Total");
      sheet.insertRowBefore(endIndex - 1);
      sheet.getRange(endIndex - 2, 3).setValue(appData[i].approved);
      sheet.getRange(endIndex - 2, 4).setValue(appData[i].requested);
      sheet.getRange(endIndex - 2, 5).setValue(today);
      sheet.getRange(endIndex, 4).setFormula(`=SUM(D${rowIndex + 1}:D${endIndex - 2})`);
    }

    // If org name is new, find the correct position
    else {
      rowIndex = 4;

      // Find the correct position alphabetically
      for (var j = 0; j < columnBValues.length; j++) {
        if (columnBValues[j][0] > organizationName) {
          rowIndex = j + 1;
          break;
        }
      }

      Logger.log(rowIndex);

      // Insert new rows
      for (var j = 0; j < 5; j++) {
        sheet.insertRowBefore(rowIndex);
      }

      sheet.getRange(rowIndex + 1, 2).setValue(organizationName).setFontWeight('bold');
      sheet.getRange(rowIndex + 2, 3).setValue(appData[i].approved);
      sheet.getRange(rowIndex + 2, 4).setValue(appData[i].requested);
      sheet.getRange(rowIndex + 2, 5).setValue(today);
      sheet.getRange(rowIndex + 4, 2).setValue(organizationName + " Total");
      sheet.getRange(rowIndex + 4, 4).setFormula(`=SUM(D${rowIndex + 2})`);

      if (applicationType == "Special Projects") {
        sheet.getRange(rowIndex + 2, 8).setValue(appData[i].projecTitle);
      }
      if (applicationType == "Conference Funding") {
        sheet.getRange(rowIndex + 2, 8).setValue(appData[i].conferenceName);
      }

      for (var col = 2; col <= 8; col++) {
        var cell = sheet.getRange(rowIndex + 4, col);

        // Apply background color and border
        cell.setBackground("#fff2cc") // Light yellow
          .setBorder(true, null, null, null, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      }
    }
  }
}
