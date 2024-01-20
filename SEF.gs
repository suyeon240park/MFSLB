//Define a linked list node structure
class Node {
  constructor(data) {
    this.data = data;
    this.next = null;
  }
}

//Create a LinkedList class
class LinkedList {
  constructor() {
    this.head = null;
  }

  
  //Add a new node to the linked list
  add(data) {
    const newNode = new Node(data);
    if (!this.head) { //if there's no node
      this.head = newNode //create a new node as a head
    } else {
      let current = this.head;
      while (current.next) {
        current = current.next;
      }
      current.next = newNode;
    }
  }

  //Sort the linked list based on the name of an organization alphabetically
  sort() {
    if (!this.head) return;

    let swapped;
    do {
      swapped = false;
      let current = this.head;
      let prev = null;

      while (current.next) {
        const next = current.next;

        if (current.data.organization > next.data.organization) {
          // Swap organization, requested, and totalAmount fields
          [current.data.organization, next.data.organization] = [next.data.organization, current.data.organization];
          [current.data.requested, next.data.requested] = [next.data.requested, current.data.requested];
          [current.data.totalAmount, next.data.totalAmount] = [next.data.totalAmount, current.data.totalAmount];

          swapped = true;
        }
        prev = current;
        current = next;
      }
    } while (swapped);
  }


  toArray() {
    const result = [];
    let current = this.head;
    while (current) {
      result.push(current.data);
      current = current.next;
    }
    return result;
  }
}



//Due to unavailability to use excel files in Google Apps Script, I need to convert them to Google Spreadsheets
function convertExcelToGoogleSS() {
  var sourceFolderId = "1WCyfgaZCZp2S1h6T1ietzWODcT1a7vTD";
  var sourceFolder = DriveApp.getFolderById(sourceFolderId);

  var files = sourceFolder.getFiles();

  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();

    if (fileName.toLowerCase().endsWith('.xlsx')) {
      var blob = file.getBlob();
      var convertedFile = Drive.Files.insert({ title: fileName, parents: [{id: sourceFolderId}] }, blob, {
        convert: true
      });

      if (fileName.startsWith("Copy of ")) {
        var newFileName = fileName.substring("Copy of ".length);
        convertedFile.setTitle(newFileName);
      }

      file.setTrashed(true);
      Logger.log(fileName + ' successfully converted to Google Spreadsheet.');
    }
    else {
      Logger.log(fileName + ' does not need conversion.');
    }
  }
  projectDirectorship();
}

function projectDirectorship() {
  var sourceFolderId = '1WCyfgaZCZp2S1h6T1ietzWODcT1a7vTD';
  var sourceFolder = DriveApp.getFolderById(sourceFolderId);
  var i = 1;

  var BudgetApplication = new LinkedList();

  var files = sourceFolder.getFiles();

  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();

    Logger.log(i + ": " + fileName);

    var requested = Number(-findRequested(file)).toFixed(2);
    var organization = findOrganization(file);

    Logger.log("Total Amount: " + totalAmount);
    Logger.log("Requested: " + requested);
    Logger.log("Organization: " + organization);

    if (organization !== undefined && totalAmount !== undefined && requested !== undefined) {
      BudgetApplication.add({
        organization: organization,
        totalAmount: totalAmount,
        requested: requested
      });
    } else {
      Logger.log('Skipping file ' + fileName + ' due to missing data.');
    }
    i++;
  }
  BudgetApplication.sort();
  const sortedArray = BudgetApplication.toArray();

  pasteToDocs(sortedArray);
}

function copyValuesFromSheetToLL() {
  var sourceFolderId = '1WCyfgaZCZp2S1h6T1ietzWODcT1a7vTD';
  var sourceFolder = DriveApp.getFolderById(sourceFolderId);
  var i = 1;

  var BudgetApplication = new LinkedList();

  var files = sourceFolder.getFiles();

  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();

    Logger.log(i + ": " + fileName);

    var totalAmount = Number(-findTotalAmount(file)).toFixed(2);
    var requested = Number(-findRequested(file)).toFixed(2);
    var organization = findOrganization(file);

    Logger.log("Total Amount: " + totalAmount);
    Logger.log("Requested: " + requested);
    Logger.log("Organization: " + organization);

    if (organization !== undefined && totalAmount !== undefined && requested !== undefined) {
      BudgetApplication.add({
        organization: organization,
        totalAmount: totalAmount,
        requested: requested
      });
    } else {
      Logger.log('Skipping file ' + fileName + ' due to missing data.');
    }
    i++;
  }
  BudgetApplication.sort();
  const sortedArray = BudgetApplication.toArray();

  pasteToDocs(sortedArray);
}

//SEF Summary of Funding Recommendations
function summary(sortedArray) {
  //fill out the organization from first column of second row of the table to the end
  //fill out the totalAmount from second column of second row of the table to the end
  //fill out the requested from third column of second row of the table to the end
  var minutes = DocumentApp.openByUrl('https://docs.google.com/document/d/1IKPfAUJPOwukva0z6bJ39VqY2aP6DigsbMJid9VVxLE/edit');
  var body = minutes.getBody();
  var tables = body.getTables();

  if (tables.length > 0) {
    var table = tables[0]; // Assuming you have only one table in the document
    var numRows = sortedArray.length;

    // Ensure the table has at least one row
    if (numRows > 0) {
      for (var i = 0; i < numRows; i++) {
        // Append a new row for each entry in sortedArray
        var row = table.appendTableRow();
        
        // Set "Organization" in the first column
        row.appendTableCell(sortedArray[i].organization);
        
        // Set "Total Amount" in the second column
        row.appendTableCell(sortedArray[i].totalAmount);
        
        // Set "Requested" in the third column
        row.appendTableCell(sortedArray[i].requested);
      }
    } else {
      Logger.log("No data in the sortedArray to populate the table.");
    }
  } else {
    Logger.log("No table found in the document.");
  }
  //find the value beside a word "Approved" in a table repeatedly until the end - store it to a variable approved
  //fill out approved from fourth column of second row of the table to the end
}


function pasteToDocs(sortedArray) {
  var minutes = DocumentApp.openByUrl('https://docs.google.com/document/d/19i_TyDVukwqX4Dad4kvmIN3xtIj1cNlP4Ggh5_YIrdU/edit');
  var body = minutes.getBody();

  //table of contents
  var text = body.editAsText();
  var found = false;

  if (!found) {
    var textElement = body.findText("2.0 Applications");
    if (textElement) {
      var textPosition = textElement.getStartOffset() + 1; // Start after the "2.0 Applications" text
      var formattedText;

      for (var i = sortedArray.length - 1; i >= 0; i--) {
        formattedText = "2." + (i + 1) + " " + sortedArray[i].organization + "\n";
        text.insertText(textPosition, formattedText);
      }
    }
    found = true;
  }

  //fill out the tables
  var textToReplace = "2.";

  // Find all occurrences of "2." in the document
  var textElements = body.findText(textToReplace);
  var replacementIndex = 1;

  while (textElements != null && sortedArray.length >= replacementIndex) {
    var element = textElements.getElement();
    var startOffset = textElements.getStartOffset();
    var endOffset = textElements.getEndOffsetInclusive();
    var content = element.getText();

    if (startOffset >= 0 && endOffset < content.length) {
      var textToReplaceExact = content.substring(startOffset, endOffset + 1);
      
      if (textToReplaceExact === "2.") {
        element.deleteText(startOffset, endOffset);
        element.insertText(startOffset, "2." + replacementIndex + " " + sortedArray[replacementIndex - 1].organization);
        replacementIndex++;
      }
    }

    textElements = body.findText(textToReplace, textElements);
  }

  var tables = minutes.getBody().getTables();
  for (var j = 0; j < sortedArray.length; j++) {
    var table = tables[j];
    // Set total amount and requested
    var totalAmountCell = table.getRow(1).getCell(2);
    var requestedCell = table.getRow(2).getCell(2);

    if (sortedArray[j].totalAmount < 0) {
      totalAmountCell.setText("-$" + Math.abs(sortedArray[j].totalAmount).toFixed(2));
    } else {
      totalAmountCell.setText("$" + Math.abs(sortedArray[j].totalAmount).toFixed(2));
    }

    if (sortedArray[j].requested < 0) {
      requestedCell.setText("-$" + Math.abs(sortedArray[j].requested).toFixed(2));
    } else {
      requestedCell.setText("$" + Math.abs(sortedArray[j].requested).toFixed(2));
    }
  }

}



function findTotalAmount(file) {
  var application = SpreadsheetApp.openById(file.getId());
  var budget = application.getSheetByName("Budget");
  if (!budget) {
    return "Error";
  }
  var columnData = budget.getRange("D:D").getValues();
  var rowIndex = -1;

  for (var i = 0; i < columnData.length; i++) {
    if (columnData[i][0] === "Total Amount Requested") {
      rowIndex = i;
      break;
    }
  }

  if (rowIndex !== -1 && rowIndex < columnData.length - 1) {
    var requested = columnData[rowIndex + 1][0];
    return requested;
  } else {
    return "Not Found";
  }
}

function findRequested(file) {
  var application = SpreadsheetApp.openById(file.getId());
  var budget = application.getSheetByName("Budget");
  if (!budget) {
    return "Error";
  }
  var columnData = budget.getRange("D:D").getValues();
  var rowIndex = -1;

  for (var i = 0; i < columnData.length; i++) {
    if (columnData[i][0] === "Total Amount Requested") {
      rowIndex = i;
      break;
    }
  }

  if (rowIndex !== -1 && rowIndex >= 4) {
    var totalAmount = columnData[rowIndex - 3][0];
    return totalAmount;
  } else {
    return "Not Found";
  }
}

function findOrganization(file) {
  var application = SpreadsheetApp.openById(file.getId());
  var budget = application.getSheetByName("Application");
  if (!budget) {
    return "Error";
  }
  var organization = budget.getRange('F15').getValue();
  if (!organization) {
    return budget.getRange('F14').getValue();
  }
  return organization;
}
