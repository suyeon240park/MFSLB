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
          [current.data.organization, next.data.organization] = [next.data.organization, current.data.organization];
          [current.data.requested, next.data.requested] = [next.data.requested, current.data.requested];
          [current.data.previous, next.data.previous] = [next.data.previous, current.data.previous];

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
  var sourceFolderId = "11X3Zf5Bdm5p7LUAEuUt3SORXo-jhyny0";
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
}

//Remove "Copy to" in the file name
function removeCopyTo() {
  var sourceFolderId = "11X3Zf5Bdm5p7LUAEuUt3SORXo-jhyny0";
  var sourceFolder = DriveApp.getFolderById(sourceFolderId);

  var files = sourceFolder.getFiles();

  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();
    var fileId = file.getId();

    // Rename the files if they contain "Copy of"
    if (fileName.startsWith("Copy of ")) {
      var newFileName = fileName.substring("Copy of ".length);
      DriveApp.getFileById(fileId).setName(newFileName);
      Logger.log(fileName + ' has been renamed.');
    }
  }
}



function copyValuesFromSheetToLL() {
  var sourceFolderId = '1PHUnuD-HO0dqQAOSjiU2RkwZAVy7YGPY';
  var sourceFolder = DriveApp.getFolderById(sourceFolderId);
  var i = 1;

  var BudgetApplication = new LinkedList();

  var files = sourceFolder.getFiles();

  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();

    Logger.log(i + ": " + fileName);

    var organization = findOrganization(file);
    var requested = Number(findRequestedAndPrevious(file)[0]).toFixed(2);
    var previous = Number(findRequestedAndPrevious(file)[1]).toFixed(2);

    Logger.log("Organization: " + organization);
    Logger.log("Requested: " + requested);
    Logger.log("Previous: " + previous);

    if (organization !== undefined && previous !== undefined && requested !== undefined) {
      BudgetApplication.add({
        organization: organization,
        requested: requested,
        previous: previous
      });
    } else {
      Logger.log('Skipping file ' + fileName + ' due to missing data.');
    }
    i++;
  }
  BudgetApplication.sort();
  const sortedArray = BudgetApplication.toArray();
  Logger.log(sortedArray);
  pasteToDocs(sortedArray);
}


function pasteToDocs(sortedArray) {
  var minutes = DocumentApp.openByUrl('https://docs.google.com/document/d/1hP5pMl-gWC3VTi28PHfPrEaA3YTE0ROudPJFJe79334/edit');
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

function findOrganization(file) {
  var application = SpreadsheetApp.openById(file.getId());
  var budget = application.getSheetByName("Information Sheet");
  if (!budget) {
    return "Error";
  }
  var organization = budget.getRange('C13').getValue();
  if (!organization) {
    organization = budget.getRange('C12').getValue();
    if (!organization) {
      organization = budget.getRange('C11').getValue();
    }
  }
  return organization;
}

function findRequestedAndPrevious(file) {
  var application = SpreadsheetApp.openById(file.getId());
  var sheets = application.getSheets();
  var budget = sheets[1];

  if (!budget) {
    return "Organization - Wrong Sheet Name.";
  }

  var columnData = budget.getRange("C:D").getValues(); // Adjust the range to include both columns C and D
  var rowIndex = -1;

  for (var i = 0; i < columnData.length; i++) {
    if (columnData[i][0] === "Amount Requested") {
      rowIndex = i;
      break;
    }
  }

  if (rowIndex !== -1 && rowIndex < columnData.length - 1) {
    var requested = columnData[rowIndex + 1][0];
  } else {
    return "Requested - Not Found.";
  }

  var resultArray = [requested];

  for (var j = 1; j <= 5; j++) { // Loop through columns D to H
    var nextColumnValue = columnData[rowIndex + 1][j];

    if (nextColumnValue !== 0) {
      // Handle the case where nextColumnValue is not 0
      resultArray.push(nextColumnValue);
      break;
    } else if (j === 5) {
      // Handle the case where nextColumnValue is 0 in all columns D to H
      resultArray.push(0);
    }
  }

  return resultArray;
}
