// Tasks
// 1. Read the date from the minutes (don't want to type it each time)

// top module function
function readMinutes() {
  var minutes = DocumentApp.openByUrl('https://docs.google.com/document/d/1hP5pMl-gWC3VTi28PHfPrEaA3YTE0ROudPJFJe79334/edit');

  var projectDirectorshipsList = readPD(minutes);
  var specialProjectsList = readSP(minutes);
  var conferneceFundingList = readCF(minutes);

  var tracker = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1fMkXSI_BtdLkwQIGCA4q2deJqmuW07iS/');

  record(tracker, "Operating Budget", projectDirectorshipsList);
  record(tracker, "Special Projects (UTSU)", specialProjectsList);
  record(tracker, "Conference Funding (UTSU)", conferneceFundingList);
}

// read Project Directorships
function readPD(minutes) {
  var paragraphs = minutes.getBody().getParagraphs();
  var projectDirectorshipsList = new LinkedList();
  var accountIncrement = 1;
  var startReading = false;
  var searchTerm = "3.0 Budget Submissions";
  var startIndex = findStartIndex(minutes, searchTerm);

  // Loop through each paragraph
  for (var i = startIndex; i < paragraphs.length; i++) {
    var text = paragraphs[i].getText();

    // Check if it's the starting point
    if (text.indexOf(searchTerm) !== -1) {
      startReading = true;
      continue;
    }

    // Check if it's the ending point
    if (text.indexOf('4.0 Special Projects') !== -1) {
      break;
    }

    // Process paragraphs between start and end points
    if (startReading) {
      // Check for 'account'
      var accountMatch = text.match(/3\.(\d+)\s+([^0-9]+)/);
      if (accountMatch) {
        var currentAccountIncrement = parseInt(accountMatch[1], 10);
        var account = accountMatch[2].trim();

        // Check if the account matches the expected increment
        if (currentAccountIncrement === accountIncrement) {
          Logger.log("Account: " + account);
          accountIncrement++;
        } else {
          Logger.log('Skipping:', text);
        }
      }

      if (text.includes('Requested')) {
        var requested = paragraphs[i + 1].getText();
        Logger.log("Requested: " + requested);
      }

      if (text.includes('Approved')) {
        var approved = paragraphs[i + 1].getText();
        Logger.log("Approved: " + approved);
      }
    }

    if (account !== undefined && requested !== undefined && approved !== undefined) {
      projectDirectorshipsList.add({
        account: account,
        requested: requested,
        approved: approved
      });
      account = undefined;
      requested = undefined;
      approved = undefined;
    }
  }
  Logger.log(projectDirectorshipsList.toArray());
  return projectDirectorshipsList;
}

//read Special Projects
function readSP(minutes) {
  var paragraphs = minutes.getBody().getParagraphs();
  var specialProjectsList = new LinkedList();
  var accountIncrement = 1;
  var startReading = false;
  var searchTerm = "4.0 Special Projects Funding";
  var startIndex = findStartIndex(minutes, searchTerm);

  for (var i = startIndex; i < paragraphs.length; i++) {
    var text = paragraphs[i].getText();

    if (text.indexOf(searchTerm) !== -1) {
      startReading = true;
      continue;
    }

    if (text.indexOf('5.0 Conference Funding') !== -1) {
      break;
    }

    if (startReading) {
      var accountMatch = text.match(/4\.(\d+)\s+([^0-9]+)/);
      if (accountMatch) {
        var currentAccountIncrement = parseInt(accountMatch[1], 10);
        var account = accountMatch[2].trim();

        if (currentAccountIncrement === accountIncrement) {
          Logger.log("Account: " + account);
          accountIncrement++;
        } else {
          Logger.log('Skipping:', text);
        }
      }

      if (text.includes('Requested')) {
        var requested = paragraphs[i + 1].getText();
        Logger.log("Requested: " + requested);
      }

      if (text.includes('Approved')) {
        var approved = paragraphs[i + 1].getText();
        Logger.log("Approved: " + approved);
      }
    }

    if (account !== undefined && requested !== undefined && approved !== undefined) {
      specialProjectsList.add({
        account: account,
        requested: requested,
        approved: approved
      });
      account = undefined;
      requested = undefined;
      approved = undefined;
    }
  }
  Logger.log(specialProjectsList.toArray());
  return specialProjectsList;
}

// read Conference Funding
function readCF(minutes) {
  var paragraphs = minutes.getBody().getParagraphs();
  var conferneceFundingList = new LinkedList();
  var accountIncrement = 1;
  var startReading = false;
  var searchTerm = "5.0 Conference Funding";
  var startIndex = findStartIndex(minutes, searchTerm);

  for (var i = startIndex; i < paragraphs.length; i++) {
    var text = paragraphs[i].getText();

    if (text.indexOf(searchTerm) !== -1) {
      startReading = true;
      continue;
    }

    if (text.indexOf('6.0 Levy Funds') !== -1) {
      break;
    }

    if (startReading) {
      var accountMatch = text.match(/5\.(\d+)\s+([^0-9]+)/);
      if (accountMatch) {
        var currentAccountIncrement = parseInt(accountMatch[1], 10);
        var account = accountMatch[2].trim();

        if (currentAccountIncrement === accountIncrement) {
          Logger.log("Account: " + account);
          accountIncrement++;
        } else {
          Logger.log('Skipping:', text);
        }
      }

      if (text.includes('Requested')) {
        var requested = paragraphs[i + 1].getText();
        Logger.log("Requested: " + requested);
      }

      if (text.includes('Approved')) {
        var approved = paragraphs[i + 1].getText();
        Logger.log("Approved: " + approved);
      }
    }

    if (account !== undefined && requested !== undefined && approved !== undefined) {
      conferneceFundingList.add({
        account: account,
        requested: requested,
        approved: approved
      });
      account = undefined;
      requested = undefined;
      approved = undefined;
    }
  }
  Logger.log(conferneceFundingList.toArray());
  return conferneceFundingList;
}

function findStartIndex(minutes, searchTerm) {
  var paragraphs = minutes.getBody().getParagraphs();
  var occurrenceCount = 0;

  for (var i = 0; i < paragraphs.length; i++) {
    var text = paragraphs[i].getText();

    if (text.indexOf(searchTerm) !== -1) {
      occurrenceCount++;
      if (occurrenceCount === 2) {
        return i; // Return the index of the second occurrence
      }
    }
  }
  return -1;
}

// record Project Directorships on the tracker
function recordPD(tracker, projectDirectorshipsList) {
  var sheetName = "Operating Budget";
  var sheet = tracker.getSheetByName(sheetName);

  if (!sheet) {
    Logger.log("Sheet '" + sheetName + "' not found.");
    return;
  }

  var startRow = 1;
  var date = "Nov-11";

  var list = projectDirectorshipsList.toArray();

  for (var i = 0; i < list.length; i++) {
    var rowData = list[i];

    // Write 'account' in column B
    sheet.getRange(startRow, 2).setValue(rowData.account).setFontWeight("bold");

    sheet.getRange(startRow + 3, 2).setValue(rowData.account + " Total");

    // Write 'approved' in column C
    sheet.getRange(startRow + 1, 3).setValue(rowData.approved);

    // Write 'requested' in column D
    sheet.getRange(startRow + 1, 4).setValue(rowData.requested);

    // Write 'date' in column E
    sheet.getRange(startRow + 1, 5).setValue(date);

    sheet.getRange(startRow + 3, 3).setFormula('=SUM(C' + (startRow) + ':C' + (startRow + 2) + ')');
    sheet.getRange(startRow + 3, 4).setFormula('=SUM(D' + (startRow) + ':D' + (startRow + 2) + ')');

    for (var j = 0; j < 10; j++) {
      sheet.getRange(startRow + 3, 2+j).setBackgroundColor("#fff2cc");
      sheet.getRange(startRow + 2, 2+j, 1, 1).setBorder(false, false, true, false, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }

    startRow += 5;
  }

  Logger.log("Data recorded successfully.");
}

function record(tracker, sheetName, appList) {
  var sheet = tracker.getSheetByName(sheetName);

  if (!sheet) {
    Logger.log("Sheet '" + sheetName + "' not found.");
    return;
  }

  var list = appList.toArray();

  var startRow = 1;
  var date = "Jan-20";

  for (var i = 0; i < list.length; i++) {
    var rowData = list[i];

    // Write 'account' in column B
    sheet.getRange(startRow, 2).setValue(rowData.account).setFontWeight("bold");

    sheet.getRange(startRow + 3, 2).setValue(rowData.account + " Total");

    // Write 'approved' in column C
    sheet.getRange(startRow + 1, 3).setValue(rowData.approved);

    // Write 'requested' in column D
    sheet.getRange(startRow + 1, 4).setValue(rowData.requested);

    // Write 'date' in column E
    sheet.getRange(startRow + 1, 5).setValue(date);

    sheet.getRange(startRow + 3, 3).setFormula('=SUM(C' + (startRow) + ':C' + (startRow + 2) + ')');
    sheet.getRange(startRow + 3, 4).setFormula('=SUM(D' + (startRow) + ':D' + (startRow + 2) + ')');

    for (var j = 0; j < 10; j++) {
      sheet.getRange(startRow + 3, 2+j).setBackgroundColor("#fff2cc");
      sheet.getRange(startRow + 2, 2+j, 1, 1).setBorder(false, false, true, false, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }

    startRow += 5;
  }

  Logger.log("Data recorded successfully.");

}
