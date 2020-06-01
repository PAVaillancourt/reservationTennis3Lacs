// Global vars
var TIME_TO_DELETION = 72*60*60*1000; // Three days
var RESERVATION_WINDOW_START = 18; // 6:00 p.m.
var RESERVATION_RANGE = "B3:D45";
var DATE_ROW = 1;
var DATE_COL = 2;
var ss = SpreadsheetApp.getActiveSpreadsheet();

// Main function that creates and deletes sheets if necessary
function sheetRoutine() {
  var currentDateStr = Utilities.formatDate(new Date(), "GMT-5", "yyyy-MM-dd" );
  var currentDate = new Date();
  var tomorrowDate = new Date();
  tomorrowDate.setDate(tomorrowDate.getDate()+1);
  var tomorrowDateStr = Utilities.formatDate(tomorrowDate, "GMT-5", "yyyy-MM-dd" );


  ss.toast("Création des nouvelles feuilles, ça ne sera pas très long ... Ne rien toucher avant le retrait de ce message. Merci!", "Attention!", 10);
  
  deleteIrrelevantSheets(ss.getSheets());
  deleteFutureSheets(ss.getSheets(), currentDate);

  // Creates a sheet if there is no current date sheet
  if (!checkIfSheetExists(ss.getSheets(), currentDateStr)) {
    createSheet(ss.getSheets(), currentDateStr);
  }
    
  createSheet(ss.getSheets(), tomorrowDateStr);

  // Removes old sheets
  deleteOldSheets(ss.getSheets(), currentDate);
}

// Deletes sheets that are older than the current date
function deleteOldSheets(sheets, currentDate) {
  sheets.forEach(sheet => {
    var sheetDateStr = sheet.getSheetName();
    var sheetDate = new Date(sheetDateStr);
    var timeDiff = currentDate - sheetDate;
    if (currentDate - sheetDate > TIME_TO_DELETION) {
      ss.deleteSheet(sheet);
    }
  });
}

// Deletes sheets that are irrelevant (added by other users)
function deleteIrrelevantSheets(sheets) {
  let dateRegEx = /^(19|20)\d\d[- /.](0[1-9]|1[012])[- /.](0[1-9]|[12][0-9]|3[01])$/;
  sheets.forEach(sheet => {
    var sheetDateStr = sheet.getSheetName();
    var regexTest = sheetDateStr.match(dateRegEx);
    if (sheetDateStr.match(dateRegEx) == null) {
      ss.deleteSheet(sheet);
    }
  });
}


function deleteFutureSheets(sheets, currentDate) {
  sheets.forEach(sheet => {
    var sheetDateStr = sheet.getSheetName();
    var sheetDateHNE = new Date(sheetDateStr);
    var sheetDate = new Date(sheetDateHNE.getTime() + Math.abs(sheetDateHNE.getTimezoneOffset()*60000));
    var delay = sheetDate - currentDate;
    if (sheetDate > currentDate) {
      ss.deleteSheet(sheet);
    }
  });
}


function createSheet(sheets, currentDateStr) {
  // Create a sheet for the current date
  var oldSheet = sheets[0];
  var newSheet = oldSheet.copyTo(ss);
  newSheet.setName(currentDateStr);

  var reservationsRange = newSheet.getRange(RESERVATION_RANGE);
  reservationsRange.clearContent();

  var dateCell = newSheet.getRange(DATE_ROW, DATE_COL);
  dateCell.setValue(currentDateStr);

  copyPermissions(oldSheet, newSheet);
  
  ss.setActiveSheet(newSheet);
  ss.moveActiveSheet(1);
}

// Returns true if a sheet matches the date
function checkIfSheetExists(sheets, dateStr) {
  for (let sheet of sheets) {
    var targetSheetName = sheet.getSheetName();
    var localeCompareResult = targetSheetName.localeCompare(dateStr);
    if (targetSheetName.localeCompare(dateStr) == 0) {
      return true;
    }
  }

  return false;
}

// Copies the permissions of an old sheet to a new sheet
// Source: https://webapps.stackexchange.com/questions/86984/in-google-sheets-how-do-i-duplicate-a-sheet-along-with-its-permission/87000#87000
function copyPermissions(oldSheet, newSheet) {
  var p = oldSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];
  var p2 = newSheet.protect();
  p2.setDescription(p.getDescription());
  p2.setWarningOnly(p.isWarningOnly());  
  if (!p.isWarningOnly()) {
    p2.removeEditors(p2.getEditors());
    p2.addEditors(p.getEditors());
    // p2.setDomainEdit(p.canDomainEdit()); //  only if using an Apps domain 
  }
  var ranges = p.getUnprotectedRanges();
  var newRanges = [];
  for (var i = 0; i < ranges.length; i++) {
    newRanges.push(newSheet.getRange(ranges[i].getA1Notation()));
  } 
  p2.setUnprotectedRanges(newRanges);
}
