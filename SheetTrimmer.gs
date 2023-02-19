/*
    This script cleans up Google spreadsheets:
        * it deletes unused columns and rows
        * it auto-resizes columns and rows to existing text
        * it freezes rows and columns

    It can be run against the active sheet or all sheets
    It can be run standalone (externally) or within a single spreadsheet
    It can run one or more of the functions described above

    To run in standalond mode:
      The text of this script needs to be saved in google drive as a Google Apps Script. 
      If you don't see that option in the New dropdown menu, then select +Connect More Apps.

    To run in a specific sheet:
      Select from the menu Extensions/Apps Script
      Past the text of this script to replace the contents of the new script
      Run the function onOpen() or reload the spreadsheet
      Choose functions from the new menu option Colby Enterprises

*/

/*
    Standalone mode. The following constants and functions are for the standalone service
    if this code is run from a top-level external script file within Google docs
*/

// true: trim empty rows/columns. false: do not.
const TRIM = true;

// 0: no rows/columns frozen. Any positive value indicates how many to freeze.
const FROZEN_ROWS = 1;
const FROZEN_COLS = 1;

// true: autoresize. false: do not.
const AUTO_RESIZE = true;

// the id of the spreadsheet to trim/format.
const SPREADSHEET_ID =  "1K1BtPI9uZxbaNKyK1_rBNFxQeDMGCzEdCqIg6DCFNsQ";

function trimSheetsExternally() {
  let allSheets = true; let trim = TRIM; let resize = AUTO_RESIZE; 
  let freezeLeft = (FROZEN_COLS == 0) ? false : true;
  let freezeTop = (FROZEN_ROWS == 0) ? false : true;
  let leftQ = FROZEN_COLS; let topQ = FROZEN_ROWS;
  let spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  doTrim(spreadsheet, allSheets, trim, resize, freezeTop, topQ, freezeLeft, leftQ);
}

/*

    Embedded mode. The following functions can ONLY be run from within a spreadsheet. 
*/

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('Colby Enterprises').addSeparator();
  var scopeOptions = ["All sheets", "Current sheet"];
  var trimOptions = ["Trim", "Do not trim"];
  var resizeOptions = ["Auto resize", "Do not auto resize"];
  var freezeOptions = ["Freeze top row", "Freeze left column", "Freeze row and column"];

  for (var i=0; i < scopeOptions.length; i++) {
    var scopeOption = scopeOptions[i];
    var subMenuScope = ui.createMenu(scopeOption); 
    for (var j=0; j < trimOptions.length; j++) {
      var trimOption = [trimOptions[j]];
      var subMenuTrim = ui.createMenu(trimOption);
      for (var k=0; k < resizeOptions.length; k++) {
        var resizeOption = resizeOptions[k];
        var subMenuResize = ui.createMenu(resizeOption); 
        for (var l=0; l < freezeOptions.length; l++) {
          var freezeOption = freezeOptions[l];
          var fName = (scopeOption + trimOption + resizeOption + freezeOption).replaceAll(" ", "").toLowerCase();
          subMenuResize.addItem(freezeOption, fName);
        }
        subMenuTrim.addSubMenu(subMenuResize);
      }
      subMenuScope.addSubMenu(subMenuTrim);
    }
    menu.addSubMenu(subMenuScope);
  }

  menu.addToUi();
}

// This the function that does all the work.
function doTrim(ss,allSheets, trim, resize, freezeTop, topQ, freezeLeft, leftQ) {
  var sheets = null;
  if (allSheets) {
    sheets = ss.getSheets();
  }
  else {
    sheets = [ss.getActiveSheet()];
  }
  
  for (var j=0; j < sheets.length; j++) {
    var sheet = sheets[j];
    var lastColumn = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();
    if (trim) {
      var colsToDelete =  sheet.getMaxColumns() - lastColumn;
      if (colsToDelete > 0) {
        sheet.deleteColumns(lastColumn+1,colsToDelete);
      }
      var rowsToDelete =  sheet.getMaxRows() - lastRow;
      if (rowsToDelete > 0) {
        sheet.deleteRows(lastRow+1, rowsToDelete);
      }
    }
    if (resize) {
      sheet.autoResizeColumns(1, lastColumn);
      sheet.autoResizeRows(1, lastRow);
    }
    if (freezeTop) {
      sheet.setFrozenRows(topQ);
    }
    if (freezeLeft) {
      sheet.setFrozenColumns(leftQ);
    }
  }
}

/*
    The following functions are all the permutations of the options on the menu. 
*/

function allsheetstrimautoresizefreezetoprow() {
  allSheets = true; trim = true; resize = true; freezeTop = true; topQ = 1; freezeLeft = false; leftQ = null;
  var spreadsheet = SpreadsheetApp.getActive();
  doTrim(spreadsheet, allSheets, trim, resize, freezeTop, topQ, freezeLeft, leftQ);
}

function allsheetstrimautoresizefreezeleftcolumn() {
  allSheets = true; trim = true; resize = true; freezeTop = false; topQ = null; freezeLeft = true; leftQ = 1;
  var spreadsheet = SpreadsheetApp.getActive();
  doTrim(spreadsheet, allSheets, trim, resize, freezeTop, topQ, freezeLeft, leftQ);
}

function allsheetstrimautoresizefreezerowandcolumn() {
  allSheets = true; trim = true; resize = true; freezeTop = true; topQ = 1; freezeLeft = true; leftQ = 1;
  var spreadsheet = SpreadsheetApp.getActive();
  doTrim(spreadsheet, allSheets, trim, resize, freezeTop, topQ, freezeLeft, leftQ);
}

function allsheetstrimdonotautoresizefreezetoprow() {
  allSheets = true; trim = true; resize = false; freezeTop = true; topQ = 1; freezeLeft = false; leftQ = 0;
  var spreadsheet = SpreadsheetApp.getActive();
  doTrim(spreadsheet, allSheets, trim, resize, freezeTop, topQ, freezeLeft, leftQ);
}

function allsheetstrimdonotautoresizefreezeleftcolumn() {
  allSheets = true; trim = true; resize = false; freezeTop = false; topQ = null; freezeLeft = true; leftQ = 1;
  var spreadsheet = SpreadsheetApp.getActive();
  doTrim(spreadsheet, allSheets, trim, resize, freezeTop, topQ, freezeLeft, leftQ);
}

function allsheetstrimdonotautoresizefreezerowandcolumn() {
  allSheets = true; trim = true; resize = false; freezeTop = true; topQ = 1; freezeLeft = true; leftQ = 1;
  var spreadsheet = SpreadsheetApp.getActive();
  doTrim(spreadsheet, allSheets, trim, resize, freezeTop, topQ, freezeLeft, leftQ);
}

function allsheetsdonottrimautoresizefreezetoprow() {
  allSheets = true; trim = false; resize = true; freezeTop = true; topQ = 1; freezeLeft = false; leftQ = null;
  var spreadsheet = SpreadsheetApp.getActive();
  doTrim(spreadsheet, allSheets, trim, resize, freezeTop, topQ, freezeLeft, leftQ);
}

function allsheetsdonottrimautoresizefreezeleftcolumn() {
  allSheets = true; trim = false; resize = true; freezeTop = false; topQ = null; freezeLeft = true; leftQ = 1;
  var spreadsheet = SpreadsheetApp.getActive();
  doTrim(spreadsheet, allSheets, trim, resize, freezeTop, topQ, freezeLeft, leftQ);
}

function allsheetsdonottrimautoresizefreezerowandcolumn() {
  allSheets = true; trim = false; resize = true; freezeTop = true; topQ = 1; freezeLeft = true; leftQ = 1;
  var spreadsheet = SpreadsheetApp.getActive();
  doTrim(spreadsheet, allSheets, trim, resize, freezeTop, topQ, freezeLeft, leftQ);
}

function allsheetsdonottrimdonotautoresizefreezetoprow() {
  allSheets = true; trim = false; resize = false; freezeTop = true; topQ = 1; freezeLeft = false; leftQ = null;
  var spreadsheet = SpreadsheetApp.getActive();
  doTrim(spreadsheet, allSheets, trim, resize, freezeTop, topQ, freezeLeft, leftQ);
}

function allsheetsdonottrimdonotautoresizefreezeleftcolumn() {
  allSheets = true; trim = false; resize = false; freezeTop = false; topQ = null; freezeLeft = true; leftQ = 1;
  var spreadsheet = SpreadsheetApp.getActive();
  doTrim(spreadsheet, allSheets, trim, resize, freezeTop, topQ, freezeLeft, leftQ);
}

function allsheetsdonottrimdonotautoresizefreezerowandcolumn() {
  allSheets = true; trim = false; resize = false; freezeTop = true; topQ = 1; freezeLeft = true; leftQ = 1;
   var spreadsheet = SpreadsheetApp.getActive();
  doTrim(spreadsheet, allSheets, trim, resize, freezeTop, topQ, freezeLeft, leftQ);
}

function currentsheettrimautoresizefreezetoprow() {
  allSheets = false; trim = true; resize = true; freezeTop = true; topQ = 1; freezeLeft = false; leftQ = null;
  var spreadsheet = SpreadsheetApp.getActive();
  doTrim(spreadsheet, allSheets, trim, resize, freezeTop, topQ, freezeLeft, leftQ);
}


function currentsheettrimautoresizefreezeleftcolumn() {
  allSheets = false; trim = true; resize = true; freezeTop = false; topQ = null; freezeLeft = true; leftQ = 1;
  var spreadsheet = SpreadsheetApp.getActive();
  doTrim(spreadsheet, allSheets, trim, resize, freezeTop, topQ, freezeLeft, leftQ);
}

function currentsheettrimautoresizefreezerowandcolumn() {
  allSheets = false; trim = true; resize = true; freezeTop = true; topQ = 1; freezeLeft = true; leftQ = 1;
  var spreadsheet = SpreadsheetApp.getActive();
  doTrim(spreadsheet, allSheets, trim, resize, freezeTop, topQ, freezeLeft, leftQ);
}


function currentsheettrimdonotautoresizefreezetoprow() {
  allSheets = false; trim = true; resize = false; freezeTop = true; topQ = 1; freezeLeft = false; leftQ = 0;
  var spreadsheet = SpreadsheetApp.getActive();
  doTrim(spreadsheet, allSheets, trim, resize, freezeTop, topQ, freezeLeft, leftQ);
}


function currentsheettrimdonotautoresizefreezeleftcolumn() {
  allSheets = false; trim = true; resize = false; freezeTop = false; topQ = null; freezeLeft = true; leftQ = 1;
  var spreadsheet = SpreadsheetApp.getActive();
  doTrim(spreadsheet, allSheets, trim, resize, freezeTop, topQ, freezeLeft, leftQ);
}


function currentsheettrimdonotautoresizefreezerowandcolumn() {
  allSheets = false; trim = true; resize = false; freezeTop = true; topQ = 1; freezeLeft = true; leftQ = 1;
  var spreadsheet = SpreadsheetApp.getActive();
  doTrim(spreadsheet, allSheets, trim, resize, freezeTop, topQ, freezeLeft, leftQ);
}


function currentsheetdonottrimautoresizefreezetoprow() {
  allSheets = false; trim = false; resize = true; freezeTop = true; topQ = 1; freezeLeft = false; leftQ = null;
  var spreadsheet = SpreadsheetApp.getActive();
  doTrim(spreadsheet, allSheets, trim, resize, freezeTop, topQ, freezeLeft, leftQ);
}


function currentsheetdonottrimautoresizefreezeleftcolumn() {
  allSheets = false; trim = false; resize = true; freezeTop = false; topQ = null; freezeLeft = true; leftQ = 1;
  var spreadsheet = SpreadsheetApp.getActive();
  doTrim(spreadsheet, allSheets, trim, resize, freezeTop, topQ, freezeLeft, leftQ);
}


function currentsheetdonottrimautoresizefreezerowandcolumn() {
  allSheets = false; trim = false; resize = true; freezeTop = true; topQ = 1; freezeLeft = true; leftQ = 1;
  var spreadsheet = SpreadsheetApp.getActive();
  doTrim(spreadsheet, allSheets, trim, resize, freezeTop, topQ, freezeLeft, leftQ);
}

function currentsheetdonottrimdonotautoresizefreezetoprow() {
  allSheets = false; trim = false; resize = false; freezeTop = true; topQ = 1; freezeLeft = false; leftQ = null;
   var spreadsheet = SpreadsheetApp.getActive();
  doTrim(spreadsheet, allSheets, trim, resize, freezeTop, topQ, freezeLeft, leftQ);
}


function currentsheetdonottrimdonotautoresizefreezeleftcolumn() {
  allSheets = false; trim = false; resize = false; freezeTop = false; topQ = null; freezeLeft = true; leftQ = 1;
  var spreadsheet = SpreadsheetApp.getActive();
  doTrim(spreadsheet, allSheets, trim, resize, freezeTop, topQ, freezeLeft, leftQ);
}


function currentsheetdonottrimdonotautoresizefreezerowandcolumn() {
  allSheets = false; trim = false; resize = false; freezeTop = true; topQ = 1; freezeLeft = true; leftQ = 1;
  var spreadsheet = SpreadsheetApp.getActive();
  doTrim(spreadsheet, allSheets, trim, resize, freezeTop, topQ, freezeLeft, leftQ);
}
