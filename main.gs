const TEMPLATE_SHEET_NAME = 'Template';
const NAMES_SHEET_NAME = 'Names';
const SIDEBAR_DATA_NAME = 'Sidebar Data';

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Actions')
      .addItem('Duplicate Template', 'duplicateTemplate')
      .addItem('Show Sidebar', 'showSidebar')
      .addSeparator()
      .addItem('Delete Grades', 'deleteGrades')
      .addToUi();
}

function showSidebar() {

  const title = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SIDEBAR_DATA_NAME).getRange(2, 3).getValue();

  var html = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle(title);
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showSidebar(html);

}
// ##################

function getSidebarElementsMS() {
  return getVTableOfElementsMS().map((element) => {
    return element[0];
  });
}



//#############

function getSidebarElementsMA() {
  return getVTableOfElementsMA().map((element) => {
    return element[0];
  });
}

//###################

function duplicateTemplate() {

  //get Vertical Table of names
  const names = getVTableOfNames();

  //for each name in list of names
  for (rowIndex in names) {

    //duplicate the template
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TEMPLATE_SHEET_NAME).copyTo(SpreadsheetApp.getActiveSpreadsheet());

    //rename the copied template to the student name
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Copy of ' + TEMPLATE_SHEET_NAME).setName(names[rowIndex][0]);

  }

}

function deleteGrades() {

  //if they don't want to delete
  if (!confirmDeletion()) {
    //exit
    return;
  }

  //get Vertical Table of names
  const names = getVTableOfNames();

  for (rowIndex in names) {

    const sheetToBeDeleted = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(names[rowIndex][0]);

    //check if the sheet to be deleted couldn't be found
    if (sheetToBeDeleted == null) {
      continue;
    }

    SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheetToBeDeleted);

  }

}

function confirmDeletion() {

    const result = SpreadsheetApp.getUi().alert(
     'Please confirm',
     'Are you sure you want to delete all grades?',
      SpreadsheetApp.getUi().ButtonSet.YES_NO);

    return result == SpreadsheetApp.getUi().Button.YES;

}
//#############
function getNumRowsOfElementsMS() {
  return parseInt(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SIDEBAR_DATA_NAME).getRange(2, 6).getValue());
}

function getVTableOfElementsMS() {

  //get num of rows of elements
  const numRows = getNumRowsOfElementsMS();

  //returning verticle table of elements
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SIDEBAR_DATA_NAME).getRange(2, 5, numRows, 1).getValues();
}


//#############
function getNumRowsOfElementsMA() {
  return parseInt(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SIDEBAR_DATA_NAME).getRange(2, 2).getValue());
}

function getVTableOfElementsMA() {

  //get num of rows of elements
  const numRows = getNumRowsOfElementsMA();

  //returning verticle table of elements
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SIDEBAR_DATA_NAME).getRange(2, 1, numRows, 1).getValues();
}


//#############
function getNumRowsOfNames() {
  return parseInt(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NAMES_SHEET_NAME).getRange(2, 2).getValue());
}

function getVTableOfNames() {

  //get num of rows of names
  const numRows = getNumRowsOfNames();

  //returning verticle table of names
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NAMES_SHEET_NAME).getRange(2, 1, numRows, 1).getValues();
}