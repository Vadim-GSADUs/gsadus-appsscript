/**
 * Adds a custom menu when the spreadsheet is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ADU Tools')
    .addItem('Convert CSV to Table', 'convertCsvToTable')
    .addToUi();
}

/**
 * Finds the active sheet’s data range, freezes and bolds the header row,
 * applies a filter, and creates a named range “ADU_Catalog” for easy references.
 */
function convertCsvToTable() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No data found—but make sure you’ve pasted your CSV with headers in row 1.');
    return;
  }

  // 1) Freeze the first (header) row
  sheet.setFrozenRows(1);

  // 2) Bold the header row
  sheet.getRange(1, 1, 1, lastCol).setFontWeight('bold');

  // 3) Apply a filter if one isn’t already there
  if (!sheet.getFilter()) {
    sheet.getRange(1, 1, lastRow, lastCol).createFilter();
  }

  // 4) Create (or overwrite) a named range for structured references
  //    covering the whole table including headers
  ss.setNamedRange(
    'ADU_Catalog',
    sheet.getRange(1, 1, lastRow, lastCol)
  );

  SpreadsheetApp.getUi().alert(
    'Done! Your CSV is now a “table”:\n• Header row frozen & bolded\n• Filter applied\n• Named range “ADU_Catalog” created'
  );
}
