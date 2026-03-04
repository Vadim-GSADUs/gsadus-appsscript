/**
 * GSADUs Formatting Tools (Permanent)
 * Creates a custom menu to apply UI formatting across any active tab.
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('GSADUs Tools')
    .addItem('Auto-Format Dropdown Columns', 'formatActiveSheetColumns')
    .addToUi();
}

/**
 * Scans the active sheet's first row. 
 * Shrinks "(Select)" columns and expands "(Preview)" columns.
 */
function formatActiveSheetColumns() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastCol = sheet.getLastColumn();
  
  if (lastCol === 0) return; // Exit if the sheet is completely empty

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  for (let i = 0; i < headers.length; i++) {
    let headerName = String(headers[i]);
    let colIndex = i + 1; // Columns are 1-indexed in Apps Script

    if (headerName.includes("(Select)")) {
      // Shrink to just show the dropdown arrow (approx 25-30 pixels)
      sheet.setColumnWidth(colIndex, 20);
    } else if (headerName.includes("(Preview)") || headerName.includes("(Image URL)")) {
      // Expand to ensure the hyperlink text is readable
      sheet.setColumnWidth(colIndex, 250);
    }
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast('Column formatting applied to ' + sheet.getName(), 'Success');
}