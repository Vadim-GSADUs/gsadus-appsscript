/**
 * Mirrors the "Folders" tab to the "GSDE Projects.csv" file.
 * Format: Standard CSV (Comma Separated)
 * Fixes "All data in Column A" issue.
 */
function mirrorFoldersToGSDE() {
  // --- CONFIGURATION ---
  const CONFIG = {
    SOURCE_SHEET: "Folders",
    FOLDER_ID: "1XmlzdvNyom6Vab6hR3NYHodR2-XDCfCE",
    FILE_NAME: "GSDE Projects.csv",
    DELIMITER: "," // CHANGED: Using Comma to ensure columns split correctly
  };

  // 1. Get Data from Sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SOURCE_SHEET);
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Sheet "${CONFIG.SOURCE_SHEET}" not found.`);
    return;
  }
  
  const data = sheet.getDataRange().getDisplayValues();

  // 2. Convert to Standard CSV Format
  const csvContent = data.map(row => {
    return row.map(field => {
      let stringValue = field.toString();
      
      // Standard CSV Rule:
      // If the cell contains a Comma, Quote, or Newline, we must wrap it in quotes.
      // We also escape existing quotes by doubling them (e.g. " becomes "")
      if (/[",\n\r]/.test(stringValue)) {
        stringValue = '"' + stringValue.replace(/"/g, '""') + '"';
      }
      return stringValue;
    }).join(CONFIG.DELIMITER);
  }).join("\n");

  // 3. Save to Google Drive
  try {
    const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    const files = folder.getFilesByName(CONFIG.FILE_NAME);
    
    if (files.hasNext()) {
      // Update existing file
      const file = files.next();
      file.setContent(csvContent);
      SpreadsheetApp.getUi().alert(`Success! Updated existing file:\n${CONFIG.FILE_NAME}\n\nIt should now open correctly with separate columns.`);
    } else {
      // Create new file if missing
      folder.createFile(CONFIG.FILE_NAME, csvContent, MimeType.CSV);
      SpreadsheetApp.getUi().alert(`Success! Created new file:\n${CONFIG.FILE_NAME}`);
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert(`Error accessing Drive:\n${e.message}`);
    Logger.log(e);
  }
}