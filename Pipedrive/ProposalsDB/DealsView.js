// ---------------------------
// Deals - User-facing sheet with field protection
// ---------------------------

/**
 * Fast data-only refresh: updates data rows without touching formatting.
 * Use this for routine pulls - it's 10-20x faster than full formatting refresh.
 * Call refreshDealsView() only when column structure changes or formatting needs reset.
 */
function refreshDealsViewDataOnly() {
  const ss = SpreadsheetApp.getActive();
  const sourceSheet = ss.getSheetByName('API_Deals');
  
  if (!sourceSheet) {
    throw new Error('API_Deals sheet not found. Run PullFromPipedrive() first.');
  }
  
  let viewSheet = ss.getSheetByName('Deals');
  const sourceData = sourceSheet.getDataRange().getValues();
  const numRows = sourceData.length;
  const numCols = sourceData[0].length;
  
  // If sheet doesn't exist or structure changed, do full refresh
  if (!viewSheet || viewSheet.getLastColumn() !== numCols) {
    Logger.log('Structure change detected, performing full refresh with formatting...');
    refreshDealsView();
    return;
  }
  
  // Fast path: just update the data
  viewSheet.getRange(1, 1, numRows, numCols).setValues(sourceData);
  
  // Resize sheet if row count changed
  const currentRows = viewSheet.getMaxRows();
  if (currentRows < numRows) {
    viewSheet.insertRowsAfter(currentRows, numRows - currentRows);
  } else if (currentRows > numRows && currentRows > 1000) {
    // Clean up excess rows if we have way too many (keep at least 1000)
    viewSheet.deleteRows(numRows + 1, currentRows - numRows);
  }
  
  Logger.log('Deals sheet data refreshed (fast): ' + (numRows - 1) + ' deals, ' + numCols + ' columns');
  
  if (typeof logEvent_ === 'function') {
    logEvent_('VIEW', 'refreshDealsViewDataOnly', 'Rows: ' + (numRows - 1) + ', Cols: ' + numCols);
  }
}

/**
 * Full refresh with formatting: creates/recreates the Deals sheet with all color coding and protection.
 * This is SLOW (40-50 seconds) due to column-by-column formatting. 
 * Only use when:
 * - First time creating the sheet
 * - Column structure has changed
 * - Formatting needs to be reset
 * 
 * For routine data updates, use refreshDealsViewDataOnly() instead.
 */
function refreshDealsView() {
  const ss = SpreadsheetApp.getActive();
  const sourceSheet = ss.getSheetByName('API_Deals');
  
  if (!sourceSheet) {
    throw new Error('API_Deals sheet not found. Run PullFromPipedrive() first.');
  }
  
  // Get or create Deals sheet
  let viewSheet = ss.getSheetByName('Deals');
  if (!viewSheet) {
    viewSheet = ss.insertSheet('Deals');
  }
  
  // Copy all data from API_Deals to Deals
  const sourceData = sourceSheet.getDataRange().getValues();
  const numRows = sourceData.length;
  const numCols = sourceData[0].length;
  
  // Clear existing content
  viewSheet.clear();
  
  // Write data
  viewSheet.getRange(1, 1, numRows, numCols).setValues(sourceData);
  
  // Get field classification map
  const classMap = getFieldClassificationMap();
  const headers = sourceData[0];
  
  // Apply formatting and protection (SLOW - this is where the 40-50 seconds come from)
  applyFieldFormatting_(viewSheet, headers, classMap);
  
  Logger.log('Deals sheet fully refreshed with formatting: ' + (numRows - 1) + ' deals, ' + numCols + ' columns');
  
  if (typeof logEvent_ === 'function') {
    logEvent_('VIEW', 'refreshDealsView', 'Rows: ' + (numRows - 1) + ', Cols: ' + numCols);
  }
}

/**
 * Apply color coding and protection to columns based on field classification.
 * @private
 */
function applyFieldFormatting_(sheet, headers, classMap) {
  const numRows = sheet.getMaxRows();
  const numCols = headers.length;
  
  // Remove all existing protections on this sheet
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  protections.forEach(function(p) {
    p.remove();
  });
  
  // Process each column
  for (let col = 1; col <= numCols; col++) {
    const header = headers[col - 1];
    const fieldInfo = classMap[header];
    
    if (!fieldInfo) {
      // Unknown field - mark as warning
      sheet.getRange(1, col, numRows, 1)
        .setBackground('#FFCDD2') // Light red
        .setNote('Unknown field - not classified');
      continue;
    }
    
    const category = fieldInfo.category;
    const color = getCategoryColor(category);
    
    // Apply background color to entire column
    sheet.getRange(1, col, numRows, 1).setBackground(color);
    
    // Add note to header explaining editability
    let note = 'Category: ' + category;
    if (fieldInfo.notes) {
      note += '\n' + fieldInfo.notes;
    }
    
    if (category === 'SYSTEM_CALCULATED') {
      note += '\n\n⚠️ READ-ONLY: This field is calculated by Pipedrive and cannot be edited.';
      
      // Protect the entire column (except header for sorting)
      if (numRows > 1) {
        const protection = sheet.getRange(2, col, numRows - 1, 1).protect();
        protection.setDescription('System-calculated field: ' + header);
        protection.setWarningOnly(false);
        
        // Remove all editors (only owner can edit)
        const me = Session.getEffectiveUser();
        protection.addEditor(me);
        protection.removeEditors(protection.getEditors());
      }
    } else if (category === 'SPECIAL_HANDLING') {
      note += '\n\n⚠️ CAUTION: Editing this field requires special handling. See documentation.';
    } else if (category === 'USER_EDITABLE') {
      note += '\n\n✓ EDITABLE: This field can be safely edited and pushed back to Pipedrive.';
    }
    
    sheet.getRange(1, col).setNote(note);
  }
  
  // Freeze header row
  sheet.setFrozenRows(1);
  
  // Bold and center align header row
  sheet.getRange(1, 1, 1, numCols)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
}

/**
 * Creates or refreshes the API_Deals_Shadow sheet for change detection.
 * This is a snapshot of API_Deals at the time of last pull.
 * Should be called after each successful PullFromPipedrive() or PushToPipedrive().
 */
function refreshShadowSheet() {
  const ss = SpreadsheetApp.getActive();
  const apiDealsSheet = ss.getSheetByName('API_Deals');
  
  if (!apiDealsSheet) {
    Logger.log('API_Deals sheet not found, cannot create shadow.');
    return;
  }
  
  let shadowSheet = ss.getSheetByName('API_Deals_Shadow');
  
  // Delete existing shadow sheet if it exists
  if (shadowSheet) {
    ss.deleteSheet(shadowSheet);
  }
  
  // Create fresh shadow as copy of API_Deals
  shadowSheet = apiDealsSheet.copyTo(ss);
  shadowSheet.setName('API_Deals_Shadow');
  
  // Hide the shadow sheet (it's for internal use only)
  shadowSheet.hideSheet();
  
  Logger.log('API_Deals_Shadow refreshed successfully.');
}

/**
 * Add menu items for Deals sheet management.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Pipedrive Sync')
    .addItem('Pull From Pipedrive', 'PullFromPipedrive')
    .addItem('Pull From Drive', 'refreshAllDriveFolders')
    .addItem('Refresh & Audit', 'refreshAndAudit')
    .addSeparator()
    .addItem('Detect Changes', 'detectAndShowChanges')
    .addItem('Push Changes To Pipedrive', 'PushToPipedrive')
    .addSeparator()
    .addSubMenu(ui.createMenu('Advanced')
      .addItem('Verify Folder Mappings', 'verifyFolderMappings')
      .addItem('Replenish Placeholders', 'replenishPlaceholders')
      .addSeparator()
      .addItem('Rebuild Formatting (SLOW)', 'refreshDealsView'))
    .addToUi();
}

function refreshAndAudit() {
  const ui = SpreadsheetApp.getUi();
  
  ui.alert('Refresh & Audit', 'Starting full refresh and validation...', ui.ButtonSet.OK);
  
  try {
    refreshAllDriveFolders();
    PullFromPipedrive();
    verifyFolderMappings();
  } catch (e) {
    ui.alert('Error', 'Refresh & Audit failed: ' + e.message, ui.ButtonSet.OK);
  }
}
