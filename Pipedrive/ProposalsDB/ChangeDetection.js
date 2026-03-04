// ---------------------------
// Change Detection - Compare Deals vs API_Deals_Shadow
// ---------------------------

/**
 * Detects changes made by users in the Deals sheet compared to the last API pull.
 * Only detects changes in USER_EDITABLE fields (white columns).
 * @returns {Array} Array of change objects: {dealId, rowIndex, fieldKey, fieldName, oldValue, newValue, category}
 */
function detectChanges() {
  const ss = SpreadsheetApp.getActive();
  const dealsSheet = ss.getSheetByName('Deals');
  const shadowSheet = ss.getSheetByName('API_Deals_Shadow');
  
  if (!dealsSheet) {
    throw new Error('Deals sheet not found. Run PullFromPipedrive() first.');
  }
  if (!shadowSheet) {
    throw new Error('API_Deals_Shadow sheet not found. Run PullFromPipedrive() first to create shadow.');
  }
  
  // Get data from both sheets
  const dealsData = dealsSheet.getDataRange().getValues();
  const shadowData = shadowSheet.getDataRange().getValues();
  
  if (dealsData.length < 2 || shadowData.length < 2) {
    Logger.log('No data rows to compare.');
    return [];
  }
  
  const headers = dealsData[0];
  const shadowHeaders = shadowData[0];
  
  // Verify headers match
  if (headers.length !== shadowHeaders.length) {
    throw new Error('Header mismatch between Deals and API_Deals_Shadow. Run PullFromPipedrive() to sync.');
  }
  
  // Get field classification map
  const classMap = getFieldClassificationMap();
  
  // Build index of Deal ID column
  const dealIdColIndex = headers.indexOf('Deal - ID');
  if (dealIdColIndex === -1) {
    throw new Error('Deal - ID column not found in Deals sheet.');
  }
  
  // Build shadow map: dealId -> row data
  const shadowMap = {};
  for (let r = 1; r < shadowData.length; r++) {
    const dealId = shadowData[r][dealIdColIndex];
    if (dealId) {
      shadowMap[dealId] = shadowData[r];
    }
  }
  
  // Detect changes
  const changes = [];
  
  for (let r = 1; r < dealsData.length; r++) {
    const currentRow = dealsData[r];
    const dealId = currentRow[dealIdColIndex];
    
    if (!dealId) continue; // Skip empty rows
    
    const shadowRow = shadowMap[dealId];
    if (!shadowRow) {
      Logger.log('Deal ID ' + dealId + ' not found in shadow (new deal?)');
      continue;
    }
    
    // Compare each column
    for (let c = 0; c < headers.length; c++) {
      const header = headers[c];
      const fieldInfo = classMap[header];
      
      if (!fieldInfo) continue; // Skip unknown fields
      
      // Only detect changes in USER_EDITABLE fields
      if (fieldInfo.category !== 'USER_EDITABLE') continue;
      
      const currentValue = normalizeValue_(currentRow[c]);
      const shadowValue = normalizeValue_(shadowRow[c]);
      
      // Detect change
      if (currentValue !== shadowValue) {
        changes.push({
          dealId: dealId,
          rowIndex: r + 1, // 1-based for sheet reference
          colIndex: c + 1, // 1-based for sheet reference
          fieldKey: fieldInfo.key,
          fieldName: header,
          oldValue: shadowValue,
          newValue: currentValue,
          category: fieldInfo.category,
          notes: fieldInfo.notes || ''
        });
      }
    }
  }
  
  Logger.log('Detected ' + changes.length + ' changes in USER_EDITABLE fields.');
  return changes;
}

/**
 * Normalize cell values for comparison (handle empty strings, nulls, etc.)
 * @private
 */
function normalizeValue_(val) {
  if (val === null || val === undefined || val === '') return '';
  if (typeof val === 'number') return String(val);
  if (typeof val === 'boolean') return String(val);
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  return String(val).trim();
}

/**
 * Detects changes and displays them to the user in a dialog.
 * Called from menu: "Detect Changes"
 */
function detectAndShowChanges() {
  const changes = detectChanges();
  
  if (changes.length === 0) {
    SpreadsheetApp.getUi().alert(
      'No Changes Detected',
      'No editable fields have been changed since the last pull from Pipedrive.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  // Write changes to a temporary sheet for review
  writeChangesToSheet_(changes);
  
  SpreadsheetApp.getUi().alert(
    'Changes Detected',
    changes.length + ' change(s) detected in editable fields.\n\n' +
    'A "Detected_Changes" sheet has been created showing all changes.\n' +
    'Review the changes before pushing to Pipedrive.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Writes detected changes to a review sheet.
 * @private
 */
function writeChangesToSheet_(changes) {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName('Detected_Changes');
  
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet('Detected_Changes');
  }
  
  // Header row
  const headers = [
    'Deal ID',
    'Row',
    'Field Name',
    'Old Value',
    'New Value',
    'Notes',
    'Status'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground('#4285F4')
    .setFontColor('#FFFFFF');
  
  // Data rows
  const rows = changes.map(function(change) {
    return [
      change.dealId,
      change.rowIndex,
      change.fieldName,
      change.oldValue || '(empty)',
      change.newValue || '(empty)',
      change.notes,
      'Pending'
    ];
  });
  
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  
  // Format
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  
  // Highlight changes
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    // Alternate row colors for readability
    for (let r = 2; r <= lastRow; r++) {
      const bgColor = (r % 2 === 0) ? '#F8F9FA' : '#FFFFFF';
      sheet.getRange(r, 1, 1, headers.length).setBackground(bgColor);
    }
  }
  
  // Add timestamp
  sheet.getRange(lastRow + 2, 1).setValue('Generated: ' + new Date().toLocaleString())
    .setFontStyle('italic')
    .setFontColor('#666666');
  
  // Activate the sheet to show it to the user
  sheet.activate();
  
  Logger.log('Changes written to Detected_Changes sheet.');
}

/**
 * Updates the Status column in Detected_Changes sheet after push.
 * @param {Object} results - Push results with successful/failed/skipped arrays
 */
function updateChangeStatuses_(results) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Detected_Changes');
  
  if (!sheet) {
    Logger.log('Detected_Changes sheet not found, skipping status update.');
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return; // No data rows
  
  const headers = data[0];
  const dealIdCol = headers.indexOf('Deal ID');
  const statusCol = headers.indexOf('Status');
  
  if (dealIdCol === -1 || statusCol === -1) {
    Logger.log('Required columns not found in Detected_Changes sheet.');
    return;
  }
  
  // Build lookup maps
  const successfulDeals = {};
  const failedDeals = {};
  
  results.successful.forEach(function(item) {
    successfulDeals[item.dealId] = true;
  });
  
  results.failed.forEach(function(item) {
    failedDeals[item.dealId] = item.reason;
  });
  
  results.skipped.forEach(function(item) {
    failedDeals[item.dealId] = item.reason;
  });
  
  // Update status column for each row
  for (let r = 1; r < data.length; r++) {
    const dealId = data[r][dealIdCol];
    const currentStatus = data[r][statusCol];
    
    // Only update if still Pending
    if (currentStatus === 'Pending') {
      let newStatus = 'Pending';
      let bgColor = '#FFFFFF';
      
      if (successfulDeals[dealId]) {
        newStatus = 'Applied';
        bgColor = '#D4EDDA'; // Light green
      } else if (failedDeals[dealId]) {
        newStatus = 'Failed: ' + failedDeals[dealId];
        bgColor = '#F8D7DA'; // Light red
      }
      
      // Update the cell
      const cell = sheet.getRange(r + 1, statusCol + 1);
      cell.setValue(newStatus);
      cell.setBackground(bgColor);
    }
  }
  
  Logger.log('Updated change statuses in Detected_Changes sheet.');
}
