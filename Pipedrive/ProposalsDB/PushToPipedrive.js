// ---------------------------
// Push Changes Back to Pipedrive
// ---------------------------

/**
 * Pushes detected changes from Deals sheet back to Pipedrive.
 * Only pushes USER_EDITABLE fields. Validates and logs all operations.
 */
function PushToPipedrive() {
  const ui = SpreadsheetApp.getUi();
  
  // Detect changes first
  const changes = detectChanges();
  
  if (changes.length === 0) {
    ui.alert(
      'No Changes to Push',
      'No editable fields have been changed since the last pull from Pipedrive.',
      ui.ButtonSet.OK
    );
    return;
  }
  
  // Show confirmation dialog
  const response = ui.alert(
    'Push Changes to Pipedrive?',
    changes.length + ' change(s) detected in editable fields.\n\n' +
    'This will update Pipedrive deals via API. Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    Logger.log('Push cancelled by user.');
    return;
  }
  
  // Group changes by deal ID for efficient API calls
  const changesByDeal = groupChangesByDeal_(changes);
  
  Logger.log('Pushing changes for ' + Object.keys(changesByDeal).length + ' deals...');
  
  const results = {
    successful: [],
    failed: [],
    skipped: []
  };
  
  // Process each deal
  for (const dealId in changesByDeal) {
    const dealChanges = changesByDeal[dealId];
    
    try {
      // Build PUT payload for this deal
      const payload = buildPushPayload_(dealChanges);
      
      // Validate payload
      const validation = validatePushPayload_(payload);
      if (!validation.valid) {
        results.skipped.push({
          dealId: dealId,
          reason: validation.reason
        });
        Logger.log('Skipped deal ' + dealId + ': ' + validation.reason);
        continue;
      }
      
      // Push to Pipedrive
      const success = pushDealToPipedrive_(dealId, payload);
      
      if (success) {
        results.successful.push({
          dealId: dealId,
          fieldCount: dealChanges.length
        });
        Logger.log('Successfully pushed ' + dealChanges.length + ' changes for deal ' + dealId);
      } else {
        results.failed.push({
          dealId: dealId,
          reason: 'API call failed'
        });
      }
      
    } catch (e) {
      results.failed.push({
        dealId: dealId,
        reason: e.message
      });
      Logger.log('Error pushing deal ' + dealId + ': ' + e.message);
    }
  }
  
  // Log results
  if (typeof logEvent_ === 'function') {
    logEvent_('PUSH', 'PushToPipedrive', 
      'Success: ' + results.successful.length + 
      ', Failed: ' + results.failed.length + 
      ', Skipped: ' + results.skipped.length);
  }
  
  // Update Detected_Changes sheet with push results
  if (typeof updateChangeStatuses_ === 'function') {
    updateChangeStatuses_(results);
  }
  
  // Show summary to user
  showPushResults_(results);
  
  // If any successful pushes, refresh data from Pipedrive
  if (results.successful.length > 0) {
    ui.alert(
      'Refresh Required',
      'Changes pushed successfully. Running PullFromPipedrive() to sync data...',
      ui.ButtonSet.OK
    );
    
    // PullFromPipedrive() now uses fast data-only refresh automatically
    PullFromPipedrive();
  }
}

/**
 * Groups changes by deal ID for efficient API calls.
 * @private
 */
function groupChangesByDeal_(changes) {
  const grouped = {};
  
  changes.forEach(function(change) {
    if (!grouped[change.dealId]) {
      grouped[change.dealId] = [];
    }
    grouped[change.dealId].push(change);
  });
  
  return grouped;
}

/**
 * Builds PUT payload for Pipedrive API from change list.
 * @private
 */
function buildPushPayload_(dealChanges) {
  const payload = {};
  
  dealChanges.forEach(function(change) {
    const fieldKey = change.fieldKey;
    let value = change.newValue;
    
    // Type conversions based on field key
    if (fieldKey === 'value' || fieldKey === 'probability') {
      // Numeric fields
      value = (value === '' || value === null) ? null : Number(value);
    } else if (fieldKey === 'stage_id' || fieldKey === 'user_id' || fieldKey === 'person_id' || fieldKey === 'org_id') {
      // ID fields - must be numbers
      value = (value === '' || value === null) ? null : Number(value);
    } else if (fieldKey === 'visible_to') {
      // Visibility codes: convert text back to number
      if (value === 'Item owner') value = 1;
      else if (value === 'All users') value = 3;
      else if (value === 'Owner only') value = 7;
      else value = Number(value); // fallback
    }
    // String fields (title, currency, lost_reason, custom fields) stay as-is
    
    payload[fieldKey] = value;
  });
  
  return payload;
}

/**
 * Validates push payload before sending to API.
 * @private
 */
function validatePushPayload_(payload) {
  // Check if payload is empty
  if (Object.keys(payload).length === 0) {
    return { valid: false, reason: 'No fields to update' };
  }
  
  // Validate Folder URL if present
  const folderUrlKey = CONFIG.PIPEDRIVE.FIELD_KEYS.FOLDER_URL;
  if (folderUrlKey && payload[folderUrlKey]) {
    const validation = validateFolderUrl_(payload[folderUrlKey]);
    if (!validation.valid) {
      return { 
        valid: false, 
        reason: 'Invalid Folder URL: ' + validation.error 
      };
    }
  }
  
  // Additional validations can be added here
  // e.g., required fields, value ranges, etc.
  
  return { valid: true };
}

/**
 * Pushes a single deal update to Pipedrive via PUT API.
 * @private
 */
function pushDealToPipedrive_(dealId, payload) {
  const token = getPipedriveToken_();
  const url = CONFIG.PIPEDRIVE.BASE_URL + '/deals/' + dealId + '?api_token=' + encodeURIComponent(token);
  
  const options = {
    method: 'put',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    
    if (code >= 200 && code < 300) {
      return true;
    } else {
      Logger.log('PUT /deals/' + dealId + ' failed: ' + code + ' - ' + response.getContentText());
      return false;
    }
  } catch (e) {
    Logger.log('Exception pushing deal ' + dealId + ': ' + e.message);
    return false;
  }
}

/**
 * Shows push results to user in a dialog.
 * @private
 */
function showPushResults_(results) {
  const ui = SpreadsheetApp.getUi();
  
  let message = 'Push Results:\n\n';
  message += '✓ Successful: ' + results.successful.length + ' deals\n';
  message += '✗ Failed: ' + results.failed.length + ' deals\n';
  message += '⊘ Skipped: ' + results.skipped.length + ' deals\n';
  
  if (results.failed.length > 0) {
    message += '\nFailed Deals:\n';
    results.failed.forEach(function(item) {
      message += '- Deal ' + item.dealId + ': ' + item.reason + '\n';
    });
  }
  
  if (results.skipped.length > 0) {
    message += '\nSkipped Deals:\n';
    results.skipped.forEach(function(item) {
      message += '- Deal ' + item.dealId + ': ' + item.reason + '\n';
    });
  }
  
  ui.alert('Push Complete', message, ui.ButtonSet.OK);
}
