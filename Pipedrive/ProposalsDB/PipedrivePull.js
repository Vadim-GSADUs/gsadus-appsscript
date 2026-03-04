// ---------------------------
// Pipedrive → API_Deals sync
// ---------------------------

/**
 * Manual entry point: refresh API_Deals sheet from Pipedrive API.
 * Builds headers dynamically from deal field metadata and writes all deals.
 */
function PullFromPipedrive() {
  const ss = SpreadsheetApp.getActive();
  
  // Delete stale Detected_Changes sheet to prevent false change detection
  const detectedChangesSheet = ss.getSheetByName('Detected_Changes');
  if (detectedChangesSheet) {
    ss.deleteSheet(detectedChangesSheet);
    Logger.log('Deleted stale Detected_Changes sheet.');
  }
  
  const sheetName = 'API_Deals';
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  // Fetch field metadata, pipeline/stage metadata, and deals
  const fieldMeta = fetchPipedriveDealFields_();
  const pipelineMap = fetchPipelineMap_();
  const stageMap = fetchStageMap_();
  const deals = fetchAllDealsFromPipedrive_();

  if (!deals.length) {
    sheet.clearContents();
    Logger.log('PullFromPipedrive: no deals returned from Pipedrive.');
    return;
  }

  // Build column order from field metadata so we get human-friendly names
  // similar to Pipedrive CSV exports.
  const headers = buildDealHeadersFromMeta_(fieldMeta, pipelineMap, stageMap);

  // Clear existing content but keep sheet formatting/layout
  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();
  sheet.getRange(1, 1, maxRows, maxCols).clearContent();

  // Write header row
  sheet.getRange(1, 1, 1, headers.length).setValues([headers.map(h => h.name)]);

  // Build data rows using extractor functions
  const rows = deals.map(function (deal) {
    return headers.map(function (h) {
      try {
        return h.extractor(deal);
      } catch (e) {
        Logger.log('Error extracting ' + h.key + ' from deal ' + deal.id + ': ' + e);
        return '';
      }
    });
  });

  // Write all deal rows starting at row 2
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  Logger.log('PullFromPipedrive complete. Deals written: ' + rows.length + ', columns: ' + headers.length);
  if (typeof logEvent_ === 'function') {
    logEvent_('PULL', 'PullFromPipedrive', 'Rows: ' + rows.length + ', Cols: ' + headers.length);
  }
  
  // Auto-refresh Deals and shadow sheet after successful pull
  // Use fast data-only refresh to avoid 40-50 second formatting delays
  if (typeof refreshDealsViewDataOnly === 'function') {
    Logger.log('Auto-refreshing Deals (data only)...');
    refreshDealsViewDataOnly();
  }
  if (typeof refreshShadowSheet === 'function') {
    Logger.log('Auto-refreshing API_Deals_Shadow...');
    refreshShadowSheet();
  }
}

/**
 * Fetch deal field metadata from Pipedrive.
 * Returns array of field objects including internal key and human-readable name.
 */
function fetchPipedriveDealFields_() {
  const token = getPipedriveToken_();
  const url = CONFIG.PIPEDRIVE.BASE_URL + '/dealFields?api_token=' + encodeURIComponent(token);

  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const code = resp.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Pipedrive GET /dealFields failed: ' + code + ' → ' + resp.getContentText());
  }

  const json = JSON.parse(resp.getContentText());
  if (!json.data || !json.data.length) {
    return [];
  }

  return json.data;
}

/**
 * Fetch pipeline metadata and return a map of pipeline_id → pipeline_name.
 */
function fetchPipelineMap_() {
  const token = getPipedriveToken_();
  const url = CONFIG.PIPEDRIVE.BASE_URL + '/pipelines?api_token=' + encodeURIComponent(token);

  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const code = resp.getResponseCode();
  if (code < 200 || code >= 300) {
    Logger.log('Warning: Pipedrive GET /pipelines failed: ' + code);
    return {};
  }

  const json = JSON.parse(resp.getContentText());
  if (!json.success || !json.data) {
    return {};
  }

  const map = {};
  json.data.forEach(function(pipeline) {
    map[pipeline.id] = pipeline.name;
  });
  return map;
}

/**
 * Fetch stage metadata and return a map of stage_id → stage_name.
 */
function fetchStageMap_() {
  const token = getPipedriveToken_();
  const url = CONFIG.PIPEDRIVE.BASE_URL + '/stages?api_token=' + encodeURIComponent(token);

  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const code = resp.getResponseCode();
  if (code < 200 || code >= 300) {
    Logger.log('Warning: Pipedrive GET /stages failed: ' + code);
    return {};
  }

  const json = JSON.parse(resp.getContentText());
  if (!json.success || !json.data) {
    return {};
  }

  const map = {};
  json.data.forEach(function(stage) {
    // Trim stage names to remove trailing spaces
    map[stage.id] = (stage.name || '').trim();
  });
  return map;
}

/**
 * Build headers matching the manual Pipedrive CSV export structure.
 * Returns an array of { key, name, extractor } objects where:
 * - key: internal reference (may not be a direct deal property)
 * - name: the header label for Sheets
 * - extractor: function(deal) that returns the value for this column
 */
function buildDealHeadersFromMeta_(fieldMeta, pipelineMap, stageMap) {
  // Define extraction helpers
  const extractName = function(obj) {
    return (obj && obj.name) ? obj.name : '';
  };
  
  const extractValue = function(obj) {
    return (obj && obj.value !== undefined && obj.value !== null) ? obj.value : '';
  };
  
  const formatDateTime = function(ts) {
    // Pipedrive returns dates already formatted as 'YYYY-MM-DD HH:MM:SS'
    if (!ts || ts === null) return '';
    return String(ts);
  };
  
  const formatDate = function(ts) {
    // Pipedrive returns dates already formatted
    if (!ts || ts === null) return '';
    const str = String(ts);
    // If it's a full timestamp, extract just the date part
    if (str.indexOf(' ') !== -1) {
      return str.split(' ')[0];
    }
    return str;
  };
  
  const capitalizeStatus = function(status) {
    if (!status) return '';
    return status.charAt(0).toUpperCase() + status.slice(1);
  };
  
  const mapVisibility = function(visCode) {
    if (visCode === null || visCode === undefined || visCode === '') return '';
    const code = String(visCode);
    if (code === '1') return 'Item owner';
    if (code === '3') return 'All users';
    if (code === '7') return 'Owner only';
    return code;
  };
  
  const mapArchiveStatus = function(isArchived) {
    if (isArchived === true) return 'Archived';
    if (isArchived === false || isArchived === undefined) return 'Not archived';
    return '';
  };
  
  const mapSourceOrigin = function(origin) {
    if (!origin) return '';
    if (origin === 'ManuallyCreated') return 'Manually created';
    return origin;
  };
  
  // Define the exact column order and extraction logic per the mapping spec
  const headers = [
    { key: 'id', name: 'Deal - ID', extractor: function(d) { return d.id || ''; } },
    { key: 'title', name: 'Deal - Title', extractor: function(d) { return d.title || ''; } },
    { key: 'creator', name: 'Deal - Creator', extractor: function(d) { return extractName(d.creator_user_id); } },
    { key: 'owner', name: 'Deal - Owner', extractor: function(d) { return extractName(d.user_id); } },
    { key: 'value', name: 'Deal - Value', extractor: function(d) { return d.value !== undefined && d.value !== null ? d.value : ''; } },
    { key: 'currency', name: 'Deal - Currency of Value', extractor: function(d) { return d.currency || ''; } },
    { key: 'weighted_value', name: 'Deal - Weighted value', extractor: function(d) { 
      const val = d.weighted_value;
      if (val === undefined || val === null || val === '') return '';
      // Convert to string to prevent Sheets from interpreting as date serial
      return String(val);
    } },
    { key: 'weighted_value_currency', name: 'Deal - Currency of Weighted value', extractor: function(d) { return d.weighted_value_currency || ''; } },
    { key: 'probability', name: 'Deal - Probability', extractor: function(d) { return d.probability !== undefined && d.probability !== null ? d.probability : ''; } },
    { key: 'organization', name: 'Deal - Organization', extractor: function(d) { return extractName(d.org_id); } },
    { key: 'organization_id', name: 'Deal - Organization ID', extractor: function(d) { return extractValue(d.org_id); } },
    { key: 'pipeline', name: 'Deal - Pipeline', extractor: function(d) { 
      const pipelineId = d.pipeline_id;
      if (pipelineId === undefined || pipelineId === null) return '';
      return pipelineMap[pipelineId] || String(pipelineId);
    } },
    { key: 'person', name: 'Deal - Contact person', extractor: function(d) { return extractName(d.person_id); } },
    { key: 'person_id', name: 'Deal - Contact person ID', extractor: function(d) { return extractValue(d.person_id); } },
    { key: 'stage', name: 'Deal - Stage', extractor: function(d) { 
      const stageId = d.stage_id;
      if (stageId === undefined || stageId === null) return '';
      return stageMap[stageId] || String(stageId);
    } },
    { key: 'label', name: 'Deal - Label', extractor: function(d) { return d.label || ''; } },
    { key: 'status', name: 'Deal - Status', extractor: function(d) { return capitalizeStatus(d.status); } },
    { key: 'add_time', name: 'Deal - Deal created', extractor: function(d) { return formatDateTime(d.add_time); } },
    { key: 'update_time', name: 'Deal - Update time', extractor: function(d) { return formatDateTime(d.update_time); } },
    { key: 'stage_change_time', name: 'Deal - Last stage change', extractor: function(d) { return formatDateTime(d.stage_change_time); } },
    { key: 'next_activity_date', name: 'Deal - Next activity date', extractor: function(d) { return formatDate(d.next_activity_date); } },
    { key: 'last_activity_date', name: 'Deal - Last activity date', extractor: function(d) { return formatDate(d.last_activity_date); } },
    { key: 'won_time', name: 'Deal - Won time', extractor: function(d) { return formatDateTime(d.won_time); } },
    { key: 'last_incoming_mail_time', name: 'Deal - Last email received', extractor: function(d) { return formatDateTime(d.last_incoming_mail_time); } },
    { key: 'last_outgoing_mail_time', name: 'Deal - Last email sent', extractor: function(d) { return formatDateTime(d.last_outgoing_mail_time); } },
    { key: 'lost_time', name: 'Deal - Lost time', extractor: function(d) { return formatDateTime(d.lost_time); } },
    { key: 'close_time', name: 'Deal - Deal closed on', extractor: function(d) { return formatDateTime(d.close_time); } },
    { key: 'lost_reason', name: 'Deal - Lost reason', extractor: function(d) { return d.lost_reason || ''; } },
    { key: 'visible_to', name: 'Deal - Visible to', extractor: function(d) { return mapVisibility(d.visible_to); } },
    { key: 'activities_count', name: 'Deal - Total activities', extractor: function(d) { return d.activities_count !== undefined && d.activities_count !== null ? d.activities_count : ''; } },
    { key: 'done_activities_count', name: 'Deal - Done activities', extractor: function(d) { return d.done_activities_count !== undefined && d.done_activities_count !== null ? d.done_activities_count : ''; } },
    { key: 'undone_activities_count', name: 'Deal - Activities to do', extractor: function(d) { return d.undone_activities_count !== undefined && d.undone_activities_count !== null ? d.undone_activities_count : ''; } },
    { key: 'email_messages_count', name: 'Deal - Email messages count', extractor: function(d) { return d.email_messages_count !== undefined && d.email_messages_count !== null ? d.email_messages_count : ''; } },
    { key: 'expected_close_date', name: 'Deal - Expected close date', extractor: function(d) { return formatDate(d.expected_close_date); } },
    { key: 'products_count', name: 'Deal - Product quantity', extractor: function(d) { return d.products_count !== undefined && d.products_count !== null ? d.products_count : ''; } },
    { key: 'product_amount', name: 'Deal - Product amount', extractor: function(d) { return ''; } }, // TBD if needed
    { key: 'product_name', name: 'Deal - Product name', extractor: function(d) { return ''; } }, // TBD if needed
    { key: 'source_origin', name: 'Deal - Source origin', extractor: function(d) { return mapSourceOrigin(d.origin); } },
    { key: 'source_origin_id', name: 'Deal - Source origin ID', extractor: function(d) { return d.origin_id || ''; } },
    { key: 'source_channel', name: 'Deal - Source channel', extractor: function(d) { return d.channel || ''; } },
    { key: 'source_channel_id', name: 'Deal - Source channel ID', extractor: function(d) { return d.channel_id || ''; } },
    { key: 'archived', name: 'Deal - Archive status', extractor: function(d) { return mapArchiveStatus(d.is_archived); } },
    { key: 'archive_time', name: 'Deal - Archive time', extractor: function(d) { return formatDateTime(d.archive_time); } },
    { key: 'sequence_enrollment', name: 'Deal - Sequence enrollment', extractor: function(d) { return ''; } }, // Not in API
    // Address fields (custom field keys from CONFIG with structured subfields)
    { key: 'address', name: 'Deal - Address', extractor: function(d) { return d[CONFIG.PIPEDRIVE.FIELD_KEYS.ADDRESS] || ''; } },
    // NOTE: Lat/Long are NOT returned by GET /deals API, even though they're editable fields
    // Pipedrive auto-geocodes addresses in UI but doesn't expose coordinates via API
    // We can WRITE to these fields via PUT /deals/{id}, but can't READ them
    // Use geocodeDeals() to generate coordinates, which can then be pushed to Pipedrive
    { key: 'address_apt', name: 'Deal - Apartment/suite no of Address', extractor: function(d) { return d[CONFIG.PIPEDRIVE.FIELD_KEYS.ADDRESS + '_subpremise'] || ''; } },
    { key: 'address_house_number', name: 'Deal - House number of Address', extractor: function(d) { return d[CONFIG.PIPEDRIVE.FIELD_KEYS.ADDRESS + '_street_number'] || ''; } },
    { key: 'address_street', name: 'Deal - Street/road name of Address', extractor: function(d) { return d[CONFIG.PIPEDRIVE.FIELD_KEYS.ADDRESS + '_route'] || ''; } },
    { key: 'address_district', name: 'Deal - District/sublocality of Address', extractor: function(d) { return d[CONFIG.PIPEDRIVE.FIELD_KEYS.ADDRESS + '_sublocality'] || ''; } },
    { key: 'address_city', name: 'Deal - City/town/village/locality of Address', extractor: function(d) { return d[CONFIG.PIPEDRIVE.FIELD_KEYS.ADDRESS + '_locality'] || ''; } },
    { key: 'address_state', name: 'Deal - State/county of Address', extractor: function(d) { return d[CONFIG.PIPEDRIVE.FIELD_KEYS.ADDRESS + '_admin_area_level_1'] || ''; } },
    { key: 'address_region', name: 'Deal - Region of Address', extractor: function(d) { return d[CONFIG.PIPEDRIVE.FIELD_KEYS.ADDRESS + '_admin_area_level_2'] || ''; } },
    { key: 'address_country', name: 'Deal - Country of Address', extractor: function(d) { return d[CONFIG.PIPEDRIVE.FIELD_KEYS.ADDRESS + '_country'] || ''; } },
    { key: 'address_zip', name: 'Deal - ZIP/Postal code of Address', extractor: function(d) { return d[CONFIG.PIPEDRIVE.FIELD_KEYS.ADDRESS + '_postal_code'] || ''; } },
    { key: 'address_full', name: 'Deal - Full/combined address of Address', extractor: function(d) { return d[CONFIG.PIPEDRIVE.FIELD_KEYS.ADDRESS + '_formatted_address'] || ''; } },
    { key: 'proposal', name: 'Deal - Proposal #', extractor: function(d) { return d[CONFIG.PIPEDRIVE.FIELD_KEYS.PROPOSAL] || ''; } },
    { key: 'folder_url', name: 'Deal - Folder URL', extractor: function(d) { return d[CONFIG.PIPEDRIVE.FIELD_KEYS.FOLDER_URL] || ''; } }
  ];
  
  return headers;
}

/**
 * Fetch all deals from Pipedrive using pagination.
 * For now, no filters; pulls the full collection.
 */
function fetchAllDealsFromPipedrive_() {
  const token   = getPipedriveToken_(); // from PipedriveWebhookHelpers.js
  const baseUrl = CONFIG.PIPEDRIVE.BASE_URL + '/deals';
  const limit   = 500; // good for ~300 deals in one or two pages

  let start = 0;
  let more  = true;
  const all = [];

  while (more) {
    const url = baseUrl +
      '?start=' + start +
      '&limit=' + limit +
      '&api_token=' + encodeURIComponent(token);

    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const code = resp.getResponseCode();
    if (code < 200 || code >= 300) {
      throw new Error('Pipedrive GET /deals failed: ' + code + ' → ' + resp.getContentText());
    }

    const json = JSON.parse(resp.getContentText());
    if (!json.data || !json.data.length) {
      break;
    }

    all.push.apply(all, json.data);

    const pg = json.additional_data && json.additional_data.pagination;
    if (pg && pg.more_items_in_collection) {
      more  = true;
      start = pg.next_start != null ? pg.next_start : (start + limit);
    } else {
      more = false;
    }
  }

  return all;
}

/**
 * Comprehensive validation of Deal Folder URLs and Proposal integrity.
 * Checks for: invalid URLs, trashed folders, duplicate PP#, folder name mismatches.
 * Returns object with validation results and issues found.
 */
function validateDealFolderUrls_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dealsSheet = ss.getSheetByName(CONFIG.SHEET_DEALS);
  
  if (!dealsSheet) {
    Logger.log('validateDealFolderUrls_: Deals sheet not found');
    return { invalidCount: 0, issues: [] };
  }
  
  const data = dealsSheet.getDataRange().getValues();
  if (data.length <= 1) {
    Logger.log('validateDealFolderUrls_: No data rows in Deals sheet');
    return { invalidCount: 0, issues: [] };
  }
  
  const headers = data[0];
  const dealIdIdx = headers.indexOf('Deal - ID');
  const folderUrlIdx = headers.indexOf('Deal - Folder URL');
  const proposalIdx = headers.indexOf('Deal - Proposal #');
  
  if (dealIdIdx === -1 || folderUrlIdx === -1) {
    Logger.log('validateDealFolderUrls_: Required columns not found');
    return { invalidCount: 0, issues: [] };
  }
  
  const issues = [];
  const proposalMap = {}; // Track PP# usage for duplicate detection
  
  // First pass: Track all PP# assignments (regardless of Folder URL)
  for (let i = 1; i < data.length; i++) {
    const dealId = data[i][dealIdIdx];
    const proposalNum = proposalIdx !== -1 ? data[i][proposalIdx] : null;
    const folderUrl = data[i][folderUrlIdx];
    
    if (proposalNum && String(proposalNum).trim() !== '') {
      if (!proposalMap[proposalNum]) {
        proposalMap[proposalNum] = [];
      }
      proposalMap[proposalNum].push({ dealId: dealId, folderUrl: folderUrl, rowIndex: i });
    }
  }
  
  // Second pass: Validate Folder URLs
  for (let i = 1; i < data.length; i++) {
    const dealId = data[i][dealIdIdx];
    const folderUrl = data[i][folderUrlIdx];
    const proposalNum = proposalIdx !== -1 ? data[i][proposalIdx] : null;
    
    if (!folderUrl || String(folderUrl).trim() === '') continue; // Skip deals without folder URL
    
    // Basic URL validation
    const validation = validateFolderUrl_(folderUrl);
    if (!validation.valid) {
      issues.push({
        dealId: dealId,
        proposalNum: proposalNum,
        type: 'INVALID_URL',
        error: validation.error,
        folderUrl: folderUrl,
        rowIndex: i
      });
      
      logEvent_(
        'VALIDATION_ERROR', 
        'Deal ' + dealId + ' has invalid Folder URL',
        'Proposal#: ' + proposalNum + ', Error: ' + validation.error + ', URL: ' + folderUrl
      );
      continue;
    }
    
    // Enhanced validation: Check if folder is trashed
    try {
      const folderId = getFolderIdFromUrl_(folderUrl);
      const folder = DriveApp.getFolderById(folderId);
      
      if (folder.isTrashed()) {
        issues.push({
          dealId: dealId,
          proposalNum: proposalNum,
          type: 'TRASHED_FOLDER',
          error: 'Folder is in trash',
          folderUrl: folderUrl,
          rowIndex: i
        });
        
        logEvent_(
          'VALIDATION_ERROR',
          'Deal ' + dealId + ' folder is trashed',
          'Proposal#: ' + proposalNum + ', URL: ' + folderUrl
        );
        continue;
      }
      
      // Folder name validation: PP# in folder name should match Pipedrive field
      if (proposalNum) {
        const folderName = folder.getName();
        const folderPP = extractProposalNumber_(folderName);
        
        if (folderPP !== proposalNum) {
          issues.push({
            dealId: dealId,
            proposalNum: proposalNum,
            type: 'PP_MISMATCH',
            error: 'Folder name "' + folderName + '" has PP# "' + folderPP + '" but deal has "' + proposalNum + '"',
            folderUrl: folderUrl,
            rowIndex: i
          });
          
          logEvent_(
            'VALIDATION_WARN',
            'Deal ' + dealId + ' PP# mismatch',
            'Pipedrive: ' + proposalNum + ', Folder name: ' + folderPP
          );
        }
      }
      
    } catch (e) {
      issues.push({
        dealId: dealId,
        proposalNum: proposalNum,
        type: 'FOLDER_ACCESS_ERROR',
        error: e.message,
        folderUrl: folderUrl,
        rowIndex: i
      });
      
      logEvent_(
        'VALIDATION_ERROR',
        'Deal ' + dealId + ' folder access error',
        'Proposal#: ' + proposalNum + ', Error: ' + e.message
      );
    }
  }
  
  // Check for duplicate PP# assignments
  for (const ppNum in proposalMap) {
    const deals = proposalMap[ppNum];
    if (deals.length > 1) {
      deals.forEach(function(deal) {
        issues.push({
          dealId: deal.dealId,
          proposalNum: ppNum,
          type: 'DUPLICATE_PP',
          error: 'PP# ' + ppNum + ' assigned to ' + deals.length + ' deals',
          folderUrl: deal.folderUrl,
          rowIndex: deal.rowIndex,
          duplicateWith: deals.map(function(d) { return d.dealId; }).filter(function(id) { return id !== deal.dealId; })
        });
      });
      
      logEvent_(
        'VALIDATION_ERROR',
        'Duplicate PP# detected: ' + ppNum,
        'Assigned to deals: ' + deals.map(function(d) { return d.dealId; }).join(', ')
      );
    }
  }
  
  if (issues.length > 0) {
    Logger.log('⚠️ ' + issues.length + ' issue(s) found - check Logs sheet');
  } else {
    Logger.log('✓ All validations passed successfully');
  }
  
  return { invalidCount: issues.length, issues: issues };
}

/**
 * Menu-accessible function to manually trigger folder validation with auto-fix option.
 */
function verifyFolderMappings() {
  const ui = SpreadsheetApp.getUi();
  
  ui.alert('Verifying Folder Mappings', 'Checking all Deal Folder URLs for integrity issues...', ui.ButtonSet.OK);
  
  const result = validateDealFolderUrls_();
  
  if (result.invalidCount === 0) {
    ui.alert(
      'Verification Complete',
      '✓ All Folder URLs are valid!\n\n' +
      'No issues found.',
      ui.ButtonSet.OK
    );
    return;
  }
  
  // Categorize issues
  const trashedCount = result.issues.filter(function(i) { return i.type === 'TRASHED_FOLDER'; }).length;
  const invalidCount = result.issues.filter(function(i) { return i.type === 'INVALID_URL' || i.type === 'FOLDER_ACCESS_ERROR'; }).length;
  const duplicateCount = result.issues.filter(function(i) { return i.type === 'DUPLICATE_PP'; }).length / 2; // Each duplicate counted twice
  const mismatchCount = result.issues.filter(function(i) { return i.type === 'PP_MISMATCH'; }).length;
  
  const summary = 
    '⚠️ Found ' + result.invalidCount + ' issue(s):\n\n' +
    (trashedCount > 0 ? '• ' + trashedCount + ' trashed folder(s)\n' : '') +
    (invalidCount > 0 ? '• ' + invalidCount + ' invalid/inaccessible URL(s)\n' : '') +
    (duplicateCount > 0 ? '• ' + duplicateCount + ' duplicate PP# assignment(s)\n' : '') +
    (mismatchCount > 0 ? '• ' + mismatchCount + ' PP# mismatch(es)\n' : '') +
    '\nCheck Logs sheet for details.\n\n' +
    'Clear invalid Proposal# and Folder URL from affected deals?';
  
  const response = ui.alert(
    'Verification Complete',
    summary,
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    autoFixInvalidMappings_(result.issues);
  }
}

/**
 * Auto-fix: Clear Proposal# and Folder URL from deals with invalid/trashed folders.
 * Only fixes INVALID_URL, TRASHED_FOLDER, and FOLDER_ACCESS_ERROR types.
 * Does NOT auto-fix duplicates or mismatches (requires manual review).
 */
function autoFixInvalidMappings_(issues) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dealsSheet = ss.getSheetByName(CONFIG.SHEET_DEALS);
  
  if (!dealsSheet) return;
  
  const headers = dealsSheet.getDataRange().getValues()[0];
  const folderUrlIdx = headers.indexOf('Deal - Folder URL');
  const proposalIdx = headers.indexOf('Deal - Proposal #');
  
  if (folderUrlIdx === -1 || proposalIdx === -1) {
    ui.alert('Error', 'Could not find required columns', ui.ButtonSet.OK);
    return;
  }
  
  let fixedCount = 0;
  const fixableTypes = ['INVALID_URL', 'TRASHED_FOLDER', 'FOLDER_ACCESS_ERROR'];
  
  issues.forEach(function(issue) {
    if (fixableTypes.indexOf(issue.type) !== -1) {
      const row = issue.rowIndex + 1; // Convert to 1-based
      
      // Clear Proposal #
      if (proposalIdx !== -1) {
        dealsSheet.getRange(row, proposalIdx + 1).clearContent();
      }
      
      // Clear Folder URL
      if (folderUrlIdx !== -1) {
        dealsSheet.getRange(row, folderUrlIdx + 1).clearContent();
      }
      
      fixedCount++;
      
      logEvent_(
        'AUTO_FIX',
        'Cleared invalid mapping for deal ' + issue.dealId,
        'Type: ' + issue.type + ', Proposal#: ' + issue.proposalNum
      );
    }
  });
  
  ui.alert(
    'Auto-Fix Complete',
    'Cleared ' + fixedCount + ' invalid mapping(s).\n\n' +
    'You can now:\n' +
    '1. Pull from Pipedrive to sync changes\n' +
    '2. Or run "Check for Changes" and push updates',
    ui.ButtonSet.OK
  );
  
  Logger.log('autoFixInvalidMappings_: Fixed ' + fixedCount + ' issues');
}

// onOpen() menu moved to DealsView.js to avoid duplicate function definitions
