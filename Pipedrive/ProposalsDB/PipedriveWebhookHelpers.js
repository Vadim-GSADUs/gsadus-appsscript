// ---------------------------
// Pipedrive Webhook Helpers
// ---------------------------

// Get Pipedrive API Token from Script Properties
function getPipedriveToken_() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('PIPEDRIVE_API_TOKEN');
  if (!token) throw new Error('Missing script property PIPEDRIVE_API_TOKEN');
  return token;
}

// GET deal data from Pipedrive
function fetchDealFromPipedrive_(dealId) {
  const token = getPipedriveToken_();
  const url = CONFIG.PIPEDRIVE.BASE_URL + '/deals/' + encodeURIComponent(dealId) +
              '?api_token=' + encodeURIComponent(token);

  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const code = resp.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Pipedrive GET /deals failed: ' + code + ' → ' + resp.getContentText());
  }

  const json = JSON.parse(resp.getContentText());
  return json.data;
}

// PUT update to a deal
function updateDealFields_(dealId, body) {
  const token = getPipedriveToken_();
  const url = CONFIG.PIPEDRIVE.BASE_URL + '/deals/' + encodeURIComponent(dealId) +
              '?api_token=' + encodeURIComponent(token);

  const resp = UrlFetchApp.fetch(url, {
    method: 'put',
    contentType: 'application/json',
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  const code = resp.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Pipedrive PUT /deals failed: ' + code + ' → ' + resp.getContentText());
  }
}

// Create a proposal folder under ROOT_PROPOSAL_FOLDER_ID
function createProposalFolder_(proposal, streetOnly) {
  const root = DriveApp.getFolderById(CONFIG.ROOT_PROPOSAL_FOLDER_ID);
  const template = DriveApp.getFolderById(CONFIG.TEMPLATE_PROPOSAL_FOLDER_ID);

  const safeStreet = sanitizeFolderNamePart_(streetOnly);
  const newName = proposal + (safeStreet ? ' ' + safeStreet : '');

  const newFolder = root.createFolder(newName);

  // Copy template contents recursively
  copyFolderContents_(template, newFolder);

  return newFolder;
}

// Recursively copy files + folders
function copyFolderContents_(src, dest) {
  // Files
  const files = src.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    f.makeCopy(f.getName(), dest);
  }

  // Folders
  const subs = src.getFolders();
  while (subs.hasNext()) {
    const sf = subs.next();
    const newSub = dest.createFolder(sf.getName());
    copyFolderContents_(sf, newSub);
  }
}

// Sanitize address part for folder name
function sanitizeFolderNamePart_(name) {
  let s = String(name || '').trim();
  s = s.replace(/[\\/:*?"<>|]/g, '-');
  if (s.length > 80) s = s.substring(0, 80).trim();
  return s;
}

// Extract folder ID from Drive URL
function getFolderIdFromUrl_(url) {
  if (!url || typeof url !== 'string') return null;
  
  // Handle both formats:
  // - https://drive.google.com/drive/folders/ABC123
  // - https://drive.google.com/drive/folders/ABC123?usp=sharing
  const parts = url.split('/');
  const idPart = parts[parts.length - 1];
  return idPart.split('?')[0]; // Strip query params if present
}

// Validate that Folder URL points to a folder under ROOT_PROPOSAL_FOLDER_ID
function validateFolderUrl_(url) {
  /**
   * Returns: { valid: boolean, folderId: string, error: string }
   */
  if (!url || typeof url !== 'string') {
    return { valid: false, error: 'Empty or invalid URL' };
  }
  
  try {
    // Extract folder ID
    const folderId = getFolderIdFromUrl_(url);
    if (!folderId) {
      return { valid: false, error: 'Could not extract folder ID from URL' };
    }
    
    // Check if folder exists and is under root
    const folder = DriveApp.getFolderById(folderId);
    const rootFolder = DriveApp.getFolderById(CONFIG.ROOT_PROPOSAL_FOLDER_ID);
    
    // Walk up parent chain to verify it's under root
    let current = folder;
    let depth = 0;
    const MAX_DEPTH = 10; // Prevent infinite loops
    
    while (depth < MAX_DEPTH) {
      const parents = current.getParents();
      if (!parents.hasNext()) {
        return { valid: false, error: 'Folder not under proposal root' };
      }
      
      const parent = parents.next();
      if (parent.getId() === rootFolder.getId()) {
        // Found root, validation passed
        return { valid: true, folderId: folderId };
      }
      
      current = parent;
      depth++;
    }
    
    return { valid: false, error: 'Folder structure too deep or not under root' };
    
  } catch (e) {
    return { valid: false, error: 'Folder not accessible: ' + e.message };
  }
}

// Determine next proposal number from Proposals sheet
function getNextProposalNumber_() {
  refreshProposalsFromDrive();   // from Proposals.js

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CONFIG.SHEET_PROPOSALS);

  const last = sh.getLastRow();
  if (last <= 1) return 'PP0';

  const keyRange = sh.getRange(2, 1, last - 1, 1).getValues();

  let maxKey = 0;
  keyRange.forEach(row => {
    const v = row[0];
    if (typeof v === 'number' && v > maxKey) maxKey = v;
  });

  const nextKey = maxKey + 1;
  return 'PP' + nextKey;
}

/**
 * Fast version: Get next proposal number from cached Proposals sheet.
 * Used in webhook to avoid expensive refreshProposalsFromDrive() call.
 * Falls back to Drive scan if sheet unavailable.
 * @private
 */
function getNextProposalNumberFast_() {
  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(CONFIG.SHEET_PROPOSALS);
    
    if (!sheet || sheet.getLastRow() <= 1) {
      Logger.log('getNextProposalNumberFast_: Proposals sheet empty, scanning Drive instead');
      return getNextProposalNumberFromDrive_();
    }
    
    // Read Key column (fastest - just one column)
    const lastRow = sheet.getLastRow();
    const keyRange = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    
    let maxKey = 0;
    keyRange.forEach(function(row) {
      const v = row[0];
      if (typeof v === 'number' && v > maxKey) {
        maxKey = v;
      }
    });
    
    const nextKey = maxKey + 1;
    Logger.log('getNextProposalNumberFast_: Next PP# = PP' + nextKey + ' (from sheet, max key: ' + maxKey + ')');
    return 'PP' + nextKey;
    
  } catch (e) {
    Logger.log('getNextProposalNumberFast_: Error reading sheet: ' + e.message + ', falling back to Drive scan');
    return getNextProposalNumberFromDrive_();
  }
}

/**
 * Fallback: Scan Drive directly to find next proposal number.
 * Used when Proposals sheet is unavailable or outdated.
 * @private
 */
function getNextProposalNumberFromDrive_() {
  try {
    const parentFolder = DriveApp.getFolderById(CONFIG.ROOT_PROPOSAL_FOLDER_ID);
    const folders = parentFolder.getFolders();
    
    let maxKey = 0;
    
    while (folders.hasNext()) {
      const folder = folders.next();
      const folderName = folder.getName();
      
      // Extract PP number from folder name
      const match = folderName.match(/^PP(\d+)/);
      if (match) {
        const key = parseInt(match[1], 10);
        if (key > maxKey) {
          maxKey = key;
        }
      }
    }
    
    const nextKey = maxKey + 1;
    Logger.log('getNextProposalNumberFromDrive_: Next PP# = PP' + nextKey + ' (from Drive scan, max: ' + maxKey + ')');
    return 'PP' + nextKey;
    
  } catch (e) {
    Logger.log('getNextProposalNumberFromDrive_: Error: ' + e.message);
    throw new Error('Cannot determine next proposal number: ' + e.message);
  }
}

// Log events into Logs sheet
function logEvent_(code, msg, detail) {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(CONFIG.SHEET_LOGS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_LOGS);
    sheet.appendRow(['Timestamp', 'Code', 'Message', 'Detail']);
  }
  sheet.appendRow([new Date(), code, msg, detail]);
}

/**
 * Maintains a pool of preallocated placeholder proposal folders.
 * Keeps TARGET_PLACEHOLDERS (default: 1) ready at all times.
 * Run on time-based trigger (e.g., every 30 minutes).
 */
function replenishPlaceholders() {
  const TARGET_PLACEHOLDERS = 3;
  
  Logger.log('replenishPlaceholders: Starting...');
  logEvent_('REPLENISH_START', 'Placeholder replenishment started', 'Target: ' + TARGET_PLACEHOLDERS);
  
  // Refresh Proposals sheet from Drive first
  refreshProposalsFromDrive();
  
  const parentFolder = DriveApp.getFolderById(CONFIG.ROOT_PROPOSAL_FOLDER_ID);
  const templateFolder = DriveApp.getFolderById(CONFIG.TEMPLATE_PROPOSAL_FOLDER_ID);
  
  // Count existing placeholders by checking folder names
  let placeholderCount = 0;
  const folders = parentFolder.getFolders();
  while (folders.hasNext()) {
    const folder = folders.next();
    if (folder.getName().includes('[PLACEHOLDER]')) {
      placeholderCount++;
    }
  }
  
  Logger.log('replenishPlaceholders: Found ' + placeholderCount + ' existing placeholders');
  
  // Calculate how many to create
  const needed = TARGET_PLACEHOLDERS - placeholderCount;
  
  if (needed <= 0) {
    Logger.log('replenishPlaceholders: Pool sufficient (' + placeholderCount + '/' + TARGET_PLACEHOLDERS + ')');
    return;
  }
  
  Logger.log('replenishPlaceholders: Creating ' + needed + ' new placeholder(s)...');
  
  // Get starting PP# once, then increment locally
  let nextPPNum = getNextProposalNumberFast_();
  let nextKey = parseInt(nextPPNum.replace('PP', ''), 10);
  
  // Create new placeholders
  for (let i = 0; i < needed; i++) {
    try {
      const proposalNum = 'PP' + nextKey;
      const placeholderName = proposalNum + ' [PLACEHOLDER]';
      
      // Create folder from template
      const newFolder = parentFolder.createFolder(placeholderName);
      
      // Copy template contents
      copyFolderContents_(templateFolder, newFolder);
      
      const folderUrl = newFolder.getUrl();
      
      Logger.log('replenishPlaceholders: Created ' + placeholderName + ' (' + folderUrl + ')');
      logEvent_('PLACEHOLDER_CREATED', 'Created ' + placeholderName, folderUrl);
      
      nextKey++; // Increment for next iteration
      
    } catch (e) {
      Logger.log('replenishPlaceholders: Error creating placeholder: ' + e.message);
      logEvent_('PLACEHOLDER_ERROR', 'Failed to create placeholder', e.message);
    }
  }
  
  // Refresh Proposals sheet to reflect new placeholders
  refreshProposalsFromDrive();
  
  // Refresh Deals sheet to sync any updated data from Pipedrive
  try {
    PullFromPipedrive();
    Logger.log('replenishPlaceholders: Refreshed Deals from Pipedrive');
  } catch (pullErr) {
    Logger.log('replenishPlaceholders: Failed to refresh Deals: ' + pullErr.message);
  }
  
  Logger.log('replenishPlaceholders: Complete. Pool now has ' + (placeholderCount + needed) + ' placeholders');
  logEvent_('REPLENISH_COMPLETE', 'Placeholder replenishment completed', 'Created: ' + needed + ', Total: ' + (placeholderCount + needed));
}



/**
 * Finds and claims an available placeholder folder for a deal.
 * Uses cached Proposals sheet (fast) instead of scanning Drive (slow).
 * Returns the folder with the LOWEST PP number to avoid skipping numbers.
 * @private
 */
function findAndClaimPlaceholder_() {
  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(CONFIG.SHEET_PROPOSALS);
    
    if (!sheet || sheet.getLastRow() <= 1) {
      Logger.log('findAndClaimPlaceholder_: Proposals sheet empty, falling back to Drive scan');
      return findAndClaimPlaceholderFromDrive_();
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find column indices
    const folderNameIdx = headers.indexOf('Folder Name');
    const folderUrlIdx = headers.indexOf('Folder URL');
    const keyIdx = headers.indexOf('Key');
    
    if (folderNameIdx === -1 || folderUrlIdx === -1) {
      Logger.log('findAndClaimPlaceholder_: Required columns not found, falling back to Drive scan');
      return findAndClaimPlaceholderFromDrive_();
    }
    
    // Find lowest-numbered placeholder from sheet
    let lowestPlaceholder = null;
    let lowestKey = Infinity;
    let lowestUrl = null;
    
    for (let i = 1; i < data.length; i++) {
      const folderName = data[i][folderNameIdx];
      const folderUrl = data[i][folderUrlIdx];
      const key = keyIdx !== -1 ? data[i][keyIdx] : null;
      
      if (folderName && folderName.toString().includes('[PLACEHOLDER]')) {
        const keyNum = key || extractKeyFromProposalNumber_(folderName);
        
        if (keyNum < lowestKey) {
          lowestKey = keyNum;
          lowestPlaceholder = folderName;
          lowestUrl = folderUrl;
        }
      }
    }
    
    if (lowestPlaceholder && lowestUrl) {
      Logger.log('findAndClaimPlaceholder_: Found lowest placeholder from sheet: ' + lowestPlaceholder + ' (key=' + lowestKey + ')');
      
      // Get folder from URL (fast - direct access)
      const folderId = getFolderIdFromUrl_(lowestUrl);
      const folder = DriveApp.getFolderById(folderId);
      
      return folder;
    }
    
    Logger.log('findAndClaimPlaceholder_: No placeholders in sheet');
    return null;
    
  } catch (e) {
    Logger.log('findAndClaimPlaceholder_: Error reading sheet: ' + e.message + ', falling back to Drive scan');
    return findAndClaimPlaceholderFromDrive_();
  }
}

/**
 * Fallback: Scan Drive directly for placeholders.
 * Used when sheet is unavailable or outdated.
 * @private
 */
function findAndClaimPlaceholderFromDrive_() {
  try {
    const parentFolder = DriveApp.getFolderById(CONFIG.ROOT_PROPOSAL_FOLDER_ID);
    const folders = parentFolder.getFolders();
    
    let lowestPlaceholder = null;
    let lowestKey = Infinity;
    
    while (folders.hasNext()) {
      const folder = folders.next();
      const folderName = folder.getName();
      
      if (folderName.includes('[PLACEHOLDER]')) {
        const match = folderName.match(/^PP(\d+)/);
        if (match) {
          const key = parseInt(match[1], 10);
          if (key < lowestKey) {
            lowestKey = key;
            lowestPlaceholder = folder;
          }
        }
      }
    }
    
    if (lowestPlaceholder) {
      Logger.log('findAndClaimPlaceholderFromDrive_: Found lowest placeholder: ' + lowestPlaceholder.getName());
      return lowestPlaceholder;
    }
    
    return null;
    
  } catch (e) {
    Logger.log('findAndClaimPlaceholderFromDrive_: Error: ' + e.message);
    return null;
  }
}

/**
 * Extract numeric key from proposal number (e.g., "PP170" -> 170)
 * @private
 */
function extractKeyFromProposalNumber_(proposalOrFolderName) {
  const match = String(proposalOrFolderName).match(/^PP(\d+)/);
  return match ? parseInt(match[1], 10) : Infinity;
}
