function doGet(e) {
  return ContentService
    .createTextOutput('Pipedrive webhook endpoint is alive')
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    var raw = (e && e.postData && e.postData.contents) || '';
    if (!raw) return okResponse_('NO_BODY');

    var payload;
    try {
      payload = JSON.parse(raw);
    } catch (parseErr) {
      logEvent_('BAD_JSON', 'Failed to parse webhook JSON', String(parseErr));
      return okResponse_('BAD_JSON');
    }

    var meta = payload.meta || {};
    if (meta.change_source !== 'app') {
      return okResponse_('IGNORED_CHANGE_SOURCE');
    }

    var createId = String(CONFIG.PIPEDRIVE.LABELS.CREATE_PP);
    var currentSnap = payload.data || {};
    if (!snapshotHasLabelId_(currentSnap, createId)) {
      return okResponse_('NO_CREATE_LABEL');
    }

    var dealId = getDealIdFromWebhook_(payload);
    if (!dealId) {
      logEvent_('NO_DEAL_ID', 'Webhook payload missing deal id', JSON.stringify(payload));
      return okResponse_('NO_DEAL_ID');
    }

    var result = runWithGlobalLock_(function () {
      var deal;
      try {
        deal = fetchDealFromPipedrive_(dealId);
      } catch (fetchErr) {
        logEvent_('DEAL_FETCH_ERROR', 'Pipedrive GET /deals failed', 'Deal ID: ' + dealId + ' | ' + String(fetchErr));
        return 'DEAL_FETCH_ERROR';
      }

      if (!deal) {
        logEvent_('DEAL_NOT_FOUND', 'Could not fetch deal from Pipedrive', 'Deal ID: ' + dealId);
        return 'DEAL_NOT_FOUND';
      }

      return handleDealChange_(deal);
    });

    return okResponse_(result || null);

  } catch (err) {
    logEvent_('ERROR', 'doPost exception', err && err.stack ? err.stack : String(err));
    return okResponse_('HANDLER_EXCEPTION');
  }
}

function okResponse_(result) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', result: result }))
    .setMimeType(ContentService.MimeType.JSON);
}

function runWithGlobalLock_(fn) {
  var lock = LockService.getScriptLock();
  var got = false;

  try {
    got = lock.tryLock(30 * 1000);
  } catch (err) {
    logEvent_('LOCK_ERROR', 'Error acquiring script lock', String(err));
    return 'LOCK_ERROR';
  }

  if (!got) {
    logEvent_('LOCK_TIMEOUT', 'Failed to acquire script lock in time', '');
    return 'LOCK_TIMEOUT';
  }

  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function getDealIdFromWebhook_(payload) {
  if (!payload) return null;

  function asInt(v) {
    if (v === null || v === undefined) return null;
    var n = parseInt(v, 10);
    return Number.isNaN(n) ? null : n;
  }

  if (payload.data && payload.data.id !== undefined) {
    var n1 = asInt(payload.data.id);
    if (n1 !== null) return n1;
  }

  if (payload.meta && payload.meta.entity_id !== undefined) {
    var n2 = asInt(payload.meta.entity_id);
    if (n2 !== null) return n2;
  }

  if (payload.meta && payload.meta.id !== undefined) {
    var n3 = asInt(payload.meta.id);
    if (n3 !== null) return n3;
  }

  return null;
}

function handleDealChange_(deal) {
  var fieldKeys = CONFIG.PIPEDRIVE.FIELD_KEYS;
  var labelsCfg = CONFIG.PIPEDRIVE.LABELS;
  var dealId = deal.id;

  var proposalVal = deal[fieldKeys.PROPOSAL];
  var folderUrlVal = deal[fieldKeys.FOLDER_URL];

  if ((proposalVal && String(proposalVal).trim()) || (folderUrlVal && String(folderUrlVal).trim())) {
    return 'SKIPPED_EXISTING_PROPOSAL';
  }

  var rawAddress = deal[fieldKeys.ADDRESS] || '';
  var addressStr = String(rawAddress).trim();

  if (!addressStr) {
    var bodyNeedsAddr = {
      label: toPipedriveLabelValue_(String(labelsCfg.NEEDS_ADDR))
    };

    try {
      updateDealFields_(dealId, bodyNeedsAddr);
      logEvent_('NEEDS_ADDRESS', 'Deal lacked address; set Needs Address label', 'Deal ID: ' + dealId);
    } catch (putErr) {
      logEvent_('PUT_ERROR_NEEDS_ADDR', 'Failed to set Needs Address label', 'Deal ID: ' + dealId + ' | ' + String(putErr));
    }

    return 'NEEDS_ADDRESS';
  }

  var streetOnly = extractStreetFromAddress_(addressStr);
  var folder = null;
  var proposalNum = null;
  var folderUrl = null;
  var usedPlaceholder = false;

  var placeholderFolder = findAndClaimPlaceholder_();

  if (placeholderFolder) {
    try {
      var placeholderName = placeholderFolder.getName();
      proposalNum = extractProposalNumber_(placeholderName);

      var safeStreet = sanitizeFolderNamePart_(streetOnly);
      var newName = proposalNum + (safeStreet ? ' ' + safeStreet : '');

      placeholderFolder.setName(newName);
      folder = placeholderFolder;
      folderUrl = folder.getUrl();
      usedPlaceholder = true;

      logEvent_('PLACEHOLDER_ASSIGNED', 'Assigned placeholder to deal', 'Deal ID: ' + dealId + ', Proposal: ' + proposalNum);

    } catch (renameErr) {
      logEvent_('PLACEHOLDER_RENAME_ERROR', 'Failed to rename placeholder, falling back to on-demand creation', 'Deal ID: ' + dealId + ' | ' + String(renameErr));
      placeholderFolder = null;
    }
  }

  if (!folder) {
    logEvent_('PLACEHOLDER_DEPLETED', 'No placeholders available, creating on-demand', 'Deal ID: ' + dealId);

    proposalNum = getNextProposalNumberFast_();
    folder = createProposalFolder_(proposalNum, streetOnly);
    folderUrl = folder.getUrl();
    usedPlaceholder = false;
  }

  var updateBody = {};
  updateBody[fieldKeys.PROPOSAL] = proposalNum;
  updateBody[fieldKeys.FOLDER_URL] = folderUrl;
  updateBody.label = null;

  try {
    updateDealFields_(dealId, updateBody);
    logEvent_('PROPOSAL_COMPLETE', usedPlaceholder ? 'Assigned placeholder' : 'Created on-demand', 'Deal ID: ' + dealId + ', Proposal: ' + proposalNum);
  } catch (putErr) {
    logEvent_('PUT_ERROR_CREATE_PROPOSAL', 'Failed to update deal with proposal/folder/label', 'Deal ID: ' + dealId + ' | ' + String(putErr));
  }

  if (usedPlaceholder) {
    triggerAsyncReplenish_();
  }

  return {
    action: usedPlaceholder ? 'ASSIGNED_PLACEHOLDER' : 'CREATED_PROPOSAL',
    proposal: proposalNum,
    folderUrl: folderUrl
  };
}

function triggerAsyncReplenish_() {
  try {
    ScriptApp.newTrigger('replenishPlaceholders')
      .timeBased()
      .after(1000)
      .create();
    Logger.log('triggerAsyncReplenish_: Trigger created successfully (will execute in 1 second)');
    logEvent_('ASYNC_REPLENISH_TRIGGER', 'Created trigger to replenish placeholder pool', 'Delay: 1 second');
  } catch (e) {
    Logger.log('triggerAsyncReplenish_: Failed to create trigger: ' + e.message);
    logEvent_('ASYNC_REPLENISH_ERROR', 'Failed to create async replenish trigger', e.message);
  }
}

function snapshotHasLabelId_(snap, targetId) {
  var ids = extractLabelIdsFromSnapshot_(snap);
  var target = String(targetId);
  return ids.indexOf(target) !== -1;
}

function extractLabelIdsFromSnapshot_(snap) {
  var out = [];
  if (!snap) return out;

  if (Array.isArray(snap.label_ids)) {
    snap.label_ids.forEach(function (id) {
      if (id !== null && id !== undefined && id !== '') {
        out.push(String(id));
      }
    });
  } else if (snap.label !== null && snap.label !== undefined && snap.label !== '') {
    var val = String(snap.label);
    val.split(',').forEach(function (part) {
      var p = part.trim();
      if (p) out.push(p);
    });
  }

  return out;
}

function toPipedriveLabelValue_(idStr) {
  var n = Number(idStr);
  return isNaN(n) ? null : n;
}

/**
 * Diagnostic function to check and clean up old replenishment triggers.
 * Run manually from menu if triggers are accumulating.
 */
function cleanupReplenishmentTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let cleaned = 0;
  
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'replenishPlaceholders') {
      ScriptApp.deleteTrigger(trigger);
      cleaned++;
    }
  });
  
  Logger.log('Cleaned up ' + cleaned + ' replenishment triggers');
  logEvent_('TRIGGER_CLEANUP', 'Cleaned up replenishment triggers', 'Count: ' + cleaned);
  
  const ui = SpreadsheetApp.getUi();
  ui.alert('Trigger Cleanup', 'Removed ' + cleaned + ' old replenishment triggers.', ui.ButtonSet.OK);
}
