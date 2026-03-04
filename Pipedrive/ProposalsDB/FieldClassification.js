// ---------------------------
// Field Editability Classification
// ---------------------------

/**
 * Field classification for push-back safety.
 * 
 * Category 1 (USER_EDITABLE): Safe to push via PUT /deals/{id}
 * Category 2 (SYSTEM_CALCULATED): Never push - computed by Pipedrive
 * Category 3 (SPECIAL_HANDLING): Requires special API endpoints or validation
 */

/**
 * Returns classification for a given field key.
 * @param {string} fieldKey - The Pipedrive field key (e.g., 'title', 'stage_id', 'weighted_value')
 * @returns {string} - 'USER_EDITABLE', 'SYSTEM_CALCULATED', 'SPECIAL_HANDLING', or 'UNKNOWN'
 */
function classifyField(fieldKey) {
  // Check if it's a custom field (long hash key)
  if (fieldKey && fieldKey.length > 30 && fieldKey.match(/^[a-f0-9]+$/)) {
    return 'USER_EDITABLE'; // All custom fields are editable
  }

  // Category 1: User-Editable
  const userEditableFields = [
    'title',
    'value',
    'currency',
    'user_id',
    'person_id',
    'org_id',
    'stage_id',
    'status',
    'probability',
    'expected_close_date',
    'visible_to',
    'lost_reason'
  ];

  if (userEditableFields.indexOf(fieldKey) !== -1) {
    return 'USER_EDITABLE';
  }

  // Category 2: System-Calculated (Never Push)
  const systemCalculatedFields = [
    'id',
    'creator_user_id',
    'weighted_value',
    'weighted_value_currency',
    'activities_count',
    'done_activities_count',
    'undone_activities_count',
    'email_messages_count',
    'next_activity_date',
    'last_activity_date',
    'add_time',
    'update_time',
    'stage_change_time',
    'won_time',
    'lost_time',
    'close_time',
    'last_incoming_mail_time',
    'last_outgoing_mail_time',
    'archive_time',
    'is_archived',
    'origin',
    'origin_id',
    'channel',
    'channel_id',
    'sequence_enrollment',
    'product_quantity',
    'product_amount',
    'product_name'
  ];

  if (systemCalculatedFields.indexOf(fieldKey) !== -1) {
    return 'SYSTEM_CALCULATED';
  }

  // Category 3: Special Handling Required
  const specialHandlingFields = [
    'pipeline', // Requires stage_id validation
    'label'     // Use label management API
  ];

  if (specialHandlingFields.indexOf(fieldKey) !== -1) {
    return 'SPECIAL_HANDLING';
  }

  return 'UNKNOWN';
}

/**
 * Returns a map of column header → field key for all Deal columns.
 * This allows looking up field classification by column header name.
 * @returns {Object} Map of header name → {key, category, notes}
 */
function getFieldClassificationMap() {
  return {
    'Deal - ID': { key: 'id', category: 'SYSTEM_CALCULATED', notes: 'Primary key' },
    'Deal - Title': { key: 'title', category: 'USER_EDITABLE', notes: '' },
    'Deal - Creator': { key: 'creator_user_id', category: 'SYSTEM_CALCULATED', notes: 'Immutable' },
    'Deal - Owner': { key: 'user_id', category: 'USER_EDITABLE', notes: '' },
    'Deal - Value': { key: 'value', category: 'USER_EDITABLE', notes: '' },
    'Deal - Currency of Value': { key: 'currency', category: 'USER_EDITABLE', notes: '' },
    'Deal - Weighted value': { key: 'weighted_value', category: 'SYSTEM_CALCULATED', notes: 'Calculated: value × probability' },
    'Deal - Currency of Weighted value': { key: 'weighted_value_currency', category: 'SYSTEM_CALCULATED', notes: '' },
    'Deal - Probability': { key: 'probability', category: 'USER_EDITABLE', notes: 'Percent (0-100)' },
    'Deal - Organization': { key: 'org_id', category: 'USER_EDITABLE', notes: 'Reference to org' },
    'Deal - Organization ID': { key: 'org_id', category: 'USER_EDITABLE', notes: 'Org ID value' },
    'Deal - Pipeline': { key: 'pipeline', category: 'SPECIAL_HANDLING', notes: 'Must update stage_id when changing' },
    'Deal - Contact person': { key: 'person_id', category: 'USER_EDITABLE', notes: 'Reference to person' },
    'Deal - Contact person ID': { key: 'person_id', category: 'USER_EDITABLE', notes: 'Person ID value' },
    'Deal - Stage': { key: 'stage_id', category: 'USER_EDITABLE', notes: '' },
    'Deal - Label': { key: 'label', category: 'SPECIAL_HANDLING', notes: 'Use label management API' },
    'Deal - Status': { key: 'status', category: 'USER_EDITABLE', notes: 'open/won/lost/deleted' },
    'Deal - Deal created': { key: 'add_time', category: 'SYSTEM_CALCULATED', notes: 'Immutable creation time' },
    'Deal - Update time': { key: 'update_time', category: 'SYSTEM_CALCULATED', notes: 'Auto-updated by server' },
    'Deal - Last stage change': { key: 'stage_change_time', category: 'SYSTEM_CALCULATED', notes: 'Auto-updated on stage change' },
    'Deal - Next activity date': { key: 'next_activity_date', category: 'SYSTEM_CALCULATED', notes: 'Computed from activities' },
    'Deal - Last activity date': { key: 'last_activity_date', category: 'SYSTEM_CALCULATED', notes: 'Computed from activities' },
    'Deal - Won time': { key: 'won_time', category: 'SYSTEM_CALCULATED', notes: 'Auto-set when status → won' },
    'Deal - Last email received': { key: 'last_incoming_mail_time', category: 'SYSTEM_CALCULATED', notes: '' },
    'Deal - Last email sent': { key: 'last_outgoing_mail_time', category: 'SYSTEM_CALCULATED', notes: '' },
    'Deal - Lost time': { key: 'lost_time', category: 'SYSTEM_CALCULATED', notes: 'Auto-set when status → lost' },
    'Deal - Deal closed on': { key: 'close_time', category: 'SYSTEM_CALCULATED', notes: 'Auto-set on deal close' },
    'Deal - Lost reason': { key: 'lost_reason', category: 'USER_EDITABLE', notes: 'Required when marking deal as lost' },
    'Deal - Visible to': { key: 'visible_to', category: 'USER_EDITABLE', notes: '1=Item owner, 3=All users, 7=Owner only' },
    'Deal - Total activities': { key: 'activities_count', category: 'SYSTEM_CALCULATED', notes: 'Computed count' },
    'Deal - Done activities': { key: 'done_activities_count', category: 'SYSTEM_CALCULATED', notes: 'Computed count' },
    'Deal - Activities to do': { key: 'undone_activities_count', category: 'SYSTEM_CALCULATED', notes: 'Computed count' },
    'Deal - Email messages count': { key: 'email_messages_count', category: 'SYSTEM_CALCULATED', notes: 'Computed count' },
    'Deal - Expected close date': { key: 'expected_close_date', category: 'USER_EDITABLE', notes: 'Date format: YYYY-MM-DD' },
    'Deal - Product quantity': { key: 'product_quantity', category: 'SYSTEM_CALCULATED', notes: 'Use /deals/{id}/products API' },
    'Deal - Product amount': { key: 'product_amount', category: 'SYSTEM_CALCULATED', notes: 'Use /deals/{id}/products API' },
    'Deal - Product name': { key: 'product_name', category: 'SYSTEM_CALCULATED', notes: 'Use /deals/{id}/products API' },
    'Deal - Source origin': { key: 'origin', category: 'SYSTEM_CALCULATED', notes: 'Set on creation, immutable' },
    'Deal - Source origin ID': { key: 'origin_id', category: 'SYSTEM_CALCULATED', notes: '' },
    'Deal - Source channel': { key: 'channel', category: 'SYSTEM_CALCULATED', notes: 'Set on creation' },
    'Deal - Source channel ID': { key: 'channel_id', category: 'SYSTEM_CALCULATED', notes: '' },
    'Deal - Archive status': { key: 'is_archived', category: 'SYSTEM_CALCULATED', notes: 'Use archive/unarchive endpoints' },
    'Deal - Archive time': { key: 'archive_time', category: 'SYSTEM_CALCULATED', notes: 'Auto-set on archive' },
    'Deal - Sequence enrollment': { key: 'sequence_enrollment', category: 'SYSTEM_CALCULATED', notes: '' },
    'Deal - Address': { key: CONFIG.PIPEDRIVE.FIELD_KEYS.ADDRESS, category: 'USER_EDITABLE', notes: 'Custom field' },
    // NOTE: Lat/Long removed from sheet - Pipedrive doesn't return them in GET /deals
    // They exist as editable fields but are write-only via API
    // Use geocodeDeals() + push to populate them in Pipedrive if needed
    'Deal - Apartment/suite no of Address': { key: CONFIG.PIPEDRIVE.FIELD_KEYS.ADDRESS + '_subpremise', category: 'USER_EDITABLE', notes: '' },
    'Deal - House number of Address': { key: CONFIG.PIPEDRIVE.FIELD_KEYS.ADDRESS + '_street_number', category: 'USER_EDITABLE', notes: '' },
    'Deal - Street/road name of Address': { key: CONFIG.PIPEDRIVE.FIELD_KEYS.ADDRESS + '_route', category: 'USER_EDITABLE', notes: '' },
    'Deal - District/sublocality of Address': { key: CONFIG.PIPEDRIVE.FIELD_KEYS.ADDRESS + '_sublocality', category: 'USER_EDITABLE', notes: '' },
    'Deal - City/town/village/locality of Address': { key: CONFIG.PIPEDRIVE.FIELD_KEYS.ADDRESS + '_locality', category: 'USER_EDITABLE', notes: '' },
    'Deal - State/county of Address': { key: CONFIG.PIPEDRIVE.FIELD_KEYS.ADDRESS + '_admin_area_level_1', category: 'USER_EDITABLE', notes: '' },
    'Deal - Region of Address': { key: CONFIG.PIPEDRIVE.FIELD_KEYS.ADDRESS + '_admin_area_level_2', category: 'USER_EDITABLE', notes: '' },
    'Deal - Country of Address': { key: CONFIG.PIPEDRIVE.FIELD_KEYS.ADDRESS + '_country', category: 'USER_EDITABLE', notes: '' },
    'Deal - ZIP/Postal code of Address': { key: CONFIG.PIPEDRIVE.FIELD_KEYS.ADDRESS + '_postal_code', category: 'USER_EDITABLE', notes: '' },
    'Deal - Full/combined address of Address': { key: CONFIG.PIPEDRIVE.FIELD_KEYS.ADDRESS + '_formatted_address', category: 'USER_EDITABLE', notes: '' },
    'Deal - Proposal #': { key: CONFIG.PIPEDRIVE.FIELD_KEYS.PROPOSAL, category: 'USER_EDITABLE', notes: 'Custom field' },
    'Deal - Folder URL': { key: CONFIG.PIPEDRIVE.FIELD_KEYS.FOLDER_URL, category: 'USER_EDITABLE', notes: 'Custom field' }
  };
}

/**
 * Returns color code for a field category (for sheet formatting).
 * @param {string} category - Field category
 * @returns {string} Hex color code
 */
function getCategoryColor(category) {
  const colors = {
    'USER_EDITABLE': '#FFFFFF',        // White - editable
    'SYSTEM_CALCULATED': '#F3F3F3',    // Light gray - locked
    'SPECIAL_HANDLING': '#FFF9C4',     // Light yellow - warning
    'UNKNOWN': '#FFCDD2'               // Light red - error
  };
  return colors[category] || colors['UNKNOWN'];
}

/**
 * Returns whether a field category allows push-back.
 * @param {string} category - Field category
 * @returns {boolean}
 */
function isFieldPushable(category) {
  return category === 'USER_EDITABLE';
}
