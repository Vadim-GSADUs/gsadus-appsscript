/**
 * Maintenance / Diagnostics (manual-only; non-critical)
 */

// One-time export of all Persons with id, name, email (from Pipedrive to a new sheet)
function exportPersonsToSheet() {
  const pd = new PipedriveClient();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('PersonsExport') || ss.insertSheet('PersonsExport');
  sh.clear();
  sh.appendRow(['id','name','email']);

  let start = 0;
  const LIMIT = 200;

  while (true) {
    const res = pd.listPersonsPage(start, LIMIT);
    const items = res?.data || [];
    if (!items.length) break;

    for (const p of items) {
      const emails = Array.isArray(p.email) ? p.email : [];
      const addr = (emails[0] && emails[0].value) || '';
      sh.appendRow([p.id, p.name, addr]);
    }

    if (res?.additional_data?.pagination?.more_items_in_collection) {
      start = res.additional_data.pagination.next_start;
    } else break;
  }

  console.log('Export completed.');
}

// Inspect cache size (email cache count)
function debugPersonsCacheCount() {
  const raw = PropertiesService.getScriptProperties().getProperty('pd_persons_cache_v1') || '{}';
  const obj = JSON.parse(raw);
  console.log(`cache persons = ${Object.keys(obj).length}`);
}

