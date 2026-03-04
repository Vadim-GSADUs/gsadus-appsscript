function geocodeDeals() {
  const SHEET_NAME = 'Deals';

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" not found.`);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const headerRow = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];

  function col(name) {
    const i = headerRow.indexOf(name);
    if (i === -1) throw new Error(`Header not found: ${name}`);
    return i;
  }

  const addrCol     = col('Deal - Address');
  const fullCol     = col('Deal - Full/combined address of Address');
  const latCol      = col('Deal - Latitude of Address');
  const lngCol      = col('Deal - Longitude of Address');

  const dataRange = sheet.getRange(2,1,lastRow-1,headerRow.length);
  const values = dataRange.getValues();

  for (let r = 0; r < values.length; r++) {
    const row = values[r];

    // Skip rows already geocoded
    if (row[latCol] && row[lngCol]) continue;

    // Choose address
    let addr = row[addrCol];
    if (!addr) addr = row[fullCol];
    if (!addr) continue;

    try {
      const result = Maps.newGeocoder().geocode(addr);

      if (result.status === 'OK' && result.results.length > 0) {
        const loc = result.results[0].geometry.location;
        row[latCol] = loc.lat;
        row[lngCol] = loc.lng;
      }

      Utilities.sleep(150);
    } catch (err) {
      Logger.log(`Error for "${addr}": ${err}`);
    }
  }

  dataRange.setValues(values);
}
