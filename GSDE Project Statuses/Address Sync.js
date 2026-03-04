/**
 * ADDRESS SYNC ENGINE
 * Replaces volatile formulas with stable "Push" updates.
 */

// ---------------------------------------
// 1. AUTOMATIC TRIGGER (The "Smart" part)
// ---------------------------------------
function handleAddressEdit(e) {
  // Safety checks
  var sheet = e.source.getActiveSheet();
  if (sheet.getName() !== "Folders") return;
  if (e.range.getHeight() > 1 || e.range.getWidth() > 1) return; // Ignore bulk pastes for safety

  // Map Columns dynamically
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var colMap = getColumnMap(headers);
  
  if (!colMap.Address) return; // Exit if Address column missing

  // Only run if the USER edited the Address column
  if (e.range.getColumn() === colMap.Address) {
    var address = e.value;
    var row = e.range.getRow();

    // If user deleted the address, clear the details
    if (!address) {
      clearAddressFields(sheet, row, colMap);
      return;
    }

    // Fetch and Write
    var details = fetchAddressDetails(address);
    if (details) {
      writeAddressToRow(sheet, row, colMap, details);
    }
  }
}

// ---------------------------------------
// 2. BULK MENU TOOL (The "Cleanup" part)
// ---------------------------------------
function fillMissingAddresses() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Folders");
  var ui = SpreadsheetApp.getUi();

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var colMap = getColumnMap(headers);

  if (!colMap.Address) {
    ui.alert("Error: 'Address' column not found.");
    return;
  }

  var updates = 0;
  var errors = 0;

  // Loop through all rows (Skip header)
  for (var i = 1; i < data.length; i++) {
    var row = i + 1;
    var address = data[i][colMap.Address - 1];
    var lat = colMap.Lat ? data[i][colMap.Lat - 1] : "HasValue"; // Check if Lat is empty

    // If Address exists but Lat/Long is empty -> FETCH
    if (address && (!lat || lat === "")) {
      try {
        var details = fetchAddressDetails(address);
        if (details) {
          writeAddressToRow(sheet, row, colMap, details);
          updates++;
          // Pause briefly to be nice to Google Maps API during bulk runs
          Utilities.sleep(200); 
        }
      } catch (e) {
        errors++;
      }
    }
  }

  ui.alert("Bulk Update Complete.\nUpdated: " + updates + " rows.\nErrors: " + errors);
}


// ---------------------------------------
// HELPER FUNCTIONS
// ---------------------------------------

// 1. Fetch from Google Maps
function fetchAddressDetails(address) {
  try {
    var response = Maps.newGeocoder().geocode(address);
    if (response.status === 'OK' && response.results.length > 0) {
      var res = response.results[0];
      var comps = res.address_components;
      var geo = res.geometry;

      var parsed = {
        street_number: "", route: "", locality: "", 
        administrative_area_level_1: "", postal_code: "", administrative_area_level_2: ""
      };

      for (var i = 0; i < comps.length; i++) {
        var types = comps[i].types;
        for (var t = 0; t < types.length; t++) {
          if (parsed.hasOwnProperty(types[t])) {
            parsed[types[t]] = comps[i].long_name;
            if (types[t] === "administrative_area_level_1") parsed[types[t]] = comps[i].short_name; // State as "CA"
          }
        }
      }

      return {
        street: (parsed.street_number + " " + parsed.route).trim(),
        city: parsed.locality,
        county: parsed.administrative_area_level_2,
        state: parsed.administrative_area_level_1,
        zip: parsed.postal_code,
        lat: geo.location.lat,
        lng: geo.location.lng
      };
    }
  } catch (e) {
    Logger.log("Map Error: " + e.message);
  }
  return null;
}

// 2. Write Data to Sheet
function writeAddressToRow(sheet, row, colMap, data) {
  if (colMap.Street) sheet.getRange(row, colMap.Street).setValue(data.street);
  if (colMap.City) sheet.getRange(row, colMap.City).setValue(data.city);
  if (colMap.County) sheet.getRange(row, colMap.County).setValue(data.county);
  if (colMap.State) sheet.getRange(row, colMap.State).setValue(data.state);
  if (colMap.Zip) sheet.getRange(row, colMap.Zip).setValue(data.zip);
  if (colMap.Lat) sheet.getRange(row, colMap.Lat).setValue(data.lat);
  if (colMap.Long) sheet.getRange(row, colMap.Long).setValue(data.lng);
}

// 3. Clear Data (If address deleted)
function clearAddressFields(sheet, row, colMap) {
  var colsToClear = ["Street", "City", "County", "State", "Zip", "Lat", "Long"];
  colsToClear.forEach(function(c) {
    if (colMap[c]) sheet.getRange(row, colMap[c]).clearContent();
  });
}

// 4. Map Columns Dynamically (Finds column numbers by name)
function getColumnMap(headers) {
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    map[headers[i]] = i + 1; // 1-based index
  }
  return map;
}