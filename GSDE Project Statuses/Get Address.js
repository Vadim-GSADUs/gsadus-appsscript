/**
 * Returns address components, geolocation, elevation, or map links.
 * @param {string} address The address to search.
 * @param {string} type The data to return.
 * Options: "street", "city", "county", "zip", "state", "full", 
 * "lat", "lng", "elevation", "elev", "link", "place_id", "type"
 * @customfunction
 */
function GET_ADDRESS(address, type) {
  if (!address || address === "") return "";
  
  // 1. GENERATE A UNIQUE KEY FOR THIS REQUEST
  // We clean the address to ensure "123 Main" and "123 Main " are treated the same
  var cleanAddress = address.toString().trim().toLowerCase();
  var cacheKey = Utilities.base64Encode(cleanAddress + "_" + type);
  
  // 2. CHECK CACHE (The "Speed Guard")
  // If we already looked this up recently, return the answer instantly.
  var cache = CacheService.getScriptCache();
  var cachedResult = cache.get(cacheKey);
  
  if (cachedResult != null) {
    // Return number if it looks like a number (for Lat/Long)
    return isNaN(cachedResult) ? cachedResult : Number(cachedResult);
  }

  // 3. IF NOT IN CACHE, CALL GOOGLE MAPS API (The "Heavy Lift")
  try {
    var response = Maps.newGeocoder().geocode(address);
    var output = "Not found"; // Default value

    if (response.status === 'OK' && response.results.length > 0) {
      var result = response.results[0];
      var components = result.address_components;
      var geometry = result.geometry;
      
      var streetNum = "";
      var route = "";
      var city = "";
      var county = "";
      var state = "";
      var zip = "";
      
      for (var i = 0; i < components.length; i++) {
        var c = components[i];
        if (c.types.indexOf("street_number") > -1) streetNum = c.long_name;
        if (c.types.indexOf("route") > -1) route = c.long_name;
        if (c.types.indexOf("locality") > -1) city = c.long_name;
        if (c.types.indexOf("administrative_area_level_2") > -1) county = c.long_name;
        if (c.types.indexOf("administrative_area_level_1") > -1) state = c.short_name;
        if (c.types.indexOf("postal_code") > -1) zip = c.long_name;
      }

      // Determine output based on requested type
      if (type === "lat") output = geometry.location.lat;
      else if (type === "lng") output = geometry.location.lng; // Note: 'long' is a reserved word in JS, usually 'lng' is safer
      else if (type === "long") output = geometry.location.lng;
      else if (type === "link") output = "https://www.google.com/maps/search/?api=1&query=" + encodeURIComponent(result.formatted_address);
      else if (type === "place_id") output = result.place_id;
      else if (type === "type") output = geometry.location_type;
      else if (type === "street") output = (streetNum + " " + route).trim();
      else if (type === "city") output = city;
      else if (type === "county") output = county;
      else if (type === "state") output = state;
      else if (type === "zip") output = zip;
      else if (type === "full") {
        var streetPart = (streetNum + " " + route).trim();
        output = streetPart + ", " + city + ", " + state + " " + zip;
      }
    }

    // 4. SAVE RESULT TO CACHE
    // Store it for 6 hours (21600 seconds), the maximum allowed.
    cache.put(cacheKey, String(output), 21600);
    
    return output;

  } catch (e) {
    return "Error: " + e.message;
  }
}