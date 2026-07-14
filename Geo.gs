// ========================================
// GEO.gs - Geolocation Helper
// ========================================

function getGeoLocation(lat, lon) {
  try {
    var url = "https://nominatim.openstreetmap.org/reverse?format=json&lat=" + lat + "&lon=" + lon + "&zoom=18&addressdetails=1";
    
    // Header User-Agent WAJIB ada agar tidak diblokir Nominatim
    var options = {
      "headers": { "User-Agent": "HifdzulApp/6.0 (id.my.app)" } 
    };
    
    var response = UrlFetchApp.fetch(url, options);
    var json = JSON.parse(response.getContentText());
    return json;
  } catch (e) {
    return null;
  }
}
