// ============================================================================
// STEP 1: MIN-STAY VOR BUCHUNG AUF 1 SETZEN (TAB "RE")
// ============================================================================

/**
 * HAUPTFUNKTION SCHRITT 1
 */
function AAA_0_setMinStayToOne() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetRE = ss.getSheetByName("RE");

  if (!sheetRE) throw new Error("Das Tabellenblatt 'RE' wurde nicht gefunden.");

  var apartmentMap = getApartmentMapping();
  var lastRow = sheetRE.getLastRow();
  if (lastRow < 2) {
    Logger.log("Keine Daten im Sheet 'RE' vorhanden.");
    return;
  }

  // Liest Daten ab Zeile 2 von Spalte A (1) bis Spalte AG (33)
  var values = sheetRE.getRange(2, 1, lastRow - 1, 33).getValues();

  values.forEach(function(row, index) {
    var rowIndex = index + 2;
    var lodgifyStatus = row[18];   // Spalte S (Index 18)
    var minStayStatus = row[32];   // Spalte AG (Index 32)
    var apartmentCodeRaw = row[13];// Spalte N (Index 13)

    // Nur ausführen, wenn Spalte S LEER ist UND Spalte AG LEER ist
    if (lodgifyStatus && lodgifyStatus.toString().trim() !== "") return;
    if (minStayStatus && minStayStatus.toString().trim() !== "") return;

    var apartmentCode = apartmentCodeRaw ? apartmentCodeRaw.toString().toLowerCase().trim() : "";
    if (!apartmentCode) return;

    var apartment = apartmentMap[apartmentCode];
    if (!apartment) return;

    // Raten-Änderung auf 1 Nacht an Lodgify senden
    var success = executeMinStayChange(apartment.property_id, apartment.room_type_id, 1);

    if (success) {
      sheetRE.getRange(rowIndex, 33).setValue("auf 1 geändert"); // Nur Spalte AG beschreiben
      Logger.log("ERFOLG: Zeile " + rowIndex + " | Property " + apartment.property_id + " min_stay auf 1 gesetzt.");
    } else {
      Logger.log("FEHLER: Zeile " + rowIndex + " | Raten-Update auf 1 fehlgeschlagen.");
    }
  });
}

/**
 * Sendet den cURL-äquivalenten POST Request an /v1/rates/savewithoutavailability
 */
function executeMinStayChange(propertyId, roomTypeId, minStayValue) {
  var props = PropertiesService.getScriptProperties();
  var loKey = props.getProperty("loKey");
  var fixieUrl = props.getProperty("fixieUrl");

  if (!loKey || !fixieUrl) throw new Error("Script Properties 'loKey' oder 'fixieUrl' fehlen.");

  var targetUrl = "https://api.lodgify.com/v1/rates/savewithoutavailability";
  var authCredentials = fixieUrl.split("@")[0].replace("http://", "").replace("https://", "");

  var payload = {
    "property_id": parseInt(propertyId, 10),
    "room_type_id": parseInt(roomTypeId, 10),
    "rates": [
      {
        "price_per_day": 199,
        "min_stay": minStayValue,
        "is_default": true
      }
    ]
  };

  var options = {
    "method": "post",
    "headers": {
      "X-ApiKey": loKey,
      "Proxy-Authorization": "Basic " + Utilities.base64Encode(authCredentials),
      "Accept": "application/json",
      "Content-Type": "application/json-patch+json"
    },
    "muteHttpExceptions": true,
    "payload": JSON.stringify(payload)
  };

  try {
    var response = UrlFetchApp.fetch(targetUrl, options);
    var code = response.getResponseCode();
    return code >= 200 && code < 300;
  } catch (e) {
    Logger.log("Exception bei executeMinStayChange: " + e.toString());
    return false;
  }
}

/**
 * Holt das Apartment-Mapping aus Tab 'd'
 */
function getApartmentMapping() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetD = ss.getSheetByName("d");

  if (!sheetD) throw new Error("Das Tabellenblatt 'd' wurde nicht gefunden.");

  var lastRow = sheetD.getLastRow();
  if (lastRow < 2) return {};

  var data = sheetD.getRange(2, 1, lastRow - 1, 4).getValues();
  var mapping = {};

  data.forEach(function(row) {
    var propertyId = row[0];
    var roomTypeId = row[1];
    var link = row[3];

    if (!propertyId || !roomTypeId || !link) return;

    var match = link.toString().trim().match(/\/a\/([^\/\?]+)/i);
    if (!match || !match[1]) return;

    var apartmentCode = match[1].toLowerCase().trim();
    mapping[apartmentCode] = {
      property_id: parseInt(propertyId, 10),
      room_type_id: parseInt(roomTypeId, 10)
    };
  });

  return mapping;
}