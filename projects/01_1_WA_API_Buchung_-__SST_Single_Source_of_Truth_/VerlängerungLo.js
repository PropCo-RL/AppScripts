// ============================================================================
// LODGIFY BUCHUNGEN UPDATE AUTOMATISIERUNG (TAB "VL")
// ============================================================================

/**
 * HAUPTFUNKTION für Updates (Sheet 'VL')
 */
function AAA_updateLodgifyBookingVL() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetVL = ss.getSheetByName("VL");

  if (!sheetVL) throw new Error("Das Tabellenblatt 'VL' wurde nicht gefunden.");

  var lastRow = sheetVL.getLastRow();
  if (lastRow < 2) {
    Logger.log("Keine Daten im Sheet 'VL' vorhanden.");
    return;
  }

  // Liest Daten ab Zeile 2 von Spalte A (1) bis Spalte X (24)
  var values = sheetVL.getRange(2, 1, lastRow - 1, 24).getValues();

  values.forEach(function(row, index) {
    var rowIndex = index + 2;
    
    var arrivalRaw = row[10];     // Spalte K (Anreise Verlängerung)
    var departureRaw = row[11];   // Spalte L (Abreise Neu)
    var bookingIdRaw = row[18];   // Spalte S (Lodgify ID)
    var updateStatus = row[23];   // Spalte X (Update Status)

    // Nur ausführen, wenn Spalte X LEER ist
    if (updateStatus && updateStatus.toString().trim() !== "") return;

    // Abbruch, falls keine Booking ID in Spalte S vorhanden ist
    if (!bookingIdRaw || bookingIdRaw.toString().trim() === "") {
      sheetVL.getRange(rowIndex, 24).setValue("ERROR: Keine Booking ID (Spalte S)");
      return;
    }

    var bookingId = parseInt(bookingIdRaw.toString().trim(), 10);

    // Pflichtfelder Datum prüfen
    if (!arrivalRaw || !departureRaw) {
      sheetVL.getRange(rowIndex, 24).setValue("ERROR: Anreise (K) oder Abreise (L) fehlt");
      return;
    }

    var arrival = formatDateOnlyVL(arrivalRaw);
    var departure = formatDateOnlyVL(departureRaw);

    if (!arrival || !departure) {
      sheetVL.getRange(rowIndex, 24).setValue("ERROR: Ungültiges Datum");
      return;
    }

    // ------------------------------------------------------------------
    // SCHRITT 1: VORAB-PRÜFUNG IN LODGIFY (Verhindert Error 666)
    // ------------------------------------------------------------------
    try {
      var currentBooking = getLodgifyBookingVL(bookingId);
      
      if (currentBooking && currentBooking.arrival && currentBooking.departure) {
        var currentArrival = formatDateOnlyVL(currentBooking.arrival);
        var currentDeparture = formatDateOnlyVL(currentBooking.departure);

        // Wenn die Daten in Lodgify bereits EXAKT mit dem Ziel übereinstimmen -> Erfolgreich abschließen
        if (currentArrival === arrival && currentDeparture === departure) {
          sheetVL.getRange(rowIndex, 24).setValue("Bereits aktualisiert (" + arrival + " - " + departure + ")");
          Logger.log("INFO: Zeile " + rowIndex + " | Booking ID " + bookingId + " war in Lodgify bereits auf dieses Datum gesetzt.");
          return;
        }
      }
    } catch (e) {
      Logger.log("Hinweis bei Vorab-Prüfung Booking ID " + bookingId + ": " + e.message);
      // Falls der GET fehlschlägt, wird normal mit dem PUT fortgefahren.
    }

    // ------------------------------------------------------------------
    // SCHRITT 2: PUT REQUEST (Falls Daten noch nicht übereinstimmen)
    // ------------------------------------------------------------------
    var payload = {
      "id": bookingId,
      "arrival": arrival,
      "departure": departure,
      "rooms": [],
      "status": "Booked"
    };

    var endpoint = "/v1/reservation/booking/" + bookingId;

    try {
      var result = sendLodgifyPutRequestVL(endpoint, payload);

      if (result.statusCode === 200 || result.statusCode === 204) {
        sheetVL.getRange(rowIndex, 24).setValue("Updated (" + arrival + " - " + departure + ")");
        Logger.log("ERFOLG! Zeile " + rowIndex + " | Booking ID " + bookingId + " aktualisiert.");
      } else {
        var errorText = result.raw || JSON.stringify(result.body);
        if (errorText.length > 300) errorText = errorText.substring(0, 300);
        sheetVL.getRange(rowIndex, 24).setValue("ERROR " + result.statusCode + ": " + errorText);
      }
    } catch (error) {
      sheetVL.getRange(rowIndex, 24).setValue("ERROR: " + error.message);
    }
  });
}

/**
 * Liest den aktuellen Status einer Buchung von Lodgify aus (GET-Abruf)
 */
function getLodgifyBookingVL(bookingId) {
  var props = PropertiesService.getScriptProperties();
  var loKey = props.getProperty("loKey");
  var fixieUrl = props.getProperty("fixieUrl");

  if (!loKey || !fixieUrl) return null;

  var targetUrl = "https://api.lodgify.com/v1/reservation/booking/" + bookingId;
  var authCredentials = fixieUrl.split("@")[0].replace("http://", "").replace("https://", "");

  var options = {
    method: "get",
    headers: {
      "X-ApiKey": loKey,
      "Proxy-Authorization": "Basic " + Utilities.base64Encode(authCredentials),
      "Accept": "application/json"
    },
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(targetUrl, options);
  if (response.getResponseCode() === 200) {
    return JSON.parse(response.getContentText());
  }
  return null;
}

/**
 * Hilfsfunktion für VL: Konvertiert ein Datum strikt in YYYY-MM-DD
 */
function formatDateOnlyVL(dateValue) {
  if (!dateValue) return "";
  
  var year, monthStr, dayStr;

  if (dateValue instanceof Date) {
    year = dateValue.getFullYear();
    var m = dateValue.getMonth() + 1;
    var d = dateValue.getDate();
    monthStr = m < 10 ? "0" + m : "" + m;
    dayStr = d < 10 ? "0" + d : "" + d;
  } else {
    var str = dateValue.toString().trim();
    var partsDE = str.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})/);
    if (partsDE) {
      dayStr = partsDE[1].length === 1 ? "0" + partsDE[1] : partsDE[1];
      monthStr = partsDE[2].length === 1 ? "0" + partsDE[2] : partsDE[2];
      year = partsDE[3];
    } else {
      var partsISO = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
      if (partsISO) {
        year = partsISO[1];
        monthStr = partsISO[2].length === 1 ? "0" + partsISO[2] : partsISO[2];
        dayStr = partsISO[3].length === 1 ? "0" + partsISO[3] : partsISO[3];
      } else {
        return "";
      }
    }
  }
  return year + "-" + monthStr + "-" + dayStr;
}

/**
 * HTTP PUT Request an Lodgify über den Fixie-Proxy für VL
 */
function sendLodgifyPutRequestVL(endpoint, payload) {
  var props = PropertiesService.getScriptProperties();
  var loKey = props.getProperty("loKey");
  var fixieUrl = props.getProperty("fixieUrl");

  if (!loKey) throw new Error("Script Property 'loKey' wurde nicht gefunden.");
  if (!fixieUrl) throw new Error("Script Property 'fixieUrl' wurde nicht gefunden.");

  var targetUrl = "https://api.lodgify.com" + endpoint;

  var authCredentials = fixieUrl
    .split("@")[0]
    .replace("http://", "")
    .replace("https://", "");

  var options = {
    method: "put",
    headers: {
      "X-ApiKey": loKey,
      "Proxy-Authorization": "Basic " + Utilities.base64Encode(authCredentials),
      "Accept": "application/json",
      "Content-Type": "application/json"
    },
    muteHttpExceptions: true,
    payload: JSON.stringify(payload)
  };

  Logger.log("Sende PUT Request an: " + targetUrl);
  var response = UrlFetchApp.fetch(targetUrl, options);
  var responseCode = response.getResponseCode();
  var responseText = response.getContentText();

  var responseBody;
  try {
    responseBody = responseText ? JSON.parse(responseText) : null;
  } catch (e) {
    responseBody = responseText;
  }

  return {
    success: responseCode >= 200 && responseCode < 300,
    statusCode: responseCode,
    body: responseBody,
    raw: responseText
  };
}