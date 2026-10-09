// ============================================================================
// LODGIFY BUCHUNGEN AUTOMATISIERUNG
// ============================================================================

/**
 * HAUPTFUNKTION (steht alphabetisch ganz oben)
 */
function AAA_startLodgifyBooking() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetRE = ss.getSheetByName("RE");

  if (!sheetRE) throw new Error("Das Tabellenblatt 'RE' wurde nicht gefunden.");

  var apartmentMap = getApartmentMapping();
  var lastRow = sheetRE.getLastRow();
  if (lastRow < 2) {
    Logger.log("Keine Daten im Sheet 'RE' vorhanden.");
    return;
  }

  var values = sheetRE.getRange(2, 1, lastRow - 1, 19).getValues();

  values.forEach(function(row, index) {
    var rowIndex = index + 2;
    var lodgifyStatus = row[18]; // Spalte S (Lodgify Status)

    // Nur ausführen, wenn Spalte S LEER ist
    if (lodgifyStatus && lodgifyStatus.toString().trim() !== "") return;

    var firma = row[1];          // B
    var name = row[2];           // C
    var email = row[8];          // I
    var phoneRaw = row[0];       // A (Nummer / Phone)
    var arrivalRaw = row[10];    // K
    var departureRaw = row[11];  // L
    var apartmentCodeRaw = row[13]; // N
    var totalNightRaw = row[14];  // O (Preis pro Nacht)
    var personNightRaw = row[15]; // P (Preis pro Person / Nacht)

    var apartmentCode = apartmentCodeRaw ? apartmentCodeRaw.toString().toLowerCase().trim() : "";

    // Falls Pflichtfelder fehlen, Rückmeldung ins Sheet
    if (!arrivalRaw || !departureRaw || !apartmentCode) {
      if (!apartmentCode && (arrivalRaw || departureRaw)) {
        sheetRE.getRange(rowIndex, 19).setValue("ERROR: Apartment (Spalte N) fehlt");
      }
      return;
    }

    // Gastname ermitteln
    var guestName = "Gast";
    if (firma && name) {
      guestName = firma.toString().trim() + " (" + name.toString().trim() + ")";
    } else if (firma) {
      guestName = firma.toString().trim();
    } else if (name) {
      guestName = name.toString().trim();
    }

    // Präzises Datum OHNE Zeitzonen-Shift
    var arrival = parseAndFormatDate(arrivalRaw);
    var departure = parseAndFormatDate(departureRaw);

    if (!arrival || !departure) {
      sheetRE.getRange(rowIndex, 19).setValue("ERROR: Ungültiges Datum");
      return;
    }

    var apartment = apartmentMap[apartmentCode];
    if (!apartment) {
      sheetRE.getRange(rowIndex, 19).setValue("ERROR: Mapping in 'd' fehlt");
      return;
    }

    // Telefonnummer bereinigen
    var cleanPhone = sanitizePhoneNumber(phoneRaw);

    // Personenzahl berechnen (z.B. 50€ / 25€ = 2 Pax)
    var peopleCount = 1;
    var priceNight = parseFloat(totalNightRaw.toString().replace(/[^0-9\.]/g, ''));
    var pricePerson = parseFloat(personNightRaw.toString().replace(/[^0-9\.]/g, ''));
    if (priceNight && pricePerson && pricePerson > 0) {
      peopleCount = Math.round(priceNight / pricePerson);
    }

    var payload = {
      "guest": {
        "name": guestName,
        "email": email ? email.toString().trim() : "",
        "phone": cleanPhone,
        "country_code": "DE"
      },
      "source_text": "api",
      "arrival": arrival,
      "departure": departure,
      "property_id": apartment.property_id,
      "status": "Booked",
      "rooms": [
        {
          "room_type_id": apartment.room_type_id,
          "people": peopleCount
        }
      ],
      "currency_code": "EUR"
    };

    try {
      var result = sendLodgifyRequest("/v1/reservation/booking", "POST", payload);

      if (result.statusCode === 201) {
        var bookingId = result.body;
        sheetRE.getRange(rowIndex, 19).setValue(bookingId !== null && bookingId !== undefined ? bookingId : "Done");
        Logger.log("ERFOLG! Zeile " + rowIndex + " | Booking ID: " + bookingId);
      } else {
        var errorText = result.raw || JSON.stringify(result.body);
        if (errorText.length > 300) errorText = errorText.substring(0, 300);
        sheetRE.getRange(rowIndex, 19).setValue("ERROR " + result.statusCode + ": " + errorText);
      }
    } catch (error) {
      sheetRE.getRange(rowIndex, 19).setValue("ERROR: " + error.message);
    }
  });
}

/**
 * Bereinigt Telefonnummern und repariert doppelte Vorwahlen (+4949... -> +49...)
 */
function sanitizePhoneNumber(phoneValue) {
  if (!phoneValue) return "";
  
  // Nur Zahlen behalten
  var cleaned = phoneValue.toString().replace(/[^0-9]/g, '');
  if (!cleaned) return "";

  // Repariere doppelte Vorwahlen (z. B. 4949173... -> 49173...)
  if (cleaned.startsWith("4949")) {
    cleaned = cleaned.substring(2);
  } else if (cleaned.startsWith("4141")) { // Für Schweiz
    cleaned = cleaned.substring(2);
  } else if (cleaned.startsWith("4343")) { // Für Österreich
    cleaned = cleaned.substring(2);
  }

  // Internationales Format mit Pluszeichen zurückgeben
  return "+" + cleaned;
}

/**
 * Robuste Datumskonvertierung OHNE Zeitzonen-Shift (verhindert 1-Tag-zu-früh-Fehler)
 */
function parseAndFormatDate(dateValue) {
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
    
    // Format DD.MM.YYYY
    var partsDE = str.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})/);
    if (partsDE) {
      dayStr = partsDE[1].length === 1 ? "0" + partsDE[1] : partsDE[1];
      monthStr = partsDE[2].length === 1 ? "0" + partsDE[2] : partsDE[2];
      year = partsDE[3];
    } else {
      // Format YYYY-MM-DD
      var partsISO = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
      if (partsISO) {
        year = partsISO[1];
        monthStr = partsISO[2].length === 1 ? "0" + partsISO[2] : partsISO[2];
        dayStr = partsISO[3].length === 1 ? "0" + partsISO[3] : partsISO[3];
      } else {
        var dTemp = new Date(str);
        if (isNaN(dTemp.getTime())) return "";
        year = dTemp.getFullYear();
        var mTemp = dTemp.getMonth() + 1;
        var dayTemp = dTemp.getDate();
        monthStr = mTemp < 10 ? "0" + mTemp : "" + mTemp;
        dayStr = dayTemp < 10 ? "0" + dayTemp : "" + dayTemp;
      }
    }
  }

  // Gibt reines YYYY-MM-DD mit Z-Offset zurück (Lodgify akzeptiert YYYY-MM-DDTHH:mm:ssZ)
  return year + "-" + monthStr + "-" + dayStr + "T00:00:00Z";
}

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

function sendLodgifyRequest(endpoint, method, payload) {
  var props = PropertiesService.getScriptProperties();
  var loKey = props.getProperty("loKey");
  var fixieUrl = props.getProperty("fixieUrl");

  if (!loKey) throw new Error("Script Property 'loKey' wurde nicht gefunden.");
  if (!fixieUrl) throw new Error("Script Property 'fixieUrl' wurde nicht gefunden.");
  if (!endpoint) throw new Error("Der Parameter 'endpoint' ist leer oder undefined.");

  method = method || "POST";
  var targetUrl = "https://api.lodgify.com" + endpoint;

  var authCredentials = fixieUrl
    .split("@")[0]
    .replace("http://", "")
    .replace("https://", "");

  var options = {
    method: method.toLowerCase(),
    headers: {
      "X-ApiKey": loKey,
      "Proxy-Authorization": "Basic " + Utilities.base64Encode(authCredentials),
      "Accept": "application/json",
      "Content-Type": "application/json-patch+json"
    },
    muteHttpExceptions: true
  };

  if (payload !== null && payload !== undefined) {
    options.payload = JSON.stringify(payload);
  }

  Logger.log("Sende Request an: " + targetUrl);
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