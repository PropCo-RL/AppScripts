// ============================================================================
// LODGIFY PREIS & MINSTAY UPDATE (TAB 'cal_advertise')
// ============================================================================

/**
 * HAUPTFUNKTION: Aktualisiert Preise & MinStay in Lodgify basierend auf Sheet 'cal_advertise'
 * Bedingung: Es werden NUR Zeilen verarbeitet, bei denen A, B, C UND D befüllt sind!
 * Schreibt den Status/Fehler direkt in Spalte E.
 */
function updateLodgifyPricesFromAdvertise() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("cal_advertise");

  if (!sheet) throw new Error("Das Tabellenblatt 'cal_advertise' wurde nicht gefunden.");

  const propertyMap = getLodgifyPropertyMapAdv();

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("Keine Daten in 'cal_advertise' vorhanden.");
    return;
  }

  // Header für Spalte E prüfen/setzen
  sheet.getRange(1, 5).setValue("Status / Rückmeldung").setFontWeight("bold");

  // Liest Spalte A (Name), B (Preis), C (Tage), D (MinStay)
  const range = sheet.getRange(2, 1, lastRow - 1, 4);
  const values = range.getValues();

  const today = new Date();
  today.setHours(12, 0, 0, 0);

  values.forEach((row, index) => {
    const rowIndex = index + 2;
    const propertyName = sanitizeStringAdv(row[0]); // Spalte A
    const newPriceRaw = row[1];                     // Spalte B
    const daysCountRaw = row[2];                    // Spalte C
    const minStayRaw = row[3];                      // Spalte D

    // Wenn alle 4 Felder komplett leer sind -> leise überspringen
    if (!propertyName && newPriceRaw === "" && daysCountRaw === "" && minStayRaw === "") {
      return;
    }

    // STRIKTE PRÜFUNG: Wenn mindestens eines der 4 Felder fehlt -> Fehler in Spalte E eintragen
    if (!propertyName || newPriceRaw === "" || newPriceRaw === null || newPriceRaw === undefined || !daysCountRaw || minStayRaw === "" || minStayRaw === null || minStayRaw === undefined) {
      sheet.getRange(rowIndex, 5).setValue("ÜBERSPRUNGEN: Spalten A-D müssen komplett befüllt sein");
      Logger.log(`[Zeile ${rowIndex}] Übersprungen. Unvollständige Eingaben (A: '${propertyName}', B: '${newPriceRaw}', C: '${daysCountRaw}', D: '${minStayRaw}')`);
      return;
    }

    // Werte parsen & validieren
    const promoPrice = parseFloat(newPriceRaw.toString().replace(/[^0-9\.]/g, ''));
    const daysCount = parseInt(daysCountRaw, 10);
    const promoMinStay = parseInt(minStayRaw, 10);

    if (isNaN(promoPrice) || promoPrice <= 0) {
      sheet.getRange(rowIndex, 5).setValue("ERROR: Ungültiger Preis (Spalte B)");
      return;
    }

    if (isNaN(daysCount) || daysCount <= 0) {
      sheet.getRange(rowIndex, 5).setValue("ERROR: Ungültige Tage (Spalte C)");
      return;
    }

    if (isNaN(promoMinStay) || promoMinStay < 1) {
      sheet.getRange(rowIndex, 5).setValue("ERROR: Ungültiger MinStay (Spalte D)");
      return;
    }

    // Enddatum berechnen: Heute + X Tage aus Spalte C
    const endDate = new Date(today.getTime());
    endDate.setDate(endDate.getDate() + daysCount);

    const mappedObj = propertyMap[propertyName.toLowerCase()];
    if (!mappedObj || !mappedObj.propertyId || !mappedObj.roomTypeId) {
      sheet.getRange(rowIndex, 5).setValue("ERROR: Mapping in 'd' fehlt");
      Logger.log(`[Zeile ${rowIndex}] Property '${propertyName}' in Blatt 'd' nicht gefunden.`);
      return;
    }

    // API-Call ausführen
    const result = sendLodgifyPriceUpdateAdv(
      mappedObj.propertyId, 
      mappedObj.roomTypeId, 
      mappedObj.defaultRate,
      mappedObj.defaultMinStay,
      promoPrice, 
      promoMinStay,
      today, 
      endDate
    );

    if (result.success) {
      const startStr = formatDateGermanAdv(today);
      const endStr = formatDateGermanAdv(endDate);
      const successMsg = `ERFOLG (${startStr} - ${endStr})`;
      
      sheet.getRange(rowIndex, 5).setValue(successMsg);
      Logger.log(`[Zeile ${rowIndex}] ${successMsg}`);
    } else {
      const errorMsg = `ERROR ${result.code}: ${result.raw.substring(0, 150)}`;
      sheet.getRange(rowIndex, 5).setValue(errorMsg);
      Logger.log(`[Zeile ${rowIndex}] ${errorMsg}`);
    }
  });

  // Änderungen im Sheet sofort sichern
  SpreadsheetApp.flush();
}

// ============================================================================
// HILFSFUNKTIONEN FÜR PREIS-UPDATE
// ============================================================================

function getLodgifyPropertyMapAdv() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetD = ss.getSheetByName("d");

  if (!sheetD) throw new Error("Das Tabellenblatt 'd' wurde nicht gefunden.");

  const lastRow = sheetD.getLastRow();
  if (lastRow < 2) return {};

  const data = sheetD.getRange(2, 1, lastRow - 1, 6).getValues();
  const map = {};

  data.forEach(row => {
    const propertyId = row[0];   // Spalte A
    const roomTypeId = row[1];   // Spalte B
    const propertyName = sanitizeStringAdv(row[2]); // Spalte C
    const defaultRateRaw = row[4];   // Spalte E
    const defaultMinStayRaw = row[5];// Spalte F

    if (propertyName && propertyId && roomTypeId) {
      const defaultRate = parseFloat(defaultRateRaw.toString().replace(/[^0-9\.]/g, '')) || 199;
      const defaultMinStay = parseInt(defaultMinStayRaw, 10) || 1;

      map[propertyName.toLowerCase()] = {
        propertyId: parseInt(propertyId, 10),
        roomTypeId: parseInt(roomTypeId, 10),
        defaultRate: defaultRate,
        defaultMinStay: defaultMinStay
      };
    }
  });

  return map;
}

function sanitizeStringAdv(val) {
  if (val === null || val === undefined) return "";
  return val.toString().replace(/[\r\n]+/g, " ").trim();
}

function formatDateLodgifyAdv(dateObj) {
  return Utilities.formatDate(dateObj, "Europe/Berlin", "yyyy-MM-dd");
}

function formatDateGermanAdv(dateObj) {
  return Utilities.formatDate(dateObj, "Europe/Berlin", "dd.MM.yyyy");
}

/**
 * Sendet das Update an Lodgify
 */
function sendLodgifyPriceUpdateAdv(propertyId, roomTypeId, defaultRate, defaultMinStay, promoPrice, promoMinStay, startDateObj, endDateObj) {
  const props = PropertiesService.getScriptProperties();
  const loKey = props.getProperty("loKey");
  const fixieUrl = props.getProperty("fixieUrl");

  if (!loKey) throw new Error("Script Property 'loKey' fehlt!");

  const targetUrl = "https://api.lodgify.com/v1/rates/savewithoutavailability";

  // Enddatum um +1 Tag erhöhen, da end_date exklusiv behandelt wird
  const lodgifyEndDate = new Date(endDateObj.getTime());
  lodgifyEndDate.setDate(lodgifyEndDate.getDate() + 1);

  const payload = {
    "property_id": propertyId,
    "room_type_id": roomTypeId,
    "rates": [
      {
        "price_per_day": defaultRate,
        "min_stay": defaultMinStay,
        "is_default": true
      },
      {
        "start_date": formatDateLodgifyAdv(startDateObj),
        "end_date": formatDateLodgifyAdv(lodgifyEndDate),
        "price_per_day": promoPrice,
        "min_stay": promoMinStay,
        "is_default": false
      }
    ]
  };

  const headers = {
    "X-ApiKey": loKey,
    "Accept": "application/json",
    "Content-Type": "application/json-patch+json"
  };

  if (fixieUrl) {
    const authCredentials = fixieUrl.split("@")[0].replace("http://", "").replace("https://", "");
    headers["Proxy-Authorization"] = "Basic " + Utilities.base64Encode(authCredentials);
  }

  const options = {
    "method": "post",
    "headers": headers,
    "muteHttpExceptions": true,
    "payload": JSON.stringify(payload)
  };

  try {
    const response = UrlFetchApp.fetch(targetUrl, options);
    const code = response.getResponseCode();
    const rawText = response.getContentText();

    return {
      success: code >= 200 && code < 300,
      code: code,
      raw: rawText
    };
  } catch (e) {
    return {
      success: false,
      code: "EXCEPTION",
      raw: e.toString()
    };
  }
}