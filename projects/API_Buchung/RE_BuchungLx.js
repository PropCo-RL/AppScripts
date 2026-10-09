// ============================================================================
// LEXOFFICE RECHNUNGSERSTELLUNG (EINZELRECHNUNGEN)
// ============================================================================

/**
 * HAUPTFUNKTION: Erstellt Einzelrechnungen in Lexoffice für offene Zeilen in Sheet 'RE'
 */
function BBB_createLexofficeInvoices() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetRE = ss.getSheetByName("RE");

  if (!sheetRE) throw new Error("Das Tabellenblatt 'RE' wurde nicht gefunden.");

  const CLEANING_ARTICLE_UUID = "fb05d566-9590-4741-a361-916b49fe2499";
  const apartmentMap = getSingleInvoiceApartmentMapping();
  
  const lastRow = sheetRE.getLastRow();
  if (lastRow < 2) return;

  // Liest ab Zeile 2 von Spalte A (1) bis Spalte AC (29)
  const range = sheetRE.getRange(2, 1, lastRow - 1, 29);
  const values = range.getValues();

  values.forEach((row, index) => {
    const rowIndex = index + 2;
    const lxInvoiceId = row[19]; // Spalte T (Index 19)

    // Nur ausführen, wenn Spalte T noch LEER ist
    if (lxInvoiceId && lxInvoiceId.toString().trim() !== "") return;

    const firma = sanitizeStringSingle(row[1]);         // B: Firma
    const name = sanitizeStringSingle(row[2]);          // C: Name
    const strasse = sanitizeStringSingle(row[3]);       // D: Straße
    const plz = sanitizeStringSingle(row[4]);           // E: PLZ
    let stadt = sanitizeStringSingle(row[5]);           // F: Stadt
    const land = sanitizeStringSingle(row[6]);          // G: Land
    const vatNo = sanitizeStringSingle(row[7]);         // H: VAT / USt-IdNr
    
    const arrivalRaw = row[10];                   // K: Anreise Buchung
    const departureRaw = row[11];                 // L: Abreise Buchung
    const daysRaw = row[12];                      // M: Abrechnungs-Tage
    const apartmentCodeRaw = row[13];            // N: Apartment Kürzel
    const pricePerNightRaw = row[14];            // O: Preis pro Nacht
    const customStartRaw = row[27];               // AB: Manueller Rechnungs-Start
    const customEndRaw = row[28];                 // AC: Manuelles Rechnungs-Ende

    const apartmentCode = extractApartmentCodeSingle(apartmentCodeRaw);

    if (!arrivalRaw || !departureRaw || !apartmentCode) return;

    // 1. ADRESSE PREPARIEREN
    if (land && land.toLowerCase() !== "deutschland" && land.toLowerCase() !== "de") {
      stadt = `${stadt} (${land})`.trim();
    }

    let recipientName = firma || "Gast";
    
    let supplementParts = [];
    if (vatNo) {
      supplementParts.push("USt-IdNr.: " + vatNo);
    }
    let recipientSupplement = supplementParts.join(", ");

    // 2. APARTMENT / ARTIKEL MAPPING
    const apartment = apartmentMap[apartmentCode];
    if (!apartment) {
      sheetRE.getRange(rowIndex, 20).setValue("ERROR: Apartment '" + apartmentCode + "' in Tab 'd' nicht gefunden");
      return;
    }

    // 3. DYNAMISCHER RECHNUNGSZEITRAUM (AB/AC > M > K/L)
    const dateRange = determineInvoiceDatesSingle(arrivalRaw, departureRaw, daysRaw, customStartRaw, customEndRaw);

    if (!dateRange.start || !dateRange.end) {
      sheetRE.getRange(rowIndex, 20).setValue("ERROR: Ungültige Datumsangaben für Rechnung");
      return;
    }

    const voucherDate = formatToLexofficeDateSingle(new Date());
    const shippingDate = formatToLexofficeDateSingle(dateRange.start);
    const shippingEndDate = formatToLexofficeDateSingle(dateRange.end);

    const nightsCount = calculateNightsBetweenSingle(dateRange.start, dateRange.end);
    if (nightsCount <= 0) {
      sheetRE.getRange(rowIndex, 20).setValue("ERROR: Rechnungs-Ende muss nach Start liegen");
      return;
    }

    const netUnitPrice = parseFloat(pricePerNightRaw.toString().replace(/[^0-9\.]/g, '')) || 0.00;

    const startStr = formatGermanDateSingle(dateRange.start);
    const endStr = formatGermanDateSingle(dateRange.end);

    // 4. MEHRZEILIGER ADRESSTEXT DIREKT UND SICHER ERZEUGT
    const addressText = `Adresse der Wohnung:\n${apartment.formattedAddress}`;

    // Pos 1: Monteurwohnung
    const item1 = {
      "type": "custom",
      "name": "Monteurwohnung",
      "description": `${addressText}\n\nZeitraum: ${startStr} bis ${endStr}`,
      "quantity": nightsCount,
      "unitName": "Tag(e)",
      "unitPrice": {
        "currency": "EUR",
        "netAmount": netUnitPrice,
        "taxRatePercentage": 7
      },
      "discountPercentage": 0
    };

    // Pos 2: Reinigung
    const item2 = {
      "id": CLEANING_ARTICLE_UUID,
      "type": "service",
      "name": "Reinigung",
      "quantity": 1,
      "unitName": "Stück",
      "unitPrice": {
        "currency": "EUR",
        "netAmount": netUnitPrice,
        "taxRatePercentage": 7
      },
      "discountPercentage": 0
    };

    const lineItems = [item1, item2];

    // 5. PAYLOAD AUFBAUEN
    const addressObj = {
      "name": recipientName,
      "countryCode": "DE"
    };

    if (recipientSupplement) addressObj["supplement"] = recipientSupplement;
    if (strasse) addressObj["street"] = strasse;
    if (stadt) addressObj["city"] = stadt;
    if (plz) addressObj["zip"] = plz;

    const payload = {
      "archived": false,
      "voucherDate": voucherDate,
      "address": addressObj,
      "lineItems": lineItems,
      "totalPrice": { "currency": "EUR" },
      "taxConditions": { "taxType": "net" },
      "shippingConditions": {
        "shippingDate": shippingDate,
        "shippingEndDate": shippingEndDate,
        "shippingType": "serviceperiod"
      },
      "paymentConditions": {
        "paymentTermLabel": "Zahlungsziel: Vorkasse vor Beginn der Reservierung oder Verlängerung. \n\nBitte beachten Sie, dass Buchungen und Verlängerungen ohne Zahlungseingang nach drei Werktagen automatisch storniert werden können. Die Stornierung ist ggf. kostenpflichtig.\n\nEine allgemeine Anleitung für den Check-in finden Sie im Anhang.",
        "paymentTermDuration": 0,
        "paymentDiscountConditions": { "discountPercentage": 0, "discountRange": 0 }
      },
      "title": "Rechnung",
      "introduction": "",
      "remark": "Vielen Dank für Ihre Zusammenarbeit mit L8 Street!"
    };

    // 6. API-CALL
    try {
      Logger.log(`Sende Lexoffice Invoice Request für Zeile ${rowIndex}: ${recipientName}`);
      const response = sendLexofficeRequestSingle("/v1/invoices?finalize=true", payload);

      if (response.statusCode === 201 || response.statusCode === 200) {
        const invoiceId = response.body.id;
        sheetRE.getRange(rowIndex, 20).setValue(invoiceId);        // Spalte T (20)
        sheetRE.getRange(rowIndex, 21).setValue("Fertig");         // Spalte U (21)
        
        sheetRE.getRange(rowIndex, 28).setValue(dateRange.start);  // Spalte AB (28)
        sheetRE.getRange(rowIndex, 29).setValue(dateRange.end);    // Spalte AC (29)

        Logger.log(`ERFOLG! Lexoffice Invoice ID für Zeile ${rowIndex}: ${invoiceId}`);
      } else {
        const errorMsg = response.raw || JSON.stringify(response.body);
        sheetRE.getRange(rowIndex, 20).setValue(`ERROR ${response.statusCode}: ${errorMsg.substring(0, 150)}`);
      }
    } catch (err) {
      sheetRE.getRange(rowIndex, 20).setValue(`ERROR: ${err.message}`);
    }
  });
}

// ============================================================================
// HILFSFUNKTIONEN FÜR EINZELRECHNUNGEN
// ============================================================================

function extractApartmentCodeSingle(rawVal) {
  if (!rawVal) return "";
  const str = rawVal.toString().toLowerCase().trim();
  const parts = str.split(/\s+/);
  return parts[parts.length - 1];
}

function determineInvoiceDatesSingle(arrivalRaw, departureRaw, daysRaw, customStartRaw, customEndRaw) {
  const arrivalParsed = parseDateObjectSingle(arrivalRaw);
  const departureParsed = parseDateObjectSingle(departureRaw);
  const customStartParsed = parseDateObjectSingle(customStartRaw);
  const customEndParsed = parseDateObjectSingle(customEndRaw);

  if (customStartParsed && customEndParsed) {
    return { start: customStartParsed, end: customEndParsed };
  }

  if (arrivalParsed && daysRaw !== null && daysRaw !== undefined && daysRaw.toString().trim() !== "") {
    const cleanDaysStr = daysRaw.toString().replace(/[^0-9]/g, '');
    const days = parseInt(cleanDaysStr, 10);
    
    if (!isNaN(days) && days > 0) {
      const computedEnd = new Date(arrivalParsed.getTime());
      computedEnd.setDate(computedEnd.getDate() + days);
      return { start: arrivalParsed, end: computedEnd };
    }
  }

  return { start: arrivalParsed, end: departureParsed };
}

function getSingleInvoiceApartmentMapping() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetD = ss.getSheetByName("d");

  if (!sheetD) throw new Error("Das Tabellenblatt 'd' wurde nicht gefunden.");

  const lastRow = sheetD.getLastRow();
  if (lastRow < 2) return {};

  const data = sheetD.getRange(2, 1, lastRow - 1, 13).getValues();
  const mapping = {};

  data.forEach(row => {
    const articleName = sanitizeStringSingle(row[2]); // C: Name
    const link = sanitizeStringSingle(row[3]);        // D: Link
    
    const strasseVal = sanitizeStringSingle(row[7]);    // H: Straße
    const nrVal = sanitizeStringSingle(row[8]);         // I: Nr
    const plzVal = sanitizeStringSingle(row[9]);        // J: PLZ
    const stadtVal = sanitizeStringSingle(row[10]);     // K: Stadt
    const etageVal = sanitizeStringSingle(row[11]);     // L: Etage
    const seiteVal = sanitizeStringSingle(row[12]);     // M: Seite

    if (!link) return;

    const match = link.match(/\/a\/([^\/\?]+)/i);
    if (!match || !match[1]) return;

    const apartmentCode = match[1].toLowerCase().trim();

    let line1 = [strasseVal, nrVal].filter(Boolean).join(" ").trim();
    let line2 = [etageVal, seiteVal].filter(Boolean).join(" ").trim();
    let line3 = [plzVal, stadtVal].filter(Boolean).join(" ").trim();

    let lines = [];
    if (line1) lines.push(line1);
    if (line2) lines.push(line2);
    if (line3) lines.push(line3);

    const formattedAddress = lines.length > 0 ? lines.join("\n") : (articleName || "Unterkunft");

    mapping[apartmentCode] = {
      articleName: articleName || "Unterkunft",
      formattedAddress: formattedAddress
    };
  });

  return mapping;
}

function sanitizeStringSingle(val) {
  if (val === null || val === undefined) return "";
  return val.toString().replace(/[\r\n]+/g, " ").trim();
}

function parseDateObjectSingle(dateVal) {
  if (!dateVal) return null;
  if (dateVal instanceof Date) return dateVal;

  const str = dateVal.toString().trim();
  const parts = str.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (parts) {
    return new Date(parseInt(parts[3], 10), parseInt(parts[2], 10) - 1, parseInt(parts[1], 10), 12, 0, 0);
  }

  const d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
}

function calculateNightsBetweenSingle(d1, d2) {
  const t1 = new Date(d1.getFullYear(), d1.getMonth(), d1.getDate(), 12, 0, 0).getTime();
  const t2 = new Date(d2.getFullYear(), d2.getMonth(), d2.getDate(), 12, 0, 0).getTime();
  const diff = t2 - t1;
  return Math.round(diff / (1000 * 60 * 60 * 24));
}

function formatGermanDateSingle(dateObj) {
  if (!dateObj) return "";
  return Utilities.formatDate(dateObj, "Europe/Berlin", "dd.MM.yyyy");
}

function formatToLexofficeDateSingle(dateObj) {
  const dFinal = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate(), 12, 0, 0);
  const datePart = Utilities.formatDate(dFinal, "Europe/Berlin", "yyyy-MM-dd");
  const timezoneOffset = Utilities.formatDate(dFinal, "Europe/Berlin", "XXX");
  return `${datePart}T00:00:00.000${timezoneOffset}`;
}

function sendLexofficeRequestSingle(endpoint, payload) {
  const props = PropertiesService.getScriptProperties();
  const lxKey = props.getProperty("lxKey");

  if (!lxKey) throw new Error("Script Property 'lxKey' fehlt!");

  const targetUrl = "https://api.lexoffice.io" + endpoint;

  const options = {
    "method": "post",
    "headers": {
      "Authorization": "Bearer " + lxKey,
      "Content-Type": "application/json",
      "Accept": "application/json"
    },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  const response = UrlFetchApp.fetch(targetUrl, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  let responseBody;
  try {
    responseBody = responseText ? JSON.parse(responseText) : null;
  } catch (e) {
    responseBody = responseText;
  }

  return {
    statusCode: responseCode,
    body: responseBody,
    raw: responseText
  };
}