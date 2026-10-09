// ============================================================================
// LEXOFFICE SAMMELRECHNUNGEN (MULTI-BOOKINGS)
// ============================================================================

/**
 * HAUPTFUNKTION: Erstellt Sammelrechnungen für Multi-Bookings (Spalte AI)
 */
function BBB_createLexofficeMultiInvoices() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  ["RE", "VL"].forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    processMultiInvoicesForSheet(sheet);
  });
}

function processMultiInvoicesForSheet(sheet) {
  const CLEANING_ARTICLE_UUID = "fb05d566-9590-4741-a361-916b49fe2499";
  const apartmentMap = getMultiInvoiceApartmentMapping();
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // Liest von Spalte A (1) bis Spalte AI (35)
  const range = sheet.getRange(2, 1, lastRow - 1, 35);
  const values = range.getValues();

  // 1. DUPLETTEN / GRUPPEN FINDEN (Gleicher Code in Spalte AI)
  const groupMap = {};

  values.forEach((row, index) => {
    const rowIndex = index + 2;
    const lxInvoiceId = row[19]; // Spalte T (Index 19)
    const multiCode = row[34] ? row[34].toString().trim() : ""; // Spalte AI (Index 34)

    if ((!lxInvoiceId || lxInvoiceId.toString().trim() === "") && multiCode !== "") {
      if (!groupMap[multiCode]) {
        groupMap[multiCode] = [];
      }
      groupMap[multiCode].push({
        rowIndex: rowIndex,
        row: row
      });
    }
  });

  // 2. JEDE GRUPPE VERARBEITEN
  Object.keys(groupMap).forEach(multiCode => {
    const groupEntries = groupMap[multiCode];
    if (groupEntries.length === 0) return;

    const firstRow = groupEntries[0].row;

    const firma = sanitizeStringMulti(firstRow[1]);         // B
    const name = sanitizeStringMulti(firstRow[2]);          // C
    const strasse = sanitizeStringMulti(firstRow[3]);       // D
    const plz = sanitizeStringMulti(firstRow[4]);           // E
    let stadt = sanitizeStringMulti(firstRow[5]);           // F
    const land = sanitizeStringMulti(firstRow[6]);          // G
    const vatNo = sanitizeStringMulti(firstRow[7]);         // H

    if (land && land.toLowerCase() !== "deutschland" && land.toLowerCase() !== "de") {
      stadt = `${stadt} (${land})`.trim();
    }

    let recipientName = firma || "Gast";
    let supplementParts = [];
    if (vatNo) supplementParts.push("USt-IdNr.: " + vatNo);
    let recipientSupplement = supplementParts.join(", ");

    const lineItems = [];
    let minArrival = null;
    let maxDeparture = null;
    let errorInGroup = false;

    // 3. ALLE APARTMENTS DER GRUPPE ERFASSEN
    groupEntries.forEach(entry => {
      const row = entry.row;
      const rowIndex = entry.rowIndex;

      const arrivalRaw = row[10];        // K: Anreise
      const departureRaw = row[11];      // L: Abreise
      const daysRaw = row[12];           // M: Tage (z.B. 14)
      const apartmentCodeRaw = row[13]; // N: Kürzel
      const pricePerNightRaw = row[14]; // O: Preis pro Nacht
      const customStartRaw = row[27];    // AB: Rechnungs-Start
      const customEndRaw = row[28];      // AC: Rechnungs-Ende

      const apartmentCode = extractApartmentCodeMulti(apartmentCodeRaw);
      const apartment = apartmentMap[apartmentCode];

      if (!apartment) {
        sheet.getRange(rowIndex, 20).setValue("ERROR: Apartment-Mapping in 'd' fehlt (" + apartmentCode + ")");
        errorInGroup = true;
        return;
      }

      // Rechnungszeitraum bestimmen (Priorität: AB/AC > M > K/L)
      const dateRange = determineInvoiceDatesMulti(arrivalRaw, departureRaw, daysRaw, customStartRaw, customEndRaw);

      if (!dateRange.start || !dateRange.end) {
        sheet.getRange(rowIndex, 20).setValue("ERROR: Ungültiges Datum");
        errorInGroup = true;
        return;
      }

      if (!minArrival || dateRange.start < minArrival) minArrival = dateRange.start;
      if (!maxDeparture || dateRange.end > maxDeparture) maxDeparture = dateRange.end;

      const nightsCount = calculateNightsBetweenMulti(dateRange.start, dateRange.end);
      const netUnitPrice = parseFloat(pricePerNightRaw.toString().replace(/[^0-9\.]/g, '')) || 0.00;

      const startStr = formatGermanDateMulti(dateRange.start);
      const endStr = formatGermanDateMulti(dateRange.end);

      // Position 1: Übernachtung mit mehrzeiligem Adressblock
      lineItems.push({
        "type": "custom",
        "name": "Monteurwohnung",
        "description": `Adresse der Wohnung:\n${apartment.formattedAddress}\n\nZeitraum: ${startStr} bis ${endStr}`,
        "quantity": nightsCount,
        "unitName": "Tag(e)",
        "unitPrice": {
          "currency": "EUR",
          "netAmount": netUnitPrice,
          "taxRatePercentage": 7
        },
        "discountPercentage": 0
      });

      // Position 2: Reinigung (sauber ohne Adressanhang)
      lineItems.push({
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
      });

      // Rechnungszeitraum ins Sheet eintragen (Spalten AB und AC)
      sheet.getRange(rowIndex, 28).setValue(dateRange.start); // AB
      sheet.getRange(rowIndex, 29).setValue(dateRange.end);   // AC
    });

    if (errorInGroup || lineItems.length === 0) {
      Logger.log(`Achtung: Multi-Rechnung '${multiCode}' wegen Fehlern abgebrochen.`);
      return;
    }

    // 4. PAYLOAD ERSTELLEN
    const voucherDate = formatToLexofficeDateMulti(new Date());
    const shippingDate = formatToLexofficeDateMulti(minArrival);
    const shippingEndDate = formatToLexofficeDateMulti(maxDeparture);

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

    // 5. API-CALL & RECHNUNGS-ID EINTRAGEN
    try {
      Logger.log(`Sende Multi-Invoice Request für Code '${multiCode}' (${groupEntries.length} Zeilen)`);
      const response = sendLexofficeRequestMulti("/v1/invoices?finalize=true", payload);

      if (response.statusCode === 201 || response.statusCode === 200) {
        const invoiceId = response.body.id;

        groupEntries.forEach(entry => {
          const rIndex = entry.rowIndex;
          sheet.getRange(rIndex, 20).setValue(invoiceId); // Spalte T
          sheet.getRange(rIndex, 21).setValue("Fertig");  // Spalte U
        });

        Logger.log(`ERFOLG! Sammelrechnung ${invoiceId} für Code '${multiCode}' erstellt.`);
      } else {
        const errorMsg = response.raw || JSON.stringify(response.body);
        groupEntries.forEach(entry => {
          sheet.getRange(entry.rowIndex, 20).setValue(`ERROR ${response.statusCode}: ${errorMsg.substring(0, 150)}`);
        });
      }
    } catch (err) {
      groupEntries.forEach(entry => {
        sheet.getRange(entry.rowIndex, 20).setValue(`ERROR: ${err.message}`);
      });
    }
  });
}

// ============================================================================
// HILFSFUNKTIONEN FÜR MULTI-RECHNUNGEN (ISOLIERT MIT Multi-SUFFIX)
// ============================================================================

function extractApartmentCodeMulti(rawVal) {
  if (!rawVal) return "";
  const str = rawVal.toString().toLowerCase().trim();
  const parts = str.split(/\s+/);
  return parts[parts.length - 1];
}

function determineInvoiceDatesMulti(arrivalRaw, departureRaw, daysRaw, customStartRaw, customEndRaw) {
  const arrivalParsed = parseDateObjectMulti(arrivalRaw);
  const departureParsed = parseDateObjectMulti(departureRaw);
  const customStartParsed = parseDateObjectMulti(customStartRaw);
  const customEndParsed = parseDateObjectMulti(customEndRaw);

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

function getMultiInvoiceApartmentMapping() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetD = ss.getSheetByName("d");

  if (!sheetD) throw new Error("Das Tabellenblatt 'd' wurde nicht gefunden.");

  const lastRow = sheetD.getLastRow();
  if (lastRow < 2) return {};

  const data = sheetD.getRange(2, 1, lastRow - 1, 13).getValues();
  const mapping = {};

  data.forEach(row => {
    const articleName = sanitizeStringMulti(row[2]); // C: Name
    const link = sanitizeStringMulti(row[3]);        // D: Link
    
    const strasseVal = sanitizeStringMulti(row[7]);    // H: Straße
    const nrVal = sanitizeStringMulti(row[8]);         // I: Nr
    const plzVal = sanitizeStringMulti(row[9]);        // J: PLZ
    const stadtVal = sanitizeStringMulti(row[10]);     // K: Stadt
    const etageVal = sanitizeStringMulti(row[11]);     // L: Etage
    const seiteVal = sanitizeStringMulti(row[12]);     // M: Seite

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

function sanitizeStringMulti(val) {
  if (val === null || val === undefined) return "";
  return val.toString().replace(/[\r\n]+/g, " ").trim();
}

function parseDateObjectMulti(dateVal) {
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

function calculateNightsBetweenMulti(d1, d2) {
  const t1 = new Date(d1.getFullYear(), d1.getMonth(), d1.getDate(), 12, 0, 0).getTime();
  const t2 = new Date(d2.getFullYear(), d2.getMonth(), d2.getDate(), 12, 0, 0).getTime();
  const diff = t2 - t1;
  return Math.round(diff / (1000 * 60 * 60 * 24));
}

function formatGermanDateMulti(dateObj) {
  if (!dateObj) return "";
  return Utilities.formatDate(dateObj, "Europe/Berlin", "dd.MM.yyyy");
}

function formatToLexofficeDateMulti(dateObj) {
  const dFinal = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate(), 12, 0, 0);
  const datePart = Utilities.formatDate(dFinal, "Europe/Berlin", "yyyy-MM-dd");
  const timezoneOffset = Utilities.formatDate(dFinal, "Europe/Berlin", "XXX");
  return `${datePart}T00:00:00.000${timezoneOffset}`;
}

function sendLexofficeRequestMulti(endpoint, payload) {
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