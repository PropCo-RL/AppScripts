// ============================================================================
// LEXOFFICE SAMMELRECHNUNGEN - VERLÄNGERUNGEN (TAB "VL", MULTI-BOOKINGS)
// ============================================================================

/**
 * HAUPTFUNKTION: Erstellt Sammelrechnungen für Verlängerungen (Spalte AI im Blatt 'VL')
 */
function BBB_createLexofficeMultiInvoicesVL() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetVL = ss.getSheetByName("VL");
  const sheetRE = ss.getSheetByName("RE");

  if (!sheetVL) throw new Error("Das Tabellenblatt 'VL' wurde nicht gefunden.");

  processMultiInvoicesVLForSheet(sheetVL, sheetRE);
}

function processMultiInvoicesVLForSheet(sheetVL, sheetRE) {
  const apartmentMap = getMultiVLApartmentMapping();
  
  const lastRow = sheetVL.getLastRow();
  if (lastRow < 2) return;

  // Liest von Spalte A (1) bis Spalte AI (35)
  const maxColsVL = Math.max(sheetVL.getLastColumn(), 35);
  const rangeVL = sheetVL.getRange(2, 1, lastRow - 1, maxColsVL);
  const valuesVL = rangeVL.getValues();

  // RE-Daten für vorherige Abreisedaten sichern
  let valuesRE = [];
  if (sheetRE && sheetRE.getLastRow() >= 2) {
    const maxColsRE = Math.max(sheetRE.getLastColumn(), 29);
    valuesRE = sheetRE.getRange(2, 1, sheetRE.getLastRow() - 1, maxColsRE).getValues();
  }

  // 1. DUPLETTEN / GRUPPEN FINDEN (Spalte T leer & Spalte AI befüllt)
  const groupMap = {};

  valuesVL.forEach((row, index) => {
    const rowIndex = index + 2;
    const lxInvoiceId = row[19]; // Spalte T (Index 19)
    const multiCode = row[34] ? row[34].toString().trim() : ""; // Spalte AI (Index 34)

    if ((!lxInvoiceId || lxInvoiceId.toString().trim() === "") && multiCode !== "") {
      if (!groupMap[multiCode]) {
        groupMap[multiCode] = [];
      }
      groupMap[multiCode].push({
        rowIndex: rowIndex,
        row: row,
        indexInValues: index
      });
    }
  });

  // 2. JEDE GRUPPE VERARBEITEN
  Object.keys(groupMap).forEach(multiCode => {
    const groupEntries = groupMap[multiCode];
    if (groupEntries.length === 0) return;

    const firstRow = groupEntries[0].row;

    const firma = sanitizeStringMultiVL(firstRow[1]);         // B: Firma
    const name = sanitizeStringMultiVL(firstRow[2]);          // C: Name
    const strasse = sanitizeStringMultiVL(firstRow[3]);       // D: Straße
    const plz = sanitizeStringMultiVL(firstRow[4]);           // E: PLZ
    let stadt = sanitizeStringMultiVL(firstRow[5]);           // F: Stadt
    const land = sanitizeStringMultiVL(firstRow[6]);          // G: Land
    const vatNo = sanitizeStringMultiVL(firstRow[7]);         // H: VAT / USt-IdNr

    if (land && land.toLowerCase() !== "deutschland" && land.toLowerCase() !== "de") {
      stadt = `${stadt} (${land})`.trim();
    }

    let recipientName = firma || "Gast";
    if (!firma && name) recipientName = name;

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
      const indexInValues = entry.indexInValues;

      const origArrivalRaw = row[10];                // K: Anreise (Ursprung)
      const newDepartureRaw = row[11];               // L: Abreise (Neu)
      const cycleRaw = row[12];                      // M: Zyklus
      const apartmentCodeRaw = row[13];              // N: Kürzel
      const priceNightRaw = row[14];                 // O: Preis Nacht
      const priceMonthRaw = row[16];                 // Q: Preis Monat
      const lodgifyId = row[18] ? row[18].toString().trim() : ""; // S: Lodgify ID
      const manualStartRaw = row[27];                // AB: Lx RE Start
      const manualEndRaw = row[28];                  // AC: Lx RE Ende

      const apartmentCode = extractApartmentCodeMultiVL(apartmentCodeRaw);
      const apartment = apartmentMap[apartmentCode];

      if (!apartment) {
        sheetVL.getRange(rowIndex, 20).setValue("ERROR: Apartment-Mapping in 'd' fehlt (" + apartmentCode + ")");
        errorInGroup = true;
        return;
      }

      // Rechnungszeitraum bestimmen (AB/AC Vorrang, sonst automatische Ketten-Ermittlung)
      let billingStartDate = parseDateObjectMultiVL(manualStartRaw);
      let billingEndDate = parseDateObjectMultiVL(manualEndRaw);

      if (!billingStartDate) {
        billingStartDate = getPreviousDepartureDateMultiVL(valuesVL, valuesRE, indexInValues, lodgifyId, origArrivalRaw);
      }

      if (!billingEndDate) {
        const targetEndDate = parseDateObjectMultiVL(newDepartureRaw);
        if (!targetEndDate || !billingStartDate) {
          sheetVL.getRange(rowIndex, 20).setValue("ERROR: Ungültiges Datum");
          errorInGroup = true;
          return;
        }

        billingEndDate = new Date(targetEndDate.getTime());
        const cycleDays = parseCycleDaysMultiVL(cycleRaw);
        if (cycleDays > 0) {
          let maxCycleEnd = new Date(billingStartDate.getTime());
          maxCycleEnd.setDate(maxCycleEnd.getDate() + cycleDays);
          if (maxCycleEnd < targetEndDate) {
            billingEndDate = maxCycleEnd;
          }
        }
      }

      if (!billingStartDate || !billingEndDate) {
        sheetVL.getRange(rowIndex, 20).setValue("ERROR: Datum fehlt");
        errorInGroup = true;
        return;
      }

      const nightsCount = calculateNightsBetweenMultiVL(billingStartDate, billingEndDate);
      if (nightsCount <= 0) {
        sheetVL.getRange(rowIndex, 20).setValue("ERROR: Keine Nächte zu berechnen");
        errorInGroup = true;
        return;
      }

      if (!minArrival || billingStartDate < minArrival) minArrival = billingStartDate;
      if (!maxDeparture || billingEndDate > maxDeparture) maxDeparture = billingEndDate;

      // Positionen für dieses Apartment im neuen Format aufbauen
      const apartmentItems = buildLineItemsMultiVL(
        apartment.formattedAddress,
        nightsCount,
        priceNightRaw,
        priceMonthRaw,
        billingStartDate,
        billingEndDate
      );

      if (!apartmentItems || apartmentItems.length === 0) {
        sheetVL.getRange(rowIndex, 20).setValue("ERROR: Kein Preis angegeben (Spalte O/Q)");
        errorInGroup = true;
        return;
      }

      apartmentItems.forEach(item => lineItems.push(item));

      // 📌 GEÄNDERT: Formatiertes Datum in Spalte AB (28) und AC (29) schreiben
      sheetVL.getRange(rowIndex, 28).setValue(formatGermanDateMultiVL(billingStartDate)); // AB
      sheetVL.getRange(rowIndex, 29).setValue(formatGermanDateMultiVL(billingEndDate));   // AC
    });

    if (errorInGroup || lineItems.length === 0) {
      Logger.log(`Achtung: Multi-Verlängerung '${multiCode}' wegen Fehlern abgebrochen.`);
      return;
    }

    // 4. PAYLOAD ERSTELLEN
    const voucherDate = formatToLexofficeDateMultiVL(new Date());
    const shippingDate = formatToLexofficeDateMultiVL(minArrival);
    const shippingEndDate = formatToLexofficeDateMultiVL(maxDeparture);

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
        "paymentTermLabel": "Zahlungsziel: Vorkasse vor Beginn der Verlängerung.\n\nBitte beachten Sie, dass Verlängerungen ohne Zahlungseingang nach drei Werktagen automatisch storniert werden können. Die Stornierung ist ggf. kostenpflichtig.",
        "paymentTermDuration": 0,
        "paymentDiscountConditions": { "discountPercentage": 0, "discountRange": 0 }
      },
      "title": "Rechnung",
      "introduction": "",
      "remark": "Vielen Dank für die Verlängerung Ihres Aufenthals bei L8 Street!"
    };

    // 5. API-CALL & RECHNUNGS-ID EINTRAGEN
    try {
      Logger.log(`Sende Multi-Verlängerung Request für Code '${multiCode}' (${groupEntries.length} Zeilen)`);
      const response = sendLexofficeRequestMultiVL("/v1/invoices?finalize=true", payload);

      if (response.statusCode === 201 || response.statusCode === 200) {
        const invoiceId = response.body.id;

        groupEntries.forEach(entry => {
          const rIndex = entry.rowIndex;
          sheetVL.getRange(rIndex, 20).setValue(invoiceId); // Spalte T
          sheetVL.getRange(rIndex, 21).setValue("Fertig");  // Spalte U
        });

        Logger.log(`ERFOLG! Sammelverlängerung ${invoiceId} für Code '${multiCode}' erstellt.`);
      } else {
        const errorMsg = response.raw || JSON.stringify(response.body);
        groupEntries.forEach(entry => {
          sheetVL.getRange(entry.rowIndex, 20).setValue(`ERROR ${response.statusCode}: ${errorMsg.substring(0, 150)}`);
        });
      }
    } catch (err) {
      groupEntries.forEach(entry => {
        sheetVL.getRange(entry.rowIndex, 20).setValue(`ERROR: ${err.message}`);
      });
    }
  });
}

// ============================================================================
// HILFSFUNKTIONEN FÜR MULTI-VERLÄNGERUNGEN (ISOLIERT MIT MultiVL-SUFFIX)
// ============================================================================

function extractApartmentCodeMultiVL(rawVal) {
  if (!rawVal) return "";
  const str = rawVal.toString().toLowerCase().trim();
  const parts = str.split(/\s+/);
  return parts[parts.length - 1];
}

function parseCycleDaysMultiVL(cycleRaw) {
  if (!cycleRaw) return 0;
  const str = cycleRaw.toString().toLowerCase().trim();
  const num = parseInt(str, 10);
  if (!isNaN(num) && num > 0) return num;

  if (str.includes("wöchentlich") || str.includes("woechentlich")) return 7;
  if (str.includes("monatlich") || str.includes("monat")) return 30;

  return 0;
}

function getPreviousDepartureDateMultiVL(valuesVL, valuesRE, currentIndex, lodgifyId, defaultArrival) {
  if (lodgifyId) {
    for (let i = currentIndex - 1; i >= 0; i--) {
      const rowLodgifyId = valuesVL[i][18] ? valuesVL[i][18].toString().trim() : "";
      if (rowLodgifyId === lodgifyId) {
        const prevAC = parseDateObjectMultiVL(valuesVL[i][28]);
        if (prevAC) return prevAC;

        const prevL = parseDateObjectMultiVL(valuesVL[i][11]);
        if (prevL) return prevL;
      }
    }

    for (let j = valuesRE.length - 1; j >= 0; j--) {
      const reLodgifyId = valuesRE[j][18] ? valuesRE[j][18].toString().trim() : "";
      if (reLodgifyId === lodgifyId) {
        const reAC = valuesRE[j].length > 28 ? parseDateObjectMultiVL(valuesRE[j][28]) : null;
        if (reAC) return reAC;

        const reL = parseDateObjectMultiVL(valuesRE[j][11]);
        if (reL) return reL;
      }
    }
  }

  return parseDateObjectMultiVL(defaultArrival);
}

function buildLineItemsMultiVL(formattedAddress, nightsCount, priceNight, priceMonth, startDate, endDate) {
  const parseAmount = (val) => {
    if (!val) return 0;
    const clean = val.toString().replace(/[^0-9,\.]/g, '').replace(',', '.');
    return parseFloat(clean) || 0;
  };

  const pNight = parseAmount(priceNight);
  const pMonth = parseAmount(priceMonth);

  const startStr = formatGermanDateMultiVL(startDate);
  const endStr = formatGermanDateMultiVL(endDate);
  const descriptionText = `Adresse der Wohnung:\n${formattedAddress}\n\nZeitraum: ${startStr} bis ${endStr}`;

  if (pMonth > 0) {
    if (nightsCount >= 28 && nightsCount <= 31) {
      return [{
        "type": "custom",
        "name": "Monteurwohnung (Monatspauschale Verlängerung)",
        "description": descriptionText,
        "quantity": 1,
        "unitName": "Monat",
        "unitPrice": { "currency": "EUR", "netAmount": pMonth, "taxRatePercentage": 7 },
        "discountPercentage": 0
      }];
    } else {
      const dailyPriceFromMonth = Math.round((pMonth / 30) * 100) / 100;
      return [{
        "type": "custom",
        "name": "Monteurwohnung (Verlängerung)",
        "description": descriptionText,
        "quantity": nightsCount,
        "unitName": "Tag(e)",
        "unitPrice": { "currency": "EUR", "netAmount": dailyPriceFromMonth, "taxRatePercentage": 7 },
        "discountPercentage": 0
      }];
    }
  }

  if (pNight > 0) {
    return [{
      "type": "custom",
      "name": "Monteurwohnung (Verlängerung)",
      "description": descriptionText,
      "quantity": nightsCount,
      "unitName": "Tag(e)",
      "unitPrice": { "currency": "EUR", "netAmount": pNight, "taxRatePercentage": 7 },
      "discountPercentage": 0
    }];
  }

  return [];
}

function getMultiVLApartmentMapping() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetD = ss.getSheetByName("d");

  if (!sheetD) throw new Error("Das Tabellenblatt 'd' wurde nicht gefunden.");

  const lastRow = sheetD.getLastRow();
  if (lastRow < 2) return {};

  const data = sheetD.getRange(2, 1, lastRow - 1, 13).getValues();
  const mapping = {};

  data.forEach(row => {
    const articleName = sanitizeStringMultiVL(row[2]); // C: Name
    const link = sanitizeStringMultiVL(row[3]);        // D: Link
    
    const strasseVal = sanitizeStringMultiVL(row[7]);    // H: Straße
    const nrVal = sanitizeStringMultiVL(row[8]);         // I: Nr
    const plzVal = sanitizeStringMultiVL(row[9]);        // J: PLZ
    const stadtVal = sanitizeStringMultiVL(row[10]);     // K: Stadt
    const etageVal = sanitizeStringMultiVL(row[11]);     // L: Etage
    const seiteVal = sanitizeStringMultiVL(row[12]);     // M: Seite

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

function sanitizeStringMultiVL(val) {
  if (val === null || val === undefined) return "";
  return val.toString().replace(/[\r\n]+/g, " ").trim();
}

function parseDateObjectMultiVL(dateVal) {
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

function calculateNightsBetweenMultiVL(d1, d2) {
  const t1 = new Date(d1.getFullYear(), d1.getMonth(), d1.getDate(), 12, 0, 0).getTime();
  const t2 = new Date(d2.getFullYear(), d2.getMonth(), d2.getDate(), 12, 0, 0).getTime();
  const diff = t2 - t1;
  return Math.round(diff / (1000 * 60 * 60 * 24));
}

function formatGermanDateMultiVL(dateObj) {
  if (!dateObj) return "";
  return Utilities.formatDate(dateObj, "Europe/Berlin", "dd.MM.yyyy");
}

function formatToLexofficeDateMultiVL(dateObj) {
  const dFinal = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate(), 12, 0, 0);
  const datePart = Utilities.formatDate(dFinal, "Europe/Berlin", "yyyy-MM-dd");
  const timezoneOffset = Utilities.formatDate(dFinal, "Europe/Berlin", "XXX");
  return `${datePart}T00:00:00.000${timezoneOffset}`;
}

function sendLexofficeRequestMultiVL(endpoint, payload) {
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