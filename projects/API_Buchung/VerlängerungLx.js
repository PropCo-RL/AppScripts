// ============================================================================
// LEXOFFICE RECHNUNGSERSTELLUNG - VERLÄNGERUNGEN (TAB "VL")
// ============================================================================

/**
 * HAUPTFUNKTION: Erstellt Verlängerungs-Rechnungen in Lexoffice für das Blatt 'VL'
 */
function BBB_createLexofficeInvoicesVL() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetVL = ss.getSheetByName("VL");
  const sheetRE = ss.getSheetByName("RE");

  if (!sheetVL) throw new Error("Das Tabellenblatt 'VL' wurde nicht gefunden.");

  const apartmentMap = getLexofficeApartmentMappingVL();
  
  const lastRow = sheetVL.getLastRow();
  if (lastRow < 2) return;

  // Max-Spalte ermitteln (Sicherstellen, dass mindestens bis Spalte AC / 29 gelesen wird)
  const maxColsVL = Math.max(sheetVL.getLastColumn(), 29);
  const rangeVL = sheetVL.getRange(2, 1, lastRow - 1, maxColsVL);
  const valuesVL = rangeVL.getValues();

  // Daten aus RE holen
  let valuesRE = [];
  if (sheetRE && sheetRE.getLastRow() >= 2) {
    const maxColsRE = Math.max(sheetRE.getLastColumn(), 29);
    valuesRE = sheetRE.getRange(2, 1, sheetRE.getLastRow() - 1, maxColsRE).getValues();
  }

  valuesVL.forEach((row, index) => {
    const rowIndex = index + 2;
    const lxInvoiceId = row[19]; // Spalte T (Index 19)

    // Nur ausführen, wenn Spalte T noch LEER ist
    if (lxInvoiceId && lxInvoiceId.toString().trim() !== "") return;

    const firma = sanitizeStringVL(row[1]);         // B: Firma
    const name = sanitizeStringVL(row[2]);          // C: Name
    const strasse = sanitizeStringVL(row[3]);       // D: Straße
    const plz = sanitizeStringVL(row[4]);           // E: PLZ
    let stadt = sanitizeStringVL(row[5]);           // F: Stadt
    const land = sanitizeStringVL(row[6]);          // G: Land
    const vatNo = sanitizeStringVL(row[7]);         // H: VAT / USt-IdNr
    const origArrivalRaw = row[10];                // K: Anreise (Ursprung)
    const newDepartureRaw = row[11];               // L: Abreise (Neu)
    const cycleRaw = row[12];                      // M: Abrechnungs-Zyklus
    const apartmentCodeRaw = row[13];              // N: Apartment Kürzel
    
    // Preise
    const priceNightRaw = row[14];                 // O: Preis pro Nacht
    const priceMonthRaw = row[16];                 // Q: Preis pro Monat
    
    const lodgifyId = row[18] ? row[18].toString().trim() : ""; // S: Lodgify ID

    // Manuelle Datumsangaben in AB (Index 27) und AC (Index 28)
    const manualStartRaw = row[27];                // AB: Lx RE Start
    const manualEndRaw = row[28];                  // AC: Lx RE Ende

    const apartmentCode = extractApartmentCodeVL(apartmentCodeRaw);
    if (!apartmentCode) return;

    // 1. START- UND ENDDATUM BESTIMMEN (AB & AC haben Vorrang!)
    let billingStartDate = parseDateObjectVL(manualStartRaw);
    let billingEndDate = parseDateObjectVL(manualEndRaw);

    // Falls AB leer -> Automatisch ermitteln (prüft AC in vorherigen Zeilen von VL & RE)
    if (!billingStartDate) {
      billingStartDate = getPreviousDepartureDateVL(valuesVL, valuesRE, index, lodgifyId, origArrivalRaw);
    }

    // Falls AC leer -> Über Zyklus/Enddatum ermitteln
    if (!billingEndDate) {
      const targetEndDate = parseDateObjectVL(newDepartureRaw);
      if (!targetEndDate || !billingStartDate) {
        sheetVL.getRange(rowIndex, 20).setValue("ERROR: Ungültiges Datum");
        return;
      }

      billingEndDate = new Date(targetEndDate.getTime());
      
      const cycleDays = parseCycleDaysVL(cycleRaw);
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
      return;
    }

    const nightsCount = calculateNightsBetweenVL(billingStartDate, billingEndDate);
    if (nightsCount <= 0) {
      sheetVL.getRange(rowIndex, 20).setValue("ERROR: Keine zu berechnenden Nächte");
      return;
    }

    // 2. ADRESSE & GAST
    if (land && land.toLowerCase() !== "deutschland" && land.toLowerCase() !== "de") {
      stadt = `${stadt} (${land})`.trim();
    }

    let recipientName = firma || "Gast";
    if (!firma && name) recipientName = name;
    
    let supplementParts = [];
    if (vatNo) supplementParts.push("USt-IdNr.: " + vatNo);
    let recipientSupplement = supplementParts.join(", ");

    // 3. APARTMENT MAPPING
    const apartment = apartmentMap[apartmentCode];
    if (!apartment) {
      sheetVL.getRange(rowIndex, 20).setValue("ERROR: Apartment '" + apartmentCode + "' in Tab 'd' nicht gefunden");
      return;
    }

    const voucherDate = formatToLexofficeDateVL(new Date());
    const shippingDate = formatToLexofficeDateVL(billingStartDate);
    const shippingEndDate = formatToLexofficeDateVL(billingEndDate);

    // 4. LINE ITEMS ERSTELLEN (Mit strukturiertem Adress- und Datumsformat)
    const lineItems = buildLineItemsVL(
      apartment.formattedAddress, 
      nightsCount, 
      priceNightRaw, 
      priceMonthRaw,
      billingStartDate,
      billingEndDate
    );

    if (!lineItems || lineItems.length === 0) {
      sheetVL.getRange(rowIndex, 20).setValue("ERROR: Kein Preis angegeben (Spalte O oder Q)");
      return;
    }

    // 5. PAYLOAD ERSTELLEN
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
      "totalPrice": {
        "currency": "EUR"
      },
      "taxConditions": {
        "taxType": "net"
      },
      "shippingConditions": {
        "shippingDate": shippingDate,
        "shippingEndDate": shippingEndDate,
        "shippingType": "serviceperiod"
      },
      "paymentConditions": {
        "paymentTermLabel": "Zahlungsziel: Vorkasse vor Beginn der Verlängerung.\n\nBitte beachten Sie, dass Verlängerungen ohne Zahlungseingang nach drei Werktagen automatisch storniert werden können. Die Stornierung ist ggf. kostenpflichtig.",
        "paymentTermDuration": 0,
        "paymentDiscountConditions": {
          "discountPercentage": 0,
          "discountRange": 0
        }
      },
      "title": "Rechnung",
      "introduction": "",
      "remark": "Vielen Dank für die Verlängerung Ihres Aufenthalts bei L8 Street!"
    };

    // 6. API-CALL & DATEN EINTRAGEN
    try {
      Logger.log(`Sende Lexoffice Request für Zeile ${rowIndex}: ${recipientName} (${nightsCount} Nächte)`);
      const response = sendLexofficeRequestVL("/v1/invoices?finalize=true", payload);

      if (response.statusCode === 201 || response.statusCode === 200) {
        const invoiceId = response.body.id;
        sheetVL.getRange(rowIndex, 20).setValue(invoiceId); // Spalte T (20): Invoice ID
        sheetVL.getRange(rowIndex, 21).setValue("Fertig");   // Spalte U (21): Status
        
        // Trägt den tatsächlich genutzten Leistungszeitraum in AB und AC ein
        sheetVL.getRange(rowIndex, 28).setValue(billingStartDate); // Spalte AB (28)
        sheetVL.getRange(rowIndex, 29).setValue(billingEndDate);   // Spalte AC (29)

        Logger.log(`ERFOLG! Lexoffice Invoice ID für Zeile ${rowIndex}: ${invoiceId}`);
      } else {
        const errorMsg = response.raw || JSON.stringify(response.body);
        sheetVL.getRange(rowIndex, 20).setValue(`ERROR ${response.statusCode}: ${errorMsg.substring(0, 150)}`);
      }
    } catch (err) {
      sheetVL.getRange(rowIndex, 20).setValue(`ERROR: ${err.message}`);
    }
  });
}

/**
 * HILFSFUNKTION: Findet das vorherige Abreisedatum für dieselbe Lodgify-ID
 */
function getPreviousDepartureDateVL(valuesVL, valuesRE, currentIndex, lodgifyId, defaultArrival) {
  if (lodgifyId) {
    // 1. Suche oberhalb in 'VL' nach derselben Lodgify ID
    for (let i = currentIndex - 1; i >= 0; i--) {
      const rowLodgifyId = valuesVL[i][18] ? valuesVL[i][18].toString().trim() : "";
      if (rowLodgifyId === lodgifyId) {
        const prevAC = parseDateObjectVL(valuesVL[i][28]);
        if (prevAC) return prevAC;

        const prevL = parseDateObjectVL(valuesVL[i][11]);
        if (prevL) return prevL;
      }
    }

    // 2. Falls in 'VL' nicht vorhanden, suche im Sheet 'RE'
    for (let j = valuesRE.length - 1; j >= 0; j--) {
      const reLodgifyId = valuesRE[j][18] ? valuesRE[j][18].toString().trim() : "";
      if (reLodgifyId === lodgifyId) {
        const reAC = valuesRE[j].length > 28 ? parseDateObjectVL(valuesRE[j][28]) : null;
        if (reAC) return reAC;

        const reL = parseDateObjectVL(valuesRE[j][11]);
        if (reL) return reL;
      }
    }
  }

  return parseDateObjectVL(defaultArrival);
}

// ============================================================================
// HILFSFUNKTIONEN FÜR VERLÄNGERUNGEN (ISOLIERT MIT VL-SUFFIX)
// ============================================================================

function extractApartmentCodeVL(rawVal) {
  if (!rawVal) return "";
  const str = rawVal.toString().toLowerCase().trim();
  const parts = str.split(/\s+/);
  return parts[parts.length - 1];
}

function parseCycleDaysVL(cycleRaw) {
  if (!cycleRaw) return 0;
  const str = cycleRaw.toString().toLowerCase().trim();
  
  const num = parseInt(str, 10);
  if (!isNaN(num) && num > 0) {
    return num;
  }

  if (str.includes("wöchentlich") || str.includes("woechentlich")) return 7;
  if (str.includes("monatlich") || str.includes("monat")) return 30;

  return 0;
}

function buildLineItemsVL(formattedAddress, nightsCount, priceNight, priceMonth, startDate, endDate) {
  const parseAmount = (val) => {
    if (!val) return 0;
    const clean = val.toString().replace(/[^0-9,\.]/g, '').replace(',', '.');
    return parseFloat(clean) || 0;
  };

  const pNight = parseAmount(priceNight);
  const pMonth = parseAmount(priceMonth);

  const startStr = formatGermanDateVL(startDate);
  const endStr = formatGermanDateVL(endDate);
  const descriptionText = `Adresse der Wohnung:\n${formattedAddress}\n\nZeitraum: ${startStr} bis ${endStr}`;

  // 1. MONATSPREIS (Spalte Q) HAT VORRANG
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

  // 2. PREIS PRO NACHT (Spalte O)
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

function getLexofficeApartmentMappingVL() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetD = ss.getSheetByName("d");

  if (!sheetD) throw new Error("Das Tabellenblatt 'd' wurde nicht gefunden.");

  const lastRow = sheetD.getLastRow();
  if (lastRow < 2) return {};

  const data = sheetD.getRange(2, 1, lastRow - 1, 13).getValues();
  const mapping = {};

  data.forEach(row => {
    const articleName = sanitizeStringVL(row[2]); // C: Name
    const link = sanitizeStringVL(row[3]);        // D: Link
    
    const strasseVal = sanitizeStringVL(row[7]);    // H: Straße
    const nrVal = sanitizeStringVL(row[8]);         // I: Nr
    const plzVal = sanitizeStringVL(row[9]);        // J: PLZ
    const stadtVal = sanitizeStringVL(row[10]);     // K: Stadt
    const etageVal = sanitizeStringVL(row[11]);     // L: Etage
    const seiteVal = sanitizeStringVL(row[12]);     // M: Seite

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

function sanitizeStringVL(val) {
  if (val === null || val === undefined) return "";
  return val.toString().replace(/[\r\n]+/g, " ").trim();
}

function parseDateObjectVL(dateVal) {
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

function calculateNightsBetweenVL(d1, d2) {
  const t1 = new Date(d1.getFullYear(), d1.getMonth(), d1.getDate(), 12, 0, 0).getTime();
  const t2 = new Date(d2.getFullYear(), d2.getMonth(), d2.getDate(), 12, 0, 0).getTime();
  const diff = t2 - t1;
  return Math.round(diff / (1000 * 60 * 60 * 24));
}

function formatGermanDateVL(dateObj) {
  if (!dateObj) return "";
  return Utilities.formatDate(dateObj, "Europe/Berlin", "dd.MM.yyyy");
}

function formatToLexofficeDateVL(dateObj) {
  const dFinal = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate(), 12, 0, 0);
  const datePart = Utilities.formatDate(dFinal, "Europe/Berlin", "yyyy-MM-dd");
  const timezoneOffset = Utilities.formatDate(dFinal, "Europe/Berlin", "XXX");
  return `${datePart}T00:00:00.000${timezoneOffset}`;
}

function sendLexofficeRequestVL(endpoint, payload) {
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