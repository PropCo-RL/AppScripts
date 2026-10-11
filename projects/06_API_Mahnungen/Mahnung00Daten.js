/**
 * Liest offene Rechnungen aus Lexoffice ab.
 * - Erfasst neu fällige Rechnungen der letzten 3 Tage (ohne Overdue-Einschränkung)
 * - Holt Start- & Ende des Leistungszeitraums DIREKT über /v1/invoices/{id} (shippingConditions)
 * - BERECHNET ÜBERFÄLLIGKEIT (Spalte E) STRIKT BASIEREND AUF DEM STARTDATUM (Spalte C / shippingDate)
 * - ERWEITERTER PORTAL-FILTER: Prüft remark, introduction, title, contactName UND paymentConditions
 * - Nutzt 'VL' & 'RE' als Fallback für fehlende Leistungszeiträume oder E-Mails
 */
function fillMahnwesenTabFromAPI() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheetName = "mah";

  let mahSheet = ss.getSheetByName(targetSheetName);
  if (!mahSheet) {
    mahSheet = ss.insertSheet(targetSheetName);
    Logger.log(`⚠️ Tabellenblatt '${targetSheetName}' wurde neu erstellt.`);
  }

  const token = PropertiesService.getScriptProperties().getProperty("lxKey");
  if (!token) {
    Logger.log("🚨 FEHLER: Script Property 'lxKey' fehlt!");
    SpreadsheetApp.getUi().alert("Fehler: API-Key 'lxKey' fehlt in den Script Properties!");
    return;
  }

  // Kopfzeilen setzen
  const headers = [
    "Firma / Name", 
    "Rechnungsnummer", 
    "Start Leistungszeitraum", 
    "Ende Leistungszeitraum", 
    "Überfällig seit (Tage)",
    "E-Mail-Adresse(n)",
    "Lexoffice ID"
  ];

  if (mahSheet.getLastRow() === 0) {
    mahSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    mahSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  }

  // Bestandsdaten aus 'mah' auslesen (InvoiceID -> Zeilennummer)
  const existingLastRow = mahSheet.getLastRow();
  const existingMap = new Map();

  if (existingLastRow >= 2) {
    const existingValues = mahSheet.getRange(2, 1, existingLastRow - 1, 7).getValues();
    existingValues.forEach((row, index) => {
      const existingId = row[6] ? row[6].toString().trim() : "";
      if (existingId) {
        existingMap.set(existingId, index + 2);
      }
    });
  }

  // Fallback-Maps aus 'VL' und 'RE' aufbauen
  const localEmailMap = buildLocalEmailMap(ss, ["VL", "RE"]);
  const localServicePeriodMap = buildLocalServicePeriodMap(ss, ["VL", "RE"]);

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // Zeitraum für Voucherlist Abruf (letzte 3 Tage)
  const threeDaysAgo = new Date();
  threeDaysAgo.setDate(today.getDate() - 3);
  const createdDateFromStr = Utilities.formatDate(threeDaysAgo, Session.getScriptTimeZone(), "yyyy-MM-dd");

  let page = 0;
  let totalPages = 1;
  const newRowsToInsert = [];
  let updatedCount = 0;
  let skippedPortalCount = 0;

  const emailCache = new Map();
  const invoiceDetailsCache = new Map();

  Logger.log(`Starte API-Abruf für offene Rechnungen (Erstellt seit ${createdDateFromStr})...`);

  while (page < totalPages) {
    const url = `https://api.lexware.io/v1/voucherlist?voucherType=salesinvoice,invoice&voucherStatus=open,sepadebit&createdDateFrom=${createdDateFromStr}&page=${page}&size=100`;

    const response = safeApiFetch(url, token);
    if (!response || response.getResponseCode() !== 200) {
      Logger.log(`🚨 API-Fehler bei Voucherlist Seite ${page}`);
      break;
    }

    const data = JSON.parse(response.getContentText());
    totalPages = data.totalPages || 1;
    const vouchers = data.content || [];

    vouchers.forEach(voucher => {
      const invoiceId = voucher.id || "";
      if (!invoiceId) return;

      // Rechnungsdetails abrufen (/v1/invoices/{id})
      const invoiceDetails = fetchInvoiceDetails(invoiceId, token, invoiceDetailsCache);
      if (!invoiceDetails) return;

      // ERWEITERTER PORTAL-FILTER: inkl. Zahlungsbedingungen (paymentConditions)
      const paymentTermText = (invoiceDetails.paymentConditions && invoiceDetails.paymentConditions.paymentTermLabel) 
        ? invoiceDetails.paymentConditions.paymentTermLabel 
        : "";

      const fullTextToCheck = [
        invoiceDetails.remark || "",
        invoiceDetails.introduction || "",
        invoiceDetails.title || "",
        voucher.contactName || "",
        paymentTermText
      ].join(" ").toLowerCase();

      if (
        fullTextToCheck.includes("airbnb") || 
        fullTextToCheck.includes("booking") || 
        fullTextToCheck.includes("buchungsportal")
      ) {
        skippedPortalCount++;
        return;
      }

      // 1. Leistungszeitraum direkt aus shippingConditions in Lexoffice auslesen
      let shippingDateStartFormatted = "-";
      let shippingDateEndFormatted = "-";
      let baseStartDate = null; // Stichtag für die Überfälligkeit (STARTDATUM)

      if (invoiceDetails.shippingConditions) {
        if (invoiceDetails.shippingConditions.shippingDate) {
          baseStartDate = new Date(invoiceDetails.shippingConditions.shippingDate);
          shippingDateStartFormatted = formatLexofficeDate(invoiceDetails.shippingConditions.shippingDate);
        }
        if (invoiceDetails.shippingConditions.shippingEndDate) {
          shippingDateEndFormatted = formatLexofficeDate(invoiceDetails.shippingConditions.shippingEndDate);
        } else if (shippingDateStartFormatted !== "-") {
          shippingDateEndFormatted = shippingDateStartFormatted;
        }
      }

      // Fallback Leistungszeitraum aus 'VL' & 'RE' (falls bei Lexoffice leer)
      if ((shippingDateStartFormatted === "-" || shippingDateEndFormatted === "-") && localServicePeriodMap.has(invoiceId)) {
        const localPeriod = localServicePeriodMap.get(invoiceId);
        if (shippingDateStartFormatted === "-") shippingDateStartFormatted = localPeriod.start;
        if (shippingDateEndFormatted === "-") shippingDateEndFormatted = localPeriod.end;

        if (!baseStartDate && localPeriod.start !== "-") {
          baseStartDate = parseDateString(localPeriod.start);
        }
      }

      // Fallback: Falls weder Lexoffice noch VL/RE ein Startdatum hatten -> dueDate als Reserve
      if (!baseStartDate && voucher.dueDate) {
        baseStartDate = new Date(voucher.dueDate);
      }

      // BERECHNUNG DER ÜBERFÄLLIGKEITSTAGE IN SPALTE E (Basierend auf STARTDATUM Spalte C)
      let overdueDays = 0;
      if (baseStartDate) {
        baseStartDate.setHours(0, 0, 0, 0);
        const diffTime = today.getTime() - baseStartDate.getTime();
        overdueDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));
      }

      // FALL 1: BEREITS BESTEHENDE ZEILE -> Spalte E (Tage) aktualisieren
      if (existingMap.has(invoiceId)) {
        const targetRow = existingMap.get(invoiceId);
        mahSheet.getRange(targetRow, 5).setValue(overdueDays < 0 ? 0 : overdueDays);
        updatedCount++;
        return;
      }

      // FALL 2: NEUE RECHNUNG (Wird immer aufgenommen, da kein Overdue-Filter)
      const companyName = voucher.contactName || (invoiceDetails.address ? invoiceDetails.address.name : "Unbekannter Kundenname");
      const voucherNumber = voucher.voucherNumber || invoiceDetails.voucherNumber || "OHNE NUMMER";

      // E-Mail abrufen (API + Fallback auf 'VL' & 'RE')
      let emailAddresses = "-";
      if (voucher.contactId) {
        if (emailCache.has(voucher.contactId)) {
          emailAddresses = emailCache.get(voucher.contactId);
        } else {
          emailAddresses = fetchContactEmails(voucher.contactId, token);
          emailCache.set(voucher.contactId, emailAddresses);
        }
      }

      if ((!emailAddresses || emailAddresses === "-") && localEmailMap.has(invoiceId)) {
        emailAddresses = localEmailMap.get(invoiceId);
      }

      newRowsToInsert.push([
        companyName,
        voucherNumber,
        shippingDateStartFormatted,
        shippingDateEndFormatted,
        overdueDays < 0 ? 0 : overdueDays,
        emailAddresses,
        invoiceId
      ]);

      existingMap.set(invoiceId, true);
    });

    page++;
  }

  // Neue Zeilen einfügen
  if (newRowsToInsert.length > 0) {
    const startRow = mahSheet.getLastRow() + 1;
    mahSheet.getRange(startRow, 1, newRowsToInsert.length, 7).setValues(newRowsToInsert);
    Logger.log(`✅ ${newRowsToInsert.length} neue Rechnung(en) in '${targetSheetName}' angefügt.`);
  }

  // ALLGEMEINER NACHPFLEGE-SCHRITT FÜR ALLE ZEILEN (Neu-Berechnung Tage STRIKT nach Startdatum in Spalte C)
  let retroFilledCount = 0;
  const finalLastRow = mahSheet.getLastRow();

  if (finalLastRow >= 2) {
    const allRows = mahSheet.getRange(2, 1, finalLastRow - 1, 7).getValues();

    allRows.forEach((row, index) => {
      const rowIndex = index + 2;
      const startC = row[2] ? row[2].toString().trim() : "";
      const endD = row[3] ? row[3].toString().trim() : "";
      const emailF = row[5] ? row[5].toString().trim() : "";
      const invoiceId = row[6] ? row[6].toString().trim() : "";

      if (!invoiceId) return;

      let changed = false;

      // 1. Rechnungsdetails von Lexoffice nachladen, wenn C oder D fehlen
      if ((!startC || startC === "" || startC === "-") || (!endD || endD === "" || endD === "-")) {
        const lxDetails = fetchInvoiceDetails(invoiceId, token, invoiceDetailsCache);

        if (lxDetails && lxDetails.shippingConditions) {
          if ((!startC || startC === "" || startC === "-") && lxDetails.shippingConditions.shippingDate) {
            const startDateStr = formatLexofficeDate(lxDetails.shippingConditions.shippingDate);
            if (startDateStr !== "-") {
              mahSheet.getRange(rowIndex, 3).setValue(startDateStr);
              changed = true;
            }
          }
          if ((!endD || endD === "" || endD === "-") && lxDetails.shippingConditions.shippingEndDate) {
            const endDateStr = formatLexofficeDate(lxDetails.shippingConditions.shippingEndDate);
            if (endDateStr !== "-") {
              mahSheet.getRange(rowIndex, 4).setValue(endDateStr);
              changed = true;
            }
          }
        }
      }

      // 2. Fallback auf 'VL' und 'RE' für Leistungszeitraum
      if (localServicePeriodMap.has(invoiceId)) {
        const localData = localServicePeriodMap.get(invoiceId);
        const currentC = mahSheet.getRange(rowIndex, 3).getValue().toString().trim();
        const currentD = mahSheet.getRange(rowIndex, 4).getValue().toString().trim();

        if ((!currentC || currentC === "" || currentC === "-") && localData.start !== "-") {
          mahSheet.getRange(rowIndex, 3).setValue(localData.start);
          changed = true;
        }
        if ((!currentD || currentD === "" || currentD === "-") && localData.end !== "-") {
          mahSheet.getRange(rowIndex, 4).setValue(localData.end);
          changed = true;
        }
      }

      // 3. EXAKTE NEU-BERECHNUNG SPALTE E (BASIEREND AUF SPALTE C: STARTDATUM)
      const finalStartC = mahSheet.getRange(rowIndex, 3).getValue().toString().trim();
      let baseStartDate = parseDateString(finalStartC);

      // Fallback falls Spalte C immer noch leer ist: Spalte D versuchen
      if (!baseStartDate) {
        const finalEndD = mahSheet.getRange(rowIndex, 4).getValue().toString().trim();
        baseStartDate = parseDateString(finalEndD);
      }

      if (baseStartDate) {
        baseStartDate.setHours(0, 0, 0, 0);
        const diffTime = today.getTime() - baseStartDate.getTime();
        const recalculatedOverdueDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));
        mahSheet.getRange(rowIndex, 5).setValue(recalculatedOverdueDays < 0 ? 0 : recalculatedOverdueDays);
      }

      // 4. E-Mail (Spalte F) nachpflegen
      if ((!emailF || emailF === "" || emailF === "-") && localEmailMap.has(invoiceId)) {
        const localEmail = localEmailMap.get(invoiceId);
        if (localEmail && localEmail !== "-") {
          mahSheet.getRange(rowIndex, 6).setValue(localEmail);
          changed = true;
        }
      }

      if (changed) retroFilledCount++;
    });
  }

  if (retroFilledCount > 0) {
    Logger.log(`🔄 Bei ${retroFilledCount} bestehenden Zeilen wurden Daten ergänzt / neu berechnet.`);
  }

  if (updatedCount > 0) {
    Logger.log(`🔄 Bei ${updatedCount} bereits vorhandenen Rechnung(en) wurden die Überfälligkeitstage aktualisiert.`);
  }

  if (skippedPortalCount > 0) {
    Logger.log(`🚫 Es wurden ${skippedPortalCount} Rechnung(en) wegen Airbnb/Booking.com-Fußnote ignoriert.`);
  }
}

// ============================================================================
// HILFSFUNKTIONEN
// ============================================================================

function parseDateString(dateStr) {
  if (!dateStr || dateStr === "-") return null;
  const match = dateStr.match(/(\d{2})\.(\d{2})\.(\d{4})/);
  if (match) {
    return new Date(parseInt(match[3], 10), parseInt(match[2], 10) - 1, parseInt(match[1], 10));
  }
  return null;
}

function fetchInvoiceDetails(invoiceId, token, cache) {
  if (cache.has(invoiceId)) return cache.get(invoiceId);

  try {
    const url = `https://api.lexware.io/v1/invoices/${invoiceId}`;
    const response = safeApiFetch(url, token);

    if (response && response.getResponseCode() === 200) {
      const invoiceData = JSON.parse(response.getContentText());
      cache.set(invoiceId, invoiceData);
      return invoiceData;
    }
  } catch (e) {
    Logger.log(`Fehler beim Abrufen der Rechnungsdetails für ID ${invoiceId}: ${e.message}`);
  }

  cache.set(invoiceId, null);
  return null;
}

function formatLexofficeDate(isoStr) {
  if (!isoStr) return "-";
  try {
    const dateObj = new Date(isoStr);
    if (isNaN(dateObj.getTime())) return "-";
    return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "dd.MM.yyyy");
  } catch (e) {
    return "-";
  }
}

function buildLocalServicePeriodMap(spreadsheet, sheetNames) {
  const map = new Map();

  sheetNames.forEach(name => {
    const sheet = spreadsheet.getSheetByName(name);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const values = sheet.getRange(2, 1, lastRow - 1, 29).getValues();

    values.forEach(row => {
      const lxId = row[19] ? row[19].toString().trim() : "";      // Spalte T
      const startDate = row[27] ? formatSheetDate(row[27]) : "";  // Spalte AB
      const endDate = row[28] ? formatSheetDate(row[28]) : "";    // Spalte AC

      if (lxId && !lxId.startsWith("ERROR") && (startDate !== "" || endDate !== "")) {
        map.set(lxId, { 
          start: startDate || "-", 
          end: endDate || startDate || "-" 
        });
      }
    });
  });

  return map;
}

function buildLocalEmailMap(spreadsheet, sheetNames) {
  const map = new Map();

  sheetNames.forEach(name => {
    const sheet = spreadsheet.getSheetByName(name);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const values = sheet.getRange(2, 1, lastRow - 1, 20).getValues();

    values.forEach(row => {
      const email = row[8] ? row[8].toString().trim() : "";   // Spalte I
      const lxId = row[19] ? row[19].toString().trim() : "";  // Spalte T

      if (lxId && email && !lxId.startsWith("ERROR")) {
        map.set(lxId, email);
      }
    });
  });

  return map;
}

function formatSheetDate(val) {
  if (!val || val === "-") return "";
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return "";
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "dd.MM.yyyy");
  }
  const strVal = val.toString().trim();
  if (strVal === "" || strVal === "-") return "";
  return strVal;
}

function safeApiFetch(url, token) {
  const options = {
    method: "get",
    headers: {
      "Authorization": "Bearer " + token,
      "Accept": "application/json"
    },
    muteHttpExceptions: true
  };

  Utilities.sleep(600);
  return UrlFetchApp.fetch(url, options);
}

function fetchContactEmails(contactId, token) {
  try {
    const url = `https://api.lexware.io/v1/contacts/${contactId}`;
    const response = safeApiFetch(url, token);

    if (response && response.getResponseCode() === 200) {
      const contact = JSON.parse(response.getContentText());
      const emails = [];

      if (contact.emailAddresses) {
        if (contact.emailAddresses.business) emails.push(...contact.emailAddresses.business);
        if (contact.emailAddresses.office) emails.push(...contact.emailAddresses.office);
        if (contact.emailAddresses.other) emails.push(...contact.emailAddresses.other);
      }

      return emails.length > 0 ? emails.join(", ") : "-";
    }
  } catch (e) {
    Logger.log(`Fehler beim Abrufen der E-Mail für Contact-ID ${contactId}: ${e.message}`);
  }
  return "-";
}