// ============================================================================
// SKRIPT: RECHNUNGSSTATUS VOM LEXOFFICE ABRUFEN & IN SPALTE AF EINTRAGEN
// ============================================================================

/**
 * Hauptfunktion: Führt das Status-Update sowohl für 'VL' als auch für 'RE' aus.
 */
function updateLexofficeInvoiceStatusAll() {
  const sheetsToUpdate = ["VL", "RE"];

  sheetsToUpdate.forEach(sheetName => {
    updateLexofficeInvoiceStatusForSheet(sheetName);
  });
}

/**
 * Aktualisiert den Rechnungsstatus für ein spezifisches Tabellenblatt.
 * Überspringt Zeilen, in denen Spalte AF bereits 'Bezahlt' oder 'Storniert' enthält.
 */
function updateLexofficeInvoiceStatusForSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    Logger.log(`⚠️ Das Tabellenblatt '${sheetName}' wurde nicht gefunden.`);
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log(`Keine Daten im Sheet '${sheetName}' vorhanden.`);
    return;
  }

  // Liest Daten ab Zeile 2 bis Spalte AF (32 Spalten), um Spalte AF direkt mitzulesen
  const range = sheet.getRange(2, 1, lastRow - 1, 32);
  const values = range.getValues();

  Logger.log(`Starte Status-Abruf für Sheet '${sheetName}' (${values.length} Zeilen)...`);

  values.forEach((row, index) => {
    const rowIndex = index + 2;
    const lxInvoiceId = row[19] ? row[19].toString().trim() : ""; // Spalte T (Index 19): Invoice ID
    const currentStatus = row[31] ? row[31].toString().trim() : ""; // Spalte AF (Index 31): Aktueller Status

    // 1. Nur verarbeiten, wenn eine gültige ID vorhanden ist
    if (!lxInvoiceId || lxInvoiceId.startsWith("ERROR")) return;

    // 2. Überspringen, wenn bereits "Bezahlt" oder "Storniert" eingetragen ist
    const normalizedStatus = currentStatus.toLowerCase();
    if (normalizedStatus === "bezahlt" || normalizedStatus === "storniert") {
      Logger.log(`⏭️ [Sheet ${sheetName} | Zeile ${rowIndex}] ID ${lxInvoiceId} übersprungen (Status ist bereits '${currentStatus}').`);
      return;
    }

    // Abfrage bei Lexoffice per Hilfsfunktion
    const invoiceData = getLexofficeInvoiceDetailsVL(lxInvoiceId);

    if (invoiceData && invoiceData.voucherStatus) {
      const rawStatus = invoiceData.voucherStatus.toLowerCase();
      let statusDE = rawStatus;

      // Status auf Deutsch übersetzen
      switch (rawStatus) {
        case "open":
          statusDE = "Offen";
          break;
        case "paid":
          statusDE = "Bezahlt";
          break;
        case "voided":
          statusDE = "Storniert";
          break;
        case "draft":
          statusDE = "Entwurf";
          break;
      }

      // Schreiben in Spalte AF (Spalte 32)
      sheet.getRange(rowIndex, 32).setValue(statusDE);
      Logger.log(`[Sheet ${sheetName} | Zeile ${rowIndex}] ID ${lxInvoiceId}: Status -> ${statusDE}`);

    } else {
      Logger.log(`🚨 [Sheet ${sheetName} | Zeile ${rowIndex}] Status konnte für ID ${lxInvoiceId} nicht abgerufen werden.`);
    }

    Utilities.sleep(200); // API-Puffer für Lexoffice Ratelimit
  });

  Logger.log(`=== STATUS-UPDATE FÜR '${sheetName}' BEENDET ===`);
}

/**
 * HILFSFUNKTION: Fragt die Rechnungsdetails direkt von der Lexoffice-API ab.
 */
function getLexofficeInvoiceDetailsVL(invoiceId) {
  const props = PropertiesService.getScriptProperties();
  const lxKey = props.getProperty("lxKey");
  
  if (!lxKey) {
    Logger.log("🚨 Fehler: Script Property 'lxKey' fehlt unter den Skripteigenschaften!");
    return null;
  }

  const url = `https://api.lexoffice.io/v1/invoices/${invoiceId}`;
  
  const options = {
    "method": "get",
    "headers": {
      "Authorization": "Bearer " + lxKey,
      "Accept": "application/json"
    },
    "muteHttpExceptions": true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() === 200) {
      return JSON.parse(response.getContentText());
    } else {
      Logger.log(`🚨 API-Fehler (${response.getResponseCode()}) für ID ${invoiceId}: ${response.getContentText()}`);
      return null;
    }
  } catch (e) {
    Logger.log(`🚨 Ausnahmefehler bei Abruf für ID ${invoiceId}: ${e.toString()}`);
    return null;
  }
}