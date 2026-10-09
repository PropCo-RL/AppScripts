// ============================================================================
// AUTOMATISCHE OPT-IN / OPT-OUT AUSWERTUNG VIA PRÄZISER JSON-ANALYSE
// ============================================================================

/**
 * HAUPTFUNKTION:
 * 1. Liest periodisch (z. B. stündlich) das Blatt 'Log' aus.
 * 2. Filtert auf eingehende Nachrichten ("in"), die in Spalte L noch ungesehen sind.
 * 3. Analysiert das Roh-JSON aus Spalte G auf echte Button-Klicks ("Ja, das ist ok") -> STATUS: IN.
 * 4. Analysiert auf Abmeldungen ("STOP" / "STOPP") -> STATUS: STOP.
 * 5. Aktualisiert den Status im Tab 'optIn' (von "Gesendet" zu "IN" / "STOP").
 * 6. Setzt den Auswertungs-Stempel AUSSCHLIESSLICH in Spalte L im Blatt 'Log'.
 */
function processOptInEvents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetLog = ss.getSheetByName("Log");
  const sheetOptIn = ss.getSheetByName("optIn") || ss.insertSheet("optIn");

  if (!sheetLog) throw new Error("Das Tabellenblatt 'Log' wurde nicht gefunden.");

  const lastRowLog = sheetLog.getLastRow();
  if (lastRowLog < 2) return;

  // Liest Spalte A bis L (12 Spalten) aus 'Log' ein
  const rangeLog = sheetLog.getRange(2, 1, lastRowLog - 1, 12);
  const valuesLog = rangeLog.getValues();

  let processedCount = 0;

  valuesLog.forEach((row, index) => {
    const rowIndex = index + 2;

    const timestamp   = row[0];                                       // Spalte A: Zeit
    const direction   = row[1] ? row[1].toString().trim().toLowerCase() : ""; // Spalte B: InC (in/out)
    const phoneRaw    = row[2] ? row[2].toString().trim() : "";      // Spalte C: Nummer
    const msgType     = row[3] ? row[3].toString().trim() : "";      // Spalte D: Art
    const message     = row[4] ? row[4].toString().trim() : "";      // Spalte E: Nachricht
    const messageId   = row[5] ? row[5].toString().trim() : "";      // Spalte F: Nachricht ID
    const rawJson     = row[6] ? row[6].toString().trim() : "";      // Spalte G: JSON
    const colLMarker  = row[11] ? row[11].toString().trim() : "";     // Spalte L: OptIn Verarbeitet

    // Nur EINGEHENDE Nachrichten ("in") verarbeiten, die noch NICHT ausgewertet wurden
    if (direction === "in" && colLMarker === "") {
      
      let statusLabel = null; // Wird 'IN' oder 'STOP'
      
      // JSON aus Spalte G parsen und analysieren
      const jsonDetails = parseWhatsAppJsonPayload(rawJson);

      // ----------------------------------------------------------------------
      // A) PRÜFUNG OPT-IN: Echter Button-Druck auf "Ja, das ist ok"
      // ----------------------------------------------------------------------
      const isExactButtonText = jsonDetails.buttonText.trim().toLowerCase() === "ja, das ist ok";
      
      if (jsonDetails.isButtonEvent && isExactButtonText) {
        statusLabel = "IN";
      } 
      // Fallback: Falls der Nachrichtentyp 'button'/'interactive' ist & der Text exakt übereinstimmt
      else if ((msgType === "button" || msgType === "interactive") && message.trim().toLowerCase() === "ja, das ist ok") {
        statusLabel = "IN";
      }

      // ----------------------------------------------------------------------
      // B) PRÜFUNG OPT-OUT: Exakt "STOP" oder "STOPP"
      // ----------------------------------------------------------------------
      else if (message.trim().toUpperCase() === "STOP" || message.trim().toUpperCase() === "STOPP") {
        statusLabel = "STOP";
      }

      // ----------------------------------------------------------------------
      // SPEICHERN IM SHEET 'optIn' UND STEMPEL IN SPALTE L (LOG)
      // ----------------------------------------------------------------------
      if (statusLabel) {
        
        // Aktualisiert die Zeile im Tab 'optIn' von "Gesendet" auf "IN" bzw. "STOP"
        smartUpsertOptInSheet(sheetOptIn, phoneRaw, [
          timestamp,      // A: Zeit
          direction,      // B: In/Out
          phoneRaw,       // C: Nummer
          msgType,        // D: Art
          message,        // E: Nachricht
          messageId,      // F: Nachricht ID
          rawJson,        // G: JSON
          statusLabel     // H: Status (IN oder STOP)
        ]);

        // Stempel AUSSCHLIESSLICH in Spalte L setzen
        const timeStampStr = Utilities.formatDate(new Date(), "Europe/Berlin", "dd.MM.yyyy HH:mm");
        sheetLog.getRange(rowIndex, 12).setValue(`OptIn Status: ${statusLabel} (${timeStampStr})`);
        
        processedCount++;
        Logger.log(`[Zeile ${rowIndex}] Event '${statusLabel}' für Nummer ${phoneRaw} erfolgreich verarbeitet.`);
      }
    }
  });

  Logger.log(`=== OPT-IN AUSWERTUNG BEENDET: ${processedCount} Event(s) verarbeitet. ===`);
}

/**
 * HILFSFUNKTION: 
 * Prüft, ob für diese Handynummer bereits eine Zeile im Tab 'optIn' existiert 
 * (z. B. vom Versand mit Status "Gesendet"). Falls ja -> überschreibt Status mit "IN" / "STOP".
 * Falls nein -> hängt neue Zeile an.
 */
function smartUpsertOptInSheet(sheetOptIn, phoneRaw, newRowData) {
  const cleanPhone = normalizePhone(phoneRaw);
  const lastRow = sheetOptIn.getLastRow();

  if (lastRow >= 2) {
    const values = sheetOptIn.getRange(2, 1, lastRow - 1, 8).getValues();
    
    for (let i = 0; i < values.length; i++) {
      const existingPhone = normalizePhone(values[i][2]); // Spalte C: Nummer
      
      if (existingPhone === cleanPhone && cleanPhone !== "") {
        // Gefunden! Überschreibt diese Zeile mit den neuen Daten & dem neuen Status
        sheetOptIn.getRange(i + 2, 1, 1, newRowData.length).setValues([newRowData]);
        return;
      }
    }
  }

  // Nummer stand noch nicht drin -> Neu anhängen
  sheetOptIn.appendRow(newRowData);
}

/**
 * HILFSFUNKTION: Parst das Meta/360dialog Webhook JSON aus Spalte G
 */
function parseWhatsAppJsonPayload(rawJson) {
  const info = {
    isButtonEvent: false,
    buttonText: "",
    msgType: ""
  };

  if (!rawJson) return info;

  try {
    const data = JSON.parse(rawJson);
    const change = data.entry?.[0]?.changes?.[0];
    const val = change?.value;

    if (val?.messages && val.messages.length > 0) {
      const msg = val.messages[0];
      info.msgType = msg.type || "";

      // Fall 1: Quick Reply Button Event
      if (msg.type === "button" && msg.button) {
        info.isButtonEvent = true;
        info.buttonText = msg.button.text || "";
      } 
      // Fall 2: Interactive Button Reply Event
      else if (msg.type === "interactive" && msg.interactive?.button_reply) {
        info.isButtonEvent = true;
        info.buttonText = msg.interactive.button_reply.title || "";
      }
    }
  } catch (err) {
    // Falls das JSON Fehler enthält
  }

  return info;
}

/**
 * Hilfsfunktion zum Bereinigen der Telefonnummer
 */
function normalizePhone(phone) {
  return String(phone || "").replace(/\D/g, "");
}