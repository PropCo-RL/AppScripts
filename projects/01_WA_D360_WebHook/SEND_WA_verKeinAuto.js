// ============================================================================
// AUTOMATISCHER WHATSAPP VERSAND FÜR 'verKeinAuto' VIA 360DIALOG
// ============================================================================

/**
 * HAUPTFUNKTION:
 * Prüft Bedingungen:
 * - Spalte F (Service Fenster): "Ja"
 * - Spalte G (Abreise < 6 Tage): "Ja"
 * - Spalte H (Kontaktiert): "Nein" (oder leer)
 * 
 * Sendet eine natürliche WhatsApp-Textnachricht via 360dialog API und markiert Spalte H.
 */
function sendVerKeinAutoWhatsAppMessages() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetVK = ss.getSheetByName("verKeinAuto");

  if (!sheetVK) throw new Error("Das Tabellenblatt 'verKeinAuto' wurde nicht gefunden.");

  const lastRowVK = sheetVK.getLastRow();
  if (lastRowVK < 2) {
    Logger.log("Keine Daten im Blatt 'verKeinAuto' vorhanden.");
    return;
  }

  // Liest Spalte A bis H (8 Spalten)
  const rangeVK = sheetVK.getRange(2, 1, lastRowVK - 1, 8);
  const valuesVK = rangeVK.getValues();

  let sentCount = 0;

  valuesVK.forEach((row, index) => {
    const rowIndex = index + 2;

    const apartmentUrl = row[0] ? row[0].toString().trim() : "";  // Spalte A: Apartment URL
    const phoneRaw = row[1] ? row[1].toString().trim() : "";      // Spalte B: Telefonnummer
    const departureRaw = row[3];                                  // Spalte D: Abreise Datum
    const serviceWindow = row[5] ? row[5].toString().trim().toLowerCase() : ""; // Spalte F: Service Fenster
    const departureSoon = row[6] ? row[6].toString().trim().toLowerCase() : ""; // Spalte G: Abreise < 6 Tage
    const contactedStatus = row[7] ? row[7].toString().trim().toLowerCase() : ""; // Spalte H: Kontaktiert

    // --- BEDINGUNGEN PRÜFEN ---
    // 1. Service Fenster MUSS "ja" sein
    // 2. Abreise < 6 Tage MUSS "ja" sein
    // 3. Kontaktiert MUSS "nein" sein (oder leer)
    if (serviceWindow === "ja" && departureSoon === "ja" && (contactedStatus === "nein" || contactedStatus === "")) {

      if (!phoneRaw) {
        Logger.log(`Zeile ${rowIndex} übersprungen: Keine Telefonnummer vorhanden.`);
        return;
      }

      // Telefonnummer bereinigen (nur Zahlen)
      const cleanPhone = phoneRaw.replace(/[^0-9]/g, '');

      // Abreisedatum OHNE JAHR formatieren (z. B. 01.08.)
      let formattedDeparture = departureRaw;
      if (departureRaw instanceof Date) {
        formattedDeparture = Utilities.formatDate(departureRaw, "Europe/Berlin", "dd.MM.");
      } else if (departureRaw) {
        // Falls String in Form DD.MM.YYYY vorliegt -> Jahr abschneiden
        const str = departureRaw.toString().trim();
        const parts = str.split(".");
        if (parts.length >= 2) {
          formattedDeparture = `${parts[0].padStart(2, '0')}.${parts[1].padStart(2, '0')}.`;
        }
      }

      // Natürlicher Nachrichtentext ohne Klammern und mit Kurzdatum
      const messageBody = `Bleibt es bei der Abreise am ${formattedDeparture} oder geht nochmal weiter?\nFür die Wohnung ${apartmentUrl}`;

      Logger.log(`[Zeile ${rowIndex}] Sende WhatsApp-Nachricht an ${cleanPhone}...`);

      // WhatsApp Message über 360dialog API senden
      const result = send360DialogTextMessage(cleanPhone, messageBody);

      if (result.success) {
        const timeStamp = Utilities.formatDate(new Date(), "Europe/Berlin", "dd.MM.yyyy HH:mm");
        sheetVK.getRange(rowIndex, 8).setValue(`Ja (${timeStamp})`); // Spalte H aktualisieren
        sentCount++;
        Logger.log(`[ERFOLG Zeile ${rowIndex}] Nachricht an ${cleanPhone} gesendet!`);
      } else {
        sheetVK.getRange(rowIndex, 8).setValue(`ERROR: ${result.error.substring(0, 50)}`);
        Logger.log(`🚨 [FEHLER Zeile ${rowIndex}] ${result.error}`);
      }

      Utilities.sleep(500); // 0,5 Sek. Puffer
    }
  });

  Logger.log(`=== PROZESS BEENDET: ${sentCount} WhatsApp-Nachricht(en) versendet. ===`);
}

// ============================================================================
// API-CALL AN 360DIALOG API (MESSAGES ENDPOINT)
// ============================================================================

/**
 * Sendet eine einfache Text-Nachricht über die 360dialog v2 API
 */
function send360DialogTextMessage(toPhone, textBody) {
  const props = PropertiesService.getScriptProperties();
  
  // Liest das Script Property 'D360-Key' aus
  const d360ApiKey = props.getProperty("D360-Key");

  if (!d360ApiKey) {
    throw new Error("Script Property 'D360-Key' fehlt! Bitte in den Projekteinstellungen prüfen.");
  }

  const targetUrl = "https://waba-v2.360dialog.io/messages";

  const payload = {
    "messaging_product": "whatsapp",
    "recipient_type": "individual",
    "to": toPhone,
    "type": "text",
    "text": {
      "preview_url": false,
      "body": textBody
    }
  };

  const options = {
    "method": "post",
    "headers": {
      "D360-API-KEY": d360ApiKey,
      "Content-Type": "application/json",
      "Accept": "application/json"
    },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    const response = UrlFetchApp.fetch(targetUrl, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode === 200 || responseCode === 201) {
      return { success: true };
    } else {
      return { success: false, error: `HTTP ${responseCode}: ${responseText}` };
    }
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}