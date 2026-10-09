// ============================================================================
// AUTOMATISCHER OPT-IN TEMPLATE VERSAND (MANUELL & 24H-AUTOMATIK)
// ============================================================================

/**
 * HAUPTFUNKTION:
 * 1. Prüft manuellen Trigger-String (sofortiger Template-Versand).
 * 2. Prüft "y - mit gemini" in Spalte I und sendet 22h nach der Nachricht:
 *    - Erst den leicht abgeänderten Ankündigungstext (Freitext)
 *    - 3 Sekunden Pause
 *    - Dann das Opt-in Template 'opt_in'
 * 3. Lässt Spalte I & L komplett unangetastet und dokumentiert Status NUR in Spalte M!
 * 4. Verhindert Doppelversand durch weltweite Nummern-Sperre im Tab 'optIn'.
 */
function sendOptInTemplateOnTrigger() {
  const MANUAL_TRIGGER_STRING = "Ich sende gleich noch eine Anfrage ob wir Benachrichtigungen (Check-in Codes, Rückfragen, Infos) per WhatsApp senden dürfen zu dieser Buchung. Ist das ok?";
  const AUTO_ANNOUNCEMENT_TEXT = "Ich sende gleich noch eine Anfrage ob wir Benachrichtigungen (z. B. für Check-in Codes, Rückfragen, Infos) per WhatsApp senden dürfen. Ist das ok?";
  
  const TEMPLATE_NAME = "opt_in";
  const TEMPLATE_LANGUAGE = "de";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetLog = ss.getSheetByName("Log");
  const sheetOptIn = ss.getSheetByName("optIn") || ss.insertSheet("optIn");

  if (!sheetLog) throw new Error("Das Tabellenblatt 'Log' wurde nicht gefunden.");

  const lastRowLog = sheetLog.getLastRow();
  if (lastRowLog < 2) return;

  // Liest Spalte A bis M (13 Spalten) aus 'Log'
  const rangeLog = sheetLog.getRange(2, 1, lastRowLog - 1, 13);
  const valuesLog = rangeLog.getValues();

  // --------------------------------------------------------------------------
  // SPERRE: Alle bekannten Nummern aus dem Tab 'optIn' laden
  // --------------------------------------------------------------------------
  const optInKnownNumbers = new Set();
  const lastRowOptIn = sheetOptIn.getLastRow();
  if (lastRowOptIn >= 2) {
    const optInPhones = sheetOptIn.getRange(2, 3, lastRowOptIn - 1, 1).getValues(); // Spalte C: Nummer
    optInPhones.forEach(r => {
      if (r[0]) {
        optInKnownNumbers.add(normalizePhone(r[0]));
      }
    });
  }

  let sentCount = 0;
  const colMOutputs = []; // Für Ergebnisse in Spalte M

  const nowMs = new Date().getTime();
  const TWENTY_TWO_HOURS_MS = 22 * 60 * 60 * 1000; // 22 Stunden in Millisekunden
  const TWENTY_FOUR_HOURS_MS = 24 * 60 * 60 * 1000; // 24 Stunden in Millisekunden

  valuesLog.forEach((row, index) => {
    const rowIndex = index + 2;

    const timestampRaw = row[0];                                       // Spalte A: Zeit
    const direction    = row[1] ? row[1].toString().trim().toLowerCase() : ""; // Spalte B: InC (in/out)
    const phoneRaw     = row[2] ? row[2].toString().trim() : "";              // Spalte C: Nummer
    const chatText     = row[4] ? row[4].toString().trim() : "";              // Spalte E: Nachricht
    const colIValue    = row[8] ? row[8].toString().trim() : "";              // Spalte I: RE Script
    const currentColM  = row[12] ? row[12].toString().trim() : "";             // Spalte M: OptIn Template senden

    const cleanPhone = normalizePhone(phoneRaw);
    const cleanChatText = cleanTextForMatching(chatText);
    const cleanManualTrigger = cleanTextForMatching(MANUAL_TRIGGER_STRING);

    const isManualMatch = cleanChatText.includes(cleanManualTrigger);
    const isGeminiMatch = colIValue.toLowerCase().includes("y - mit gemini");

    let statusM = currentColM || "x";

    // Nur verarbeiten, wenn in Spalte M noch nicht erfolgreich gesendet wurde
    if (!currentColM.startsWith("Template gesendet")) {

      // PRÜFUNG: Wurde an diese Handynummer im Tab 'optIn' schon mal versendet?
      if (cleanPhone && optInKnownNumbers.has(cleanPhone)) {
        statusM = "x (Sperre - Bereits in optIn)";
      } 
      else {
        
        // ----------------------------------------------------------------------
        // FALL 1: MANUELLER TRIGGER (Sofortiger Versand)
        // ----------------------------------------------------------------------
        if (direction === "out" && isManualMatch && cleanPhone) {
          Logger.log(`[Zeile ${rowIndex}] Manueller Trigger erkannt! Sende Opt-In Template an ${cleanPhone}...`);

          const result = send360DialogOpenTemplate(cleanPhone, TEMPLATE_NAME, TEMPLATE_LANGUAGE);

          if (result.success) {
            statusM = "Template gesendet (Manuell)";
            
            // Im Tab 'optIn' eintragen, dass versendet wurde!
            sheetOptIn.appendRow([new Date(), "out", phoneRaw, "template", "opt_in", "", "", "Gesendet"]);
            optInKnownNumbers.add(cleanPhone); // Sofort im lokalen Speicher sperren!
            sentCount++;
          } else {
            statusM = `ERROR: ${result.error.substring(0, 30)}`;
          }
        }

        // ----------------------------------------------------------------------
        // FALL 2: AUTOMATIK-TRIGGER VIA SPALTE I ("y - mit gemini")
        // ----------------------------------------------------------------------
        else if (isGeminiMatch && cleanPhone) {
          const msgDate = new Date(timestampRaw);
          
          if (!isNaN(msgDate.getTime())) {
            const ageMs = nowMs - msgDate.getTime();

            // Prüfe, ob die Nachricht zwischen 22 und 24 Stunden alt ist (ca. 2h vor Ablauf)
            if (ageMs >= TWENTY_TWO_HOURS_MS && ageMs < TWENTY_FOUR_HOURS_MS) {
              Logger.log(`[Zeile ${rowIndex}] 22h-Fenster für 'y - mit gemini' erreicht! Sende Ankündigung + Template an ${cleanPhone}...`);

              // 1. Freitext-Ankündigung senden (im aktiven 24h-Fenster)
              const textResult = send360DialogTextMessage(cleanPhone, AUTO_ANNOUNCEMENT_TEXT);

              if (textResult.success) {
                // 2. 3 Sekunden warten
                Utilities.sleep(3000);

                // 3. Template senden
                const templateResult = send360DialogOpenTemplate(cleanPhone, TEMPLATE_NAME, TEMPLATE_LANGUAGE);

                if (templateResult.success) {
                  statusM = "Template gesendet (24h Auto)";
                  
                  // Im Tab 'optIn' eintragen!
                  sheetOptIn.appendRow([new Date(), "out", phoneRaw, "template", "opt_in", "", "", "Gesendet"]);
                  optInKnownNumbers.add(cleanPhone);
                  sentCount++;
                } else {
                  statusM = `ERROR Template: ${templateResult.error.substring(0, 30)}`;
                }
              } else {
                statusM = `ERROR Text: ${textResult.error.substring(0, 30)}`;
              }
            } else if (ageMs < TWENTY_TWO_HOURS_MS) {
              statusM = "Wartet auf 22h-Fenster";
            }
          }
        }
      }
    }

    colMOutputs.push([statusM]);
  });

  // Schreibt das Ergebnis AUSSCHLIESSLICH in Spalte M ab Zeile 2!
  sheetLog.getRange(2, 13, colMOutputs.length, 1).setValues(colMOutputs);

  Logger.log(`=== OPT-IN TEMPLATE VERSAND BEENDET: ${sentCount} Prozess(e) verarbeitet. ===`);
}

// ============================================================================
// API-CALLS AN 360DIALOG API
// ============================================================================

/**
 * Sendet eine normale Freitext-Nachricht über 360dialog
 */
function send360DialogTextMessage(toPhone, textBody) {
  const props = PropertiesService.getScriptProperties();
  const d360ApiKey = props.getProperty("D360-Key");

  const targetUrl = "https://waba-v2.360dialog.io/messages";

  const payload = {
    "messaging_product": "whatsapp",
    "recipient_type": "individual",
    "to": toPhone,
    "type": "text",
    "text": {
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
    return { success: (response.getResponseCode() === 200 || response.getResponseCode() === 201) };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Sendet das 'opt_in' Template über 360dialog
 */
function send360DialogOpenTemplate(toPhone, templateName, languageCode) {
  const props = PropertiesService.getScriptProperties();
  const d360ApiKey = props.getProperty("D360-Key");

  const targetUrl = "https://waba-v2.360dialog.io/messages";

  const payload = {
    "messaging_product": "whatsapp",
    "recipient_type": "individual",
    "to": toPhone,
    "type": "template",
    "template": {
      "name": templateName,
      "language": {
        "code": languageCode
      }
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

/**
 * Hilfsfunktion zum Bereinigen der Telefonnummer
 */
function normalizePhone(phone) {
  return String(phone || "").replace(/\D/g, "");
}

/**
 * Hilfsfunktion zum Bereinigen von Text für den Abgleich
 */
function cleanTextForMatching(text) {
  return String(text || "").toLowerCase().replace(/\s+/g, " ").trim();
}