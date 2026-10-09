// =================================================================
// HAUPTFUNKTION: SEQUENTIELLE VERARBEITUNG NEUER ANFRAGEN
// =================================================================
function processNewInquiriesSequentially() {
  const TARGET_SPREADSHEET_ID = "11gnOwdzRtb9Jng-s10tUYZ1tBfuOLYf75-pnrfGC5P8"; 
  const SOURCE_LOG_SHEET_NAME = "Log";                                        
  const TARGET_SHEET_NAME     = "A";                                          
  
  const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('Gemini-Key');

  const TRIGGER_STRING = "Ok sekunde, ich schaue kurz was frei ist";

  const sourceSs = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = sourceSs.getSheetByName(SOURCE_LOG_SHEET_NAME);
  
  let targetSs, targetSheet;
  try {
    targetSs = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    targetSheet = targetSs.getSheetByName(TARGET_SHEET_NAME);
    if (!targetSheet) {
      targetSheet = targetSs.insertSheet(TARGET_SHEET_NAME);
    }
  } catch (e) {
    Logger.log("FEHLER beim Öffnen der Ziel-Datei: " + e.toString());
    return;
  }

  if (!sourceSheet || !targetSheet) return;

  const lastSourceRow = sourceSheet.getLastRow();
  if (lastSourceRow <= 1) return;

  // Liest alle relevanten Daten bis Spalte Q (Spalte 17) ein
  const sourceRange = sourceSheet.getRange(2, 1, lastSourceRow - 1, 17);
  const sourceValues = sourceRange.getValues();
  const numbersToProcess = new Set();

  // 1. MARKIERUNG IN SPALTE Q (SPALTE 17) SETZEN
  for (let i = 0; i < sourceValues.length; i++) {
    const rowNum    = i + 2;
    const rawPhone  = String(sourceValues[i][2]).trim(); 
    const chatText  = String(sourceValues[i][4]).trim(); 
    const marker    = String(sourceValues[i][16]).trim(); // Spalte Q (Index 16)

    const cleanPhone = normalizePhone(rawPhone);

    if (marker.toLowerCase().includes("mit gemini") || marker.toLowerCase() === "x" || marker.toLowerCase() === "y" || !chatText) {
      continue;
    }

    const isTriggerMatch = cleanTextForMatching(chatText).includes(cleanTextForMatching(TRIGGER_STRING));

    if (isTriggerMatch) {
      sourceSheet.getRange(rowNum, 17).setValue("y - mit gemini");
      sourceValues[i][16] = "y - mit gemini";
      if (cleanPhone) numbersToProcess.add(cleanPhone);
    } else {
      sourceSheet.getRange(rowNum, 17).setValue("x - mit gemini");
      sourceValues[i][16] = "x - mit gemini";
    }
  }

  // Wichtig für sofortige Speicherung & Vermeidung von Race-Conditions
  SpreadsheetApp.flush();

  if (numbersToProcess.size === 0) return;

  Logger.log(`Starte Anfragen-Analyse für ${numbersToProcess.size} Nummer(n)...`);

  // 2. VERARBEITUNG PRO NUMMER
  numbersToProcess.forEach(cleanPhone => {
    // Eingrenzung streng auf 48 Stunden vor Trigger
    const { historyText, referenceDateStr, triggerTimestamp } = get48HourHistoryForPhone(sourceValues, cleanPhone, TRIGGER_STRING);
    const rawPhoneToUse = getRawPhoneForCleanPhone(sourceValues, cleanPhone);

    Logger.log(`[Anfrage-Modul] Analysiere Text für ${cleanPhone} mit Bezugsdatum ${referenceDateStr}...`);
    const inquiryData = callGeminiForInquiryDetails(historyText, GEMINI_API_KEY, referenceDateStr);

    // -------------------------------------------------------------
    // EXPORT IN SHEET "A"
    // -------------------------------------------------------------
    const newRow = [
      triggerTimestamp,                  // Spalte A: Zeitstempel des Trigger-Events
      "whatsapp api",                    // Spalte B: Anfrage Channel
      inquiryData.name || "",            // Spalte C: Name anfragende Person
      inquiryData.email || "",           // Spalte D: E-Mail
      rawPhoneToUse,                     // Spalte E: Telefonnummer
      inquiryData.stadt_ort || "",       // Spalte F: A_Stadt / Ort / PLZ
      inquiryData.personen || "",        // Spalte G: A_Personen
      inquiryData.start_datum || "",     // Spalte H: A_Start datum
      inquiryData.ende_datum || ""       // Spalte I: A_Ende datum
    ];

    targetSheet.appendRow(newRow);
    Logger.log(`Anfrage für ${cleanPhone} erfolgreich in Tab 'A' eingefügt.`);

    Utilities.sleep(200);
  });

  Logger.log("--- PROZESS NEUE ANFRAGEN ERFOLGREICH BEENDET ---");
}


// =================================================================
// KI-AUFRUF FÜR ANFRAGEDETAILS (GEMINI)
// =================================================================
function callGeminiForInquiryDetails(chatText, apiKey, referenceDateStr) {
  const fallback = { name: "", email: "", stadt_ort: "", personen: "", start_datum: "", ende_datum: "" };
  if (!apiKey) return fallback;

  const promptText = `
Du bist ein Assistent zur Extraktion von Buchungsanfragen aus Kunden-Chats.
Das Bezugsdatum (HEUTE) für relative Zeitangaben (z.B. "nächste Woche", "ab morgen", "bis November") ist: **${referenceDateStr}**.

Extrahiere die folgenden Informationen aus dem Chat-Verlauf:
- "name": Vor- und Nachname der anfragenden Person (falls genannt).
- "email": E-Mail-Adresse (falls genannt, E-Mails auf @l8street.com ignorieren).
- "stadt_ort": Stadt, Ort, Stadtteil, Adresse oder Postleitzahl, in der die Unterkunft gesucht wird.
- "personen": Anzahl der Personen/Gäste (z. B. "2", "3 Arbeiter", "1 Person").
- "start_datum": Anreisedatum / Mietbeginn in DD.MM.YYYY. Wenn nur ungenau genannt (z.B. "ab Mitte Oktober"), gib den ungefähren/genannten Text an.
- "ende_datum": Abreisedatum / Mietende in DD.MM.YYYY oder Text (z.B. "Ende November", "ca. 3 Monate", "Open End").

Antworte AUSSCHLIESSLICH als valides JSON:
{
  "name": "",
  "email": "",
  "stadt_ort": "",
  "personen": "",
  "start_datum": "",
  "ende_datum": ""
}

Chat-Verlauf:
"""
${chatText}
"""
`;

  const parts = [{ text: promptText }];
  const result = executeGeminiRequest(parts, apiKey);
  return result ? { ...fallback, ...result } : fallback;
}


// =================================================================
// HELPER FUNKTIONEN (SPEZIELL FÜR 48 STUNDEN FENSTER)
// =================================================================

function get48HourHistoryForPhone(sourceValues, cleanPhone, triggerString) {
  const FORTY_EIGHT_HOURS_MS = 48 * 60 * 60 * 1000;
  let triggerTimeMs = new Date().getTime();
  let triggerTimestamp = new Date();

  // Trigger-Zeile ermitteln
  for (let i = sourceValues.length - 1; i >= 0; i--) {
    const matchPhone = normalizePhone(sourceValues[i][2]);
    const chatText   = String(sourceValues[i][4]).trim();
    if (matchPhone === cleanPhone && cleanTextForMatching(chatText).includes(cleanTextForMatching(triggerString))) {
      const parsedDate = new Date(sourceValues[i][0]);
      if (!isNaN(parsedDate.getTime())) {
        triggerTimeMs = parsedDate.getTime();
        triggerTimestamp = sourceValues[i][0]; // Original-Zeitstempel aus Spalte A
      }
      break;
    }
  }

  const cutoffTimeMs = triggerTimeMs - FORTY_EIGHT_HOURS_MS;
  const texts = [];

  for (let i = 0; i < sourceValues.length; i++) {
    const matchPhone = normalizePhone(sourceValues[i][2]);
    const chatText   = String(sourceValues[i][4]).trim();
    const msgDate    = new Date(sourceValues[i][0]);

    if (matchPhone === cleanPhone && chatText) {
      if (isNaN(msgDate.getTime()) || (msgDate.getTime() >= cutoffTimeMs && msgDate.getTime() <= triggerTimeMs)) {
        texts.push(`[${sourceValues[i][0]}] ${chatText}`);
      }
    }
  }

  return { 
    historyText: texts.join("\n"), 
    referenceDateStr: new Date(triggerTimeMs).toLocaleDateString("de-DE"),
    triggerTimestamp: triggerTimestamp
  };
}