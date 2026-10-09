// =================================================================
// VERLÄNGERUNGEN (VL) - VOLLSTÄNDIGER, OPTIMIERTER CODE
// =================================================================
function processExtensionsSequentially() {
  const TARGET_SPREADSHEET_ID = "1mMEVsVFIx5VxIV3EyJYlba_G0fgEX9t0tR2KbFkm-nU";
  const SOURCE_LOG_SHEET_NAME = "Log";
  const TARGET_VL_SHEET_NAME = "VL";
  const TARGET_RE_SHEET_NAME = "RE";
  const BACKUP_SPREADSHEET_ID = "1gBNQOi4qCHiFfEdZwibr4rW-371JtyhRfQaae1uxDmU";
  const TARGET_BACKUP_SHEET_NAME = "Log_Backup_VL";

  const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('Gemini-Key');
  const TRIGGER_STRING = "Super, vielen Dank\nIch trage die Verlängerung so ein und sende gleich alles zu";

  const sourceSs = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = sourceSs.getSheetByName(SOURCE_LOG_SHEET_NAME);
  let targetSs, vlSheet, reSheet, backupSs, backupSheet;

  try {
    targetSs = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    vlSheet = targetSs.getSheetByName(TARGET_VL_SHEET_NAME);
    reSheet = targetSs.getSheetByName(TARGET_RE_SHEET_NAME);
    
    backupSs = SpreadsheetApp.openById(BACKUP_SPREADSHEET_ID);
    backupSheet = backupSs.getSheetByName(TARGET_BACKUP_SHEET_NAME);
    if (!backupSheet) {
      backupSheet = backupSs.insertSheet(TARGET_BACKUP_SHEET_NAME);
      backupSheet.appendRow(["Telefonnummer", "JSON_Raw_Data", "Backup_Zeitstempel"]);
    }
  } catch (e) {
    Logger.log("FEHLER beim Öffnen der Tabellen: " + e.toString());
    return;
  }

  if (!sourceSheet || !vlSheet || !reSheet) return;

  const lastSourceRow = sourceSheet.getLastRow();
  if (lastSourceRow <= 1) return;

  const sourceRange = sourceSheet.getRange(2, 1, lastSourceRow - 1, 10);
  const sourceValues = sourceRange.getValues();
  const numbersToProcess = new Set();
  const cleanedTrigger = cleanTextForMatchingVL(TRIGGER_STRING);

  // 1. MARKIERUNG IN SPALTE J SETZEN (SPALTE 10)
  for (let i = 0; i < sourceValues.length; i++) {
    const rowNum = i + 2;
    const direction = String(sourceValues[i][1]).trim().toLowerCase();
    const rawPhone = String(sourceValues[i][2]).trim();
    const chatText = String(sourceValues[i][4]).trim();
    const markerJ = String(sourceValues[i][9]).trim().toLowerCase();
    const cleanPhone = normalizePhoneVL(rawPhone);

    if (markerJ.includes("mit gemini") || markerJ === "x" || markerJ === "y" || !chatText) continue;

    const isTriggerMatch = cleanTextForMatchingVL(chatText).includes(cleanedTrigger);
    const isOutbound = (direction === "out");

    if (isTriggerMatch && isOutbound) {
      sourceSheet.getRange(rowNum, 10).setValue("y - verlängerung mit gemini");
      sourceValues[i][9] = "y - verlängerung mit gemini";
      if (cleanPhone) numbersToProcess.add(cleanPhone);
    } else {
      sourceSheet.getRange(rowNum, 10).setValue("x - verlängerung mit gemini");
      sourceValues[i][9] = "x - verlängerung mit gemini";
    }
  }

  if (numbersToProcess.size === 0) return;

  // Daten aus RE und VL laden für schnellen Abgleich im Speicher
  const reData = reSheet.getDataRange().getValues();
  const vlData = vlSheet.getDataRange().getValues();

  numbersToProcess.forEach(cleanPhone => {
    const { historyText, referenceDateStr, cutoffTimeMs, triggerMessageText } = getRecentHistoryForPhoneVL(sourceValues, cleanPhone, TRIGGER_STRING);
    const rawPhoneToUse = getRawPhoneForCleanPhoneVL(sourceValues, cleanPhone);

    // KI-Aufruf ermittelt die genauen Verlängerungszeiträume
    const vlDetailsArray = callGeminiForExtensionDetails(historyText, GEMINI_API_KEY, referenceDateStr, triggerMessageText);
    save7DayJsonBackupVL(sourceValues, backupSheet, cleanPhone, rawPhoneToUse, cutoffTimeMs);

    vlDetailsArray.forEach(vlDetail => {
      // 🎯 STEP 1: Suche vorherigen Eintrag (Zuerst in VL, sonst in RE)
      let baseRow = findPreviousBookingRow(vlData, cleanPhone, vlDetail.apartment) || 
                    findPreviousBookingRow(reData, cleanPhone, vlDetail.apartment);

      let newRow = [];

      if (baseRow) {
        // Exakte Kopie der Vorlage erstellen
        newRow = [...baseRow];
        // Stelle sicher, dass das Array mindestens 37 Spalten lang ist (Spalten A bis AK)
        while (newRow.length < 37) newRow.push("");
      } else {
        // Fallback: Neue Zeile initialisieren (37 Spalten für A bis AK)
        newRow = new Array(37).fill("");
        newRow[0] = rawPhoneToUse; // Spalte A: Telefonnummer
        newRow[6] = "Deutschland"; // Spalte G: Land
      }

      // 🎯 STEP 2: Verlängerungsrelevante Felder aktualisieren
      
      // Spalte K (Index 10): Ursprüngliche Anreise bleibt unverändert (aus Kopie übernommen)!
      
      // Spalte L (Index 11): Neues Abreisedatum
      if (vlDetail.abreise) newRow[11] = vlDetail.abreise;
      
      // Spalte M (Index 12): Rhythmus (falls neu)
      if (vlDetail.rhythmus) newRow[12] = vlDetail.rhythmus;
      
      // Spalte N (Index 13): Apartment (falls gewechselt)
      if (vlDetail.apartment) newRow[13] = vlDetail.apartment;

      // Preise anpassen falls in KI erkannt
      if (vlDetail.preis_pro_nacht) newRow[14] = vlDetail.preis_pro_nacht; // Spalte O
      if (vlDetail.preis_pro_person_nacht) newRow[15] = vlDetail.preis_pro_person_nacht; // Spalte P
      if (vlDetail.preis_pro_monat) newRow[16] = vlDetail.preis_pro_monat; // Spalte Q

      // Spalte J (Index 9): Status auf Done setzen
      newRow[9] = "Done";

      // 🎯 STEP 3: Rechnungs-Zeiträume für Verlängerung eintragen (Spalten AB & AC)
      // AB (Index 27): Start der Verlängerung (Entweder explizit genannt ODER alte Abreise ODER Bezugsdatum)
      let oldAbreiseRaw = baseRow ? baseRow[11] : "";
      let oldAbreiseFormatted = (oldAbreiseRaw instanceof Date) ? Utilities.formatDate(oldAbreiseRaw, Session.getScriptTimeZone(), "dd.MM.yyyy") : String(oldAbreiseRaw || "");

      newRow[27] = vlDetail.anreise || oldAbreiseFormatted || referenceDateStr; 
      
      // AC (Index 28): Ende der Verlängerung (neue Abreise)
      newRow[28] = vlDetail.abreise || "";

      // 🎯 STEP 4: Nachfolgende Processing-Spalten LEEREN
      // (Spalte S / Index 18 Buchungsnummer BLEIBT ERHALTEN)
      newRow[19] = ""; // Spalte T (Index 19)
      newRow[20] = ""; // Spalte U (Index 20)
      newRow[21] = ""; // Spalte V (Index 21)
      newRow[23] = ""; // Spalte X (Index 23)

      // Spalten AF bis AK ebenfalls leeren:
      newRow[31] = ""; // Spalte AF (Index 31)
      newRow[32] = ""; // Spalte AG (Index 32)
      newRow[33] = ""; // Spalte AH (Index 33)
      newRow[34] = ""; // Spalte AI (Index 34)
      newRow[35] = ""; // Spalte AJ (Index 35)
      newRow[36] = ""; // Spalte AK (Index 36)

      // Neue Zeile an Tab VL anhängen
      vlSheet.appendRow(newRow);
      Logger.log(`[VL Erfolgreich] Verlängerung für ${cleanPhone} angelegt!`);
    });
  });

  Logger.log("--- VERLÄNGERUNGSPROZESS BEENDET ---");
}

// =================================================================
// HELPER: VORHERIGE BUCHUNGSZEILE SUCHEN
// =================================================================
function findPreviousBookingRow(sheetValues, cleanPhone, apartmentCode) {
  if (!sheetValues || sheetValues.length <= 1) return null;

  // Rückwärts suchen (neuesten Eintrag zuerst finden)
  for (let i = sheetValues.length - 1; i >= 1; i--) {
    const rowPhone = normalizePhoneVL(sheetValues[i][0]); // Spalte A: Nummer
    const rowApartment = String(sheetValues[i][13] || "").trim().toLowerCase(); // Spalte N: Apartment

    if (rowPhone === cleanPhone) {
      if (!apartmentCode || rowApartment.includes(apartmentCode.toLowerCase())) {
        return sheetValues[i];
      }
    }
  }
  return null;
}

// =================================================================
// GEMINI API AUFRUF & WEITERE HILFSFUNKTIONEN
// =================================================================
function callGeminiForExtensionDetails(chatText, apiKey, referenceDateStr, triggerMessageText) {
  const fallbackItem = { anreise: "", abreise: "", rhythmus: "", apartment: "", preis_pro_nacht: "", preis_pro_person_nacht: "", preis_pro_monat: "" };
  if (!apiKey) return [fallbackItem];

  const cleanTriggerMsg = triggerMessageText.toLowerCase();
  const containsMultipleKeywords = cleanTriggerMsg.includes("folgende wohnungen") ||
                                    cleanTriggerMsg.includes("folgenden wohnungen") ||
                                    cleanTriggerMsg.includes("wohnungen:");

  let multiBookingInstruction = containsMultipleKeywords
    ? `ACHTUNG - MEHRERE WOHNUNGEN WERDEN VERLÄNGERT: Erstelle FÜR JEDE genannte Verlängerungs-Wohnung ein eigenes Objekt im JSON-Array.`
    : `ACHTUNG - EINZELNE VERLÄNGERUNG: Gib EXAKT EIN OBJEKT im JSON-Array zurück (für die primäre / zuletzt genannte Wohnung).`;

  const promptText = `
Du bist ein Assistent für Buchungsverlängerungen bei Monteurwohnungen.
Das Bezugsdatum (HEUTE / Tag der Verlängerung) ist: **${referenceDateStr}**.

PRIORITÄTEN-REGEL:
Die Angaben in der Trigger-Nachricht ("${triggerMessageText.replace(/\n/g, " ")}") haben ABSOLUTE PRIORITÄT vor älteren Nachrichten im Chat-Verlauf.

${multiBookingInstruction}

EXTRAHIERE DIE VERLÄNGERUNGSDETAILS PRO WOHNUNG AUS DEM CHATVERLAUF:
1. "anreise": Startdatum der Verlängerung (NUR falls explizit in der Verlängerungsanfrage genannt, z.B. "ab 15.10.").
2. "abreise": NEUES Enddatum der Verlängerung. Rechne relative Angaben (z.B. "um 2 Wochen verlängern", "bis Ende des Monats") basierend auf ${referenceDateStr} in ein konkretes Datum (DD.MM.YYYY) um!
3. "rhythmus": Abrechnungsrhythmus (z.B. "Monatlich", "Wöchentlich", "Fixes Enddatum").
4. "apartment": Kürzel aus dem Link mit "l8street.com/a/" (NUR der Code nach "/a/").
5. "preis_pro_person_nacht": Einzelpreis pro Person/Nacht (falls neu vereinbart, z.B. "25€").
6. "preis_pro_nacht": GESAMTPREIS der Unterkunft pro Nacht.
7. "preis_pro_monat": Vereinbarter Pauschalpreis pro Monat.

ANTWORTE AUSSCHLIESSLICH ALS EIN JSON-ARRAY VON OBJEKTEN:
[{
"anreise": "",
"abreise": "",
"rhythmus": "",
"apartment": "",
"preis_pro_nacht": "",
"preis_pro_person_nacht": "",
"preis_pro_monat": ""
}]

Chat-Verlauf:
"""${chatText}"""`;

  const parts = [{ text: promptText }];
  const result = executeGeminiRequestVL(parts, apiKey);

  if (Array.isArray(result) && result.length > 0) {
    return result.map(item => ({ ...fallbackItem, ...item }));
  } else if (result && typeof result === "object") {
    return [{ ...fallbackItem, ...result }];
  }

  return [fallbackItem];
}

function executeGeminiRequestVL(parts, apiKey) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash:generateContent`;
  const payload = {
    contents: [{ parts: parts }],
    generationConfig: { responseMimeType: "application/json", temperature: 0.1 }
  };
  const options = {
    method: "post",
    contentType: "application/json",
    headers: { "x-goog-api-key": apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  try {
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) return null;
    const json = JSON.parse(response.getContentText());
    if (json.candidates && json.candidates[0] && json.candidates[0].content && json.candidates[0].content.parts[0].text) {
      let rawText = json.candidates[0].content.parts[0].text.trim();
      const jsonMatch = rawText.match(/\[[\s\S]*\]|\{[\s\S]*\}/);
      if (jsonMatch) rawText = jsonMatch[0];
      return safeParseJSONVL(rawText);
    }
  } catch (e) {
    Logger.log("Fehler bei Gemini VL API Aufruf: " + e.toString());
  }
  return null;
}

function safeParseJSONVL(rawText) {
  try { return JSON.parse(rawText); } catch (e) {
    try {
      let cleanText = rawText
        .replace(/[\u0000-\u001F\u007F-\u009F]/g, "")
        .replace(/\r?\n/g, " ")
        .replace(/",\s*}/g, '"}')
        .replace(/",\s*]/g, '"]');
      return JSON.parse(cleanText);
    } catch (e2) { return null; }
  }
}

function normalizePhoneVL(phone) {
  return String(phone || "").replace(/\D/g, "");
}

function cleanTextForMatchingVL(text) {
  return String(text || "").toLowerCase().replace(/\s+/g, " ").trim();
}

function getRecentHistoryForPhoneVL(sourceValues, cleanPhone, triggerString) {
  const SEVEN_DAYS_MS = 7 * 24 * 60 * 60 * 1000;
  let triggerTimeMs = new Date().getTime();
  let triggerMessageText = triggerString;
  for (let i = sourceValues.length - 1; i >= 0; i--) {
    const matchPhone = normalizePhoneVL(sourceValues[i][2]);
    const chatText = String(sourceValues[i][4]).trim();
    if (matchPhone === cleanPhone && cleanTextForMatchingVL(chatText).includes(cleanTextForMatchingVL(triggerString))) {
      triggerMessageText = chatText;
      const parsedDate = new Date(sourceValues[i][0]);
      if (!isNaN(parsedDate.getTime())) triggerTimeMs = parsedDate.getTime();
      break;
    }
  }
  const cutoffTimeMs = triggerTimeMs - SEVEN_DAYS_MS;
  const texts = [];
  for (let i = 0; i < sourceValues.length; i++) {
    const matchPhone = normalizePhoneVL(sourceValues[i][2]);
    const chatText = String(sourceValues[i][4]).trim();
    const msgDate = new Date(sourceValues[i][0]);
    if (matchPhone === cleanPhone && chatText) {
      if (isNaN(msgDate.getTime()) || msgDate.getTime() >= cutoffTimeMs) {
        texts.push(`[${sourceValues[i][0]}] ${chatText}`);
      }
    }
  }
  return {
    historyText: texts.join("\n"),
    referenceDateStr: new Date(triggerTimeMs).toLocaleDateString("de-DE"),
    cutoffTimeMs: cutoffTimeMs,
    triggerMessageText: triggerMessageText
  };
}

function getRawPhoneForCleanPhoneVL(sourceValues, cleanPhone) {
  for (let i = 0; i < sourceValues.length; i++) {
    if (normalizePhoneVL(sourceValues[i][2]) === cleanPhone) {
      return String(sourceValues[i][2]).trim();
    }
  }
  return cleanPhone;
}

function save7DayJsonBackupVL(sourceValues, backupSheet, cleanPhone, rawPhoneToUse, cutoffTimeMs) {
  const timestampStr = new Date().toLocaleString("de-DE");
  let savedCount = 0;
  for (let i = 0; i < sourceValues.length; i++) {
    if (normalizePhoneVL(sourceValues[i][2]) === cleanPhone) {
      const msgDate = new Date(sourceValues[i][0]);
      if (isNaN(msgDate.getTime()) || msgDate.getTime() >= cutoffTimeMs) {
        const jsonRaw = String(sourceValues[i][6]).trim();
        if (jsonRaw) {
          backupSheet.appendRow([rawPhoneToUse, jsonRaw, timestampStr]);
          savedCount++;
        }
      }
    }
  }
}