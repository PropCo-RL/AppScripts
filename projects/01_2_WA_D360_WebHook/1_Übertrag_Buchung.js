// =================================================================
// HAUPTFUNKTION: SEQUENTIELLE VERARBEITUNG DER BUCHUNGEN
// =================================================================
function processTriggeredBookingsSequentially() {
  const TARGET_SPREADSHEET_ID = "1mMEVsVFIx5VxIV3EyJYlba_G0fgEX9t0tR2KbFkm-nU"; 
  const SOURCE_LOG_SHEET_NAME = "Log";                                        
  const TARGET_RE_SHEET_NAME  = "RE";                                         
  
  // 🎯 NEUE SEPARATE BACKUP-DATEI ID
  const BACKUP_SPREADSHEET_ID = "1gBNQOi4qCHiFfEdZwibr4rW-371JtyhRfQaae1uxDmU";
  const TARGET_BACKUP_SHEET_NAME = "Log_Backup";

  const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('Gemini-Key');
  const D360_API_KEY   = PropertiesService.getScriptProperties().getProperty('D360-Key');

  const TRIGGER_STRING = "Super, vielen Dank\nIch bereite jetzt alles vor und sende alles zu. Ca 40 minuten brauche ich";

  const sourceSs = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = sourceSs.getSheetByName(SOURCE_LOG_SHEET_NAME);
  
  let targetSs, targetSheet, backupSs, backupSheet;
  try {
    targetSs = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    targetSheet = targetSs.getSheetByName(TARGET_RE_SHEET_NAME);
    
    // 🎯 BACKUP-DATEI SEPARAT ÖFFNEN:
    backupSs = SpreadsheetApp.openById(BACKUP_SPREADSHEET_ID);
    backupSheet = backupSs.getSheetByName(TARGET_BACKUP_SHEET_NAME);
    if (!backupSheet) {
      backupSheet = backupSs.insertSheet(TARGET_BACKUP_SHEET_NAME);
      backupSheet.appendRow(["Telefonnummer", "JSON_Raw_Data", "Backup_Zeitstempel"]);
    }
  } catch (e) {
    Logger.log("FEHLER beim Öffnen der Ziel- oder Backup-Datei: " + e.toString());
    return;
  }

  if (!sourceSheet || !targetSheet) return;

  const lastSourceRow = sourceSheet.getLastRow();
  if (lastSourceRow <= 1) return;

  const sourceRange = sourceSheet.getRange(2, 1, lastSourceRow - 1, 9);
  const sourceValues = sourceRange.getValues();
  const numbersToProcess = new Set();

  // 1. MARKIERUNG IN SPALTE I SETZEN
  for (let i = 0; i < sourceValues.length; i++) {
    const rowNum    = i + 2;
    const rawPhone  = String(sourceValues[i][2]).trim(); 
    const chatText  = String(sourceValues[i][4]).trim(); 
    const marker    = String(sourceValues[i][8]).trim(); 

    const cleanPhone = normalizePhone(rawPhone);

    if (marker.toLowerCase().includes("mit gemini") || marker.toLowerCase() === "x" || marker.toLowerCase() === "y" || !chatText) {
      continue;
    }

    const isTriggerMatch = cleanTextForMatching(chatText).includes(cleanTextForMatching(TRIGGER_STRING));

    if (isTriggerMatch) {
      sourceSheet.getRange(rowNum, 9).setValue("y - mit gemini");
      sourceValues[i][8] = "y - mit gemini";
      if (cleanPhone) numbersToProcess.add(cleanPhone);
    } else {
      sourceSheet.getRange(rowNum, 9).setValue("x - mit gemini");
      sourceValues[i][8] = "x - mit gemini";
    }
  }

  if (numbersToProcess.size === 0) return;

  Logger.log(`Starte Analyse für ${numbersToProcess.size} Nummer(n)...`);

  // 2. VERARBEITUNG PRO NUMMER
  numbersToProcess.forEach(cleanPhone => {
    // Eingrenzung auf 7 Tage vor Trigger
    const { historyText, referenceDateStr, cutoffTimeMs, triggerMessageText } = getRecentHistoryForPhone(sourceValues, cleanPhone, TRIGGER_STRING);
    const rawPhoneToUse = getRawPhoneForCleanPhone(sourceValues, cleanPhone);

    // 🎯 ERZEUGUNG DER UNIQUE ID FÜR DIESEN BUCHUNGS-DURCHLAUF
    const batchUniqueID = generateUniqueID();

    // -------------------------------------------------------------
    // MODUL 1: RECHNUNGSADRESSE (Text -> Bild-Fallback falls nötig)
    // -------------------------------------------------------------
    Logger.log(`[Modul 1: Adresse] Analysiere Text für ${cleanPhone}...`);
    let invoiceData = callGeminiForInvoice(historyText, null, GEMINI_API_KEY, triggerMessageText);
    
    let isComplete = checkInvoiceIsComplete(invoiceData);

    if (!isComplete) {
      Logger.log(`[Modul 1: Adresse unvollständig] Bild-Fallback für ${cleanPhone}...`);
      const latestImageInfo = getLatestImageBeforeTrigger(sourceValues, cleanPhone, TRIGGER_STRING);

      if (latestImageInfo && latestImageInfo.mediaId && D360_API_KEY) {
        const imageBlob = downloadMediaFrom360Dialog(latestImageInfo.mediaId, D360_API_KEY);
        if (imageBlob) {
          const imageInvoiceData = callGeminiForInvoice(historyText, imageBlob, GEMINI_API_KEY, triggerMessageText);
          if (imageInvoiceData) {
            invoiceData = imageInvoiceData;
            isComplete = checkInvoiceIsComplete(invoiceData);
          }
        }
      }
    }

    // -------------------------------------------------------------
    // MODUL 2: BUCHUNGSDETAILS
    // -------------------------------------------------------------
    Logger.log(`[Modul 2: Buchungsdetails] Analysiere Text mit Bezugsdatum ${referenceDateStr}...`);
    const bookingsArray = callGeminiForBookingDetails(historyText, GEMINI_API_KEY, referenceDateStr, triggerMessageText);

    // -------------------------------------------------------------
    // SPEICHERN & EXPORT
    // -------------------------------------------------------------
    
    // Nur JSONs des 7-Tage-Fensters sichern (in das neue Backup-Sheet)
    save7DayJsonBackupForPhone(sourceValues, backupSheet, cleanPhone, rawPhoneToUse, cutoffTimeMs);

    const doneStatus = isComplete ? "Done" : "Unvollständig";

    // Für jede erkannte Wohnung eine eigene Zeile anlegen
    bookingsArray.forEach((bookingData, index) => {
      const newRow = [
        rawPhoneToUse,                             // Spalte A (1): Telefonnummer
        invoiceData.firma || "",                  // Spalte B (2): Firma
        invoiceData.name || "",                   // Spalte C (3): Name
        invoiceData.strasse || "",                // Spalte D (4): Straße
        invoiceData.plz || "",                    // Spalte E (5): PLZ
        invoiceData.stadt || "",                  // Spalte F (6): Stadt
        invoiceData.land || "Deutschland",        // Spalte G (7): Land
        invoiceData.vat || "",                    // Spalte H (8): VAT
        invoiceData.email || "",                  // Spalte I (9): E-Mail
        doneStatus,                               // Spalte J (10): Status
        bookingData.anreise || "",                // Spalte K (11): Anreise
        bookingData.abreise || "",                // Spalte L (12): Abreise
        bookingData.rhythmus || "",               // Spalte M (13): Rhythmus
        bookingData.apartment || "",              // Spalte N (14): Apartment
        bookingData.preis_pro_nacht || "",        // Spalte O (15): Gesamtpreis / Nacht
        bookingData.preis_pro_person_nacht || "", // Spalte P (16): Einzelpreis / Person / Nacht
        bookingData.preis_pro_monat || "",        // Spalte Q (17): Gesamtpreis / Monat
        "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", // Spalten R bis AH (18 bis 34) leer auffüllen
        batchUniqueID                             // 🎯 Spalte AI (35): Eindeutiger Identifier
      ];

      smartUpsertToRESheet(targetSheet, cleanPhone, newRow);
      Logger.log(`Buchung ${index + 1}/${bookingsArray.length} für ${cleanPhone} mit ID [${batchUniqueID}] eingefügt.`);
    });

    Utilities.sleep(200);
  });

  Logger.log("--- PROZESS ERFOLGREICH BEENDET ---");
}


// =================================================================
// HILFSFUNKTIONEN SPEICHERN & UNIQUE ID GENERATOR
// =================================================================

// Generiert einen eindeutigen Zufalls-String (z. B. ID-9b1e4a2c)
function generateUniqueID() {
  const chars = 'abcdefghijklmnopqrstuvwxyz0123456789';
  let randomStr = '';
  for (let i = 0; i < 8; i++) {
    randomStr += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return `ID-${randomStr}`;
}

function save7DayJsonBackupForPhone(sourceValues, backupSheet, cleanPhone, rawPhoneToUse, cutoffTimeMs) {
  const timestampStr = new Date().toLocaleString("de-DE");
  let savedCount = 0;

  for (let i = 0; i < sourceValues.length; i++) {
    if (normalizePhone(sourceValues[i][2]) === cleanPhone) {
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
  Logger.log(`[Backup] ${savedCount} JSON-Datensätze im Tab 'Log_Backup' gesichert.`);
}

function smartUpsertToRESheet(targetSheet, cleanPhone, newRowData) {
  targetSheet.appendRow(newRowData);
}


// =================================================================
// MODUL 1: KI-AUFRUF FÜR RECHNUNGSADRESSE
// =================================================================
function callGeminiForInvoice(chatText, imageBlob, apiKey, triggerMessageText) {
  const fallback = { firma: "", name: "", strasse: "", plz: "", stadt: "", land: "Deutschland", vat: "", email: "" };
  if (!apiKey) return fallback;

  const parts = [];
  if (imageBlob) {
    parts.push({
      inlineData: {
        mimeType: imageBlob.getContentType() || "image/jpeg",
        data: Utilities.base64Encode(imageBlob.getBytes())
      }
    });
  }

  const promptText = `
Du bist ein Assistent für Adress-Extraktion. 
Extrahiere die RECHNUNGSANSCHRIFT und KUNDEN-E-MAIL aus dem Chat und/oder Bild.

PRIORITÄTEN-REGEL:
1. Falls die Trigger-Nachricht ("${triggerMessageText.replace(/\n/g, " ")}") direkt Anschrifts- oder Adressdaten enthält, haben diese Angaben ABSOLUTE PRIORITÄT.
2. Rechtsformen (GmbH, UG, e.K., Sp. z o.o., etc.) gehören in "firma".
3. Vor- und Nachname der Person gehören in "name" (falls Privatperson, bei "firma" UND "name").
4. E-Mails auf @l8street.com ignorieren.

Antworte NUR als valides JSON:
{
  "firma": "",
  "name": "",
  "strasse": "",
  "plz": "",
  "stadt": "",
  "land": "Deutschland",
  "vat": "",
  "email": ""
}

Chat-Verlauf:
"""
${chatText}
"""
`;

  parts.push({ text: promptText });
  const result = executeGeminiRequest(parts, apiKey);
  return result ? { ...fallback, ...result } : fallback;
}


// =================================================================
// MODUL 2: KI-AUFRUF FÜR BUCHUNGSDETAILS
// =================================================================
function callGeminiForBookingDetails(chatText, apiKey, referenceDateStr, triggerMessageText) {
  const fallbackItem = { anreise: "", abreise: "", rhythmus: "", apartment: "", preis_pro_nacht: "", preis_pro_person_nacht: "", preis_pro_monat: "" };
  if (!apiKey) return [fallbackItem];

  const cleanTriggerMsg = triggerMessageText.toLowerCase();
  
  const containsMultipleKeywords = cleanTriggerMsg.includes("folgende wohnungen") || 
                                    cleanTriggerMsg.includes("folgenden wohnungen") ||
                                    cleanTriggerMsg.includes("wohnungen:");

  let multiBookingInstruction = "";
  if (containsMultipleKeywords) {
    multiBookingInstruction = `
ACHTUNG - MEHRERE WOHNUNGEN GEBUCHT:
In der Trigger-Nachricht wurde explizit auf meherere Wohnungen hingewiesen ("folgende wohnungen"). 
Erstelle FÜR JEDE genannte Wohnung ein eigenes Objekt im JSON-Array.
`;
  } else {
    multiBookingInstruction = `
ACHTUNG - EINZELBUCHUNG:
In der Trigger-Nachricht steht KEIN Hinweis auf meherere Wohnungen ("folgende wohnungen"). 
Gib EXAKT EIN OBJEKT im JSON-Array zurück (für die primäre / zuletzt genannte Wohnung). Ignoriere eventuelle Mehrfach-Listen aus älteren Nachrichten.
`;
  }

  const promptText = `
Du bist ein Assistent für Buchungsdaten.
Das Bezugsdatum (HEUTE) für relative Zeitangaben ist: **${referenceDateStr}**.

PRIORITÄTEN-REGEL:
Die Angaben in der Trigger-Nachricht ("${triggerMessageText.replace(/\n/g, " ")}") haben ABSOLUTE PRIORITÄT vor älteren Nachrichten im Chat-Verlauf.

${multiBookingInstruction}

FELDER PRO WOHNUNG (Bezugsdatum für relative Angaben ist ${referenceDateStr}):
- "anreise": Anreisedatum in DD.MM.YYYY.
- "abreise": Abreisedatum in DD.MM.YYYY.
- "rhythmus": Abrechnungsrhythmus (z.B. "Monatlich", "Fixes Enddatum", "Wöchentlich").
- "apartment": Kürzel aus dem Link mit "l8street.com/a/" (NUR der Code nach "/a/").
- "preis_pro_person_nacht": Einzelpreis pro Person/Nacht (z.B. "25€").
- "preis_pro_nacht": GESAMTPREIS der Unterkunft pro Nacht.
- "preis_pro_monat": Vereinbarter Pauschalpreis pro Monat.

ANTWORTE AUSSCHLIESSLICH ALS EIN JSON-ARRAY VON OBJEKTEN:
[
  {
    "anreise": "",
    "abreise": "",
    "rhythmus": "",
    "apartment": "",
    "preis_pro_nacht": "",
    "preis_pro_person_nacht": "",
    "preis_pro_monat": ""
  }
]

Chat-Verlauf:
"""
${chatText}
"""
`;

  const parts = [{ text: promptText }];
  const result = executeGeminiRequest(parts, apiKey);

  if (Array.isArray(result) && result.length > 0) {
    return result.map(item => ({ ...fallbackItem, ...item }));
  } else if (result && typeof result === "object") {
    return [{ ...fallbackItem, ...result }];
  }

  return [fallbackItem];
}


// =================================================================
// GEMINI API NETWORK & ROBUSTER JSON PARSER
// =================================================================
function executeGeminiRequest(parts, apiKey) {
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

      return safeParseJSON(rawText);
    }
  } catch (e) {
    Logger.log("Fehler bei Gemini API Aufruf: " + e.toString());
  }
  return null;
}

function safeParseJSON(rawText) {
  try {
    return JSON.parse(rawText);
  } catch (e) {
    try {
      let cleanText = rawText
        .replace(/[\u0000-\u001F\u007F-\u009F]/g, "")
        .replace(/\r?\n/g, " ")
        .replace(/",\s*}/g, '"}')
        .replace(/",\s*]/g, '"]');
      return JSON.parse(cleanText);
    } catch (e2) {
      Logger.log("🚨 Sicheres JSON-Parsing fehlgeschlagen: " + e2.toString());
      return null;
    }
  }
}


// =================================================================
// AUTOMATISCHER LÖSCH-TRIGGER (> 60 TAGE)
// =================================================================
function deleteOldLogRows() {
  const SOURCE_LOG_SHEET_NAME = "Log";
  const DAYS_TO_KEEP = 60;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SOURCE_LOG_SHEET_NAME);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const now = new Date();
  const cutoffTimeMs = now.getTime() - (DAYS_TO_KEEP * 24 * 60 * 60 * 1000);

  const dateValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let deletedCount = 0;

  for (let i = dateValues.length - 1; i >= 0; i--) {
    const rawValue = dateValues[i][0];
    if (!rawValue) continue;

    const rowDate = new Date(rawValue);
    if (!isNaN(rowDate.getTime()) && rowDate.getTime() < cutoffTimeMs) {
      sheet.deleteRow(i + 2);
      deletedCount++;
    }
  }

  Logger.log(`Cleanup beendet: ${deletedCount} Zeile(n) gelöscht.`);
}


// =================================================================
// WEITERE HELPER FUNKTIONEN
// =================================================================

function checkInvoiceIsComplete(extracted) {
  return Boolean(
    (extracted.firma || extracted.name) &&
    extracted.strasse &&
    extracted.plz &&
    extracted.stadt &&
    extracted.email
  );
}

function normalizePhone(phone) {
  return String(phone || "").replace(/\D/g, "");
}

function cleanTextForMatching(text) {
  return String(text || "").toLowerCase().replace(/\s+/g, " ").trim();
}

function getRecentHistoryForPhone(sourceValues, cleanPhone, triggerString) {
  const SEVEN_DAYS_MS = 7 * 24 * 60 * 60 * 1000;
  let triggerTimeMs = new Date().getTime();
  let triggerMessageText = triggerString;

  for (let i = sourceValues.length - 1; i >= 0; i--) {
    const matchPhone = normalizePhone(sourceValues[i][2]);
    const chatText   = String(sourceValues[i][4]).trim();
    if (matchPhone === cleanPhone && cleanTextForMatching(chatText).includes(cleanTextForMatching(triggerString))) {
      triggerMessageText = chatText;
      const parsedDate = new Date(sourceValues[i][0]);
      if (!isNaN(parsedDate.getTime())) triggerTimeMs = parsedDate.getTime();
      break;
    }
  }

  const cutoffTimeMs = triggerTimeMs - SEVEN_DAYS_MS;
  const texts = [];

  for (let i = 0; i < sourceValues.length; i++) {
    const matchPhone = normalizePhone(sourceValues[i][2]);
    const chatText   = String(sourceValues[i][4]).trim();
    const msgDate    = new Date(sourceValues[i][0]);

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

function getRawPhoneForCleanPhone(sourceValues, cleanPhone) {
  for (let i = 0; i < sourceValues.length; i++) {
    if (normalizePhone(sourceValues[i][2]) === cleanPhone) {
      return String(sourceValues[i][2]).trim();
    }
  }
  return cleanPhone;
}

function getLatestImageBeforeTrigger(sourceValues, cleanPhone, triggerString) {
  let triggerIndex = -1;
  const cleanTrigger = cleanTextForMatching(triggerString);

  for (let i = sourceValues.length - 1; i >= 0; i--) {
    const matchPhone = normalizePhone(sourceValues[i][2]);
    const chatText   = String(sourceValues[i][4]).trim();
    if (matchPhone === cleanPhone && cleanTextForMatching(chatText).includes(cleanTrigger)) {
      triggerIndex = i;
      break;
    }
  }

  if (triggerIndex === -1) triggerIndex = sourceValues.length;

  for (let i = triggerIndex; i >= 0; i--) {
    if (normalizePhone(sourceValues[i][2]) === cleanPhone) {
      const jsonString = String(sourceValues[i][6]).trim();
      if (!jsonString) continue;

      try {
        const payload = JSON.parse(jsonString);
        if (payload.entry && payload.entry[0].changes) {
          const value = payload.entry[0].changes[0].value;
          const messages = value.messages || value.message_echoes;
          
          if (messages && messages.length > 0) {
            const msg = messages[0];
            if (msg.type === "image" && msg.image && msg.image.id) {
              return { mediaId: msg.image.id, mimeType: msg.image.mime_type || "image/jpeg" };
            }
          }
        }
      } catch (e) {}
    }
  }
  return null;
}

function downloadMediaFrom360Dialog(mediaId, d360ApiKey) {
  const getUrlApi = `https://waba-v2.360dialog.io/${mediaId}`;
  try {
    const response = UrlFetchApp.fetch(getUrlApi, { headers: { "D360-API-KEY": d360ApiKey }, muteHttpExceptions: true });
    if (response.getResponseCode() === 200) {
      const json = JSON.parse(response.getContentText());
      if (json.url) {
        let downloadUrl = json.url.replace("https://lookaside.fbsbx.com", "https://waba-v2.360dialog.io").replace(/\\/g, "");
        const imageResponse = UrlFetchApp.fetch(downloadUrl, { headers: { "D360-API-KEY": d360ApiKey }, muteHttpExceptions: true });
        if (imageResponse.getResponseCode() === 200) return imageResponse.getBlob();
      }
    }
  } catch (e) {
    Logger.log("Fehler beim 360dialog Medien-Download: " + e.toString());
  }
  return null;
}