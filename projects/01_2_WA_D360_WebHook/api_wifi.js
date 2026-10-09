// ==========================================
// KONFIGURATION & IDS
// ==========================================
const TRIGGER_STRING = "Danke! Ich prüfe die wlan Daten jetzt";
const TARGET_SPREADSHEET_ID = "17MZ0xZXIPNEEQ3Ji2VMEGpRKu1HIOhVm_dIv1GygW-M";
const DRIVE_FOLDER_ID = "1S5pJ_S31KJXhig0F9Y278PrN_RkEFUVc";

/**
 * HAUPTFUNKTION: Verarbeitet neu eingetroffene Zeilen im Sheet (Alle 30 Min)
 */
function processNewWiFiEntries() {
  Logger.log("=== VERARBEITUNG GESTARTET ===");
  
  const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('Gemini-Key');
  const D360_API_KEY = PropertiesService.getScriptProperties().getProperty('D360-Key');

  if (!GEMINI_API_KEY) {
    Logger.log("FEHLER: 'Gemini-Key' wurde nicht in den Script Properties gefunden!");
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("log");
  if (!sheet) {
    Logger.log("FEHLER: Tabellenblatt 'log' wurde nicht gefunden.");
    return;
  }

  const data = sheet.getDataRange().getValues();
  Logger.log(`Gesamte Zeilen im Sheet: ${data.length}`);

  for (let i = 1; i < data.length; i++) {
    const rowIndex = i + 1;
    const messageText = data[i][4] ? String(data[i][4]).trim() : ""; // Spalte E
    const currentStatusInS = data[i][18] ? String(data[i][18]).trim() : ""; // Spalte S
    const senderPhone = data[i][2] ? String(data[i][2]).trim() : ""; // Spalte C

    if (currentStatusInS === "") {
      Logger.log(`--------------------------------------------------`);
      Logger.log(`Verarbeite Zeile ${rowIndex} | Absender: ${senderPhone}`);

      const isMatch = cleanTextForMatching(messageText).includes(cleanTextForMatching(TRIGGER_STRING));

      if (isMatch) {
        Logger.log(`-> MATCH! Schreibe 'y' in Spalte S.`);
        sheet.getRange(rowIndex, 19).setValue("y");

        processWlanCheck(sheet, senderPhone, GEMINI_API_KEY, D360_API_KEY);

      } else {
        Logger.log(`-> Kein Match. Schreibe 'x' in Spalte S.`);
        sheet.getRange(rowIndex, 19).setValue("x");
      }
    }
  }
  Logger.log("=== VERARBEITUNG ABGESCHLOSSEN ===");
}

// ==========================================
// CHAT DER LETZTEN 24H PRÜFEN & BILDER DOWNLOADEN
// ==========================================
function processWlanCheck(logSheet, targetPhone, geminiApiKey, d360ApiKey) {
  Logger.log(`--> Starte Chat-Filterung (letzte 24h) für Nummer: ${targetPhone}`);
  const now = new Date();
  const twentyFourHoursAgo = new Date(now.getTime() - (24 * 60 * 60 * 1000));

  const data = logSheet.getDataRange().getValues();
  let chatTexts = [];
  let imageParts = [];

  const driveFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  let matchCount = 0;

  for (let i = 1; i < data.length; i++) {
    const rowDate = new Date(data[i][0]); // Spalte A: Zeitstempel
    const phone = data[i][2] ? String(data[i][2]).trim() : ""; // Spalte C
    const msgContent = data[i][4] ? String(data[i][4]).trim() : ""; // Spalte E
    const jsonRaw = data[i][6] ? String(data[i][6]).trim() : ""; // Spalte G (JSON Raw)

    if (normalizePhone(phone) === normalizePhone(targetPhone) && rowDate >= twentyFourHoursAgo) {
      matchCount++;
      
      if (msgContent && msgContent !== "[Template]") {
        chatTexts.push(`[${data[i][0]}] ${msgContent}`);
      }

      if (jsonRaw) {
        try {
          const payload = JSON.parse(jsonRaw);
          if (payload.entry && payload.entry[0].changes) {
            const value = payload.entry[0].changes[0].value;
            const messages = value.messages || value.message_echoes;
            if (messages && messages.length > 0) {
              const msg = messages[0];
              if (msg.type === "image" && msg.image && msg.image.id) {
                let imgBlob = null;
                if (d360ApiKey) {
                  imgBlob = downloadMediaFrom360Dialog(msg.image.id, d360ApiKey);
                }

                if (imgBlob) {
                  const fileName = `WLAN_Foto_${targetPhone}_${rowDate.getTime()}.jpg`;
                  imgBlob.setName(fileName);
                  const savedFile = driveFolder.createFile(imgBlob);

                  imageParts.push({
                    inlineData: {
                      mimeType: imgBlob.getContentType() || "image/jpeg",
                      data: Utilities.base64Encode(imgBlob.getBytes())
                    }
                  });
                }
              }
            }
          }
        } catch (jsonErr) {}
      }
    }
  }

  // Daten aus dem Tab 'data' für den Bestandsabgleich auslesen
  const referenceDataText = getReferenceDataAsText();

  const fullChatHistory = chatTexts.join("\n");
  Logger.log(`Gefundene Verlaufseinträge: ${matchCount} | Heruntergeladene Bilder: ${imageParts.length}`);

  if (fullChatHistory.length > 0 || imageParts.length > 0) {
    callGeminiForWlanDetails(fullChatHistory, imageParts, geminiApiKey, targetPhone, referenceDataText);
  } else {
    Logger.log("Keine Daten der letzten 24h für Gemini vorhanden.");
  }
}

// ==========================================
// REFERENZ-DATEN AUS TAB 'data' FÜR GEMINI AUFBEREITEN
// ==========================================
function getReferenceDataAsText() {
  try {
    const targetSs = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const dataSheet = targetSs.getSheetByName("data");
    if (!dataSheet) return "";

    const values = dataSheet.getDataRange().getValues();
    let refText = "REFERENZ-LISTE DER BESTEHENDEN APARTMENTS & BESTANDSDATEN:\n";
    for (let i = 1; i < values.length; i++) {
      const apt = values[i][0]; // Spalte A: Apartment Name
      const imei = values[i][1]; // Spalte B: IMEI Bestand
      const sim = values[i][2]; // Spalte C: SIM Bestand
      if (apt) {
        refText += `Apt: "${apt}" | Bestands-IMEI: "${imei}" | Bestands-SIM: "${sim}"\n`;
      }
    }
    return refText;
  } catch (e) {
    Logger.log("Fehler beim Auslesen der Referenzdaten: " + e.toString());
    return "";
  }
}

// ==========================================
// GEMINI API AUFRUF MIT MITGELIEFERTER REFERENZ-LISTE
// ==========================================
function callGeminiForWlanDetails(chatText, imageParts, apiKey, targetPhone, referenceDataText) {
  Logger.log("--> Rufe Gemini API auf...");

  const parts = [];
  imageParts.forEach(part => parts.push(part));

  const promptText = `
Du bist ein Assistent zur Extraktion von WLAN-Router und SIM-Karten Daten.

${referenceDataText}

AUFGABE:
1. Extrahiere die neuen SIM/IMEI Daten sowie die Orts/Apartment-Bezeichnung aus dem Chattext/Bildern.
2. Gleiche die Orts/Apartment-Bezeichnung intelligent mit der REFERENZ-LISTE ab.
3. Wenn der Ort im Chat auf ein Apartment in der Referenz-Liste passt (z. B. "pforzheim Calwer Straße 112 Apartment F3" -> "Pforzheim CS112 F3"), übernehme die Bestands-IMEI und Bestands-SIM aus der Referenz-Liste als "sim_bestand" und "imei_bestand".
4. Falls alte Daten explizit im Chat genannt/gezeigt werden, nutze diese bevorzugt als Bestand.

ANTWORTE AUSSCHLIESSLICH ALS EIN JSON-ARRAY VON OBJEKTEN:
[
  {
    "apartment": "Name des Apartments aus dem Chat oder der Referenz-Liste",
    "sim_neu": "NEUE SIM-Karten Nummer (aus Chat/Bild)",
    "imei_neu": "NEUE IMEI Nummer des Routers (aus Chat/Bild)",
    "sim_bestand": "Bestehende/Alte SIM-Karten Nummer",
    "imei_bestand": "Bestehende/Alte IMEI Nummer des Routers"
  }
]
Falls Angaben fehlen, trage "" ein.

Chat-Verlauf:
"""${chatText}"""`;

  parts.push({ text: promptText });

  const result = executeGeminiRequest(parts, apiKey);

  if (result) {
    const extractedData = Array.isArray(result) ? result : [result];
    Logger.log(`Erfolgreich extrahiert! Anzahl Standorte: ${extractedData.length}`);

    saveToTargetSheet(extractedData, targetPhone);
  } else {
    Logger.log("FEHLER: Gemini API lieferte kein gültiges Ergebnis.");
  }
}

// ==========================================
// GEMINI API NETWORK & ROBUSTER JSON PARSER
// ==========================================
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
  try { return JSON.parse(rawText); } catch (e) {
    try {
      let cleanText = rawText.replace(/[\u0000-\u001F\u007F-\u009F]/g, "").replace(/\r?\n/g, " ");
      return JSON.parse(cleanText);
    } catch (e2) { return null; }
  }
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
  } catch (e) {}
  return null;
}

// ==========================================
// SPEICHERN IM ZIEL-GOOGLE-SHEET ("log")
// Spalten: A=Zeitstempel, B=Nummer, C=Apartment, D=SIM Neu, E=IMEI Neu, F=SIM Bestand, G=IMEI Bestand, H=Email versenden
// ==========================================
function saveToTargetSheet(extractedData, phone) {
  Logger.log("--> Speichere extrahierte Daten im Ziel-Sheet...");
  
  try {
    const targetSpreadsheet = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const targetLogSheet = targetSpreadsheet.getSheetByName("log");

    if (!targetLogSheet) return;

    const timestamp = new Date();

    extractedData.forEach(item => {
      targetLogSheet.appendRow([
        timestamp,             // Spalte A
        phone,                 // Spalte B
        item.apartment || "",  // Spalte C
        item.sim_neu || "",    // Spalte D
        item.imei_neu || "",   // Spalte E
        item.sim_bestand || "",// Spalte F
        item.imei_bestand || "",// Spalte G
        ""                     // Spalte H (bleibt leer für deine manuelle Freigabe 'x')
      ]);
    });
    Logger.log("Erfolgreich in Ziel-Sheet geschrieben.");
  } catch (e) {
    Logger.log("FEHLER beim Schreiben in das Ziel-Sheet: " + e.toString());
  }
}

function cleanTextForMatching(text) { return String(text || "").toLowerCase().replace(/\s+/g, " ").trim(); }
function normalizePhone(phone) { return String(phone || "").replace(/\D/g, ""); }