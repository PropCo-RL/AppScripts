function doPost(e) {
  try {
    const rawData = e.postData.contents;
    const d = JSON.parse(rawData);
    const change = d.entry?.[0]?.changes?.[0];
    const val = change?.value;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName("Log") || ss.insertSheet("Log");

    // 📩 Eingehende Nachrichten (Kunde schreibt/klickt)
    if (val?.messages) {
      const m = val.messages[0];
      const msgId = m.id || "";
      const bodyText = m.text?.body || m.button?.text || "";

      // NEU: Nur speichern, wenn die Nachrichten-ID noch nicht im Log steht
      if (!msgId || !isDuplicateMsgId(sh, msgId)) {
        sh.appendRow([
          new Date(),
          "in",
          m.from || "",
          m.type || "",
          bodyText,   // <--- HIER steht jetzt der saubere Text ("Hallo, vielen Dank...")
          msgId,
          rawData     // Das volle JSON am Ende für alle Fälle
        ]);
      }
    }

    // 📤 Ausgehende Nachrichten (Echo von dir)
    if (change?.field === "smb_message_echoes") {
      const m = val?.message_echoes?.[0];
      const msgId = m?.id || "";
      // Zieht den Text aus dem Body des Echos
      const echoBody = m?.text?.body || m?.button?.text || "[Template]";

      // NEU: Nur speichern, wenn die Nachrichten-ID noch nicht im Log steht
      if (!msgId || !isDuplicateMsgId(sh, msgId)) {
        sh.appendRow([
          new Date(),
          "out",
          m?.to || "",
          m?.type || "",
          echoBody,   // <--- Auch hier steht jetzt der Text lesbar
          msgId,
          rawData     // Das volle JSON am Ende
        ]);
      }
    }

  } catch (err) {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Log");
    sh.appendRow([new Date(), "ERROR", "", "", err.toString(), "", e.postData.contents]);
  }

  return ContentService.createTextOutput("OK");
}

// NEU: Hilfsfunktion prüft nur die letzten 10 Zeilen in Spalte F auf Duplikate
function isDuplicateMsgId(sheet, msgId) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return false;

  const startRow = Math.max(2, lastRow - 10);
  const numRows = lastRow - startRow + 1;
  
  const existingIds = sheet.getRange(startRow, 6, numRows, 1).getValues().flat();
  return existingIds.includes(msgId);
}