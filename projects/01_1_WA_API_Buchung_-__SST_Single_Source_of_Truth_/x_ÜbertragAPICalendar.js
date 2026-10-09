// ============================================================================
// SKRIPT: EXPORT VON RE UND VL ZUM ZIEL-SHEET ('inRE')
// ============================================================================

function exportDataToInRE() {
  // ID des Ziel-Spreadsheets (aus der Ziel-URL)
  const TARGET_SPREADSHEET_ID = "11gnOwdzRtb9Jng-s10tUYZ1tBfuOLYf75-pnrfGC5P8";
  const TARGET_SHEET_NAME = "inRE";
  const SOURCE_SHEET_NAMES = ["RE", "VL"];

  const sourceSS = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Ziel-Spreadsheet öffnen
  let targetSS;
  try {
    targetSS = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  } catch (e) {
    throw new Error("Konnte das Ziel-Spreadsheet nicht öffnen. Bitte überprüfe die ID und die Zugriffsrechte.");
  }

  const targetSheet = targetSS.getSheetByName(TARGET_SHEET_NAME);
  if (!targetSheet) {
    throw new Error(`Das Ziel-Tabellenblatt '${TARGET_SHEET_NAME}' wurde in der Zieldatei nicht gefunden.`);
  }

  const exportData = [];

  // 2. Daten aus allen definierten Quell-Sheets ("RE" und "VL") auslesen
  SOURCE_SHEET_NAMES.forEach(sheetName => {
    const sheet = sourceSS.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`⚠️ Quell-Sheet '${sheetName}' wurde nicht gefunden.`);
      return;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log(`Keine Daten im Sheet '${sheetName}' vorhanden.`);
      return;
    }

    // Wir lesen den Bereich ab Zeile 2 bis Spalte AF (Spalte 32)
    const range = sheet.getRange(2, 1, lastRow - 1, 32);
    const values = range.getValues();

    values.forEach(row => {
      const lodgifyId  = row[18]; // Spalte S  (Index 18): Lodgify ID
      const lxInvoiceId = row[19]; // Spalte T  (Index 19): LX Rechnung ID
      const reStart     = row[27]; // Spalte AB (Index 27): LX RE Start
      const reEnde      = row[28]; // Spalte AC (Index 28): LX RE Ende
      const reStatus    = row[31]; // Spalte AF (Index 31): RE Status

      // Nur Zeilen übernehmen, bei denen mindestens eine ID vorhanden ist (verhindert leere Zeilen)
      if (lodgifyId || lxInvoiceId) {
        exportData.push([
          lodgifyId,
          lxInvoiceId,
          reStart,
          reEnde,
          reStatus
        ]);
      }
    });
  });

  if (exportData.length === 0) {
    Logger.log("Keine Daten zum Übertragen gefunden.");
    return;
  }

  // 3. Alte Daten im Ziel-Sheet ab Zeile 2 löschen (Überschriften in Zeile 1 bleiben stehen)
  const targetLastRow = targetSheet.getLastRow();
  if (targetLastRow > 1) {
    targetSheet.getRange(2, 1, targetLastRow - 1, 5).clearContent();
  }

  // 4. Gesammelte Daten ab Zeile 2 im Ziel-Sheet eintragen
  targetSheet.getRange(2, 1, exportData.length, 5).setValues(exportData);

  Logger.log(`✅ Erfolgreich ${exportData.length} Zeilen in das Sheet '${TARGET_SHEET_NAME}' übertragen.`);
}