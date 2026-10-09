function deleteOldLogRows2() {

  const SHEET_NAME = "Log";
  const DAYS_TO_KEEP = 7;
  
  // BigQuery Konfiguration
  const PROJECT_ID = "l8-datenspeicherung-langzeit";
  const DATASET_ID = "whatsapp_archive";
  const TABLE_ID = "logs";

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  if (!sheet) {
    Logger.log("FEHLER: Log nicht gefunden.");
    return;
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  Logger.log("========================================");
  Logger.log("Letzte Zeile: " + lastRow);

  const now = new Date();
  const cutoff = new Date(now.getTime() - DAYS_TO_KEEP * 24 * 60 * 60 * 1000);

  Logger.log("Jetzt      : " + now);
  Logger.log("Stichtag   : " + cutoff);

  // ===== TEST DER ERSTEN 10 ZEILEN (Exakt wie in deinem Original) =====

  Logger.log("========== TEST DATUM ==========");

  for (let r = 2; r <= Math.min(lastRow, 11); r++) {

    const value = sheet.getRange(r, 1).getValue();
    const display = sheet.getRange(r, 1).getDisplayValue();

    Logger.log("----------------------------");
    Logger.log("Zeile: " + r);
    Logger.log("getValue(): " + value);
    Logger.log("Display  : " + display);
    Logger.log("Typ      : " + typeof value);
    Logger.log("Date?    : " + (value instanceof Date));

    if (value instanceof Date) {
      Logger.log("Millis   : " + value.getTime());
      Logger.log("Alt?     : " + (value.getTime() < cutoff.getTime()));
    }

  }

  Logger.log("========== BEGIN CLEANUP ==========");

  const values = sheet.getDataRange().getValues();

  const keep = [];
  keep.push(values[0]); // Header-Zeile behalten

  const rowsToInsert = [];
  let deleted = 0;

  for (let i = 1; i < values.length; i++) {

    const d = values[i][0];

    if (!(d instanceof Date)) {
      keep.push(values[i]);
      continue;
    }

    if (d.getTime() < cutoff.getTime()) {
      deleted++;

      // Spalte G (Index 6) ist der JSON-String
      let jsonRaw = values[i][6];

      if (typeof jsonRaw === "string" && jsonRaw.trim() !== "") {
        // Eventuell doppelt maskierte Anführungszeichen bereinigen
        jsonRaw = jsonRaw.replace(/""/g, '"');
        
        try {
          // Prüfen, ob es valides JSON ist und für BigQuery formatieren
          const parsed = JSON.parse(jsonRaw);
          rowsToInsert.push({
            json: {
              timestamp: d.toISOString(),
              payload: JSON.stringify(parsed)
            }
          });
        } catch (e) {
          Logger.log("Warnung in Zeile " + (i + 1) + ": Ungültiges JSON in Spalte G.");
        }
      }

    } else {
      keep.push(values[i]);
    }

  }

  Logger.log("Zu löschen / an BigQuery senden: " + deleted);

  if (deleted > 0) {

    // ===== 1. ERST IN BIGQUERY SPEICHERN =====
    if (rowsToInsert.length > 0) {
      try {
        insertIntoBigQuery(PROJECT_ID, DATASET_ID, TABLE_ID, rowsToInsert);
        Logger.log("Erfolgreich " + rowsToInsert.length + " Datensätze in BigQuery gespeichert.");
      } catch (err) {
        Logger.log("FEHLER bei BigQuery Übertragung: " + err.toString());
        Logger.log("Löschvorgang ABGEBROCHEN, um Datenverlust zu vermeiden.");
        return; // Stoppt sofort! Nichts im Sheet wird gelöscht.
      }
    }

    // ===== 2. DANN SHEET CLEANUP (Exakt wie in deinem Original) =====
    sheet.clearContents();

    sheet.getRange(1, 1, keep.length, lastCol).setValues(keep);

    Logger.log("Cleanup beendet.");

  } else {

    Logger.log("Keine Zeilen gelöscht.");

  }

}

/**
 * Hilfsfunktion: Fügt Datensätze über den BigQuery-Streaming-Dienst ein
 */
function insertIntoBigQuery(projectId, datasetId, tableId, rows) {
  const insertRequest = BigQuery.newTableDataInsertAllRequest();
  insertRequest.rows = rows;

  const response = BigQuery.Tabledata.insertAll(insertRequest, projectId, datasetId, tableId);

  if (response.insertErrors && response.insertErrors.length > 0) {
    throw new Error("BigQuery Insert Error: " + JSON.stringify(response.insertErrors));
  }
}