// ============================================================================
// EIGENSTÄNDIGES SYNC-SCRIPT:
// Überträgt E-Mail-Versendestatus (und ggf. die komplette Zeile) 
// aus dem aktuellen Sheet in das externe Webhook-Sheet (Tab 'RE').
// ============================================================================

function CCC_syncEmailStatusToExternalSheet() {
  // ID des externen Ziel-Sheets aus deiner URL:
  const EXTERNAL_TARGET_SPREADSHEET_ID = "1bMPVHjGNnwWDH0KXUyB47_VbLcq_lsFK1jSbdj_GmH0";
  
  const sourceSs = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheetRE = sourceSs.getSheetByName("RE");

  if (!sourceSheetRE) {
    Logger.log("🚨 Quellsheet 'RE' wurde nicht gefunden.");
    return;
  }

  const lastSourceRow = sourceSheetRE.getLastRow();
  if (lastSourceRow < 2) {
    Logger.log("Keine Daten im Quellsheet vorhanden.");
    return;
  }

  // Quell-Daten lesen (Spalte A bis V)
  const sourceValues = sourceSheetRE.getRange(2, 1, lastSourceRow - 1, 22).getValues();

  // Externes Sheet öffnen
  let extSheetRE;
  try {
    const extSs = SpreadsheetApp.openById(EXTERNAL_TARGET_SPREADSHEET_ID);
    extSheetRE = extSs.getSheetByName("RE");
    if (!extSheetRE) {
      Logger.log("🚨 Externes Sheet 'RE' wurde nicht gefunden.");
      return;
    }
  } catch (e) {
    Logger.log("🚨 Fehler beim Öffnen des externen Sheets: " + e.toString());
    return;
  }

  let lastExtRow = extSheetRE.getLastRow();
  
  // Wenn das externe Sheet noch leer ist (außer Header), Hilfs-Array initialisieren
  let extPhoneValues = [];
  let extStatusValues = [];

  if (lastExtRow >= 2) {
    extPhoneValues = extSheetRE.getRange(2, 1, lastExtRow - 1, 1).getValues();
    extStatusValues = extSheetRE.getRange(2, 22, lastExtRow - 1, 1).getValues();
  }

  let updatedCount = 0;
  let addedCount = 0;

  // Schleife durch alle Zeilen des Quell-Sheets
  sourceValues.forEach((row, index) => {
    const sourcePhone = row[0] ? row[0].toString().trim() : "";        // Spalte A
    const sourceStatus = row[21] ? row[21].toString().trim() : "";    // Spalte V

    // Nur verarbeiten, wenn eine Telefonnummer und ein gesendeter Status vorliegen
    if (!sourcePhone || !sourceStatus || !sourceStatus.startsWith("Gesendet an")) {
      return;
    }

    const cleanSourcePhone = sourcePhone.replace(/\D/g, "");
    if (!cleanSourcePhone) return;

    let phoneFound = false;

    // 1. Suche nach passender Telefonnummer im externen Sheet
    for (let i = 0; i < extPhoneValues.length; i++) {
      const extPhone = extPhoneValues[i][0] ? extPhoneValues[i][0].toString().trim() : "";
      const cleanExtPhone = extPhone.replace(/\D/g, "");

      if (cleanExtPhone && cleanExtPhone === cleanSourcePhone) {
        phoneFound = true;
        const targetRowIndex = i + 2;
        const currentExtStatus = extStatusValues[i][0] ? extStatusValues[i][0].toString().trim() : "";

        // Nur schreiben, wenn sich der Status unterscheidet
        if (currentExtStatus !== sourceStatus) {
          // 🎯 EXAKT NUR SPALTE 22 (V) BESCHREIBEN
          extSheetRE.getRange(targetRowIndex, 22).setValue(sourceStatus);
          
          extStatusValues[i][0] = sourceStatus; 
          updatedCount++;
          Logger.log(`[SYNC UPDATE] Nummer ${sourcePhone} -> Status in Zelle V${targetRowIndex} im externen Sheet aktualisiert.`);
        }
        break; // Passende Zeile gefunden, Schleife beenden
      }
    }

    // 2. Falls die Telefonnummer im externen Sheet NOCH NICHT existiert:
    if (!phoneFound) {
      // Fügt die komplette Zeile (Spalte A bis V) unten an den Tab 'RE' an
      extSheetRE.appendRow(row);
      
      // Lokale Arrays aktualisieren, damit bei mehrfachen Treffern in derselben Ausführung nicht doppelt angehängt wird
      extPhoneValues.push([sourcePhone]);
      extStatusValues.push([sourceStatus]);
      
      addedCount++;
      Logger.log(`[SYNC NEU] Nummer ${sourcePhone} war nicht im externen Sheet vorhanden. Gesamte Zeile neu angelegt.`);
    }
  });

  Logger.log(`--- SYNC BEENDET: ${updatedCount} Status-Einträge aktualisiert, ${addedCount} neue Zeilen angelegt. ---`);
}