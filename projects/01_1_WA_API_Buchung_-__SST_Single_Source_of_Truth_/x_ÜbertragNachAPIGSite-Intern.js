function syncTabsToPortalSheet() {
  // ID des Ziel-Sheets (aus deiner URL):
  const TARGET_SHEET_ID = "1SM7VSAq2Zz99IXKJ3bwV587wRQ7WQRMSF6j90XgYVWk";
  
  // Die 4 Tabs, die 1:1 übertragen werden sollen:
  const tabsToSync = ["RE", "VL", "wgb", "d"];

  const sourceSS = SpreadsheetApp.getActiveSpreadsheet();
  const targetSS = SpreadsheetApp.openById(TARGET_SHEET_ID);

  tabsToSync.forEach(tabName => {
    const sourceSheet = sourceSS.getSheetByName(tabName);
    if (!sourceSheet) return; // Falls ein Tab nicht existiert, überspringen

    // Daten aus der Quelldatei holen
    const data = sourceSheet.getDataRange().getValues();

    if (data.length === 0) return;

    // Ziel-Tab suchen oder neu anlegen
    let targetSheet = targetSS.getSheetByName(tabName);
    if (!targetSheet) {
      targetSheet = targetSS.insertSheet(tabName);
    }

    // Ziel-Tab vor dem Einfügen leeren (damit keine alten Reste bleiben)
    targetSheet.clearContents();

    // Daten 1:1 in das Ziel-Tab schreiben
    targetSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  });

  Logger.log("Synchronisation erfolgreich abgeschlossen!");
}