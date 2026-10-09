function syncTabsToSecondPortalSheet() {
  // ID des zweiten Ziel-Sheets (aus deiner neuen URL):
  const SECOND_TARGET_SHEET_ID = "1VAw-KR7KlxXjjaj-0-iNZ0g-VvKzwg9uKaUIuop71sE";
  
  // Die 4 Tabs, die 1:1 übertragen werden sollen:
  const tabsToSyncSecond = ["RE", "VL", "wgb", "d"];

  const sourceSSSecond = SpreadsheetApp.getActiveSpreadsheet();
  const targetSSSecond = SpreadsheetApp.openById(SECOND_TARGET_SHEET_ID);

  tabsToSyncSecond.forEach(tabName => {
    const sourceSheet = sourceSSSecond.getSheetByName(tabName);
    if (!sourceSheet) return; // Falls ein Tab nicht existiert, überspringen

    // Daten aus der Quelldatei holen
    const data = sourceSheet.getDataRange().getValues();

    if (data.length === 0) return;

    // Ziel-Tab suchen oder neu anlegen
    let targetSheet = targetSSSecond.getSheetByName(tabName);
    if (!targetSheet) {
      targetSheet = targetSSSecond.insertSheet(tabName);
    }

    // Ziel-Tab vor dem Einfügen leeren (damit keine alten Reste bleiben)
    targetSheet.clearContents();

    // Daten 1:1 in das Ziel-Tab schreiben
    targetSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  });

  Logger.log("Zweite Synchronisation erfolgreich abgeschlossen!");
}