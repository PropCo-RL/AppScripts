function logExecution(scriptName, status, details, errorMessage) {
  try {
    var sheetId = "1VB253Cq_hr4YfFrWDXNG-pgv8qOOKAzb7jN_KYciEHw"; 
    var ss = SpreadsheetApp.openById(sheetId);
    var logSheet = ss.getSheetByName("Logs");

    if (!logSheet) {
      logSheet = ss.insertSheet("Logs");
      logSheet.appendRow(["Zeitstempel", "Skript Name", "Status", "Details / Absender", "Detaillierte Fehlermeldung"]);
      logSheet.getRange(1, 1, 1, 5).setFontWeight("bold");
    }

    logSheet.appendRow([
      new Date(),
      scriptName,
      status,
      details || "",
      errorMessage || ""
    ]);
  } catch (e) {
    Logger.log("Fehler beim Schreiben des Logs im Sheet: " + e.toString());
  }
}