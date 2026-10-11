function sendDriveFilesAsEmail() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var startRow = 2; // Startet bei Zeile 2 (unter dem Header)
  var lastRow = sheet.getLastRow();
  
  // Ziel-E-Mail-Adresse
  var recipient = "l8@inbox.lexware.email";
  
  if (lastRow < startRow) {
    Logger.log("Keine Daten vorhanden.");
    return;
  }
  
  // Daten aus den Spalten D bis F abrufen (Spalte D = Index 4, Spalte F = Index 6)
  var range = sheet.getRange(startRow, 4, lastRow - startRow + 1, 3);
  var values = range.getValues();
  
  for (var i = 0; i < values.length; i++) {
    var url = values[i][0];       // Spalte D: Google Drive Link
    var status = values[i][2];    // Spalte F: Status/Erfolgsmeldung
    var currentRow = startRow + i;
    
    // Nur verarbeiten, wenn ein Link vorhanden ist UND die Zeile noch nicht verarbeitet wurde
    if (url && url.toString().trim() !== "" && (!status || status.toString().trim() === "")) {
      try {
        // ID aus dem Google Drive Link extrahieren
        var fileId = extractFileIdFromUrl(url.toString());
        
        if (fileId) {
          // Datei aus Google Drive abrufen
          var file = DriveApp.getFileById(fileId);
          var fileName = file.getName();
          
          // E-Mail mit Anhang senden
          GmailApp.sendEmail(recipient, "Rechnung: " + fileName, "Anbei befindet sich die angehängte Datei.", {
            attachments: [file.getAs(file.getMimeType())],
            name: "Automatischer Belegversand"
          });
          
          // Zeitstempel und Erfolgsmeldung in Spalte F eintragen
          var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm:ss");
          sheet.getRange(currentRow, 6).setValue("Anhang versendet am " + timestamp);
          
          Logger.log("Erfolgreich gesendet für Zeile " + currentRow);
        } else {
          sheet.getRange(currentRow, 6).setValue("Fehler: Ungültige Google Drive URL");
        }
      } catch (e) {
        // Fehler in Spalte F eintragen, falls etwas schiefgeht
        sheet.getRange(currentRow, 6).setValue("Fehler: " + e.message);
        Logger.log("Fehler bei Zeile " + currentRow + ": " + e.message);
      }
    }
  }
}

/**
 * Hilfsfunktion zum Extrahieren der Datei-ID aus verschiedenen Google Drive Link-Formaten
 */
function extractFileIdFromUrl(url) {
  var match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}
