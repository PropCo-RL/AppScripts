// Diese Funktion fängt den POST-Request von Wix ab
function doPost(e) {
  try {
    // 1. Die eingehenden JSON-Daten von Wix parsen
    var payload = JSON.parse(e.postData.contents);
    
    // 2. Das aktuelle Spreadsheet und das erste Tabellenblatt aufrufen
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0]; 
    
    // 3. Vorbereitung der Datenzeile
    // Da Wix CMS Daten je nach Webhook-Typ variieren, loggen wir das Datum 
    // und hängen das gesamte JSON-Objekt oder bestimmte Felder an.
    var timestamp = new Date();
    
    // Beispiel: Wir wandeln das JSON in Text um, um es komplett in einer Zelle zu sehen,
    // oder ziehen uns gezielt Daten heraus (z.B. payload.data oder payload.slug)
    var stringifiedPayload = JSON.stringify(payload);
    
    // 4. Daten als neue Zeile am Ende der Tabelle einfügen
    // Spalte A: Zeitstempel, Spalte B: Die Rohdaten von Wix
    sheet.appendRow([timestamp, stringifiedPayload]);
    
    // 5. Wix eine erfolgreiche Rückmeldung (Status 200) senden
    return ContentService.createTextOutput(JSON.stringify({"status": "success"}))
                         .setMimeType(ContentService.MimeType.JSON);
                         
  } catch(error) {
    // Falls etwas schiefgeht, Fehler zurückgeben
    return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": error.toString()}))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}