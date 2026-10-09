function generateCourtPDFFromRow() {
  // Welche Zeile willst du umwandeln?
  var ROW_INDEX = 2; // <--- HIER DIE ZEILE EINTRAGEN
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("sites"); // Prüfe ob dein Sheet so heißt!
  
  // Hole den CSV-String aus Spalte 14 (bzw. der Spalte, in die du ihn oben schreibst)
  var csvString = sheet.getRange(ROW_INDEX, 14).getValue();
  if (!csvString) {
    Logger.log("Kein Audit-Log gefunden.");
    return;
  }
  
  var data = csvString.split(";");
  
  var docName = "BEWEISPROTOKOLL_" + data[1] + "_" + Utilities.formatDate(new Date(), "GMT+1", "yyyyMMdd");
  var doc = DocumentApp.create(docName);
  var body = doc.getBody();
  
  body.appendParagraph("BEWEISPROTOKOLL / AUDIT-LOG").setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph("Nachweis der Online-Buchung und AGB-Zustimmung gemäß § 286 ZPO & § 312j BGB\n").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  
  body.appendParagraph("Hiermit wird dokumentiert, dass am " + data[0] + " eine rechtlich bindende Willenserklärung über das L8 Street System getätigt wurde.");
  
  var tableData = [
    ["Parameter", "Erfasster Systemwert"],
    ["Zeitstempel (UTC)", data[0]],
    ["Eindeutige Buchungs-ID", data[1]],
    ["Eingangskanal", data[2]],
    ["Firma / Name", data[3]],
    ["E-Mail Adresse", data[4]],
    ["Telefonnummer", data[5]],
    ["IP-Adresse des Nutzers", data[6]],
    ["Technischer Fingerabdruck", data[7]],
    ["Gebuchtes Objekt", data[8]],
    ["Reisezeitraum", data[9]],
    ["Willenserklärung", data[12]],
    ["AGB zugestimmt", data[10]],
    ["Akzeptierte AGB-Version", data[11]]
  ];
  
  body.appendTable(tableData);
  body.appendParagraph("\nDieser Bericht wurde maschinell und manipulationssicher extrahiert.");
  
  doc.saveAndClose();
  
  var pdfBlob = doc.getAs('application/pdf');
  var pdfFile = DriveApp.createFile(pdfBlob);
  pdfFile.setName(docName + ".pdf");
  
  DriveApp.getFileById(doc.getId()).setTrashed(true);
  Logger.log("Gerichtsprotokoll PDF erfolgreich in Drive generiert!");
}