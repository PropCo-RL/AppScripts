function fillMissingEmails() {
  const currentSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // ID des externen Sheets aus der URL
  const externalSpreadsheetId = "1mMEVsVFIx5VxIV3EyJYlba_G0fgEX9t0tR2KbFkm-nU";
  const externalSpreadsheet = SpreadsheetApp.openById(externalSpreadsheetId);
  
  // Map zum Speichern aller ID -> E-Mail Paare aus dem externen Sheet
  const emailMap = new Map();
  const sheetsToSearch = ["RE", "VL"];
  
  // 1. Daten aus den Tabs "RE" und "VL" im externen Sheet auslesen
  sheetsToSearch.forEach(sheetName => {
    const sheet = externalSpreadsheet.getSheetByName(sheetName);
    if (!sheet) return;
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return; // Wenn nur die Kopfzeile da ist
    
    // Spalte I (E-Mails) ist Spalte 9, Spalte T (Rechnungs-ID) ist Spalte 20
    const emails = sheet.getRange(2, 9, lastRow - 1, 1).getValues();
    const ids = sheet.getRange(2, 20, lastRow - 1, 1).getValues();
    
    for (let i = 0; i < ids.length; i++) {
      const id = String(ids[i][0]).trim();
      const email = String(emails[i][0]).trim();
      
      // Nur eintragen, wenn eine gültige ID und eine echte E-Mail existiert (kein "-")
      if (id && email && email !== "-") {
        emailMap.set(id, email);
      }
    }
  });
  
  // 2. Aktuelles Sheet verarbeiten
  const currentLastRow = currentSheet.getLastRow();
  if (currentLastRow < 2) return;
  
  // Spalte F (E-Mails) ist Spalte 6, Spalte G (Lexoffice ID) ist Spalte 7
  const currentEmailsRange = currentSheet.getRange(2, 6, currentLastRow - 1, 1);
  const currentEmails = currentEmailsRange.getValues();
  const currentIds = currentSheet.getRange(2, 7, currentLastRow - 1, 1).getValues();
  
  let updatedCount = 0;
  
  // 3. Fehlende / mit "-" markierte E-Mails abgleichen
  for (let i = 0; i < currentEmails.length; i++) {
    const currentEmail = String(currentEmails[i][0]).trim();
    const currentId = String(currentIds[i][0]).trim();
    
    // Prüfe, ob Spalte F leer ist ODER nur den Bindestrich "-" enthält
    const isMissing = !currentEmail || currentEmail === "-";
    
    if (isMissing && currentId) {
      if (emailMap.has(currentId)) {
        currentEmails[i][0] = emailMap.get(currentId);
        updatedCount++;
      }
    }
  }
  
  // 4. Zurück in das Sheet schreiben, falls Änderungen vorhanden sind
  if (updatedCount > 0) {
    currentEmailsRange.setValues(currentEmails);
    SpreadsheetApp.getActiveSpreadsheet().toast(`${updatedCount} E-Mail-Adressen wurden erfolgreich ergänzt!`, "Erfolg");
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast("Keine passenden E-Mail-Adressen zum Ergänzen gefunden.", "Info");
  }
}