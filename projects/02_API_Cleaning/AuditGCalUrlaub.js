function applyVacationLogicToGCal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gCalSheet = ss.getSheetByName("gCal");
  const cleanerSheet = ss.getSheetByName("Cleaners");
  // DIESE ZEILE HAT NOCH GEFEHLT:
  const actionSheet = ss.getSheetByName("Schedule"); 
  
  const gCalData = gCalSheet.getDataRange().getValues().slice(1); 
  const cleanerData = cleanerSheet.getDataRange().getValues().slice(1);
  
  // Hilfsfunktion: Setzt ein Datum auf exakt Mitternacht für sauberen Vergleich
  const zeroTime = (d) => {
    if (!d || isNaN(new Date(d).getTime())) return null;
    const tempDate = new Date(d);
    tempDate.setHours(0, 0, 0, 0);
    return tempDate.getTime();
  };

  // 1. Cleaner-Urlaubsdaten laden (Mapping Apartment -> Info)
  const cleanerMap = {};
  cleanerData.forEach(row => {
    const aptName = String(row[1]).trim().toLowerCase();
    if (aptName) {
      cleanerMap[aptName] = {
        standardEmail: String(row[4] || "").trim().toLowerCase(),    // Spalte E
        replacementEmail: String(row[8] || "").trim().toLowerCase(), // Spalte I
        vStart: zeroTime(row[9]),  // Spalte J (Start Urlaub)
        vEnd: zeroTime(row[10])     // Spalte K (Ende Urlaub)
      };
    }
  });

  // 2. gCal Blatt Zeile für Zeile prüfen und korrigieren
  gCalData.forEach((row, index) => {
    const rowIndex = index + 2;
    const terminDatum = zeroTime(row[0]); // Datum Reinigung aus Spalte A
    const aptName = String(row[1]).trim().toLowerCase(); // Apartment Name aus Spalte B
    const gCalEmailIst = String(row[6] || "").trim().toLowerCase(); // Spalte G (Importierte Email)
    
    if (!terminDatum || !cleanerMap[aptName]) return;

    const info = cleanerMap[aptName];
    let correctEmail = info.standardEmail;

    // PRÜFUNG: Fällt das Reinigungsdatum in den Urlaubszeitraum (J bis K)?
    if (info.vStart !== null && info.vEnd !== null) {
      if (terminDatum >= info.vStart && terminDatum <= info.vEnd) {
        // Falls Urlaub: Nimm Email aus Spalte I (Replacement)
        correctEmail = info.replacementEmail;
      }
    }

    // 3. Werte in gCal schreiben
    // Spalte D (Soll-Email)
    gCalSheet.getRange(rowIndex, 4).setValue(correctEmail);

    // Spalte F (Match-Status gegen Spalte G prüfen)
    const matchStatus = (gCalEmailIst === correctEmail) ? "MATCH" : "Kein Match";
    gCalSheet.getRange(rowIndex, 6).setValue(matchStatus);
  });

  // 4. WICHTIG: Daten neu einlesen für den Export nach 'Schedule'
  SpreadsheetApp.flush();
  const finalGCalData = gCalSheet.getDataRange().getValues().slice(1);

  // 5. FILTER: Nur Zeilen, in denen Spalte F (Index 5) NICHT "MATCH" ist
  const exceptions = finalGCalData
    .filter(row => String(row[5]).trim().toUpperCase() !== "MATCH")
    .map(row => [row[0], row[1], row[2], row[3], row[4]]); // Datum, Apartment, Status, Soll-Email, ID

  // 6. Nur hinzufügen (APPEND), nichts löschen!
  if (exceptions.length > 0) {
    let lastRow = actionSheet.getLastRow();
    
    // Falls das Blatt fast leer ist, fangen wir mindestens bei Zeile 7 an
    let startRow = Math.max(lastRow + 1, 7); 
    
    actionSheet.getRange(startRow, 1, exceptions.length, 5).setValues(exceptions);
    actionSheet.getRange(startRow, 1, exceptions.length, 1).setNumberFormat("dd.MM.yyyy");
  }
}