// ============================================================================
// DATEN-TRANSFER: 'verKeinAuto' AUS TAB 'VL' IN EXTERNES SHEET
// ============================================================================

/**
 * HAUPTFUNKTION: Überträgt spezifische Daten aus Sheet 'VL' in ein externes Sheet
 * Fügt in Spalte A die vollständige Apartment-URL ein (L8Street.com/a/ + Kürzel)
 * Marker wird in Spalte AE geschrieben!
 */
function transferVerKeinAuto() {
  const ssSource = SpreadsheetApp.getActiveSpreadsheet();
  const sheetVL = ssSource.getSheetByName("VL");

  if (!sheetVL) throw new Error("Das Tabellenblatt 'VL' wurde im Quell-Sheet nicht gefunden.");

  // Ziel-Spreadsheet über ID öffnen
  const targetSpreadsheetId = "1bMPVHjGNnwWDH0KXUyB47_VbLcq_lsFK1jSbdj_GmH0";
  let ssTarget;
  
  try {
    ssTarget = SpreadsheetApp.openById(targetSpreadsheetId);
  } catch (e) {
    throw new Error("Konnte das Ziel-Sheet nicht öffnen. Bitte Zugriffsrechte prüfen! Fehler: " + e.message);
  }

  const sheetTarget = ssTarget.getSheetByName("verKeinAuto");
  if (!sheetTarget) throw new Error("Das Tabellenblatt 'verKeinAuto' wurde im Ziel-Sheet nicht gefunden.");

  const lastRowSource = sheetVL.getLastRow();
  if (lastRowSource < 2) {
    Logger.log("Keine Datenzeilen im Sheet 'VL' gefunden.");
    return;
  }

  // Liest bis Spalte AE (31 Spalten)
  const rangeSource = sheetVL.getRange(2, 1, lastRowSource - 1, 31);
  const valuesVL = rangeSource.getValues();

  let rowsToExport = [];
  let rowIndicesToMark = [];

  valuesVL.forEach((row, index) => {
    const rowIndex = index + 2;
    
    const phone = row[0];           // Spalte A: Telefonnummer
    const firma = row[1];           // Spalte B: Firma
    const departure = row[11];      // Spalte L: Abreise Datum
    const apartmentCode = row[13];  // Spalte N: Apartment Kürzel
    const lodgifyId = row[18];      // Spalte S: Lodgify Buchungsnummer
    const filterCondition = row[26] ? row[26].toString().trim() : ""; // Spalte AA: 'verKeinAuto'
    const transferStatus = row[30] ? row[30].toString().trim() : "";  // Spalte AE (Index 30): Marker

    // Nur verarbeiten, wenn in Spalte AA 'verKeinAuto' steht AND in Spalte AE noch kein Marker ist
    if (filterCondition.toLowerCase() === "verkeinauto" && transferStatus === "") {
      
      // Datum sauber als DD.MM.YYYY formatieren
      let formattedDate = departure;
      if (departure instanceof Date) {
        formattedDate = Utilities.formatDate(departure, "Europe/Berlin", "dd.MM.yyyy");
      }

      // Vollständige Apartment-URL zusammenbauen
      const cleanCode = apartmentCode ? apartmentCode.toString().trim() : "";
      const fullApartmentUrl = cleanCode ? "L8Street.com/a/" + cleanCode : "";

      // Ziel-Zeile aufbauen:
      // Spalte A: Apartment-URL, B: Phone, C: Firma, D: Abreise, E: Lodgify ID
      rowsToExport.push([
        fullApartmentUrl,
        phone ? phone.toString().trim() : "",
        firma ? firma.toString().trim() : "",
        formattedDate ? formattedDate.toString().trim() : "",
        lodgifyId ? lodgifyId.toString().trim() : ""
      ]);

      rowIndicesToMark.push(rowIndex);
    }
  });

  // Falls relevante Datensätze gefunden wurden -> ins Ziel-Sheet schreiben (5 Spalten: A bis E)
  if (rowsToExport.length > 0) {
    // Nächste freie Zeile im Ziel-Sheet ermitteln
    const lastRowTarget = sheetTarget.getLastRow();
    const startRowTarget = lastRowTarget < 1 ? 2 : lastRowTarget + 1; // Falls leer, ab Zeile 2 starten

    // Schreiben der gefilterten Daten in das Ziel-Sheet (5 Spalten A bis E)
    sheetTarget.getRange(startRowTarget, 1, rowsToExport.length, 5).setValues(rowsToExport);
    Logger.log(`${rowsToExport.length} Zeile(n) erfolgreich ins Ziel-Sheet 'verKeinAuto' übertragen.`);

    // Quell-Zeilen in Spalte AE als 'Übertragen' markieren (Spalte 31)
    rowIndicesToMark.forEach(rIndex => {
      const timeStamp = Utilities.formatDate(new Date(), "Europe/Berlin", "dd.MM.yyyy HH:mm");
      sheetVL.getRange(rIndex, 31).setValue("Übertragen (" + timeStamp + ")");
    });

  } else {
    Logger.log("Keine neuen Zeilen mit 'verKeinAuto' in Spalte AA gefunden.");
  }
}