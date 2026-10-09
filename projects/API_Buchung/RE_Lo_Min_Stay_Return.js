// ============================================================================
// STEP 2: MIN-STAY NACH BUCHUNG ZURÜCK AUF 14 SETZEN (TAB "RE")
// ============================================================================

/**
 * HAUPTFUNKTION SCHRITT 2
 */
function ZZZ_resetMinStayToFourteen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetRE = ss.getSheetByName("RE");

  if (!sheetRE) throw new Error("Das Tabellenblatt 'RE' wurde nicht gefunden.");

  var apartmentMap = getApartmentMapping();
  var lastRow = sheetRE.getLastRow();
  if (lastRow < 2) {
    Logger.log("Keine Daten im Sheet 'RE' vorhanden.");
    return;
  }

  // Liest Daten ab Zeile 2 von Spalte A (1) bis Spalte AG (33)
  var values = sheetRE.getRange(2, 1, lastRow - 1, 33).getValues();

  values.forEach(function(row, index) {
    var rowIndex = index + 2;
    var lodgifyStatus = row[18];   // Spalte S (Index 18)
    var minStayStatus = row[32];   // Spalte AG (Index 32)
    var apartmentCodeRaw = row[13];// Spalte N (Index 13)

    // STRIKTE BEDINGUNGEN:
    // 1. Spalte S MUSS befüllt sein (vom Hauptskript)
    if (!lodgifyStatus || lodgifyStatus.toString().trim() === "") return;

    // 2. In Spalte AG MUSS exakt "auf 1 geändert" stehen
    // (Bricht ab wenn AG leer ist, wenn schon "auf 14 geändert" steht, oder sonst was drinsteht)
    if (!minStayStatus || minStayStatus.toString().trim() !== "auf 1 geändert") return;

    var apartmentCode = apartmentCodeRaw ? apartmentCodeRaw.toString().toLowerCase().trim() : "";
    if (!apartmentCode) return;

    var apartment = apartmentMap[apartmentCode];
    if (!apartment) return;

    // Raten-Änderung zurück auf 14 Nächte an Lodgify senden
    var success = executeMinStayChange(apartment.property_id, apartment.room_type_id, 14);

    if (success) {
      sheetRE.getRange(rowIndex, 33).setValue("auf 14 geändert"); // AG überschreiben
      Logger.log("ERFOLG: Zeile " + rowIndex + " | Property " + apartment.property_id + " min_stay zurück auf 14 gesetzt.");
    } else {
      Logger.log("FEHLER: Zeile " + rowIndex + " | Raten-Update auf 14 fehlgeschlagen.");
    }
  });
}