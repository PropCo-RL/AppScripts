// ============================================================================
// AUTOMATISCHE INTERNE VERLÄNGERUNG (TABS "RE" & "VL")
// ============================================================================

function AAA_autoExtendBookings() {

  // ==========================================================================
  // EINGEBETTETE HILFSFUNKTIONEN
  // ==========================================================================

  function parseDateAuto(dateVal) {
    if (!dateVal) return null;
    if (dateVal instanceof Date) {
      const d = new Date(dateVal.getTime());
      d.setHours(0, 0, 0, 0);
      return d;
    }

    const str = dateVal.toString().trim();
    const parts = str.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})/);
    if (parts) {
      return new Date(parseInt(parts[3], 10), parseInt(parts[2], 10) - 1, parseInt(parts[1], 10), 0, 0, 0);
    }

    const d = new Date(str);
    if (!isNaN(d.getTime())) {
      d.setHours(0, 0, 0, 0);
      return d;
    }
    return null;
  }

  function parseCycleDaysAuto(cycleRaw) {
    if (!cycleRaw) return 7;
    const str = cycleRaw.toString().toLowerCase().trim();
    
    const num = parseInt(str, 10);
    if (!isNaN(num) && num > 0) return num;

    if (str.includes("wöchentlich") || str.includes("woechentlich")) return 7;
    if (str.includes("monatlich") || str.includes("monat")) return 30;

    return 7;
  }

  function formatDisplayDateAuto(d) {
    if (!d) return "";
    const day = ("0" + d.getDate()).slice(-2);
    const month = ("0" + (d.getMonth() + 1)).slice(-2);
    return `${day}.${month}.${d.getFullYear()}`;
  }

  // ==========================================================================
  // HAUPTLOGIK & ABLAUF
  // ==========================================================================

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheets = ["RE", "VL"];

  targetSheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`Sheet '${sheetName}' wurde nicht gefunden.`);
      return;
    }

    const initialLastRow = sheet.getLastRow();
    if (initialLastRow < 2) return;

    // Lese Daten bis Spalte AH (Spalte 34)
    const maxCols = Math.max(sheet.getLastColumn(), 34);
    const range = sheet.getRange(2, 1, initialLastRow - 1, maxCols);
    const values = range.getValues();

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    // Speicher für gesammelte Aktionen (verhindert Zeilenversatz während der Schleife)
    const rowsToAppend = [];
    const rowsToUpdateAlt = [];

    for (let i = 0; i < values.length; i++) {
      const rowIndex = i + 2;
      const rowData = values[i];

      const valAH = rowData[33] ? rowData[33].toString().trim() : ""; // Spalte AH
      const valL  = rowData[11]; // Spalte L (Abreise Neu / Finales Ende)
      const valAB = rowData[27]; // Spalte AB (Lx RE Start)
      const valAC = rowData[28]; // Spalte AC (Lx RE Ende)
      const valM  = rowData[12]; // Spalte M (Zyklus/Tage)

      // ----------------------------------------------------------------------
      // BEDINGUNG 1: Spalte AH prüfen
      // ----------------------------------------------------------------------
      if (!valAH || valAH.toLowerCase() === "auto erledigt") continue;
      if (!valAH.toLowerCase().startsWith("auto")) continue;

      const dateL  = parseDateAuto(valL);
      const dateAB = parseDateAuto(valAB);
      const dateAC = parseDateAuto(valAC);

      if (!dateL || !dateAB || !dateAC) {
        Logger.log(`[${sheetName}] Zeile ${rowIndex}: Übersprungen, ein Datum (L, AB oder AC) ist ungültig.`);
        continue;
      }

      // ----------------------------------------------------------------------
      // BEDINGUNG 2: Spalte L ist weiter in der Zukunft als Spalte AC (L > AC)
      // ----------------------------------------------------------------------
      if (dateL.getTime() <= dateAC.getTime()) continue;

      // ----------------------------------------------------------------------
      // BEDINGUNG 3: Spalte AC ist ab heute in 14 Tagen oder weniger (AC <= heute + 14)
      // ----------------------------------------------------------------------
      const fourteenDaysFromNow = new Date(today.getTime());
      fourteenDaysFromNow.setDate(fourteenDaysFromNow.getDate() + 14);

      if (dateAC.getTime() > fourteenDaysFromNow.getTime()) continue;

      // ======================================================================
      // ALLE 3 BEDINGUNGEN ERFÜLLT - AKTION VORBEREITEN
      // ======================================================================
      Logger.log(`[${sheetName}] Bedingung erfüllt für Zeile ${rowIndex} (${valAH}). Merke zur Verarbeitung...`);

      // 1. Ursprungszeile für "auto erledigt" vormerken
      rowsToUpdateAlt.push(rowIndex);

      // 2. Exakte Zeilen-Kopie im Speicher anlegen
      const newRowValues = [...rowData];

      // 3. Berechnungen für die neue Zeile:
      const newStartDate = new Date(dateAC.getTime());
      const cycleDays = parseCycleDaysAuto(valM);
      let newEndDate = new Date(newStartDate.getTime());
      newEndDate.setDate(newEndDate.getDate() + cycleDays);

      // Falls neues Enddatum > Spalte L -> Deckeln auf Spalte L
      if (newEndDate.getTime() > dateL.getTime()) {
        newEndDate = new Date(dateL.getTime());
      }

      // 4. In der gekonten Array-Kopie die Felder anpassen:
      newRowValues[27] = formatDisplayDateAuto(newStartDate); // Spalte AB (28)
      newRowValues[28] = formatDisplayDateAuto(newEndDate);   // Spalte AC (29)

      // Felder leeren in der Kopie: T (19), U (20), V (21), X (23), AF (31)
      newRowValues[19] = ""; // Spalte T (Rechnungs ID)
      newRowValues[20] = ""; // Spalte U (Status Lexoffice)
      newRowValues[21] = ""; // Spalte V (E-Mail Status)
      newRowValues[23] = ""; // Spalte X (Lodgify Status)
      newRowValues[31] = ""; // Spalte AF (Status RE)

      rowsToAppend.push(newRowValues);
    }

    // ======================================================================
    // SCHREIBVORGANG (AUSSERHALB DER SCHLEIFE - ZEILENVERSATZ UNMÖGLICH)
    // ======================================================================
    if (rowsToAppend.length > 0) {
      // 1. Altzeilen auf "auto erledigt" setzen
      rowsToUpdateAlt.forEach(rIdx => {
        sheet.getRange(rIdx, 34).setValue("auto erledigt");
      });

      // 2. Alle neuen Zeilen geordnet unten anfügen
      rowsToAppend.forEach(rowArr => {
        sheet.appendRow(rowArr);
      });

      Logger.log(`[${sheetName}] Erfolgreich ${rowsToAppend.length} neue Zeile(n) angelegt.`);
    }
  });
}