/**
 * SKRIPT 1: Generiert ausschließlich das Hauptblatt "cal" UND synchronisiert
 * die benötigten Verfügbarkeiten (Spalten 1-3) ins Marketing-Sheet.
 */
function generateBaseSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. APARTMENTS LADEN
  const sheetD = ss.getSheetByName("d");
  const propertyMap = {};

  if (sheetD && sheetD.getLastRow() >= 1) {
    const valuesD = sheetD.getRange(1, 1, sheetD.getLastRow(), 4).getValues();
    valuesD.forEach(row => {
      const id = String(row[0]).trim();
      const name = String(row[2]).trim();
      const link = String(row[3]).trim();

      if (id !== "" && id.toLowerCase() !== "property_id") {
        let code = "";
        const match = link.match(/\/a\/([^\/\?]+)/i);
        if (match && match[1]) code = match[1].toLowerCase().trim();

        propertyMap[id] = { name: name || id, code: code };
      }
    });
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const pastLimit = new Date(today); pastLimit.setDate(today.getDate() - 7);
  const futureLimit = new Date(today); futureLimit.setDate(today.getDate() + 365);
  const autoLimit = new Date(today); autoLimit.setDate(today.getDate() + 30);

  let calEntries = [];

  // 2. BUCHUNGEN (db) EINLESEN
  const sheetDB = ss.getSheetByName("db");
  if (sheetDB && sheetDB.getLastRow() >= 4) {
    const valuesDB = sheetDB.getRange(4, 1, sheetDB.getLastRow() - 3, 23).getValues();

    valuesDB.forEach(row => {
      const bookingId = String(row[1] || "").trim();
      const arrival = row[3] instanceof Date ? row[3] : (row[3] ? new Date(row[3]) : null);
      const departure = row[4] instanceof Date ? row[4] : (row[4] ? new Date(row[4]) : null);
      const propId = String(row[5] || "").trim();
      const guestName = String(row[13] || "").trim();
      const phone = String(row[15] || "").trim();
      const origin = String(row[22] || "").trim();

      if (propId && arrival && departure) {
        arrival.setHours(0, 0, 0, 0);
        departure.setHours(0, 0, 0, 0);

        const apartmentInfo = propertyMap[propId] || { name: propId, code: "" };
        const isBookingCom = origin.toLowerCase().includes("bookingcom");
        const isCurrentlyActive = arrival <= today && departure > today;

        if (departure >= pastLimit && arrival <= futureLimit) {
          let vAutoStatus = (!isBookingCom && isCurrentlyActive && departure > autoLimit) ? "verAuto" : "-";
          calEntries.push({
            aptName: apartmentInfo.name,
            arrival: arrival,
            departure: departure,
            cpId: bookingId,
            guestName: guestName,
            phone: phone,
            origin: origin,
            vAuto: vAutoStatus
          });
        }
      }
    });
  }

  // 3. BLOCKIERUNGEN (dbX) EINLESEN
  const sheetDBX = ss.getSheetByName("dbX");
  if (sheetDBX && sheetDBX.getLastRow() >= 2) {
    const valuesDBX = sheetDBX.getRange(2, 1, sheetDBX.getLastRow() - 1, 4).getValues();

    valuesDBX.forEach(row => {
      const propId = String(row[0] || "").trim();
      const arrival = row[1] ? new Date(row[1]) : null;
      const departure = row[2] ? new Date(row[2]) : null;
      const cpId = String(row[3] || "").trim();

      if (propId && arrival && departure) {
        arrival.setHours(0, 0, 0, 0);
        departure.setHours(0, 0, 0, 0);

        if (departure >= pastLimit && arrival <= futureLimit) {
          const apartmentInfo = propertyMap[propId] || { name: propId, code: "" };
          calEntries.push({
            aptName: apartmentInfo.name,
            arrival: arrival,
            departure: departure,
            cpId: cpId,
            guestName: "BLOCKIERUNG",
            phone: "",
            origin: "",
            vAuto: "-"
          });
        }
      }
    });
  }

  // 4. SORTIEREN & SCHREIBEN VON "cal" (100% UNVERÄNDERT FÜR DEIN ALTES SHEET)
  calEntries.sort((a, b) => {
    const nameComp = a.aptName.localeCompare(b.aptName, 'de', { numeric: true, sensitivity: 'base' });
    return nameComp !== 0 ? nameComp : a.departure - b.departure;
  });

  let calRows = [];
  let lastAptCal = null;
  calEntries.forEach(entry => {
    if (lastAptCal !== null && entry.aptName !== lastAptCal) calRows.push(["", "", "", "", "", "", "", ""]);
    calRows.push([
      entry.aptName,
      Utilities.formatDate(entry.arrival, "GMT+2", "yyyy-MM-dd"),
      Utilities.formatDate(entry.departure, "GMT+2", "yyyy-MM-dd"),
      entry.cpId,
      entry.guestName,
      entry.phone,
      entry.origin,
      entry.vAuto
    ]);
    lastAptCal = entry.aptName;
  });

  let sheetCal = ss.getSheetByName("cal") || ss.insertSheet("cal");
  sheetCal.clear();
  const headersCal = ["property_name", "periods.start", "periods.end", "periods.closed_period.id", "guest_name", "phone", "origin", "vAUTO"];
  sheetCal.getRange(1, 1, 1, headersCal.length).setValues([headersCal]).setFontWeight("bold");
  if (calRows.length > 0) sheetCal.getRange(2, 1, calRows.length, headersCal.length).setValues(calRows);

  // ============================================================================
  // EXPORT-SCHRITT: FÜLLT DAS NEUE MARKETING-SHEET (NEUER TAB 'cal')
  // ============================================================================
  try {
    const MARKETING_SPREADSHEET_ID = "1Aq9doxAkEw435bRCpXafayRwdNeaP9LVMU1YwWG7z70";
    const targetSs = SpreadsheetApp.openById(MARKETING_SPREADSHEET_ID);
    let targetSheetCal = targetSs.getSheetByName("cal") || targetSs.insertSheet("cal");
    
    targetSheetCal.clear();
    targetSheetCal.getRange(1, 1, 1, 3).setValues([["property_name", "periods.start", "periods.end"]]).setFontWeight("bold");

    let exportRows = [];
    calEntries.forEach(entry => {
      if (entry.departure >= today) {
        exportRows.push([
          entry.aptName,
          Utilities.formatDate(entry.arrival, "GMT+2", "yyyy-MM-dd"),
          Utilities.formatDate(entry.departure, "GMT+2", "yyyy-MM-dd")
        ]);
      }
    });

    if (exportRows.length > 0) {
      targetSheetCal.getRange(2, 1, exportRows.length, 3).setValues(exportRows);
    }
  } catch (e) {
    Logger.log("Export-Info: " + e.toString());
  }
}