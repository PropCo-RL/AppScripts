/**
 * SKRIPT 2: Liest "inRE" und ergänzt im bestehenden Blatt "cal" die Spalten I & J.
 */
function enrichCalWithInvoices() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetCal = ss.getSheetByName("cal");
  const sheetInRE = ss.getSheetByName("inRE");

  if (!sheetCal || sheetCal.getLastRow() < 2) return;

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // 1. TAB "inRE" EINLESEN
  const inReMap = {};
  if (sheetInRE && sheetInRE.getLastRow() >= 2) {
    const valuesInRE = sheetInRE.getRange(2, 1, sheetInRE.getLastRow() - 1, 5).getValues();

    valuesInRE.forEach(row => {
      const lodgifyId = String(row[0] || "").trim();
      const rechnungsId = String(row[1] || "").trim();
      const rawStart = row[2];
      const rawEnd = row[3];
      const status = String(row[4] || "").trim();

      if (!lodgifyId || lodgifyId.startsWith("ERROR") || !rawStart || !rawEnd) return;

      const startDate = rawStart instanceof Date ? new Date(rawStart) : parseDateUtil(rawStart);
      const endDate = rawEnd instanceof Date ? new Date(rawEnd) : parseDateUtil(rawEnd);

      if (startDate && endDate) {
        startDate.setHours(0, 0, 0, 0);
        endDate.setHours(0, 0, 0, 0);

        if (!inReMap[lodgifyId]) inReMap[lodgifyId] = [];
        inReMap[lodgifyId].push({ rechnungsId: rechnungsId, start: startDate, end: endDate, status: status });
      }
    });
  }

  // 2. TAB "cal" ANREICHERN
  const numRows = sheetCal.getLastRow() - 1;
  const calBookingIds = sheetCal.getRange(2, 4, numRows, 1).getValues(); // Spalte D (ID)

  let colI_values = [];
  let colJ_values = [];

  for (let i = 0; i < numRows; i++) {
    const bookingId = String(calBookingIds[i][0] || "").trim();

    if (!bookingId) {
      colI_values.push(["-"]);
      colJ_values.push(["-"]);
      continue;
    }

    const invoices = inReMap[bookingId] || [];

    // Spalte I (RE bis heute?)
    const hasInvoiceForToday = invoices.some(inv => inv.status.toLowerCase() !== "storniert" && inv.start <= today && inv.end >= today);
    const reBisHeuteText = hasInvoiceForToday ? "JA" : "NEIN";

    // Spalte J (Offene RE)
    let overDueItems = [];
    invoices.forEach(inv => {
      if (inv.status.toLowerCase() === "offen" && inv.start <= today) {
        const diffDays = Math.floor((today.getTime() - inv.start.getTime()) / (1000 * 60 * 60 * 24));
        overDueItems.push(inv.rechnungsId + " (" + diffDays + ")");
      }
    });
    const offeneReText = overDueItems.length > 0 ? overDueItems.join(", ") : "-";

    colI_values.push([reBisHeuteText]);
    colJ_values.push([offeneReText]);
  }

  // 3. SCHREIBEN IN "cal" (Spalte I & J)
  sheetCal.getRange(1, 9).setValue("RE bis heute?").setFontWeight("bold");
  sheetCal.getRange(1, 10).setValue("Offene RE (Tage überfällig)").setFontWeight("bold");

  sheetCal.getRange(2, 9, numRows, 1).setValues(colI_values);
  sheetCal.getRange(2, 10, numRows, 1).setValues(colJ_values);
}

function parseDateUtil(str) {
  if (!str) return null;
  const s = String(str).trim();
  if (s.includes(".")) {
    const parts = s.split(".");
    if (parts.length === 3) return new Date(parts[2], parts[1] - 1, parts[0]);
  }
  if (s.includes("/")) {
    const parts = s.split("/");
    if (parts.length === 3) return new Date(parts[2], parts[0] - 1, parts[1]);
  }
  const parsed = new Date(s);
  return isNaN(parsed.getTime()) ? null : parsed;
}