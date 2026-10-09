/**
 * SKRIPT 3: Verteilt die Daten aus "cal" NUR für heute aktuell aktive Mieter auf:
 * - cal_re_erstellen (Spalte I == "NEIN", nur aktuell aktive Mieter + Ursprung "Manual" oder "PublicApi")
 * - cal_kickout (Spalte J hat überfällige Tage in Klammern -> Gruppiert nach Firma ab der am meisten überfälligen)
 * - cal_advertise (Frei heute oder in den kommenden 4 Tagen UND danach min. 5 Tage am Stück frei)
 */
function generateTargetSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetCal = ss.getSheetByName("cal");
  const sheetD = ss.getSheetByName("d");

  if (!sheetCal || sheetCal.getLastRow() < 2) return;

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const maxLookaheadDays = 2; // Max. 4 Tage in der Zukunft frei werden
  const requiredFreeDays = 7;  // Mindestens 5 Tage am Stück frei

  // 1. ALL APARTMENTS AUS "d" SAMMELN (ohne Header "Name" / "property_name")
  const allApartments = new Set();
  if (sheetD && sheetD.getLastRow() >= 1) {
    const valuesD = sheetD.getRange(1, 1, sheetD.getLastRow(), 3).getValues();
    valuesD.forEach((row, index) => {
      if (index === 0) return; // Header überspringen

      const id = String(row[0]).trim();
      const name = String(row[2]).trim();

      if (name && name.toLowerCase() !== "name" && name.toLowerCase() !== "property_name") {
        allApartments.add(name);
      } else if (id && id.toLowerCase() !== "property_id" && id.toLowerCase() !== "id") {
        allApartments.add(id);
      }
    });
  }

  // 2. CAL EINLESEN (Spalten A bis J)
  const lastRow = sheetCal.getLastRow();
  const calValues = sheetCal.getRange(2, 1, lastRow - 1, 10).getValues();

  let rowsReErstellen = [];
  let rowsKickoutRaw = [];
  let bookingsPerApt = {};

  calValues.forEach(row => {
    const aptName = String(row[0] || "").trim();
    const start = row[1] ? new Date(row[1]) : null;
    const end = row[2] ? new Date(row[2]) : null;
    const origin = String(row[6] || "").trim();              // Spalte G: Ursprung
    const colI = String(row[8] || "").trim().toUpperCase(); // Spalte I: RE bis heute?
    const colJ = String(row[9] || "").trim();               // Spalte J: Offene RE

    if (!aptName || !start || !end) return;

    start.setHours(0, 0, 0, 0);
    end.setHours(0, 0, 0, 0);

    // Buchungen pro Apartment für cal_advertise erfassen
    if (!bookingsPerApt[aptName]) {
      bookingsPerApt[aptName] = [];
    }
    bookingsPerApt[aptName].push({ start: start, end: end });

    // Kriterium: Ist der Mieter HEUTE aktuell in der Wohnung? (Anreise <= heute < Abreise)
    const isCurrentlyActive = (start <= today && end > today);

    // Nur verarbeiten, wenn der Mieter AKTUELL im Apartment wohnt
    if (isCurrentlyActive) {
      const baseData = row.slice(0, 8); // Spalten A-H

      // Kriterium 1: cal_re_erstellen 
      const isValidOrigin = (origin === "Manual" || origin === "PublicApi");
      if (colI === "NEIN" && isValidOrigin) {
        rowsReErstellen.push([...baseData, colI]);
      }

      // Kriterium 2: cal_kickout (Aktueller Mieter + Spalte J hat Tage in Klammern)
      const hasOverdueBracket = /\(\d+\)/.test(colJ);
      if (colJ !== "-" && colJ !== "" && hasOverdueBracket) {
        const match = colJ.match(/\((\d+)\)/);
        const overdueDays = match ? parseInt(match[1], 10) : 0;
        
        rowsKickoutRaw.push({
          data: [...baseData, colJ],
          overdueDays: overdueDays,
          guestName: String(row[4] || "").trim() || "Unbekannter Gast",
          aptName: aptName
        });
      }
    }
  });

  // 3. NEUE LOGIK FÜR CAL_ADVERTISE
  let rowsAdvertiseNames = [];

  Array.from(allApartments).forEach(aptName => {
    const aptBookings = bookingsPerApt[aptName] || [];

    // Freigabedatum ermitteln (Kettung von bestehenden/nahtlosen Buchungen verschiebt das Datum)
    let availableFrom = new Date(today.getTime());
    let changed = true;

    while (changed) {
      changed = false;
      for (let b of aptBookings) {
        if (b.start.getTime() <= availableFrom.getTime() && b.end.getTime() > availableFrom.getTime()) {
          availableFrom = new Date(b.end.getTime());
          changed = true;
        }
      }
    }

    // Prüfen, ob Freigabedatum in den nächsten 0 bis 4 Tagen liegt
    const diffDays = Math.round((availableFrom.getTime() - today.getTime()) / (1000 * 3600 * 24));

    if (diffDays >= 0 && diffDays <= maxLookaheadDays) {
      // Prüfen, ob danach 5 Tage am Stück frei sind
      const minRequiredEnd = new Date(availableFrom.getTime());
      minRequiredEnd.setDate(minRequiredEnd.getDate() + requiredFreeDays);

      let hasConflict = false;
      for (let b of aptBookings) {
        if (b.start.getTime() < minRequiredEnd.getTime() && b.end.getTime() > availableFrom.getTime()) {
          hasConflict = true;
          break;
        }
      }

      if (!hasConflict) {
        rowsAdvertiseNames.push(aptName);
      }
    }
  });

  // Alphabetisch sortieren
  rowsAdvertiseNames.sort((a, b) => a.localeCompare(b, 'de', { numeric: true }));

  let rowsAdvertise = [];
  rowsAdvertiseNames.forEach(apt => rowsAdvertise.push([apt, "", "", "", ""]));

  // 4. SORTIERUNG UND GRUPPIERUNG FÜR KICKOUT (1:1 aus deinem Code)

  // a) Maximale Überfälligkeit pro Firma/Gast ermitteln
  const maxOverduePerGuest = {};
  rowsKickoutRaw.forEach(item => {
    const key = item.guestName.toLowerCase();
    if (!maxOverduePerGuest[key] || item.overdueDays > maxOverduePerGuest[key]) {
      maxOverduePerGuest[key] = item.overdueDays;
    }
  });

  // b) Sortieren nach Firmentypen (Gruppierung)
  rowsKickoutRaw.sort((a, b) => {
    const keyA = a.guestName.toLowerCase();
    const keyB = b.guestName.toLowerCase();

    const maxA = maxOverduePerGuest[keyA];
    const maxB = maxOverduePerGuest[keyB];

    // 1. Primär: Höchste Überfälligkeit der gesamten Firma absteigend
    if (maxB !== maxA) {
      return maxB - maxA;
    }

    // 2. Sekundär: Wenn zwei Firmen dieselbe maximale Überfälligkeit haben -> alphabetisch nach Gastname
    const guestComp = a.guestName.localeCompare(b.guestName, 'de', { numeric: true, sensitivity: 'base' });
    if (guestComp !== 0) return guestComp;

    // 3. Tertiär: Buchungen derselben Firma absteigend nach ihrer individuellen Überfälligkeit
    if (b.overdueDays !== a.overdueDays) {
      return b.overdueDays - a.overdueDays;
    }

    // 4. Quartär: Nach Apartment-Name
    return a.aptName.localeCompare(b.aptName, 'de', { numeric: true, sensitivity: 'base' });
  });

  const finalRowsKickout = rowsKickoutRaw.map(item => item.data);

  // Sortierung für cal_re_erstellen (nach Gastname, 1:1 aus deinem Code)
  rowsReErstellen.sort((a, b) => {
    const guestA = String(a[4] || "").trim();
    const guestB = String(b[4] || "").trim();
    const guestComp = guestA.localeCompare(guestB, 'de', { numeric: true, sensitivity: 'base' });
    if (guestComp !== 0) return guestComp;
    
    const aptA = String(a[0] || "").trim();
    const aptB = String(b[0] || "").trim();
    return aptA.localeCompare(aptB, 'de', { numeric: true, sensitivity: 'base' });
  });

  // 5. TABS SCHREIBEN (1:1 aus deinem Code)
  const baseHeaders = [
    "property_name", 
    "periods.start", 
    "periods.end", 
    "periods.closed_period.id", 
    "guest_name", 
    "phone", 
    "origin", 
    "vAUTO"
  ];

  // Tab: cal_re_erstellen
  writeSheet(ss, "cal_re_erstellen", [...baseHeaders, "RE bis heute?"], rowsReErstellen);

  // Tab: cal_kickout
  writeSheet(ss, "cal_kickout", [...baseHeaders, "Offene RE (Tage überfällig)"], finalRowsKickout);

  // Tab: cal_advertise
  const advertiseHeaders = [
    "property_name",
    "Preis",
    "Tage ab heute",
    "Mindestaufenthalt",
    "Status / Rückmeldung"
  ];

  writeSheet(ss, "cal_advertise", advertiseHeaders, rowsAdvertise);
}

function writeSheet(ss, sheetName, headers, rows) {
  let sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
}