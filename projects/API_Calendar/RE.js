/**
 * Generiert das Tabellenblatt "RE":
 * - Basiert auf Buchungen aus Tab "db".
 * - Gekoppelt mit Rechnungen aus Tab "inRE" als Unterzeilen (->).
 * - FILTER:
 *   1. Keine BookingCom-Buchungen (origin)
 *   2. Keine Blockierungen (dbX)
 *   3. Keine Multi-Apartment-Buchungen (gleicher Name oder Tel-Nr. bei verschiedenen Apartments)
 */
function generateReSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 1. DATEN AUS TAB "inRE" IN RECHNUNGSMAP EINLESEN ---
  const sheetInRE = ss.getSheetByName("inRE");
  const invoiceMap = {};

  if (sheetInRE && sheetInRE.getLastRow() >= 2) {
    const valuesInRE = sheetInRE.getRange(2, 1, sheetInRE.getLastRow() - 1, 5).getValues();
    valuesInRE.forEach(row => {
      const lodgifyId = String(row[0] || "").trim();
      const lxInvoiceId = String(row[1] || "").trim();
      const reStart = row[2] instanceof Date ? Utilities.formatDate(row[2], "GMT+2", "yyyy-MM-dd") : String(row[2] || "").trim();
      const reEnde = row[3] instanceof Date ? Utilities.formatDate(row[3], "GMT+2", "yyyy-MM-dd") : String(row[3] || "").trim();
      const reStatus = String(row[4] || "").trim();

      if (lodgifyId) {
        if (!invoiceMap[lodgifyId]) {
          invoiceMap[lodgifyId] = [];
        }
        invoiceMap[lodgifyId].push({
          lxId: lxInvoiceId,
          start: reStart,
          ende: reEnde,
          status: reStatus
        });
      }
    });
  }

  // --- 2. APARTMENTS (ID -> NAME) AUS TAB "d" LADEN ---
  const sheetD = ss.getSheetByName("d");
  const propertyMap = {};

  if (sheetD && sheetD.getLastRow() >= 1) {
    const valuesD = sheetD.getRange(1, 1, sheetD.getLastRow(), 4).getValues();
    valuesD.forEach(row => {
      const id = String(row[0]).trim();
      const name = String(row[2]).trim();
      if (id !== "" && id.toLowerCase() !== "property_id") {
        propertyMap[id] = name || id;
      }
    });
  }

  // Datums-Grenzen (-7 bis +365 Tage)
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const pastLimit = new Date(today);
  pastLimit.setDate(today.getDate() - 7);

  const futureLimit = new Date(today);
  futureLimit.setDate(today.getDate() + 365);

  const autoLimit = new Date(today);
  autoLimit.setDate(today.getDate() + 30);

  let rawEntries = [];

  // --- 3. BUCHUNGEN AUS TAB "db" EINLESEN ---
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

      const isBookingCom = origin.toLowerCase().includes("bookingcom");

      // FILTER 1: Keine BookingCom-Buchungen
      if (propId && arrival && departure && !isBookingCom) {
        arrival.setHours(0, 0, 0, 0);
        departure.setHours(0, 0, 0, 0);

        const aptName = propertyMap[propId] || propId;
        const isCurrentlyActive = arrival <= today && departure > today;

        if (departure >= pastLimit && arrival <= futureLimit) {
          let vAutoStatus = "-";
          if (isCurrentlyActive && departure > autoLimit) {
            vAutoStatus = "verAuto";
          }

          rawEntries.push({
            aptName: aptName,
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

  // (HINWEIS: Tab "dbX" / Blockierungen wurden hier komplett weggelassen)

  // --- 4. MULTI-APARTMENT ANALYSE ---
  // Prüft, ob ein Name oder eine Telefonnummer zu MEHREREN VERSCHIEDENEN Apartments gehört
  const apartmentsByName = {};
  const apartmentsByPhone = {};

  rawEntries.forEach(entry => {
    const nameKey = entry.guestName.toLowerCase().trim();
    const phoneKey = entry.phone.trim();

    if (nameKey !== "") {
      if (!apartmentsByName[nameKey]) apartmentsByName[nameKey] = new Set();
      apartmentsByName[nameKey].add(entry.aptName);
    }
    if (phoneKey !== "") {
      if (!apartmentsByPhone[phoneKey]) apartmentsByPhone[phoneKey] = new Set();
      apartmentsByPhone[phoneKey].add(entry.aptName);
    }
  });

  // FILTER 3: Nur Buchungen behalten, die KEINE Multi-Apartment-Buchungen sind
  let filteredEntries = rawEntries.filter(entry => {
    const nameKey = entry.guestName.toLowerCase().trim();
    const phoneKey = entry.phone.trim();

    const hasMultiByName = nameKey !== "" && apartmentsByName[nameKey] && apartmentsByName[nameKey].size > 1;
    const hasMultiByPhone = phoneKey !== "" && apartmentsByPhone[phoneKey] && apartmentsByPhone[phoneKey].size > 1;

    // Wenn er mehrere Apartments gebucht hat, wird die Buchung gefiltert
    return !hasMultiByName && !hasMultiByPhone;
  });

  // Nach Apartment & Abreise sortieren
  filteredEntries.sort((a, b) => {
    const nameComp = a.aptName.localeCompare(b.aptName, 'de', { numeric: true, sensitivity: 'base' });
    if (nameComp !== 0) return nameComp;
    return a.departure - b.departure;
  });

  // --- 5. ZEILEN FÜR TAB "RE" AUFBAUEN ---
  let reRows = [];
  let lastApt = null;

  filteredEntries.forEach(entry => {
    // Leere Trennzeile bei neuem Apartment
    if (lastApt !== null && entry.aptName !== lastApt) {
      reRows.push(["", "", "", "", "", "", "", "", "", "", "", ""]);
    }

    // 1. Hauptzeile (Buchung)
    reRows.push([
      entry.aptName,
      Utilities.formatDate(entry.arrival, "GMT+2", "yyyy-MM-dd"),
      Utilities.formatDate(entry.departure, "GMT+2", "yyyy-MM-dd"),
      entry.cpId,
      entry.guestName,
      entry.phone,
      entry.origin,
      entry.vAuto,
      "", "", "", "" // Spalten I bis L bleiben bei der Buchungszeile leer
    ]);

    // 2. Unterzeilen (Rechnungen aus inRE anhängen & sortieren)
    const matchingInvoices = invoiceMap[entry.cpId];
    if (matchingInvoices && matchingInvoices.length > 0) {
      
      // SORTIERUNG: Neuestes Enddatum nach oben (absteigend)
      matchingInvoices.sort((a, b) => {
        const dateA = a.ende || a.start || "";
        const dateB = b.ende || b.start || "";
        return dateB.localeCompare(dateA);
      });

      matchingInvoices.forEach(inv => {
        reRows.push([
          "->", // Spalte A
          "", "", "", "", "", "", "", // Spalten B bis H leer
          inv.lxId,   // Spalte I: LX ID
          inv.start,  // Spalte J: RE Start
          inv.ende,   // Spalte K: RE Ende
          inv.status  // Spalte L: Status
        ]);
      });
    }

    lastApt = entry.aptName;
  });

  // --- 6. TAB "RE" SCHREIBEN ---
  let sheetRE = ss.getSheetByName("RE") || ss.insertSheet("RE");
  sheetRE.clear();

  const headersRE = [
    "property_name", 
    "periods.start", 
    "periods.end", 
    "periods.closed_period.id", 
    "guest_name", 
    "phone", 
    "origin",
    "vAUTO",
    "re_id",
    "re_start",
    "re_ende",
    "re_status"
  ];

  sheetRE.getRange(1, 1, 1, headersRE.length).setValues([headersRE]).setFontWeight("bold");

  if (reRows.length > 0) {
    sheetRE.getRange(2, 1, reRows.length, headersRE.length).setValues(reRows);
    Logger.log(reRows.length + " Zeilen erfolgreich in 'RE' geschrieben.");
  } else {
    Logger.log("Keine Einträge für 'RE' vorhanden.");
  }
}