// ==========================================
// ZUSATZ-MODULE: RE-SHEET SYNC & GMAIL LABEL
// ==========================================

/**
 * Fügt eine neue Buchungs-Zeile im Tabellenblatt "RE" der EXTERNEN Tabelle an.
 * Befüllt nur die gewünschten Spalten (A, B, D, E, F, G, H, I, K, L, M, N, O).
 * Hängt die Zeile am Ende an und verfälscht keine bestehenden Daten.
 */
function appendToReSheet(payload) {
  try {
    const externalSpreadsheetId = "1mMEVsVFIx5VxIV3EyJYlba_G0fgEX9t0tR2KbFkm-nU";
    const targetSs = SpreadsheetApp.openById(externalSpreadsheetId);
    const sheetRE = targetSs.getSheetByName("RE");
    
    if (!sheetRE) {
      console.error("Tabellenblatt 'RE' im Ziel-Spreadsheet nicht gefunden.");
      return;
    }

    // Kürzel aufbereiten (z. B. aus "/a/gera4" wird "gera4")
    let cleanAptCode = String(payload.aptCode || "").replace(/^\/a\//, "").trim();
    let cycle = payload.billingCycle || "Fixes Enddatum";

    // Array für 15 Spalten (Index 0 bis 14 = Spalte A bis O)
    let row = new Array(15).fill("");
    row[0]  = payload.phone || "";                            // A: Telefonnummer
    row[1]  = payload.company || "";                          // B: Firmenname
    // Spalte C (Index 2) bleibt leer
    row[3]  = payload.street || "";                           // D: Straße mit Hausnummer
    row[4]  = payload.zip || "";                              // E: PLZ
    row[5]  = payload.city || "";                             // F: Stadt
    row[6]  = payload.country || "";                          // G: Land
    row[7]  = payload.vatId || "";                            // H: VAT / USt-ID
    row[8]  = payload.email || "";                            // I: E-Mail
    // Spalte J (Index 9) bleibt leer
    row[10] = payload.startDate || "";                        // K: Anreise
    row[11] = payload.endDate || "";                          // L: Abreise
    row[12] = cycle;                                          // M: Abrechnungszyklus
    row[13] = cleanAptCode;                                   // N: Apartment Kürzel
    row[14] = payload.pricePerNight || payload.price || "";   // O: Preis

    // Zeile ganz unten anfügen
    sheetRE.appendRow(row);
  } catch (err) {
    console.error("Fehler beim Einfügen in externes Blatt RE: " + err.toString());
  }
}

/**
 * Sucht die gesendete E-Mail anhand des Betreffs und weist das Label "01_Online Buchung" zu.
 */
function applyBookingLabel(emailSubject) {
  try {
    const labelName = "01_Online Buchung";
    let label = GmailApp.getUserLabelByName(labelName);
    if (!label) {
      label = GmailApp.createLabel(labelName);
    }

    Utilities.sleep(1500); // 1.5s Puffer, damit Gmail die Mail verarbeitet hat
    const threads = GmailApp.search('subject:"' + emailSubject + '"', 0, 1);
    if (threads && threads.length > 0) {
      label.addToThread(threads[0]);
    } else {
      console.warn("E-Mail mit Betreff '" + emailSubject + "' wurde nicht zum Labeln gefunden.");
    }
  } catch (err) {
    console.error("Fehler beim Zuweisen des E-Mail-Labels: " + err.toString());
  }
}