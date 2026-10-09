/**
 * Master-Steuerung für Verlängerungen
 */
function runAllProcessesVL() {
  // Logger.log("--- 0. Fehlende Kundendaten aus RE / VL anreichern ---");
  // AAA_0_enrichMissingCustomerDataVL(); // DEAKTIVIERT: Verhindert das Überschreiben von Daten & Zeilenversatz

  Logger.log("--- 1. Lodgify Buchung aktualisieren ---");
  AAA_updateLodgifyBookingVL();
  
  Logger.log("--- 2a. Lexoffice Sammel-Verlängerungs-Rechnungen erstellen (Multi-Bookings zuerst) ---");
  BBB_createLexofficeMultiInvoicesVL();

  Logger.log("--- 2b. Lexoffice Einzel-Verlängerungs-Rechnungen erstellen (Verbleibende Zeilen) ---");
  BBB_createLexofficeInvoicesVL();
  
  // 5 Sekunden Puffer, damit alle Lexoffice-IDs & PDFs sicher bereitstehen
  Logger.log("Puffer: Warte 5 Sekunden vor dem E-Mail-Versand...");
  Utilities.sleep(5000);

  Logger.log("--- 3. Lexoffice Verlängerungs-Rechnungen per E-Mail versenden ---");
  CCC_sendLexofficeInvoiceEmailsVL();

  Logger.log("--- ALLE VERLÄNGERUNGS-PROZESSE ERFOLGREICH DURCHGELAUFEN ---");
}