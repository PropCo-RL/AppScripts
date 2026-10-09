/**
 * Master-Steuerung für den 15-Minuten-Trigger
 */
function runAllProcesses() {
  Logger.log("--- 0. Min-Stay vorab auf 1 setzen ---");
  AAA_0_setMinStayToOne();

  // WICHTIG: 3 Sekunden Puffer, damit Lodgify das MinStay-Update auf seinen Servern verarbeitet
  Logger.log("Puffer: Warte 3 Sekunden, damit Lodgify min_stay=1 übernimmt...");
  Utilities.sleep(3000);

  Logger.log("--- 1. Lodgify Buchung starten ---");
  AAA_startLodgifyBooking();

  Logger.log("--- 1.1. Min-Stay wieder auf 14 zurücksetzen ---");
  ZZZ_resetMinStayToFourteen();
  
  Logger.log("--- 2. Lexoffice Rechnung erstellen ---");
  BBB_createLexofficeInvoices();
  
  Logger.log("Puffer: Warte 5 Sekunden vor dem E-Mail-Versand...");
  Utilities.sleep(5000);

  Logger.log("--- 3. Lexoffice Rechnungen per E-Mail versenden ---");
  CCC_sendLexofficeInvoiceEmails();

  Logger.log("Puffer: Warte 2 Sekunden vor der Synchronisation...");
  Utilities.sleep(2000);

  Logger.log("--- 4. Versendestatus in externes Sheet synchronisieren ---");
  CCC_syncEmailStatusToExternalSheet();

  Logger.log("--- ALLE PROZESSE ERFOLGREICH DURCHGELAUFEN ---");
}