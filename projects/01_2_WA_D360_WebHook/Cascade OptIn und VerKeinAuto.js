// ============================================================================
// MASTER 2: SERVICE & OPT-IN (verKeinAuto / OptIn) - TRIGGER ALLE 15 MINUTEN
// ============================================================================
/**
 * Haupt-Steuerung für WhatsApp Service-Fenster, automatischen Versand & Opt-In.
 * Auf diesen Funktionsnamen setzt du den zweiten 15-Minuten-Zeit-Trigger.
 */
function runMasterProcessServiceAndOptIn() {
  Logger.log("=== START: MASTER-PROZESS SERVICE & OPT-IN ===");

  try {
    // 1. Service-Fenster und Abreisefrist prüfen
    Logger.log("--- Step 1: checkVerKeinAutoStatus() gestartet ---");
    checkVerKeinAutoStatus();

    Utilities.sleep(1000);

    // 2. Automatischer WhatsApp-Versand für 'verKeinAuto' über 360dialog
    Logger.log("--- Step 2: sendVerKeinAutoWhatsAppMessages() gestartet ---");
    sendVerKeinAutoWhatsAppMessages();

    Utilities.sleep(1000);

    // 3. Opt-In / Opt-Out Rückmeldungen aus dem Log verarbeiten
    Logger.log("--- Step 3: processOptInEvents() gestartet ---");
    processOptInEvents();

    Utilities.sleep(1000);

    // 4. Automatischen & manuellen Opt-In Template-Versand prüfen & ausführen
    Logger.log("--- Step 4: sendOptInTemplateOnTrigger() gestartet ---");
    sendOptInTemplateOnTrigger();

  } catch (e) {
    Logger.log("🚨 FEHLER im Master-Prozess Service & Opt-In: " + e.toString());
  }

  Logger.log("=== ENDE: MASTER-PROZESS SERVICE & OPT-IN ERFOLGREICH ===");
}