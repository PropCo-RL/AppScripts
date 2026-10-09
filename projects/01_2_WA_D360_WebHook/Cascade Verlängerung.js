// ============================================================================
// MASTER 1: VERLÄNGERUNGEN (VL) - TRIGGER ALLE 15 MINUTEN
// ============================================================================
/**
 * Haupt-Steuerung für alle Prozesse rund um Verlängerungen (VL).
 * Auf diesen Funktionsnamen setzt du den 15-Minuten-Zeit-Trigger.
 */
function runMasterProcessVL() {
  Logger.log("=== START: MASTER-PROZESS VERLÄNGERUNGEN (VL) ===");

  try {
    // 1. Sequentielle Verarbeitung der Verlängerungen (Gemini Extraktion & Log -> VL Sheet)
    Logger.log("--- Step 1: processExtensionsSequentially() gestartet ---");
    processExtensionsSequentially();
    
    // Puffer vor dem nächsten Schritt
    Utilities.sleep(1000);

    // 2. Auswertung der Antworten im 'verKeinAuto'-Prozess via Gemini
    Logger.log("--- Step 2: evaluateVerKeinAutoReplies() gestartet ---");
    evaluateVerKeinAutoReplies();

  } catch (e) {
    Logger.log("🚨 FEHLER im Master-Prozess VL: " + e.toString());
  }

  Logger.log("=== ENDE: MASTER-PROZESS VERLÄNGERUNGEN (VL) ERFOLGREICH ===");
}