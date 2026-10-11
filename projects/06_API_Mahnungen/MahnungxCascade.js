/**
 * MASTER CASCADE WORKFLOW:
 * Führt alle Mahnschritte geordnet nacheinander aus.
 * Diesen Funktionsnamen als 1x täglichen Trigger (Zeitgesteuert) einstellen!
 */
function runMasterDunningCascade() {
  Logger.log("==================================================");
  Logger.log("🚀 STARTE TÄGLICHEN MAHN-CASCADE-WORKFLOW...");
  Logger.log("==================================================");

  // 1. NEUE RECHNUNGEN Suchen & in 'mah' Eintragen
  Logger.log("\n1️⃣ SCHRITT 1: Neue überfällige Rechnungen abrufen...");
  try {
    fillMahnwesenTabFromAPI();
  } catch (e) {
    Logger.log(`🚨 Fehler in Schritt 1: ${e.message}`);
  }
  Utilities.sleep(2000);

  // 1.5 FEHLENDE E-MAILS AUS EXTERNEM SHEET ERGÄNZEN
  Logger.log("\n📧 SCHRITT 1.5: Fehlende E-Mail-Adressen ergänzen...");
  try {
    fillMissingEmails();
  } catch (e) {
    Logger.log(`🚨 Fehler beim Ergänzen der E-Mail-Adressen: ${e.message}`);
  }
  Utilities.sleep(2000);

  // 2. RECHNUNGSSTATUS Prüfen (Spalte J: Bezahlt / Storniert)
  Logger.log("\n2️⃣ SCHRITT 2: Rechnungsstatus in Spalte J aktualisieren...");
  try {
    updateMahInvoiceStatus();
  } catch (e) {
    Logger.log(`🚨 Fehler in Schritt 2: ${e.message}`);
  }
  Utilities.sleep(2000);

  // 3. ERSTE MAHNUNGEN Erstellen & Versenden (für neue Zeilen)
  Logger.log("\n3️⃣ SCHRITT 3: 1. Mahnungen verarbeiten...");
  try {
    createAndSendAllDunningsForMahTab();
  } catch (e) {
    Logger.log(`🚨 Fehler in Schritt 3: ${e.message}`);
  }
  Utilities.sleep(2000);

  // 4. ZWEITE MAHNUNGEN Erstellen & Versenden (Spalte K, falls Spalte I >= 3 Tage her)
  Logger.log("\n4️⃣ SCHRITT 4: 2. Mahnungen verarbeiten...");
  try {
    m2_sendSecondDunning();
  } catch (e) {
    Logger.log(`🚨 Fehler in Schritt 4: ${e.message}`);
  }
  Utilities.sleep(2000);

  // 5. DRITTE MAHNUNGEN Erstellen & Versenden (Spalte L, falls Spalte K >= 3 Tage her)
  Logger.log("\n5️⃣ SCHRITT 5: 3. Mahnungen verarbeiten...");
  try {
    m3_sendThirdDunning();
  } catch (e) {
    Logger.log(`🚨 Fehler in Schritt 5: ${e.message}`);
  }

  Logger.log("==================================================");
  Logger.log("✅ TÄGLICHER MAHN-WORKFLOW ERFOLGREICH BEENDET");
  Logger.log("==================================================");
}