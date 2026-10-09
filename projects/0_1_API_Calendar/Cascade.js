/**
 * Master Controller: Führt die gesamte Kaskade nacheinander aus.
 * Diesen Trigger einfach 1x pro Stunde automatisch ausführen lassen.
 */
function runLodgifySyncCycle() {
  console.log("--- START: Lodgify Sync Cycle ---");

  try {
    // 1. Buchungen von Lodgify holen -> schreibt in 'db'
    console.log("Schritt 1: Rufe Lodgify Buchungen ab (Tab 'db')...");
    getLodgifyBookings();
    
    // 2. Verfügbarkeiten / Blockierungen von Lodgify holen -> schreibt in 'dbX'
    console.log("Schritt 2: Rufe Lodgify Blockierungen ab (Tab 'dbX')...");
    getLodgifyAvailability();
    
    // 3. Beide Quellen (db & dbX) mit Namens-Mapping aus 'd' in Tab 'cal' zusammenführen
    console.log("Schritt 3: Generiere Tab 'cal' & Auswertungen...");
    generateBaseSheets(); // Bzw. generateCalAndVerSheets() / generateCalSheet()

    // 4. Rechnungsdaten aus 'inRE' abgleichen und Spalten I & J in 'cal' ergänzen
    console.log("Schritt 4: Reichere Tab 'cal' mit Rechnungsdaten an (Spalten I & J)...");
    enrichCalWithInvoices();

    // 5. Ziel-Tabs (cal_re_erstellen, cal_kickout, cal_advertise) generieren
    console.log("Schritt 5: Generiere Ziel-Tabs (cal_re_erstellen, cal_kickout, cal_advertise)...");
    generateTargetSheets();

    console.log("--- ERFOLG: Alle Daten synchronisiert, angereichert & alle Kalender-Tabs aktualisiert. ---");
  } catch (e) {
    console.error("Fehler im Master-Prozess: " + e.toString());
  }
}