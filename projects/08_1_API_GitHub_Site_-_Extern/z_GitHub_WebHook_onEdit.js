function triggerGitHubBuild() {
  const GITHUB_TOKEN = PropertiesService.getScriptProperties().getProperty('GitHub_API_Sites_Extern');
  const REPO_OWNER = "PropCo-RL";
  const REPO_NAME = "apiExtern";
  
  // 1. VORAB-PRÜFUNG: Daten aus dem Sheet analysieren & im Log protokollieren
  console.log("🔍 Starte Datenprüfung vor dem Absenden...");
  
  try {
    const allApts = getAllApartmentsForBuild();
    console.log(`📊 Gefundene Apartments gesamt: ${allApts.length}`);

    let validCount = 0;
    let missingPathCount = 0;
    const citiesSet = new Set();

    allApts.forEach((apt, index) => {
      if (apt.Apartment && apt.Apartment.trim() !== "") {
        validCount++;
        if (apt.city) citiesSet.add(apt.city.trim());
      } else {
        missingPathCount++;
        console.warn(`⚠️ Zeile ${index + 2}: Kein Pfad in Spalte G für "${apt.title || 'Unbekannt'}"`);
      }
    });

    console.log(`✅ Gültige Seiten zum Generieren: ${validCount}`);
    console.log(`🏙️ Erfasste Städte (${citiesSet.size}): ${Array.from(citiesSet).join(", ")}`);
    
    if (missingPathCount > 0) {
      console.warn(`⚠️ Achtung: ${missingPathCount} Zeilen haben keinen Pfad in Spalte G und werden übersprungen.`);
    }

  } catch (err) {
    console.error("❌ Fehler bei der Datenprüfung im Sheet: " + err.toString());
  }

  // 2. SIGNAL AN GITHUB SENDEN
  const url = `https://api.github.com/repos/${REPO_OWNER}/${REPO_NAME}/dispatches`;
  
  const options = {
    method: "post",
    headers: {
      "Authorization": "Bearer " + GITHUB_TOKEN,
      "Accept": "application/vnd.github.v3+json",
      "Content-Type": "application/json"
    },
    payload: JSON.stringify({
      event_type: "sheet_updated"
    }),
    muteHttpExceptions: true
  };
  
  console.log("🚀 Sende Build-Signal an GitHub...");
  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  
  if (responseCode === 204) {
    console.log("✅ Erfolgreich an GitHub übermittelt (Status 204 No Content). Build gestartet!");
  } else {
    console.error(`❌ Fehler beim Senden an GitHub! Status Code: ${responseCode}`);
    console.error("Antwort von GitHub: " + response.getContentText());
  }
}