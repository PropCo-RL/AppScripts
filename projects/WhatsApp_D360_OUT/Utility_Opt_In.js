function sendWhatsAppOptIn() {
  const p = PropertiesService.getScriptProperties();
  
  // 1. Hole deine Daten (Achte auf exakte Großschreibung wie in deinen Settings!)
  const apiKey = p.getProperty('D360-KEY');
  const namespace = p.getProperty('D360-Namespace');
  const targetNumber = "447311557295"; // Die Nummer, die gerade im Tool funktioniert hat
  
  const url = "https://waba-v2.360dialog.io/messages";

  // 2. Das JSON-Payload (exakt wie im erfolgreichen Test)
  const payload = {
    "messaging_product": "whatsapp",
    "recipient_type": "individual",
    "to": targetNumber,
    "type": "template",
    "template": {
      "namespace": namespace,
      "name": "utility_opt_in",
      "language": {
        "code": "de" // Falls Fehler 132001 erscheint, hier "de_DE" testen
      }
    }
  };

  const options = {
    "method": "post",
    "headers": {
      "D360-API-KEY": apiKey,
      "Content-Type": "application/json"
    },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  // 3. Ausführung und Log
  try {
    const response = UrlFetchApp.fetch(url, options);
    const resBody = response.getContentText();
    
    Logger.log("Status Code: " + response.getResponseCode());
    Logger.log("Antwort von Meta: " + resBody);

    if (response.getResponseCode() == 200 || response.getResponseCode() == 201) {
      console.log("✅ Nachricht erfolgreich an " + targetNumber + " gesendet!");
    } else {
      console.error("❌ Fehler: " + resBody);
    }
  } catch (e) {
    console.error("❌ Kritischer Fehler im Script: " + e.message);
  }
}