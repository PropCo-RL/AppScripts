/**
 * Liest die Zugangsdaten aus den Script Properties aus
 * und sendet die Test-SMS über den Twilio Irland-Endpoint.
 */
function sendTestSMS() {
  const scriptProperties = PropertiesService.getScriptProperties();
  
  // Zuordnung deiner Script Properties:
  const accountSid = scriptProperties.getProperty('TWILIO_ACCOUNT_SID'); 
  const apiKeySid  = scriptProperties.getProperty('TwilioKeySID');     
  const apiSecret  = scriptProperties.getProperty('TwilioKey');          
  const serviceSid = scriptProperties.getProperty('TwilioServiceSID');   

  // Festgelegte Zielnummer und Nachrichtentext
  const empfaengerHandy = '+4917684801295'; 
  const text = 'Hallo! Eure Buchung steht. Bei Fragen schreibt uns auf WhatsApp: https://wa.me/4915123456789';

  // Sicherheitsprüfung
  if (!accountSid || !apiKeySid || !apiSecret || !serviceSid) {
    Logger.log('❌ FEHLER: Es fehlen noch Script Properties!');
    return;
  }

  // Twilio Irland (IE1) API Endpoint – passend zu deinem IE1 API Key
  const url = 'https://api.dublin.ie1.twilio.com/2010-04-01/Accounts/' + accountSid + '/Messages.json';
  
  // Nutzdaten
  const payload = {
    'To': empfaengerHandy,
    'MessagingServiceSid': serviceSid,
    'Body': text
  };
  
  // Basic Auth: TwilioKeySID (Username) : TwilioKey (Password/Secret)
  const options = {
    'method': 'post',
    'headers': {
      'Authorization': 'Basic ' + Utilities.base64Encode(apiKeySid + ':' + apiSecret)
    },
    'payload': payload,
    'muteHttpExceptions': true
  };
  
  // API Call ausführen
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const result = JSON.parse(response.getContentText());
    
    if (responseCode === 200 || responseCode === 201) {
      Logger.log('✅ SMS erfolgreich gesendet an ' + empfaengerHandy + '! Message SID: ' + result.sid);
    } else {
      Logger.log('❌ Fehler von Twilio (Status ' + responseCode + '): ' + result.message);
    }
  } catch (error) {
    Logger.log('❌ Netzwerk-/Systemfehler: ' + error.toString());
  }
}