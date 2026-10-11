const RECIPIENT_EMAIL = "info@l8street.com"; // Hier später auf die reale Support-Mail umstellen!

/**
 * Prüft Spalte H auf 'x' und versendet die E-Mails bei Freigabe
 */
function sendApprovedWifiEmails() {
  Logger.log("=== PRÜFE MANUELLE FREIGABEN IN SPALTE H ===");
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("log");
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  
  let rowsToProcess = [];

  for (let i = 1; i < data.length; i++) {
    const rowIndex = i + 1;
    const approval = data[i][7] ? String(data[i][7]).trim().toLowerCase() : ""; // Spalte H (Index 7)

    // Nur Zeilen verarbeiten, in denen manuell ein 'x' eingetragen wurde
    if (approval === "x") {
      rowsToProcess.push({
        rowIndex: rowIndex,
        apartment: data[i][2],
        simNeu: data[i][3],
        imeiNeu: data[i][4],
        simBestand: data[i][5],
        imeiBestand: data[i][6]
      });
    }
  }

  if (rowsToProcess.length === 0) {
    Logger.log("Keine freigegebenen Zeilen (mit 'x' in Spalte H) gefunden.");
    return;
  }

  Logger.log(`Gefundene freigegebene Einträge: ${rowsToProcess.length}. Sende E-Mail...`);

  let deviceDetailsText = "";
  rowsToProcess.forEach((item, index) => {
    deviceDetailsText += `Gerät ${index + 1}:\n`;
    if (item.imeiNeu || item.simNeu) {
      deviceDetailsText += `IMEI Neu: ${item.imeiNeu || "N/A"}\n`;
      deviceDetailsText += `SIM Neu: ${item.simNeu || "N/A"}\n`;
    }
    if (item.imeiBestand || item.simBestand) {
      deviceDetailsText += `IMEI Bestand/Alt: ${item.imeiBestand || "N/A"}\n`;
      deviceDetailsText += `SIM Bestand/Alt: ${item.simBestand || "N/A"}\n`;
    }
    deviceDetailsText += `Standort: ${item.apartment || "N/A"}\n\n`;
  });

  const subject = "Meldung zu Wlan Störung";
  const emailBody = `Guten Tag,

wir melden hiermit eine WLAN-Störung für die folgenden Geräte:

${deviceDetailsText}Ein Neustart wurde bereits durchgeführt.

Mit freundlichen Grüßen / Best regards

Raul Leneweit`;

  // E-Mail senden
  GmailApp.sendEmail(RECIPIENT_EMAIL, subject, emailBody);

  // Status in Spalte H auf 'sent' aktualisieren
  rowsToProcess.forEach(item => {
    sheet.getRange(item.rowIndex, 8).setValue("sent");
  });

  Logger.log("E-Mails erfolgreich versendet und Spalte H auf 'sent' gesetzt!");
}