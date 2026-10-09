// ============================================================================
// RECHNUNGS-PDF DOWNLOAD & E-MAIL VERSAND (SPALTE V) - TAB "RE"
// ============================================================================

/**
 * HAUPTFUNKTION: Versendet alle fertigen Lexoffice-Rechnungen aus Spalte T per E-Mail
 * Liest Kunden-E-Mail aus Spalte I und schreibt Versendestatus in Spalte V.
 * NEU: Verhindert Mehrfachversand bei Sammelrechnungen (gleiche Invoice-ID).
 */
function CCC_sendLexofficeInvoiceEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetRE = ss.getSheetByName("RE");

  if (!sheetRE) throw new Error("Das Tabellenblatt 'RE' wurde nicht gefunden.");

  const lastRow = sheetRE.getLastRow();
  if (lastRow < 2) {
    Logger.log("Keine Datenzeilen im Sheet 'RE' gefunden.");
    return;
  }

  // 22 Spalten lesen (Spalte A bis V)
  const range = sheetRE.getRange(2, 1, lastRow - 1, 22);
  const values = range.getValues();

  Logger.log(`Starte E-Mail-Prüfung für ${values.length} Zeile(n)...`);

  let processedCount = 0;
  // Set zum Verfolgen von IDs, die in diesem Durchlauf bereits versendet wurden
  const processedInvoiceIds = new Set();

  values.forEach((row, index) => {
    const rowIndex = index + 2;
    const lxInvoiceId = row[19] ? row[19].toString().trim() : ""; // Spalte T (Index 19): Invoice ID
    const emailStatus = row[21] ? row[21].toString().trim() : ""; // Spalte V (Index 21): E-Mail Versendestatus

    // 1. Wenn keine Lexoffice ID in Spalte T steht -> Überspringen
    if (!lxInvoiceId || lxInvoiceId.startsWith("ERROR")) {
      Logger.log(`Zeile ${rowIndex} übersprungen: Keine gültige Lexoffice ID in Spalte T.`);
      return;
    }

    // 2. Wenn Spalte V bereits befüllt ist -> Überspringen (bereits gesendet)
    if (emailStatus !== "") {
      // ID trotzdem als bereits verarbeitet registrieren, falls vorherige Zeilen schon gesendet wurden
      processedInvoiceIds.add(lxInvoiceId);
      Logger.log(`Zeile ${rowIndex} übersprungen: Spalte V enthält bereits Status ('${emailStatus}').`);
      return;
    }

    // 3. MULTI-PRÜFUNG: Wenn für diese Invoice-ID in diesem Durchlauf bereits eine E-Mail versendet wurde
    if (processedInvoiceIds.has(lxInvoiceId)) {
      sheetRE.getRange(rowIndex, 22).setValue("Übersprungen (Multi-Position)");
      Logger.log(`Zeile ${rowIndex} übersprungen: Sammelrechnung ${lxInvoiceId} wurde in diesem Durchlauf bereits per E-Mail versendet.`);
      return;
    }

    // Ab hier: Gültige ID da und E-Mail noch nicht gesendet!
    processedCount++;

    const firma = row[1] ? row[1].toString().trim() : "";
    const name = row[2] ? row[2].toString().trim() : "";
    const recipientEmail = row[8] ? row[8].toString().trim() : ""; // Spalte I: E-Mail
    const apartmentCode = row[13] ? row[13].toString().trim() : "";

    const recipientName = firma || name || "Gast";

    // Prüfen, ob eine gültige E-Mail-Adresse vorliegt
    if (!recipientEmail || !recipientEmail.includes("@")) {
      sheetRE.getRange(rowIndex, 22).setValue("Keine gültige E-Mail vorhanden");
      Logger.log(`Zeile ${rowIndex}: Keine gültige E-Mail in Spalte I ('${recipientEmail}').`);
      return;
    }

    Logger.log(`[Zeile ${rowIndex}] Lade PDF für Invoice ID ${lxInvoiceId} herunter...`);

    // PDF aus Lexoffice laden
    const pdfBlob = downloadLexofficeInvoicePdf(lxInvoiceId);

    if (pdfBlob) {
      // Dateiname des Anhangs sauber formatieren
      const cleanFileName = recipientName.replace(/[^a-zA-Z0-9]/g, "_");
      pdfBlob.setName(`Rechnung_${cleanFileName}.pdf`);

      const subject = `Ihre Rechnung von L8 Street – Buchung ${apartmentCode.toUpperCase()}`;
      const body = `Hallo ${recipientName},\n\nvielen Dank für Ihre Buchung bei L8 Street!\n\nIm Anhang finden Sie Ihre finale Rechnung als PDF-Dokument.\n\nBei Fragen stehen wir Ihnen jederzeit zur Verfügung.\n\nMit freundlichen Grüßen\nIhr L8 Street Team`;

      try {
        // E-Mail senden mit CC und Reply-To
        MailApp.sendEmail({
          to: recipientEmail,
          cc: "info@l8street.com",          // ✉️ Immer in CC setzen
          replyTo: "info@l8street.com",     // ↩️ Antworten gehen an info@l8street.com
          subject: subject,
          body: body,
          attachments: [pdfBlob]
        });

        // Nach erfolgreichem Versand: ID im Speicher merken
        processedInvoiceIds.add(lxInvoiceId);

        const sentTimeStr = Utilities.formatDate(new Date(), "Europe/Berlin", "dd.MM.yyyy HH:mm");
        sheetRE.getRange(rowIndex, 22).setValue(`Gesendet an ${recipientEmail} (${sentTimeStr})`); // Schreiben in Spalte V
        Logger.log(`[ERFOLG Zeile ${rowIndex}] E-Mail an ${recipientEmail} (CC: info@l8street.com) versendet!`);

      } catch (mailErr) {
        Logger.log(`🚨 Fehler beim E-Mail-Versand Zeile ${rowIndex}: ${mailErr.message}`);
        sheetRE.getRange(rowIndex, 22).setValue(`ERROR Mail: ${mailErr.message.substring(0, 100)}`);
      }
    } else {
      sheetRE.getRange(rowIndex, 22).setValue("ERROR: PDF-Download fehlgeschlagen");
      Logger.log(`🚨 Zeile ${rowIndex}: PDF-Download von Lexoffice fehlgeschlagen.`);
    }

    Utilities.sleep(500); // 0,5 Sek. Pause
  });

  Logger.log(`--- PROZESS BEENDET: ${processedCount} Zeile(n) verarbeitet. ---`);
}

/**
 * Lädt das finale Rechnungs-PDF als Binary Data von Lexoffice herunter
 */
function downloadLexofficeInvoicePdf(invoiceId) {
  const props = PropertiesService.getScriptProperties();
  const lxKey = props.getProperty("lxKey");

  if (!lxKey) throw new Error("Script Property 'lxKey' fehlt!");

  const targetUrl = `https://api.lexoffice.io/v1/invoices/${invoiceId}/file`;

  const options = {
    "method": "get",
    "headers": {
      "Authorization": "Bearer " + lxKey,
      "Accept": "application/pdf"
    },
    "muteHttpExceptions": true
  };

  try {
    const response = UrlFetchApp.fetch(targetUrl, options);
    if (response.getResponseCode() === 200) {
      return response.getBlob();
    } else {
      Logger.log(`🚨 PDF-Download Fehler (HTTP ${response.getResponseCode()}): ${response.getContentText()}`);
    }
  } catch (e) {
    Logger.log("Fehler beim PDF-Download: " + e.toString());
  }

  return null;
}