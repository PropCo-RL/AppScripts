/**
 * HAUPTSKRIPT FÜR DIE 1. MAHNUNG:
 * Geht alle Zeilen im Tab 'mah' durch.
 * Erstellt eine 1. Mahnung nur wenn:
 * - Rechnungsstatus in Spalte J = 'Offen' ist
 * - Überfällig seit (Tage) in Spalte E >= 3 ist
 * - Noch keine Mahnungs-ID in Spalte H existiert
 * 
 * Erstellt in Lexoffice eine Mahnung inkl. 10 € Mahngebühr (0 % USt),
 * speichert die Mahnungs-ID in Spalte H, lädt das PDF herunter, versendet die E-Mail
 * (inkl. CC an info@l8street.com) und dokumentiert den Versandstatus in Spalte I.
 */
function createAndSendAllDunningsForMahTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mahSheet = ss.getSheetByName("mah");

  if (!mahSheet) {
    Logger.log("🚨 Das Tabellenblatt 'mah' wurde nicht gefunden!");
    SpreadsheetApp.getUi().alert("Fehler: Tabellenblatt 'mah' existiert nicht.");
    return;
  }

  const lastRow = mahSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("ℹ️ Keine Rechnungen im Tab 'mah' vorhanden.");
    return;
  }

  const token = PropertiesService.getScriptProperties().getProperty("lxKey");
  if (!token) {
    Logger.log("🚨 FEHLER: Script Property 'lxKey' fehlt!");
    SpreadsheetApp.getUi().alert("Fehler: API-Key 'lxKey' fehlt in den Script Properties!");
    return;
  }

  // Spaltenköpfe für Spalte H, I und J setzen
  mahSheet.getRange(1, 8).setValue("Mahnungs ID").setFontWeight("bold");
  mahSheet.getRange(1, 9).setValue("Status Mahnwesen").setFontWeight("bold");
  mahSheet.getRange(1, 10).setValue("Rechnungsstatus").setFontWeight("bold");

  // Liest Spalten A bis J ab Zeile 2 (10 Spalten)
  const data = mahSheet.getRange(2, 1, lastRow - 1, 10).getValues();

  Logger.log(`Starte Erstellung und Versand von 1. Mahnungen für ${data.length} Zeile(n)...`);

  let processedCount = 0;
  let skippedCount = 0;

  data.forEach((row, index) => {
    const rowIndex = index + 2;
    const companyName = row[0];
    const voucherNumber = row[1];
    const overdueDays = Number(row[4]) || 0;                        // Spalte E: Überfällig seit (Tage)
    const recipientEmail = row[5] ? row[5].toString().trim() : "";  // Spalte F: E-Mail
    const invoiceId = row[6] ? row[6].toString().trim() : "";       // Spalte G: Lexoffice Rechnungs-ID
    const existingDunningId = row[7] ? row[7].toString().trim() : "";// Spalte H: Mahnungs-ID
    const currentStatus = row[8] ? row[8].toString().trim() : "";   // Spalte I: Status Mahnwesen
    const invoiceStatus = row[9] ? row[9].toString().trim() : "";  // Spalte J: Rechnungsstatus

    // 1. DUPLETTEN-PRÜFUNG: Überspringen, falls bereits verarbeitet
    if (existingDunningId !== "" || currentStatus.toLowerCase().includes("versendet") || currentStatus.toLowerCase().includes("erstellt")) {
      skippedCount++;
      return;
    }

    // 2. RECHNUNGSSTATUS-PRÜFUNG: Bezahlt oder Storniert überspringen
    if (invoiceStatus === "Bezahlt" || invoiceStatus === "Storniert") {
      skippedCount++;
      return;
    }

    // 3. ÜBERFÄLLIGKEITS-PRÜFUNG: Nur verarbeiten wenn Spalte E >= 3 Tage her ist
    if (overdueDays < 3) {
      Logger.log(`⏳ [Zeile ${rowIndex} | ${voucherNumber}]: Erst seit ${overdueDays} Tag(en) überfällig (Spalte E). 1. Mahnung erst ab 3 Tagen.`);
      skippedCount++;
      return;
    }

    if (!invoiceId) {
      Logger.log(`⚠️ [Zeile ${rowIndex} | ${companyName}]: Keine Lexoffice ID in Spalte G vorhanden. Übersprungen.`);
      skippedCount++;
      return;
    }

    Logger.log(`🔄 Verarbeite 1. Mahnung für Zeile ${rowIndex}: ${companyName} (${voucherNumber} | ${overdueDays} Tage überfällig)...`);

    // 1. Ursprüngliche Rechnung von Lexoffice abrufen
    const invoiceDetails = getInvoiceDetailsForDunning(invoiceId, token);
    if (!invoiceDetails) {
      Logger.log(`🚨 [Zeile ${rowIndex}]: Rechnungsdetails für ID ${invoiceId} konnten nicht geladen werden.`);
      return;
    }

    // 2. Erstellung der Mahngebühr-Position (10 € ohne USt / Schadensersatz)
    const dunningFeeItem = {
      type: "custom",
      name: "Mahngebühr / Verzugsschaden",
      quantity: 1,
      unitName: "Pauschal",
      unitPrice: {
        currency: "EUR",
        netAmount: 10.00,
        grossAmount: 10.00,
        taxRatePercentage: 0
      },
      lineItemAmount: 10.00
    };

    // Positionen der Ursprungsrechnung um die Mahngebühr erweitern
    const updatedLineItems = Array.isArray(invoiceDetails.lineItems) ? [...invoiceDetails.lineItems] : [];
    updatedLineItems.push(dunningFeeItem);

    // Gesamtsummen um 10.00 € erhöhen
    const origTotalPrice = invoiceDetails.totalPrice || { currency: "EUR", totalNetAmount: 0, totalGrossAmount: 0, totalTaxAmount: 0 };
    const updatedTotalPrice = {
      currency: origTotalPrice.currency || "EUR",
      totalNetAmount: Number((origTotalPrice.totalNetAmount + 10.00).toFixed(2)),
      totalGrossAmount: Number((origTotalPrice.totalGrossAmount + 10.00).toFixed(2)),
      totalTaxAmount: Number((origTotalPrice.totalTaxAmount || 0).toFixed(2))
    };

    // 3. Mahnungs-Payload aufbauen
    const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss.SSSXXX");

    const dunningPayload = {
      voucherDate: todayStr,
      address: invoiceDetails.address,
      lineItems: updatedLineItems,
      totalPrice: updatedTotalPrice,
      taxConditions: invoiceDetails.taxConditions,
      shippingConditions: invoiceDetails.shippingConditions || { shippingType: "none" },
      title: "Mahnung",
      introduction: `Wir bitten Sie, die nachfolgend aufgelisteten Lieferungen/Leistungen zur Rechnung ${voucherNumber} inkl. der aufgelisteten Mahngebühr unverzüglich zu begleichen.`,
      remark: "Sollten Sie den offenen Betrag bereits beglichen haben, betrachten Sie dieses Schreiben als gegenstandslos."
    };

    // 4. Mahnung per POST in Lexoffice erstellen
    const dunningUrl = `https://api.lexware.io/v1/dunnings?precedingSalesVoucherId=${invoiceId}`;
    const options = {
      method: "post",
      contentType: "application/json",
      headers: {
        "Authorization": "Bearer " + token,
        "Accept": "application/json"
      },
      payload: JSON.stringify(dunningPayload),
      muteHttpExceptions: true
    };

    Utilities.sleep(600); // 600ms Delay für Rate-Limit Schutz
    const response = UrlFetchApp.fetch(dunningUrl, options);
    const resCode = response.getResponseCode();

    if (resCode === 201 || resCode === 200) {
      const result = JSON.parse(response.getContentText());
      const dunningId = result.id; // Lexoffice Mahnungs-ID
      
      // Mahnungs-ID in Spalte H eintragen
      mahSheet.getRange(rowIndex, 8).setValue(dunningId);
      Logger.log(`✅ [Zeile ${rowIndex}]: Mahnung in Lexoffice erfolgreich erstellt (Mahnungs-ID: ${dunningId})`);

      // 5. Mahnungs-PDF von Lexoffice herunterladen
      const pdfBlob = downloadDunningPdf(dunningId, voucherNumber, token);

      // 6. E-Mail versenden
      if (recipientEmail && recipientEmail !== "-") {
        const mailSubject = `1. Mahnung zur Rechnung ${voucherNumber} - ${companyName}`;
        const mailBody = `Guten Tag ${companyName},\n\n` +
          `leider konnten wir für die Rechnung ${voucherNumber} bisher keinen Zahlungseingang feststellen.\n` +
          `Im Anhang finden Sie die dazugehörige 1. Mahnung (inkl. 10,00 € Mahngebühr) mit der Bitte um zeitnahe Begleichung.\n\n` +
          `Sollten Sie den Betrag in den letzten Tagen bereits überwiesen haben, betrachten Sie diese E-Mail bitte als gegenstandslos.\n\n` +
          `Mit freundlichen Grüßen\nL8Street Team`;

        const mailOptions = {
          to: recipientEmail,
          cc: "info@l8street.com",
          subject: mailSubject,
          body: mailBody
        };

        if (pdfBlob) {
          mailOptions.attachments = [pdfBlob];
        }

        GmailApp.sendEmail(recipientEmail, mailSubject, mailBody, mailOptions);
        
        const todayFormatted = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy");
        const statusText = `1. Mahnung versendet am ${todayFormatted}`;
        
        // Status in Spalte I eintragen
        mahSheet.getRange(rowIndex, 9).setValue(statusText);
        Logger.log(`📧 [Zeile ${rowIndex}]: E-Mail erfolgreich an ${recipientEmail} (CC: info@l8street.com) versendet!`);

      } else {
        const todayFormatted = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy");
        mahSheet.getRange(rowIndex, 9).setValue(`Mahnung in Lexoffice erstellt am ${todayFormatted} (Keine E-Mail vorhanden)`);
        Logger.log(`⚠️ [Zeile ${rowIndex}]: Keine E-Mail-Adresse vorhanden. Mahnung wurde nur in Lexoffice erstellt.`);
      }

      processedCount++;

    } else {
      Logger.log(`🚨 [Zeile ${rowIndex} | ${voucherNumber}]: Fehler beim Erstellen der Mahnung (HTTP ${resCode}): ${response.getContentText()}`);
    }
  });

  Logger.log(`=== 1. MAHNUNGSERSTELLUNG UND VERSAND BEENDET ===`);
  Logger.log(`📊 Übersicht: ${processedCount} Mahnung(en) neu verarbeitet | ${skippedCount} Zeile(n) übersprungen.`);
}

/**
 * Hilfsfunktion: Lädt Rechnungsdetails von Lexoffice
 */
function getInvoiceDetailsForDunning(invoiceId, token) {
  try {
    const url = `https://api.lexware.io/v1/invoices/${invoiceId}`;
    const options = {
      method: "get",
      headers: {
        "Authorization": "Bearer " + token,
        "Accept": "application/json"
      },
      muteHttpExceptions: true
    };

    Utilities.sleep(600);
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() === 200) {
      return JSON.parse(response.getContentText());
    }
  } catch (e) {
    Logger.log(`Fehler beim Abrufen der Rechnung ID ${invoiceId}: ${e.message}`);
  }
  return null;
}

/**
 * Hilfsfunktion: Lädt das PDF-Dokument der Mahnung von Lexoffice herunter
 */
function downloadDunningPdf(dunningId, voucherNumber, token) {
  try {
    const url = `https://api.lexware.io/v1/dunnings/${dunningId}/file`;
    const options = {
      method: "get",
      headers: {
        "Authorization": "Bearer " + token,
        "Accept": "application/pdf"
      },
      muteHttpExceptions: true
    };

    Utilities.sleep(600);
    const response = UrlFetchApp.fetch(url, options);

    if (response.getResponseCode() === 200) {
      const blob = response.getBlob();
      blob.setName(`1_Mahnung_${voucherNumber}.pdf`);
      return blob;
    } else {
      Logger.log(`⚠️ PDF konnte nicht von Lexoffice geladen werden (HTTP ${response.getResponseCode()}).`);
    }
  } catch (e) {
    Logger.log(`Fehler beim Download des Mahnungs-PDFs: ${e.message}`);
  }
  return null;
}