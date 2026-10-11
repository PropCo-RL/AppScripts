/**
 * SKRIPT FÜR DIE 2. MAHNUNG:
 * Erstellt und versendet die 2. Mahnung (20 € Mahngebühr) nur wenn:
 * - Rechnungsstatus in Spalte J = 'Offen' ist
 * - 1. Mahnung in Spalte I >= 3 Tage her ist
 * - Spalte K noch leer ist
 */
function m2_sendSecondDunning() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mahSheet = ss.getSheetByName("mah");

  if (!mahSheet) {
    Logger.log("🚨 Das Tabellenblatt 'mah' wurde nicht gefunden!");
    SpreadsheetApp.getUi().alert("Fehler: Tabellenblatt 'mah' existiert nicht.");
    return;
  }

  const lastRow = mahSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("ℹ️ Keine Daten im Tab 'mah' vorhanden.");
    return;
  }

  const token = PropertiesService.getScriptProperties().getProperty("lxKey");
  if (!token) {
    Logger.log("🚨 FEHLER: Script Property 'lxKey' fehlt!");
    SpreadsheetApp.getUi().alert("Fehler: API-Key 'lxKey' fehlt in den Script Properties!");
    return;
  }

  // Kopfzeile Spalte K (Spalte 11) setzen
  mahSheet.getRange(1, 11).setValue("2. Mahnung").setFontWeight("bold");

  // Liest bis Spalte K (11 Spalten) ab Zeile 2
  const data = mahSheet.getRange(2, 1, lastRow - 1, 11).getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  Logger.log("Starte Prüfung für 2. Mahnungen...");

  data.forEach((row, index) => {
    const rowIndex = index + 2;
    const companyName = row[0];
    const voucherNumber = row[1];
    const recipientEmail = row[5] ? row[5].toString().trim() : ""; // Spalte F: E-Mail
    const invoiceId = row[6] ? row[6].toString().trim() : "";      // Spalte G: Lexoffice ID
    const dunning1Status = row[8] ? row[8].toString().trim() : "";// Spalte I: 1. Mahnung
    const invoiceStatus = row[9] ? row[9].toString().trim() : ""; // Spalte J: Rechnungsstatus
    const dunning2Status = row[10] ? row[10].toString().trim() : "";// Spalte K: 2. Mahnung

    // 1. Abbrechen wenn Bezahlt oder Storniert
    if (invoiceStatus === "Bezahlt" || invoiceStatus === "Storniert") {
      return;
    }

    // 2. Muss bereits eine 1. Mahnung haben
    if (!dunning1Status.includes("1. Mahnung versendet am")) {
      return;
    }

    // 3. 2. Mahnung darf in Spalte K noch NICHT existieren
    if (dunning2Status !== "") {
      return;
    }

    // 4. Datum der 1. Mahnung prüfen (muss >= 3 Tage her sein)
    const lastDate = m2_parseDateFromStatus(dunning1Status);
    if (!lastDate || m2_getDaysBetween(lastDate, today) < 3) {
      Logger.log(`⏳ Zeile ${rowIndex} [${voucherNumber}]: 1. Mahnung ist noch keine 3 Tage her.`);
      return;
    }

    Logger.log(`⚡ Erstelle 2. Mahnung für Zeile ${rowIndex}: ${companyName} (${voucherNumber})...`);

    // 5. 2. Mahnung ausführen & senden
    m2_executeDunningCreation({
      mahSheet: mahSheet,
      rowIndex: rowIndex,
      invoiceId: invoiceId,
      voucherNumber: voucherNumber,
      companyName: companyName,
      recipientEmail: recipientEmail,
      token: token,
      feeAmount: 20.00 // 2x 10 € Mahngebühr
    });
  });

  Logger.log("=== PRÜFUNG UND VERSAND DER 2. MAHNUNG BEENDET ===");
}


/**
 * HILFSFUNKTION: Legt die 2. Mahnung bei Lexoffice an, lädt das PDF, versendet die Mail
 * und speichert den Status in Spalte K.
 */
function m2_executeDunningCreation(p) {
  const invoiceDetails = m2_getInvoiceDetailsForDunning(p.invoiceId, p.token);
  if (!invoiceDetails) {
    Logger.log(`🚨 Zeile ${p.rowIndex}: Rechnungsdetails konnten nicht geladen werden.`);
    return;
  }

  // Position für 2. Mahngebühr (20 € ohne USt)
  const dunningFeeItem = {
    type: "custom",
    name: "Mahngebühr (2. Mahnung)",
    quantity: 1,
    unitName: "Pauschal",
    unitPrice: { currency: "EUR", netAmount: p.feeAmount, grossAmount: p.feeAmount, taxRatePercentage: 0 },
    lineItemAmount: p.feeAmount
  };

  const updatedLineItems = Array.isArray(invoiceDetails.lineItems) ? [...invoiceDetails.lineItems] : [];
  updatedLineItems.push(dunningFeeItem);

  const origTotalPrice = invoiceDetails.totalPrice || { currency: "EUR", totalNetAmount: 0, totalGrossAmount: 0 };
  const updatedTotalPrice = {
    currency: origTotalPrice.currency || "EUR",
    totalNetAmount: Number((origTotalPrice.totalNetAmount + p.feeAmount).toFixed(2)),
    totalGrossAmount: Number((origTotalPrice.totalGrossAmount + p.feeAmount).toFixed(2)),
    totalTaxAmount: Number((origTotalPrice.totalTaxAmount || 0).toFixed(2))
  };

  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss.SSSXXX");

  const dunningPayload = {
    voucherDate: todayStr,
    address: invoiceDetails.address,
    lineItems: updatedLineItems,
    totalPrice: updatedTotalPrice,
    taxConditions: invoiceDetails.taxConditions,
    shippingConditions: invoiceDetails.shippingConditions || { shippingType: "none" },
    title: "Mahnung",
    introduction: `Wir bitten Sie, die nachfolgend aufgelisteten Lieferungen/Leistungen zur Rechnung ${p.voucherNumber} inkl. der Mahngebühr unverzüglich zu begleichen. Dies ist die 2. Mahnung.`,
    remark: "Sollten Sie den offenen Betrag bereits beglichen haben, betrachten Sie dieses Schreiben als gegenstandslos."
  };

  const url = `https://api.lexware.io/v1/dunnings?precedingSalesVoucherId=${p.invoiceId}`;
  const options = {
    method: "post",
    contentType: "application/json",
    headers: { "Authorization": "Bearer " + p.token, "Accept": "application/json" },
    payload: JSON.stringify(dunningPayload),
    muteHttpExceptions: true
  };

  Utilities.sleep(600);
  const response = UrlFetchApp.fetch(url, options);

  if (response.getResponseCode() === 201 || response.getResponseCode() === 200) {
    const dunningId = JSON.parse(response.getContentText()).id;
    const pdfBlob = m2_downloadDunningPdf(dunningId, p.voucherNumber, p.token);

    if (p.recipientEmail && p.recipientEmail !== "-") {
      const mailSubject = `2. Mahnung zur Rechnung ${p.voucherNumber} - ${p.companyName}`;
      const mailBody = `Guten Tag ${p.companyName},\n\n` +
        `leider konnten wir für die Rechnung ${p.voucherNumber} bisher keinen Zahlungseingang feststellen.\n` +
        `Im Anhang finden Sie die dazugehörige 2. Mahnung (inkl. ${p.feeAmount.toFixed(2)} € Mahngebühr) mit der dringenden Bitte um Ausgleich.\n\n` +
        `Sollten Sie den Betrag in den letzten Tagen bereits überwiesen haben, betrachten Sie diese E-Mail bitte als gegenstandslos.\n\n` +
        `Mit freundlichen Grüßen\nL8Street Team`;

      const mailOptions = { to: p.recipientEmail, cc: "info@l8street.com", subject: mailSubject, body: mailBody };
      if (pdfBlob) mailOptions.attachments = [pdfBlob];

      GmailApp.sendEmail(p.recipientEmail, mailSubject, mailBody, mailOptions);

      const todayFormatted = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy");
      
      // Status & ID in Spalte K (Spalte 11) eintragen
      p.mahSheet.getRange(p.rowIndex, 11).setValue(`2. Mahnung versendet am ${todayFormatted} (ID: ${dunningId})`);
      Logger.log(`📧 [Zeile ${p.rowIndex}]: 2. Mahnung per E-Mail versendet und in Spalte K eingetragen.`);
    } else {
      const todayFormatted = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy");
      p.mahSheet.getRange(p.rowIndex, 11).setValue(`2. Mahnung in Lexoffice erstellt am ${todayFormatted} (Keine E-Mail) (ID: ${dunningId})`);
      Logger.log(`⚠️ [Zeile ${p.rowIndex}]: Keine E-Mail vorhanden. 2. Mahnung nur in Lexoffice erstellt.`);
    }
  } else {
    Logger.log(`🚨 [Zeile ${p.rowIndex}]: Fehler beim Erstellen der 2. Mahnung (HTTP ${response.getResponseCode()}): ${response.getContentText()}`);
  }
}


// ============================================================================
// HILFSFUNKTIONEN FÜR MAHNUNG 2 (alle mit m2_ Präfix)
// ============================================================================

function m2_parseDateFromStatus(statusStr) {
  if (!statusStr) return null;
  const match = statusStr.match(/(\d{2})\.(\d{2})\.(\d{4})/);
  if (match) {
    const day = parseInt(match[1], 10);
    const month = parseInt(match[2], 10) - 1;
    const year = parseInt(match[3], 10);
    return new Date(year, month, day);
  }
  return null;
}

function m2_getDaysBetween(date1, date2) {
  const diffTime = Math.abs(date2.getTime() - date1.getTime());
  return Math.floor(diffTime / (1000 * 60 * 60 * 24));
}

function m2_getInvoiceDetailsForDunning(invoiceId, token) {
  try {
    const url = `https://api.lexware.io/v1/invoices/${invoiceId}`;
    const options = {
      method: "get",
      headers: { "Authorization": "Bearer " + token, "Accept": "application/json" },
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

function m2_downloadDunningPdf(dunningId, voucherNumber, token) {
  try {
    const url = `https://api.lexware.io/v1/dunnings/${dunningId}/file`;
    const options = {
      method: "get",
      headers: { "Authorization": "Bearer " + token, "Accept": "application/pdf" },
      muteHttpExceptions: true
    };
    Utilities.sleep(600);
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() === 200) {
      const blob = response.getBlob();
      blob.setName(`2_Mahnung_${voucherNumber}.pdf`);
      return blob;
    }
  } catch (e) {
    Logger.log(`Fehler beim Download des PDFs für Mahnung ID ${dunningId}: ${e.message}`);
  }
  return null;
}