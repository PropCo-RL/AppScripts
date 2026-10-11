function processInvoicesAndPreFilter() {
  Logger.log("=== START: GMAIL DOWNLOAD & VORFILTERUNG ===");

  // KONFIGURATION
  const SEARCH_LABEL = "support@-rechnung@";
  const PROCESSED_LABEL = "support@-rechnung@/weitergeleitetAnLx";
  const DRIVE_FOLDER_ID = "1j8xPlrOZTFP205HG9EPpGxiQYV-UI6rm";
  const SHEET_NAME = "API Mail Eingang";
  
  const MY_DOMAIN = "l8street.com";
  const MIN_FILE_SIZE_KB = 20; // Anhänge unter 20 KB ignorieren

  // 1. Folgelabel abrufen oder erstellen
  let processedLabelObj = GmailApp.getUserLabelByName(PROCESSED_LABEL);
  if (!processedLabelObj) {
    processedLabelObj = GmailApp.createLabel(PROCESSED_LABEL);
    Logger.log("Neues Folgelabel '" + PROCESSED_LABEL + "' wurde erstellt.");
  }

  // 2. Drive & Sheet laden
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Datum / Zeit", "Absender", "Dateiname", "Link zur Datei"]);
  }

  // 3. Suche nach E-Mails im Rechnungs-Label
  const query = 'label:' + SEARCH_LABEL + ' -label:"' + PROCESSED_LABEL + '" has:attachment';
  const threads = GmailApp.search(query);

  let savedCount = 0;
  let skippedCount = 0;

  // 4. Threads verarbeiten
  for (let i = 0; i < threads.length; i++) {
    const messages = threads[i].getMessages();

    for (let j = 0; j < messages.length; j++) {
      const message = messages[j];
      const sender = message.getFrom().toLowerCase();
      const subject = message.getSubject().toLowerCase();
      const toRecipient = message.getTo().toLowerCase();
      const ccRecipient = message.getCc().toLowerCase();

      // SCHUTZ 1: Eigene gesendete Mails überspringen
      if (sender.includes(MY_DOMAIN)) {
        Logger.log(`[AUSSORTIERT] Eigene gesendete Mail: ${message.getFrom()}`);
        skippedCount++;
        continue;
      }

      // SCHUTZ 2: System- & Report-Mails (z. B. DMARC Reports an info@) ignorieren
      const isSystemReport = sender.includes("dmarc") || sender.includes("noreply") || subject.includes("report domain");
      const isInfoOnly = (toRecipient.includes("info@") || ccRecipient.includes("info@")) && !toRecipient.includes("rechnung@");
      
      if (isSystemReport || isInfoOnly) {
        Logger.log(`[AUSSORTIERT] System-/Info-Mail übersprungen: ${subject} (${message.getFrom()})`);
        skippedCount++;
        continue;
      }

      const attachments = message.getAttachments();

      for (let k = 0; k < attachments.length; k++) {
        const attachment = attachments[k];
        const fileName = attachment.getName().toLowerCase();
        const mimeType = attachment.getContentType().toLowerCase();
        const fileSizeKB = attachment.getSize() / 1024;

        // SCHUTZ 3: Mindestgröße prüfen
        if (fileSizeKB < MIN_FILE_SIZE_KB) {
          Logger.log(`[AUSSORTIERT] Anhang zu klein (${fileSizeKB.toFixed(2)} KB): '${attachment.getName()}'`);
          skippedCount++;
          continue; 
        }

        // SCHUTZ 4: Dateityp-Prüfung (Positivliste: Nur Dokumente / PDFs erlauben)
        const isPdf = mimeType.includes("pdf") || fileName.endsWith(".pdf");
        const isDoc = fileName.endsWith(".docx") || fileName.endsWith(".doc") || fileName.endsWith(".xlsx");
        
        // Expliziter Ausschluss bekannter Stördateien
        const isImage = mimeType.startsWith("image/") || fileName.endsWith(".png") || fileName.endsWith(".jpg") || fileName.endsWith(".jpeg");
        const isZip = fileName.endsWith(".zip") || mimeType.includes("zip");
        const isCalendar = fileName.endsWith(".ics") || mimeType.includes("calendar");
        const isInstructionSheet = fileName.includes("anleitung") || fileName.includes("merkblatt") || fileName.includes("agb");
        const isOwnInvoicePattern = fileName.includes("rechnung_re-") || fileName.includes("mahnung_re-");

        // Wenn es KEIN Dokument/PDF ist ODER in die Ausschlussliste fällt -> Überspringen!
        if ((!isPdf && !isDoc) || isImage || isZip || isCalendar || isInstructionSheet || isOwnInvoicePattern) {
          Logger.log(`[AUSSORTIERT] Unerwünschter Anhang/Format: '${attachment.getName()}'`);
          skippedCount++;
          continue; 
        }

        // In Google Drive speichern
        const file = folder.createFile(attachment);
        
        // In Google Sheet eintragen
        sheet.appendRow([
          message.getDate(),
          message.getFrom(),
          file.getName(),
          file.getUrl()
        ]);

        savedCount++;
      }
    }
    // Label zuweisen, damit der Thread nicht erneut verarbeitet wird
    threads[i].addLabel(processedLabelObj);
  }

  Logger.log(`=== ENDE: ${savedCount} Dateien gespeichert, ${skippedCount} Mails/Anhänge aussortiert ===`);
}