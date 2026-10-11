/**
 * SKRIPT 1: Rechnungsstatus in Spalte J aktualisieren
 * Prüft bei Lexoffice, ob die Rechnung im Tab 'mah' bezahlt oder storniert wurde.
 */
function updateMahInvoiceStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mahSheet = ss.getSheetByName("mah");
  if (!mahSheet) return;

  const lastRow = mahSheet.getLastRow();
  if (lastRow < 2) return;

  const token = PropertiesService.getScriptProperties().getProperty("lxKey");
  if (!token) return;

  // Kopfzeile in Spalte J setzen
  mahSheet.getRange(1, 10).setValue("Rechnungsstatus").setFontWeight("bold");

  // Liest Daten bis Spalte J (10 Spalten)
  const values = mahSheet.getRange(2, 1, lastRow - 1, 10).getValues();

  Logger.log("Prüfe Rechnungsstatus für Tab 'mah' in Lexoffice...");

  values.forEach((row, index) => {
    const rowIndex = index + 2;
    const invoiceId = row[6] ? row[6].toString().trim() : "";      // Spalte G: Lexoffice ID
    const currentStatus = row[9] ? row[9].toString().trim() : "";  // Spalte J: Rechnungsstatus

    if (!invoiceId || currentStatus === "Bezahlt" || currentStatus === "Storniert") return;

    const url = `https://api.lexware.io/v1/invoices/${invoiceId}`;
    const response = safeApiFetch(url, token);

    if (response && response.getResponseCode() === 200) {
      const invoiceData = JSON.parse(response.getContentText());
      const rawStatus = (invoiceData.voucherStatus || "").toLowerCase();

      let statusDE = "Offen";
      if (rawStatus === "paid") statusDE = "Bezahlt";
      else if (rawStatus === "voided") statusDE = "Storniert";

      mahSheet.getRange(rowIndex, 10).setValue(statusDE); // In Spalte J schreiben
      Logger.log(`[Zeile ${rowIndex}] ${row[1]}: Status -> ${statusDE}`);
    }
  });
}