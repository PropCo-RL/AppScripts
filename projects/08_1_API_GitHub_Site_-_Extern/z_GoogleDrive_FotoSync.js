/**
 * AUTOMATISCHER SYNC: Benennt Dateien um und speichert CDN-Links in Spalte V & AH
 */
function syncAllDriveMediaToSheet() {
  const TARGET_FOTOS_ID = "1L0CJ_p4IG3zxwG0MJkJPI43UOqyxJKYA";
  const TARGET_REVIEWS_ID = "1lnQY1WNja1py6liWCtT7wBECnezaLD6h";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("M") || ss.getSheets()[0];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // Liest Daten von Zeile 2 bis Spalte AH (Spalte 34)
  const range = sheet.getRange(2, 1, lastRow - 1, 34);
  const dataValues = range.getValues();

  // 1. Fotos synchronisieren (Spalte V = Index 22)
  processFolderMediaToColumn(sheet, dataValues, TARGET_FOTOS_ID, 22, "img");

  // 2. Bewertungs-Screenshots synchronisieren (Spalte AH = Index 34)
  processFolderMediaToColumn(sheet, dataValues, TARGET_REVIEWS_ID, 34, "review");
}

function processFolderMediaToColumn(sheet, dataValues, parentFolderId, targetColumnIndex, filePrefix) {
  const parentFolder = DriveApp.getFolderById(parentFolderId);
  const subFolders = parentFolder.getFolders();

  while (subFolders.hasNext()) {
    const folder = subFolders.next();
    const folderName = folder.getName().toLowerCase().trim(); // z.B. "k42a6"

    for (let i = 0; i < dataValues.length; i++) {
      const aptLink = String(dataValues[i][6] || "").toLowerCase().trim(); // Spalte G (Index 6)
      const currentCellVal = String(dataValues[i][targetColumnIndex - 1] || "").trim();

      if (aptLink.includes("/a/" + folderName)) {
        const rowIndex = i + 2;
        const files = folder.getFiles();
        const mediaUrls = [];
        let indexCounter = 1;

        while (files.hasNext()) {
          const file = files.next();
          const mime = file.getMimeType();

          // Bilder und PDFs zulassen
          if (mime.indexOf("image/") !== -1 || mime.indexOf("pdf") !== -1) {
            const ext = file.getName().split('.').pop() || "png";
            const newName = `${folderName}_${filePrefix}_${indexCounter}.${ext}`;

            if (file.getName() !== newName) {
              file.setName(newName);
            }

            const directUrl = `https://lh3.googleusercontent.com/d/${file.getId()}`;
            mediaUrls.push(directUrl);
            indexCounter++;
          }
        }

        // Nur schreiben, wenn echte Dateien im Ordner vorhanden sind
        if (mediaUrls.length > 0) {
          const jsonVal = JSON.stringify(mediaUrls);
          if (currentCellVal !== jsonVal) {
            sheet.getRange(rowIndex, targetColumnIndex).setValue(jsonVal);
            Logger.log(`[Sync] Spalte ${targetColumnIndex} in Zeile ${rowIndex} (${folderName}) aktualisiert.`);
          }
        }
        break;
      }
    }
  }
}