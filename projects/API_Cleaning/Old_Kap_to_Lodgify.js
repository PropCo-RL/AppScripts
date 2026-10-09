function syncVerschiebungenToLodgify() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gCalSheet = ss.getSheetByName("gCal");
  const cleanerSheet = ss.getSheetByName("Cleaners");
  const loKey = PropertiesService.getScriptProperties().getProperty('loKey'); 

  if (!gCalSheet || !cleanerSheet) return;

  const data = gCalSheet.getDataRange().getValues();
  const cleanerData = cleanerSheet.getDataRange().getValues();
  
  const apartmentMap = {};
  cleanerData.slice(1).forEach(row => {
    const aptName = String(row[1]).trim();
    if (aptName) {
      apartmentMap[aptName] = {
        propertyId: row[0],
        roomTypeId: row[16]
      };
    }
  });

  const tz = ss.getSpreadsheetTimeZone();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const apartmentName = String(row[1]).trim();
    const newDateRaw = row[0];
    const originalDateRaw = row[11];

    // ✅ Dein neuer Check (korrekt)
    if (!(newDateRaw instanceof Date) || !(originalDateRaw instanceof Date)) continue;
    if (!apartmentMap[apartmentName]) continue;
    if (newDateRaw <= originalDateRaw) continue;

    // --- dein ursprünglicher Code unverändert ---

    const startDateStr = Utilities.formatDate(originalDateRaw, tz, "yyyy-MM-dd");
    
    let endDate = new Date(newDateRaw.getTime());
    endDate.setDate(endDate.getDate() - 1);
    const endDateStr = Utilities.formatDate(endDate, tz, "yyyy-MM-dd");

    const ids = apartmentMap[apartmentName];
    
    const payload = {
      "available": 0,
      "period_start": startDateStr + "T00:00:00.000+01:00",
      "period_end": endDateStr + "T00:00:00.000+01:00"
    };

    const url = `https://api.lodgify.com/v1/availability/${ids.propertyId}/${ids.roomTypeId}/set`;
    
    const options = {
      "method": "post",
      "headers": {
        "X-ApiKey": loKey,
        "accept": "application/json",
        "content-type": "application/json"
      },
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };

    try {
      const response = UrlFetchApp.fetch(url, options);
      if (response.getResponseCode() === 200 || response.getResponseCode() === 204) {
        console.log(`Erfolg: ${apartmentName} blockiert von ${startDateStr} bis ${endDateStr}`);
      } else {
        console.error(`Fehler bei ${apartmentName}: ${response.getContentText()}`);
      }
    } catch (e) {
      console.error(`API Error bei ${apartmentName}: ${e.message}`);
    }
  }
}