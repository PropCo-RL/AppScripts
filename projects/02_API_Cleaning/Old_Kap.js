function enforceCapacityAndRulesToGCal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gCalSheet = ss.getSheetByName("gCal");
  const cleanerSheet = ss.getSheetByName("Cleaners");

  const data = gCalSheet.getDataRange().getValues();
  if (data.length < 2) return;

  const rows = data.slice(1);
  const cleanerData = cleanerSheet.getDataRange().getValues().slice(1);

  const cleanerSettings = {};
  cleanerData.forEach(r => {
    const email = String(r[4]).trim().toLowerCase();
    if (email) {
      cleanerSettings[email] = {
        maxPerDay: (r[13] !== "" && !isNaN(r[13])) ? parseInt(r[13]) : 2,
        noSunday: String(r[14] || "").toLowerCase().includes("ohne sonntag"),
        leadTime: (r[18] !== "" && !isNaN(r[18])) ? parseInt(r[18]) : 4
      };
    }
  });

  const isWorkday = d => d.getDay() !== 0 && d.getDay() !== 6;

  const addWorkdays = (startDate, days) => {
    let d = new Date(startDate.getTime());
    let added = 0;
    while (added < days) {
      d.setDate(d.getDate() + 1);
      if (isWorkday(d)) added++;
    }
    return d;
  };

  const capacityTracker = {};
  const today = new Date();
  today.setHours(0,0,0,0);

  // 🔥 SORTIERUNG: Anreise-Jobs zuerst
  const sortedRows = rows
    .map((row, i) => ({ row, i }))
    .sort((a, b) => {
      const aCheckIn = a.row[8] instanceof Date;
      const bCheckIn = b.row[8] instanceof Date;
      return (bCheckIn ? 1 : 0) - (aCheckIn ? 1 : 0);
    });

  const updatesRaw = sortedRows.map(({ row, i }) => {

    if (!row[0]) return [row[0], "", row[11] || ""];

    let date = new Date(row[0]);
    date.setHours(0,0,0,0);

    const originalDate = new Date(date.getTime());
    const status = row[2];
    const email = String(row[3]).trim().toLowerCase();

    const nextCheckIn = row[8] instanceof Date ? new Date(row[8]) : null;
    if (nextCheckIn) nextCheckIn.setHours(0,0,0,0);

    const settings = cleanerSettings[email] || {
      maxPerDay: 2,
      noSunday: false,
      leadTime: 4
    };

    let changeNote = "OK";
    let dateChanged = false;

    // 1. Lead-Time prüfen
    if (status === "NUR API") {
      const minDate = addWorkdays(today, settings.leadTime);
      if (date < minDate) {
        date = new Date(minDate.getTime());
        dateChanged = true;
        changeNote = "Verschoben: " + settings.leadTime + "-Tage-Regel";
      }
    }

    // 2. Kapazität + Sonntag + Anreise-Logik
    let foundSlot = false;
    let securityCounter = 0;

    while (!foundSlot && securityCounter < 100) {
      const isSunday = date.getDay() === 0;
      const dateKey = Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd") + "_" + email;
      const currentCount = capacityTracker[dateKey] || 0;

      if ((isSunday && settings.noSunday) || (currentCount >= settings.maxPerDay)) {

        let nextDay = new Date(date.getTime());
        nextDay.setDate(nextDay.getDate() + 1);

        if (nextCheckIn && nextDay > nextCheckIn) {
          changeNote = "⚠️ LIMIT: Anreise am " +
            Utilities.formatDate(nextCheckIn, ss.getSpreadsheetTimeZone(), "dd.MM");

          capacityTracker[dateKey] = currentCount + 1;
          foundSlot = true;

        } else {
          date = nextDay;
          dateChanged = true;
          changeNote = (isSunday && settings.noSunday)
            ? "Verschoben: Sonntagsverbot"
            : "Verschoben: Kapazität voll";
        }

      } else {
        capacityTracker[dateKey] = currentCount + 1;
        foundSlot = true;
      }

      securityCounter++;
    }

    if (status === "MATCH" && dateChanged) {
      changeNote = "⚠️ BESTAND ÄNDERN: " + changeNote;
    }

    return [date, changeNote, originalDate];
  });

  // 🔁 ZURÜCK IN ORIGINAL-REIHENFOLGE
  const updates = new Array(rows.length);
  updatesRaw.forEach((res, idx) => {
    const originalIndex = sortedRows[idx].i;
    updates[originalIndex] = res;
  });

  const numRows = updates.length;
  if (numRows === 0) return;

  gCalSheet.getRange(2, 1, numRows, 1).setValues(updates.map(r => [r[0]]));
  gCalSheet.getRange(2, 11, numRows, 1).setValues(updates.map(r => [r[1]]));
  gCalSheet.getRange(2, 12, numRows, 1).setValues(updates.map(r => [r[2]]));

  gCalSheet.getRange(2, 1, numRows, 1).setNumberFormat("dd.MM.yyyy");
  gCalSheet.getRange(2, 12, numRows, 1).setNumberFormat("dd.MM.yyyy");
}