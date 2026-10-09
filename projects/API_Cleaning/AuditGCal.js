function generateGCalAudit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sch = ss.getSheetByName("Import").getDataRange().getValues().slice(1);
  const cal = ss.getSheetByName("import gCal").getDataRange().getValues().slice(1);
  const cleanerSheet = ss.getSheetByName("Cleaners");

  // Definiert "Heute" um 00:00 Uhr
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  const cleanerData = cleanerSheet.getDataRange().getValues().slice(1);
  const cleanerMap = {};
  
  // 1. Cleaner-Daten
  cleanerData.forEach(row => {
    const aptName = String(row[1]).trim().toLowerCase(); 
    const email = row[4]; 
    const phone = row[12];
    if (aptName) {
      cleanerMap[aptName] = { email: email, phone: phone };
    }
  });

  const report = [], keys = {};
  const getK = (d, n) => {
    if (!d || isNaN(new Date(d).getTime())) return null;
    return Utilities.formatDate(new Date(d), "GMT+1", "dd.MM.yyyy") + "_" + String(n).trim().toLowerCase();
  };

  // 2. Kalender-Map aufbauen
  const calMap = {};
  cal.forEach(r => { 
    if (new Date(r[0]) < today) return; 
    const k = getK(r[0], r[1]);
    if (k) {
      calMap[k] = {
        id: r[3],
        email: String(r[5] || "").trim().toLowerCase()
      };
    }
  });

  // 3. Import verarbeiten
  sch.forEach(r => {
    if (new Date(r[0]) < today) return; 
    
    const aptName = String(r[1]).trim();
    const k = getK(r[0], aptName);
    if (!k) return;
    keys[k] = true;
    
    const abreiseDaten = r[4] || ""; 
    const anreiseDaten = r[7] instanceof Date ? r[7] : null;
    const cleanerInfo = cleanerMap[aptName.toLowerCase()] || { email: "nicht gefunden", phone: "" };
    const targetEmail = String(cleanerInfo.email || "").trim().toLowerCase();
    
    const eventData = calMap[k];
    const eventId = eventData ? eventData.id : ""; 
    const calEmail = eventData ? eventData.email : "";

    let emailMatch = "";
    if (eventData) {
      emailMatch = (calEmail === targetEmail) ? "MATCH" : "Kein Match";
    }

    report.push([
      r[0], aptName, eventData ? "MATCH" : "NUR API", 
      cleanerInfo.email, eventId, emailMatch, calEmail,
      abreiseDaten, anreiseDaten, cleanerInfo.phone
    ]);
  });

  // 4. Nur im Kalender vorhandene Termine
  cal.forEach(r => {
    if (new Date(r[0]) < today) return; 
    
    const aptName = String(r[1]).trim();
    const k = getK(r[0], aptName);
    if (k && !keys[k]) {
      const cleanerInfo = cleanerMap[aptName.toLowerCase()] || { email: "nicht gefunden", phone: "" };
      const targetEmail = String(cleanerInfo.email || "").trim().toLowerCase();
      const eventId = r[3]; 
      const calEmail = String(r[5] || "").trim().toLowerCase();
      const emailMatch = (calEmail === targetEmail) ? "MATCH" : "Kein Match";

      report.push([r[0], aptName, "NUR KALENDER", cleanerInfo.email, eventId, emailMatch, calEmail, "", "", cleanerInfo.phone]);
    }
  });

  // --- NEUE SORTIERUNG & LEERZEILEN LOGIK ---
  
  // Sortieren nach Spalte D (Cleaner Email) und dann nach Datum
  report.sort((a, b) => {
    const emailA = String(a[3] || "").toLowerCase();
    const emailB = String(b[3] || "").toLowerCase();
    if (emailA !== emailB) return emailA.localeCompare(emailB);
    return new Date(a[0]) - new Date(b[0]);
  });

  const finalData = [];
  // Erste Zeile (Start Row 2) leer lassen
  finalData.push(new Array(10).fill(""));

  // Daten mit Leerzeilen bei Cleaner-Wechsel aufbereiten
  if (report.length > 0) {
    for (let i = 0; i < report.length; i++) {
      finalData.push(report[i]);
      // Wenn der nächste Eintrag ein anderer Cleaner ist -> Leerzeile einfügen
      if (i < report.length - 1) {
        const currentCleaner = String(report[i][3] || "").toLowerCase();
        const nextCleaner = String(report[i+1][3] || "").toLowerCase();
        if (currentCleaner !== nextCleaner) {
          finalData.push(new Array(10).fill(""));
        }
      }
    }
  }

  // --- AUSGABE IN DAS SHEET ---
  const s = ss.getSheetByName("gCal");
  s.clear();
  const headers = [["Datum", "Apartment", "Status", "Cleaner Email", "Event ID", "Email Match", "gCal Email", "Abreise Daten", "Anreise Daten", "Cleaner Phone"]];
  s.getRange(1, 1, 1, 10).setValues(headers).setFontWeight("bold");
  
  if (finalData.length > 0) {
    s.getRange(2, 1, finalData.length, 10).setValues(finalData);
    s.getRange(2, 1, finalData.length, 1).setNumberFormat("dd.MM.yyyy");
  }
 
  // Schedule Update (unverändert)
  const actionSheet = ss.getSheetByName("Schedule"); 
  const nonMatches = report.filter(row => row[2] !== "MATCH").map(row => row.slice(0, 5));
  const lastRowAction = actionSheet.getLastRow();
  if (lastRowAction >= 7) {
    actionSheet.getRange(7, 1, lastRowAction - 6, 5).clearContent();
  }
  if (nonMatches.length > 0) {
    actionSheet.getRange(7, 1, nonMatches.length, 5).setValues(nonMatches);
    actionSheet.getRange(7, 1, nonMatches.length, 1).setNumberFormat("dd.MM.yyyy");
  }
}