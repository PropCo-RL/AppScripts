function sendAllCleaningSchedulesWhatsApp() {
  const p = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gCalSheet = ss.getSheetByName("gCal");
  
  if (!gCalSheet) {
    console.error("Blatt 'gCal' nicht gefunden!");
    return;
  }
  
  const data = gCalSheet.getDataRange().getValues().slice(1);
  const today = new Date();
  today.setHours(0,0,0,0);
  const nextWeek = new Date(today.getTime() + 14 * 24 * 60 * 60 * 1000);
  
  const cleanerGroups = {};
  const days = ["So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"];
  const months = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"];

  data.forEach(row => {
    const jobDate = new Date(row[0]);
    if (isNaN(jobDate.getTime())) return;

    const apartment = row[1];
    const email = String(row[3]).trim().toLowerCase();
    const phone = String(row[9] || "").replace(/\D/g, '');

    if (jobDate >= today && jobDate <= nextWeek && email && phone && email !== "nicht gefunden") {
      if (!cleanerGroups[email]) {
        cleanerGroups[email] = { phone: phone, jobs: [] };
      }
      const dayName = days[jobDate.getDay()];
      const monthName = months[jobDate.getMonth()];
      const dayOfMonth = jobDate.getDate();
      const jobText = `${dayOfMonth}. ${monthName} (${dayName}): ${apartment}`;
      
      if (!cleanerGroups[email].jobs.includes(jobText)) {
        cleanerGroups[email].jobs.push(jobText);
      }
    }
  });

  const apiKey = p.getProperty('D360-KEY');
  // HIER GELÖSCHT: waSheet Definition

  for (let email in cleanerGroups) {
    const cleaner = cleanerGroups[email];
    let jobs = cleaner.jobs;

    if (jobs.length === 0) continue;

    jobs.sort();
    jobs = jobs.slice(0, 7);
    while (jobs.length < 7) {
      jobs.push(" ");
    }

    const payload = {
      "messaging_product": "whatsapp",
      "recipient_type": "individual",
      "to": cleaner.phone,
      "type": "template",
      "template": {
        "name": "utility_cleaning",
        "language": { "code": "de" },
        "components": [{
          "type": "body",
          "parameters": jobs.map(j => ({ "type": "text", "text": j }))
        }]
      }
    };

    const options = {
      "method": "post",
      "headers": {
        "D360-API-KEY": apiKey,
        "Content-Type": "application/json"
      },
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };

    try {
      const response = UrlFetchApp.fetch("https://waba-v2.360dialog.io/messages", options);
      const resCode = response.getResponseCode();

      if (resCode == 200 || resCode == 201) {
        console.log(`✅ Gesendet an ${email} (${cleaner.phone})`);
        // HIER GELÖSCHT: waSheet.appendRow
      } else {
        console.error(`❌ Fehler bei ${email}: ` + response.getContentText());
        // HIER GELÖSCHT: waSheet.appendRow
      }
    } catch (e) {
      console.error(`❌ Kritischer Fehler bei ${email}: ` + e.message);
    }
    
    Utilities.sleep(500);
  }
}