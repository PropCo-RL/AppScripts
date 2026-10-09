function importCalendarToGCal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("import gCal");
  
  const calendarId = 'support@l8street.com'; 
  const calendar = CalendarApp.getCalendarById(calendarId);
  
  const now = new Date();
  const startDate = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0);
  const endDate = new Date(now.getTime() + (16 * 24 * 60 * 60 * 1000));
  
  const events = calendar.getEvents(startDate, endDate);
  const eventData = [];

  // Liste der zu ignorierenden E-Mails (in Kleinbuchstaben für sicheren Vergleich)
  const ignoreList = ["bastatas.jadine@gmail.com", "info@l8street.com", "support@l8street.com"];

  events.forEach(event => {
    let fullTitle = event.getTitle();
    let cleanTitle = fullTitle.split('-')[0].trim();

    let status = "Aktiv";
    if (fullTitle.toLowerCase().includes("storniert") || fullTitle.toLowerCase().includes("abgesagt")) {
      status = "Storniert";
    }

    let eventDate = new Date(event.getStartTime());
    eventDate.setHours(0, 0, 0, 0);

    // Gäste filtern: Nur behalten, was NICHT in der ignoreList steht
    let guestList = event.getGuestList();
    let cleanerEmail = "";

    for (let i = 0; i < guestList.length; i++) {
      let email = guestList[i].getEmail().toLowerCase().trim();
      if (ignoreList.indexOf(email) === -1) {
        cleanerEmail = email; 
        break; // Sobald wir den "echten" Cleaner gefunden haben, hören wir auf
      }
    }

    // Wir gehen zurück auf 6 Spalten (Datum, Titel, Status, ID, Ersteller, Cleaner-Gast)
    eventData.push([
      eventDate,
      cleanTitle, 
      status,
      event.getId(),
      event.getCreators().join(", "),
      cleanerEmail // Hier steht jetzt nur der relevante Cleaner
    ]);
  });

  sheet.clear();
  // Header auf 6 Spalten angepasst
  sheet.getRange(1, 1, 1, 6).setValues([["Datum", "Titel", "Status", "ID", "Ersteller", "Cleaner Gast"]]);

  if (eventData.length > 0) {
    sheet.getRange(2, 1, eventData.length, 6).setValues(eventData);
    
    const range = sheet.getRange(2, 1, eventData.length, 6);
    range.sort([
      {column: 1, ascending: true}, 
      {column: 2, ascending: true}
    ]);
    
    sheet.getRange(2, 1, eventData.length, 1).setNumberFormat("dd.MM.yyyy");
  }
}