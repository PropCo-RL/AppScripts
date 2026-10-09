function syncGCalWithAudit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gCalSheet = ss.getSheetByName("gCal");
  const auditData = gCalSheet.getDataRange().getValues().slice(1);
  
  const calendarId = 'support@l8street.com'; 
  const calendar = CalendarApp.getCalendarById(calendarId);

  let createdCount = 0;
  let deletedCount = 0;

  auditData.forEach((row, index) => {
    const date = new Date(row[0]);
    const apartment = row[1];
    const status = row[2];
    const cleanerEmail = row[3];
    const eventId = row[4]; // Die ID steht jetzt direkt hier in Spalte E!

    // FALL 1: NEUES EVENT ERSTELLEN
    if (status === "NUR API") {
      const event = calendar.createAllDayEvent(apartment, date, {
        guests: cleanerEmail + ",info@l8street.com",
        sendInvites: true
      });
      
      createdCount++;
      gCalSheet.getRange(index + 2, 3).setValue("MATCH (Neu erstellt)");
    }

    // FALL 2: EVENT LÖSCHEN
    else if (status === "NUR KALENDER" && eventId) {
      try {
        const event = calendar.getEventById(eventId);
        if (event) {
          event.deleteEvent();
          deletedCount++;
          gCalSheet.getRange(index + 2, 3).setValue("Gelöscht");
        }
      } catch (e) {
        console.log("Konnte Event nicht löschen: " + eventId);
      }
    }
  });
}