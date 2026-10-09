/**
 * ZENTRALE FILTER FÜR TAB "db"
 */
const FILTERS = {
  maxPastDays: 30,     // Berücksichtigt Abreisen bis zu 30 Tage in der Vergangenheit
  maxFutureDays: 365,  // Berücksichtigt Abreisen bis zu 365 Tage in der Zukunft
  status: "Booked",
  excludeDeleted: true
};

/**
 * Hauptfunktion: Holt Buchungen von Lodgify und befüllt Tab "db"
 */
function getLodgifyBookings() {
  const props = PropertiesService.getScriptProperties();
  const loKey = props.getProperty('loKey');
  const fixieUrl = props.getProperty('fixieUrl');

  if (!loKey || !fixieUrl) {
    Logger.log("FEHLER: 'loKey' oder 'fixieUrl' fehlen in Script Properties!");
    return;
  }

  // Abfrage aller Buchungen ab 2025
  const baseUrl = "https://api.lodgify.com/v2/reservations/bookings?size=100&includeCount=false&stayFilter=All&updatedSince=2025-01-01T00%3A00%3A00.000Z&includeTransactions=false&includeExternal=false&includeQuoteDetails=false";

  let currentPage = 1;
  let allItems = [];
  let keepGoing = true;

  Logger.log("Rufe Buchungsdaten von Lodgify ab...");

  while (keepGoing) {
    const fullUrlWithPage = baseUrl + "&page=" + currentPage;
    const responseData = sendThroughFixie(fullUrlWithPage, loKey, fixieUrl);

    if (responseData && responseData.items && responseData.items.length > 0) {
      allItems = allItems.concat(responseData.items);
      if (responseData.items.length < 100) {
        keepGoing = false;
      } else {
        currentPage++;
      }
    } else {
      keepGoing = false;
    }
    if (currentPage > 30) keepGoing = false; // Sicherheitsabbruch nach 30 Seiten (3000 Buchungen)
  }

  Logger.log("Insgesamt erhaltene Roh-Buchungen: " + allItems.length);
  writeToSheet(allItems);
}

/**
 * Hilfsfunktion für HTTP-Proxy über Fixie
 */
function sendThroughFixie(targetUrl, loKey, fixieUrl) {
  const authCredentials = fixieUrl.split('@')[0].replace('http://', '');
  const options = {
    "method": "get",
    "headers": {
      "X-ApiKey": loKey,
      "Proxy-Authorization": "Basic " + Utilities.base64Encode(authCredentials),
      "accept": "application/json"
    },
    "muteHttpExceptions": true
  };
  const response = UrlFetchApp.fetch(targetUrl, options);
  return JSON.parse(response.getContentText());
}

/**
 * Verarbeitet die API-Daten und schreibt sie in den Tab "db" (ab Zeile 4)
 */
function writeToSheet(items) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("db");

  if (!sheet) sheet = ss.insertSheet("db");

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const rows = items
    .filter(item => {
      if (!item.departure || !item.arrival) return false;

      const depDate = new Date(item.departure);
      depDate.setHours(0, 0, 0, 0);

      const arrDate = new Date(item.arrival);
      arrDate.setHours(0, 0, 0, 0);

      const todayCopy = new Date(today);

      // Tage von HEUTE im Vergleich zu Abreise und Anreise
      const diffDaysDep = Math.round((todayCopy - depDate) / (1000 * 60 * 60 * 24));
      const diffDaysArr = Math.round((todayCopy - arrDate) / (1000 * 60 * 60 * 24));

      // Datumsfilter nach Konfiguration oben
      const matchesDeparture =
        diffDaysDep <= FILTERS.maxPastDays &&
        diffDaysDep >= -FILTERS.maxFutureDays;

      const matchesArrival =
        diffDaysArr <= FILTERS.maxPastDays &&
        diffDaysArr >= -FILTERS.maxFutureDays;

      if (!(matchesDeparture || matchesArrival)) return false;
      if (item.status !== FILTERS.status) return false;
      if (FILTERS.excludeDeleted && item.is_deleted) return false;

      return true;
    })
    .map(item => {
      // Datum ohne Uhrzeit
      const arrival = item.arrival ? new Date(new Date(item.arrival).setHours(0,0,0,0)) : "";
      const departure = item.departure ? new Date(new Date(item.departure).setHours(0,0,0,0)) : "";
      const updated = item.updated_at ? new Date(item.updated_at) : "";
      const peopleCount = (item.guest_breakdown) ? item.guest_breakdown.adults : (item.rooms && item.rooms[0] ? item.rooms[0].people : "");

      let row = new Array(30).fill("");

      row[0] = item.property_name || (item.property ? item.property.name : ""); // Spalte A: Apartment Name
      row[1] = String(item.id).replace(/-/g, ""); // Spalte B: Booking ID
      row[3] = arrival;   // Spalte D: Ankunft
      row[4] = departure; // Spalte E: Abreise
      row[5] = String(item.property_id).replace(/-/g, ""); // Spalte F: Property ID
      row[11] = peopleCount; // Spalte L: Personen
      row[13] = (item.guest) ? item.guest.name : "";  // Spalte N: Gast Name
      row[14] = (item.guest) ? item.guest.email : ""; // Spalte O: Gast E-Mail
      row[15] = (item.guest) ? item.guest.phone : ""; // Spalte P: Gast Telefon
      row[18] = item.status || ""; // Spalte S: Status
      row[22] = item.source || ""; // Spalte W: Quelle
      row[26] = updated;          // Spalte AA: Aktualisiert am
      row[29] = item.is_deleted || false; // Spalte AD: Gelöscht Flag

      return row;
    });

  // Sortierung chronologisch nach Abreisedatum (Spalte E / Index 4)
  rows.sort((a, b) => {
    const dateA = a[4];
    const dateB = b[4];
    if (!(dateA instanceof Date)) return 1;
    if (!(dateB instanceof Date)) return -1;
    return dateA - dateB;
  });

  // Altes Blatt ab Zeile 4 leeren
  const lastRow = sheet.getLastRow();
  const maxColumns = 30;
  if (lastRow >= 4) {
    sheet.getRange(4, 1, lastRow - 3, maxColumns).clearContent();
  }

  // Neue Daten schreiben
  if (rows.length > 0) {
    const targetRange = sheet.getRange(4, 1, rows.length, maxColumns);
    targetRange.setValues(rows);
    
    // Spalte D (Ankunft) und E (Abreise) als YYYY-MM-DD formatieren
    sheet.getRange(4, 4, rows.length, 2).setNumberFormat("yyyy-mm-dd");
    Logger.log(rows.length + " Buchungen erfolgreich in Tab 'db' eingetragen.");
  } else {
    Logger.log("Keine passenden Buchungen gefunden.");
  }
}