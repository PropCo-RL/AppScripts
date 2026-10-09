/**
 * Prüft NUR API Termine gegen bestehende MATCH Termine.
 *
 * Regel:
 * - 0 oder 1 bestehender MATCH Termin → nichts tun
 * - 2 oder mehr MATCH Termine → Empfehlung in "Kap"
 *
 * Kap Ausgabe:
 * A = Timestamp
 * B = Apartment
 * C = Original Datum
 * D = Empfehlung
 * E = Cleaner Email
 */

function generateKapRecommendations() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const gCalSheet = ss.getSheetByName("gCal");
  const kapSheet = ss.getSheetByName("Kap");

  if (!gCalSheet || !kapSheet) {
    console.log("gCal oder Kap fehlt");
    return;
  }

  const data =
    gCalSheet
      .getDataRange()
      .getValues()
      .slice(1);

  const now = new Date();

  const timestamp = Utilities.formatDate(
    now,
    Session.getScriptTimeZone(),
    "dd.MM.yyyy HH 'Uhr'"
  );

  /**
   * MATCH Termine zählen
   */
  const matchCounter = {};

  data.forEach(row => {

    const status =
      String(row[2]).trim();

    if (status !== "MATCH") {
      return;
    }

    const date =
      formatDate(row[0]);

    const email =
      String(row[3])
        .trim()
        .toLowerCase();

    const key =
      date + "_" + email;

    matchCounter[key] =
      (matchCounter[key] || 0) + 1;
  });

  /**
   * NUR API prüfen
   */
  const newEntries = [];

  data.forEach(row => {

    const status =
      String(row[2]).trim();

    if (status !== "NUR API") {
      return;
    }

    const apartment =
      row[1];

    const originalDate =
      formatDate(row[0]);

    const email =
      String(row[3])
        .trim()
        .toLowerCase();

    const key =
      originalDate + "_" + email;

    const existingMatches =
      matchCounter[key] || 0;

    // Nur Empfehlung wenn bereits 2+ MATCH Termine existieren
    if (existingMatches < 2) {
      return;
    }

    const suggestedDate =
      findNextFreeDate(
        matchCounter,
        row[0],
        email
      );

    const suggestedDateStr =
      formatDate(suggestedDate);

    newEntries.push([
      timestamp,
      apartment,
      originalDate,
      suggestedDateStr,
      email
    ]);
  });

  /**
   * Nach Email + Datum sortieren
   */
  newEntries.sort((a, b) => {

    const emailCompare =
      a[4].localeCompare(b[4]);

    if (emailCompare !== 0) {
      return emailCompare;
    }

    return (
      new Date(a[3]) -
      new Date(b[3])
    );
  });

  /**
   * Schreiben
   */
  if (newEntries.length > 0) {

    const startRow =
      Math.max(
        2,
        kapSheet.getLastRow() + 1
      );

    kapSheet
      .getRange(
        startRow,
        1,
        newEntries.length,
        5
      )
      .setValues(newEntries);
  }
}

/**
 * Findet nächsten Tag mit weniger als 2 MATCH Terminen
 */
function findNextFreeDate(
  matchCounter,
  startDate,
  email
) {

  const date =
    new Date(startDate);

  date.setHours(0, 0, 0, 0);

  let found = false;
  let safety = 0;

  while (!found && safety < 30) {

    date.setDate(
      date.getDate() + 1
    );

    const key =
      formatDate(date) +
      "_" +
      email;

    const load =
      matchCounter[key] || 0;

    if (load < 2) {
      found = true;
    }

    safety++;
  }

  return date;
}

/**
 * Datum zu YYYY-MM-DD
 */
function formatDate(date) {

  if (
    !date ||
    isNaN(
      new Date(date).getTime()
    )
  ) {
    return "";
  }

  return Utilities.formatDate(
    new Date(date),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );
}