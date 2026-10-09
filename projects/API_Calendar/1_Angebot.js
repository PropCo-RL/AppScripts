// ============================================================================
// AUTOMATISCHES APARTMENT-MATCHING & ANGEBOTS-ENTWURF (TAB 'A')
// ============================================================================

function processApartmentInquiries() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetA = ss.getSheetByName("A");
  const sheetCal = ss.getSheetByName("cal");

  if (!sheetA || !sheetCal) throw new Error("Blatt 'A' oder 'cal' fehlt.");

  // URSPRUNGS-SHEET FÜR ENT WURF-EXPORT ÖFFNEN
  const LOG_SPREADSHEET_ID = "1bMPVHjGNnwWDH0KXUyB47_VbLcq_lsFK1jSbdj_GmH0";
  let logSheet = null;
  try {
    const logSs = SpreadsheetApp.openById(LOG_SPREADSHEET_ID);
    logSheet = logSs.getSheetByName("Log");
  } catch (e) {
    Logger.log("WARNUNG: Ursprungs-Sheet für Entwurf-Export konnte nicht geöffnet werden: " + e.toString());
  }

  const lastRowA = sheetA.getLastRow();
  if (lastRowA < 2) return;

  const apartmentsD = getApartmentsFromD();
  const calBookings = getCalBookings(sheetCal);

  const rangeA = sheetA.getRange(2, 1, lastRowA - 1, 10);
  const valuesA = rangeA.getValues();

  valuesA.forEach((row, index) => {
    const rowIndex = index + 2;
    const currentOfferJ = row[9] ? row[9].toString().trim() : "";

    if (currentOfferJ !== "") return; // Bereits verarbeitet

    const phoneRaw = row[4] ? row[4].toString().trim() : ""; // Spalte E (Telefonnummer)
    const targetCity = row[5] ? row[5].toString().trim() : "";
    const paxRaw = row[6];
    const startDateRaw = row[7];
    const endDateRaw = row[8];

    if (!targetCity || !paxRaw || !startDateRaw || !endDateRaw) return;

    const paxNeeded = parseInt(paxRaw, 10) || 1;
    const reqStart = parseDateInquiry(startDateRaw);
    const reqEnd = parseDateInquiry(endDateRaw);

    if (!reqStart || !reqEnd || reqStart >= reqEnd) {
      sheetA.getRange(rowIndex, 10).setValue("ERROR: Ungültiges Datumsformat");
      return;
    }

    Logger.log(`[Zeile ${rowIndex}] Suche für ${targetCity} (${paxNeeded} Pax)...`);

    // 1. APARTMENTS MIT FAHRZEIT UND KAPAZITÄT VORFILTERN
    const validApts = [];
    apartmentsD.forEach(apt => {
      if (apt.beds < paxNeeded) return;
      const driveTime = getDrivingTimeMinutes(targetCity, apt.fullAddress);
      if (driveTime !== null && driveTime <= 45) {
        validApts.push({ ...apt, driveTime: driveTime });
      }
    });

    let opt1_DirectMain = []; // <= 30 Min, durchgehend
    let opt2_DirectB = [];    // 30-45 Min, durchgehend
    let opt3_Delayed = [];    // <= 30 Min, Start 1-7 Tage später
    let opt4_Splits = [];     // Umzug

    // 2. OPTION 1 & 2: DURCHGEHENDE BEREITSCHAFTS-PRÜFUNG
    validApts.forEach(apt => {
      const isBooked = checkCalendarOverlap(apt.name, reqStart, reqEnd, calBookings);
      if (!isBooked) {
        if (apt.driveTime <= 30) {
          opt1_DirectMain.push(apt);
        } else {
          opt2_DirectB.push(apt);
        }
      }
    });

    // 3. OPTION 3: SPÄTERER START (falls keine oder wenige perfekte Treffer)
    if (opt1_DirectMain.length < 2) {
      validApts.filter(a => a.driveTime <= 30).forEach(apt => {
        for (let delayDays = 1; delayDays <= 7; delayDays++) {
          const altStart = new Date(reqStart.getTime());
          altStart.setDate(altStart.getDate() + delayDays);

          if (altStart >= reqEnd) break;

          const isBooked = checkCalendarOverlap(apt.name, altStart, reqEnd, calBookings);
          if (!isBooked) {
            opt3_Delayed.push({
              ...apt,
              delayDays: delayDays,
              startDate: altStart
            });
            break;
          }
        }
      });
    }

    // 4. OPTION 4: ZWISCHENSTOPP / SPLIT
    if (opt1_DirectMain.length === 0) {
      const mainApts = validApts.filter(a => a.driveTime <= 30);
      const bApts = validApts.filter(a => a.driveTime > 30 && a.driveTime <= 45);

      mainApts.forEach(mainApt => {
        for (let splitDayOffset = 1; splitDayOffset <= 7; splitDayOffset++) {
          const splitDate = new Date(reqStart.getTime());
          splitDate.setDate(splitDate.getDate() + splitDayOffset);

          if (splitDate >= reqEnd) break;

          const mainFree = !checkCalendarOverlap(mainApt.name, splitDate, reqEnd, calBookings);

          if (mainFree) {
            for (let bApt of bApts) {
              const bFreeFirstLeg = !checkCalendarOverlap(bApt.name, reqStart, splitDate, calBookings);
              if (bFreeFirstLeg) {
                opt4_Splits.push({
                  startApt: bApt,
                  mainApt: mainApt,
                  switchDate: splitDate,
                  daysInB: splitDayOffset
                });
                break;
              }
            }
          }
        }
      });
    }

    // 5. SORTIERUNGEN
    opt1_DirectMain.sort((a, b) => a.driveTime - b.driveTime);
    opt2_DirectB.sort((a, b) => a.driveTime - b.driveTime);
    opt3_Delayed.sort((a, b) => a.delayDays - b.delayDays || a.driveTime - b.driveTime);

    // 6. FORMATIERUNG DES ERGEBNISSES FÜR SPALTE J
    let outputParts = [];

    if (opt1_DirectMain.length > 0) {
      const txt = opt1_DirectMain.map(a => `${a.code} (${a.driveTime} Min)`).join(", ");
      outputParts.push(`PERFEKT: ${txt}`);
    }

    if (opt2_DirectB.length > 0) {
      const txt = opt2_DirectB.map(a => `${a.code} (${a.driveTime} Min)`).join(", ");
      outputParts.push(`B-LAGE: ${txt}`);
    }

    if (opt3_Delayed.length > 0) {
      const txt = opt3_Delayed.slice(0, 3).map(d => `${d.code} (${d.driveTime} Min, ab ${formatDateGerman(d.startDate)})`).join(", ");
      outputParts.push(`SPÄTERER START (+1-7T): ${txt}`);
    }

    if (opt4_Splits.length > 0) {
      const s = opt4_Splits[0];
      outputParts.push(`SPLIT/UMZUG: Zuerst ${s.startApt.code} (${s.startApt.driveTime} Min), ab ${formatDateGerman(s.switchDate)} Wechsel in ${s.mainApt.code} (${s.mainApt.driveTime} Min)`);
    }

    let finalResult = outputParts.join(" | ");
    if (!finalResult) finalResult = "Keine passenden Angebote gefunden";

    sheetA.getRange(rowIndex, 10).setValue(finalResult);
    Logger.log(`[Zeile ${rowIndex}] ${finalResult}`);

    // 7. NEU: ANGEBOTS-ENT WURF BAUEN & IN URSPRUNGS-SHEET (SPALTE R) SCHREIBEN
    if (logSheet && phoneRaw) {
      // Bestimme die besten Optionen (max 3 Alternativen)
      let selectedOffers = [...opt1_DirectMain, ...opt2_DirectB];
      if (selectedOffers.length === 0 && opt3_Delayed.length > 0) {
        selectedOffers = opt3_Delayed;
      }
      
      const offerText = buildWhatsAppOfferDraft(selectedOffers, opt4_Splits);
      if (offerText) {
        saveDraftToLogSheet(logSheet, phoneRaw, offerText);
      }
    }
  });

  SpreadsheetApp.flush();
}


// ============================================================================
// NEU: HILFSFUNKTION - NATIVE NORMALE NACHRICHTEN-ERSTELLUNG
// ============================================================================

function buildWhatsAppOfferDraft(availableApts, splitOptions) {
  if (availableApts.length === 0 && splitOptions.length === 0) {
    return "Leider haben wir für den gewünschten Zeitraum keine passenden Wohnungen frei.\n\nAGB: L8Street.com/agb";
  }

  const itemsToInclude = availableApts.slice(0, 3);
  let message = "Die folgenden Apartments wären noch frei:\n\n";

  itemsToInclude.forEach((apt, idx) => {
    if (idx > 0) {
      message += "Alternativ ginge diese Wohnung:\n\n";
    }

    const roomsText = apt.rooms ? `${apt.rooms} Zimmer` : "Mehrere Zimmer";
    const bedsText = apt.beds ? `${apt.beds} Betten` : "Betten vorhanden";
    const priceText = apt.price ? `${apt.price} Euro pro Nacht` : "Preis auf Anfrage";

    message += `Fotos und Ausstattungsdetails auf unserer Website unter folgendem Link:\n`;
    message += `L8Street.com/a/${apt.code}\n\n`;
    message += `Adresse der Wohnung:\n${apt.fullAddress}\n`;
    message += `Das ist ungefähr ${apt.driveTime} Minuten Fahrt.\n`;
    message += `Wohnung mit ${roomsText}, ${bedsText}, eigener Küche und Bad\n`;
    message += `${priceText}\n\n`;
  });

  // Falls nur ein Split als Ausweichoption existiert
  if (itemsToInclude.length === 0 && splitOptions.length > 0) {
    const s = splitOptions[0];
    message += `Fotos und Ausstattungsdetails unter folgenden Links:\n`;
    message += `1. Teil: L8Street.com/a/${s.startApt.code}\n`;
    message += `2. Teil (ab ${formatDateGerman(s.switchDate)}): L8Street.com/a/${s.mainApt.code}\n\n`;
    message += `Wohnungswechsel erforderlich am ${formatDateGerman(s.switchDate)}.\n\n`;
  }

  message += "AGB: L8Street.com/agb";
  return message;
}


// ============================================================================
// NEU: HILFSFUNKTION - SCHREIBT DEN ENT WURF IN SPALTE R VOM LOG-SHEET
// ============================================================================

function saveDraftToLogSheet(logSheet, phoneRaw, offerText) {
  const cleanPhone = String(phoneRaw).replace(/\D/g, "");
  if (!cleanPhone) return;

  const lastRow = logSheet.getLastRow();
  if (lastRow < 2) return;

  // Suche im Log-Sheet nach der aktuellsten Zeile der Telefonnummer mit 'y - mit gemini' in Spalte Q (Spalte 17)
  const data = logSheet.getRange(2, 1, lastRow - 1, 17).getValues();

  for (let i = data.length - 1; i >= 0; i--) {
    const rowPhone = String(data[i][2]).replace(/\D/g, ""); // Spalte C
    const markerQ = String(data[i][16]).trim().toLowerCase(); // Spalte Q

    if (rowPhone === cleanPhone && markerQ.includes("y - mit gemini")) {
      const targetRowIndex = i + 2;
      logSheet.getRange(targetRowIndex, 18).setValue(offerText); // Spalte R (Spalte 18)
      Logger.log(`Entwurf erfolgreich in Log-Sheet (Zeile ${targetRowIndex}, Spalte R) gespeichert.`);
      break;
    }
  }
}


// ============================================================================
// BESTEHENDE HILFSFUNKTIONEN
// ============================================================================

function getApartmentsFromD() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetD = ss.getSheetByName("d");
  if (!sheetD) throw new Error("Blatt 'd' fehlt.");

  const lastRow = sheetD.getLastRow();
  if (lastRow < 2) return [];

  const data = sheetD.getRange(2, 1, lastRow - 1, 17).getValues();
  const list = [];

  data.forEach(row => {
    const name = row[2] ? row[2].toString().trim() : "";
    const link = row[3] ? row[3].toString().trim() : "";
    const street = row[7] ? row[7].toString().trim() : "";
    const houseNo = row[8] ? row[8].toString().trim() : "";
    const zip = row[9] ? row[9].toString().trim() : "";
    const city = row[10] ? row[10].toString().trim() : "";
    const rooms = parseInt(row[15], 10) || 0; // Spalte P (Zimmer)
    const beds = parseInt(row[16], 10) || 0;  // Spalte Q (Betten)
    const price = row[11] ? row[11].toString().trim() : ""; // Beispiel Spalte L (Preis)

    if (!name || !link) return;

    let code = name;
    const match = link.match(/\/a\/([^\/\?]+)/i);
    if (match && match[1]) code = match[1].trim();

    const fullAddress = `${street} ${houseNo}, ${zip} ${city}`.trim();

    list.push({ name, code, fullAddress, beds, rooms, price });
  });

  return list;
}

function getCalBookings(sheetCal) {
  const lastRow = sheetCal.getLastRow();
  if (lastRow < 2) return {};

  const data = sheetCal.getRange(2, 1, lastRow - 1, 3).getValues();
  const bookings = {};

  data.forEach(row => {
    const aptName = row[0] ? row[0].toString().trim() : "";
    const start = parseDateInquiry(row[1]);
    const end = parseDateInquiry(row[2]);

    if (aptName && start && end) {
      if (!bookings[aptName]) bookings[aptName] = [];
      bookings[aptName].push({ start, end });
    }
  });

  return bookings;
}

function checkCalendarOverlap(aptName, reqStart, reqEnd, calBookings) {
  const aptBookings = calBookings[aptName] || [];
  for (let b of aptBookings) {
    if (reqStart.getTime() < b.end.getTime() && reqEnd.getTime() > b.start.getTime()) {
      return true;
    }
  }
  return false;
}

function getDrivingTimeMinutes(originCity, destinationAddress) {
  try {
    const directions = Maps.newDirectionFinder()
      .setOrigin(originCity)
      .setDestination(destinationAddress)
      .setMode(Maps.DirectionFinder.Mode.DRIVING)
      .getDirections();

    if (directions.status === "OK" && directions.routes.length > 0) {
      return Math.round(directions.routes[0].legs[0].duration.value / 60);
    }
  } catch (e) {
    Logger.log(`Maps Fehler (${originCity} -> ${destinationAddress}): ${e.toString()}`);
  }
  return null;
}

function parseDateInquiry(dateVal) {
  if (!dateVal) return null;
  if (dateVal instanceof Date) {
    const d = new Date(dateVal.getTime());
    d.setHours(12, 0, 0, 0);
    return isNaN(d.getTime()) ? null : d;
  }

  const str = dateVal.toString().trim();
  const parts = str.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})/);
  if (parts) {
    return new Date(parseInt(parts[3], 10), parseInt(parts[2], 10) - 1, parseInt(parts[1], 10), 12, 0, 0);
  }

  const d = new Date(str);
  if (!isNaN(d.getTime())) {
    d.setHours(12, 0, 0, 0);
    return d;
  }
  return null;
}

function formatDateGerman(dateObj) {
  return Utilities.formatDate(dateObj, "Europe/Berlin", "dd.MM.yyyy");
}