const ADMIN_EMAIL = "info@l8street.com";

// ==========================================
// 1. API ROUTING (doGet & doPost)
// ==========================================
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
  
  if (action === 'getAvailability') {
    const ids = e.parameter.ids ? e.parameter.ids.split(',') : [];
    const avail = fetchAvailabilityForApts(ids);
    return ContentService.createTextOutput(JSON.stringify(avail)).setMimeType(ContentService.MimeType.JSON);
  }
  
  if (action === 'getAllApartments') {
    const allApts = getAllApartmentsForBuild();
    return ContentService.createTextOutput(JSON.stringify(allApts)).setMimeType(ContentService.MimeType.JSON);
  }
  
  return ContentService.createTextOutput(JSON.stringify({ status: "API active" })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch(err) {
    payload = e.parameter || {};
  }
  
  const result = submitNewBooking(payload);
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// 2. DATEN FÜR GITHUB GENERATOR (getAllApartments)
// ==========================================
function getAllApartmentsForBuild() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetApts = ss.getSheets()[0];
  const data = sheetApts.getDataRange().getValues();
  let results = [];

  for (let i = 1; i < data.length; i++) {
    let status = String(data[i][20] || "").trim(); // Spalte U (Status)
    if (status.toLowerCase() === "archived" || status.toLowerCase() === "inaktiv") continue;
    
    let apartmentPath = String(data[i][6] || "").trim(); // Spalte G (/a/gera1)
    if (!apartmentPath) continue;
    
    let driveImages = parseGalleryUrls(String(data[i][21] || "").trim()); // Spalte V
    let galleryImages = parseGalleryUrls(String(data[i][11] || "").trim()); // Spalte L
    let imageList = [...driveImages, ...galleryImages].filter(url => !url.includes("placeholder"));
    
    if (imageList.length === 0) {
      imageList = ["https://via.placeholder.com/600x400?text=Apartment+L8+Street"];
    }

    let internalTitle = String(data[i][0] || "").trim(); // Spalte A
    let guestTitle = String(data[i][7] || internalTitle).trim(); // Spalte H (GuestTitle)
    let rankingVal = parseInt(data[i][27]);
    if (isNaN(rankingVal)) rankingVal = 999;
    
    const isAnonymous = String(data[i][32] || "").trim().toLowerCase() === "x";
    const cityStr = String(data[i][9] || "").trim(); // Spalte J
    const streetStr = String(data[i][13] || "").trim(); // Spalte N
    const zipStr = String(data[i][12] || "").trim(); // Spalte M
    
    let displayAddress = "";
    if (isAnonymous) {
      displayAddress = cityStr ? `${zipStr} ${cityStr}`.trim() : "Ort auf Anfrage";
    } else {
      displayAddress = [streetStr, `${zipStr} ${cityStr}`.trim()].filter(Boolean).join(", ");
    }

    results.push({
      id: String(data[i][1] || internalTitle).trim(), // Spalte B (ID)
      internalTitle: internalTitle,
      title: guestTitle,
      Apartment: apartmentPath,
      city: cityStr,
      regions: String(data[i][22] || "").trim(), // Spalte W (Regionen)
      street: isAnonymous ? "" : streetStr,
      zip: zipStr,
      isAnonymous: isAnonymous,
      displayAddress: displayAddress,
      bedrooms: data[i][14] || 1, // Spalte O
      beds: parseInt(data[i][15]) || 1, // Spalte P
      pricePerNight: data[i][17] || "49", // Spalte R
      ranking: rankingVal,
      images: imageList,
      reviews: parseGalleryUrls(String(data[i][33] || "").trim()) // Spalte AH (Reviews)
    });
  }

  results.sort((a, b) => a.ranking - b.ranking);
  return results;
}

// ==========================================
// 3. VERFÜGBARKEITEN PRÜFEN
// ==========================================
function fetchAvailabilityForApts(aptIds) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calBookings = getCalBookingsFromLocalSheet(ss);
  const sheetApts = ss.getSheets()[0];
  const data = sheetApts.getDataRange().getValues();
  
  const reqStart = new Date();
  reqStart.setHours(12, 0, 0, 0);
  const reqEnd = new Date(reqStart.getTime() + 7 * 24 * 60 * 60 * 1000);
  
  const cleanAptIds = aptIds.map(id => String(id).trim().toLowerCase());
  let availabilities = {};

  for (let i = 1; i < data.length; i++) {
    let internalName = String(data[i][0] || "").trim();
    let id = String(data[i][1] || internalName).trim();
    
    let isRequested = cleanAptIds.length === 0 || 
                      cleanAptIds.includes(id.toLowerCase()) || 
                      cleanAptIds.includes(internalName.toLowerCase());

    if (isRequested) {
      let isBooked = checkCalendarOverlap(internalName, reqStart, reqEnd, calBookings);
      let resObj = {};
      
      if (!isBooked) {
        resObj = { isDirectlyAvailable: true };
      } else {
        let availableFromDateStr = null;
        for (let delayDays = 1; delayDays <= 90; delayDays++) {
          let altStart = new Date(reqStart.getTime());
          altStart.setDate(altStart.getDate() + delayDays);
          let altEnd = new Date(altStart.getTime() + (reqEnd.getTime() - reqStart.getTime()));
          
          if (!checkCalendarOverlap(internalName, altStart, altEnd, calBookings)) {
            availableFromDateStr = formatDateGerman(altStart);
            break;
          }
        }
        resObj = { isDirectlyAvailable: false, availableFromDate: availableFromDateStr };
      }
      
      availabilities[id] = resObj;
      availabilities[internalName] = resObj;
    }
  }

  return availabilities;
}

function getCalBookingsFromLocalSheet(ss) {
  const bookings = {};
  const sheetCal = ss.getSheetByName("cal");
  if (!sheetCal) return bookings;
  
  const data = sheetCal.getRange(2, 1, Math.max(1, sheetCal.getLastRow() - 1), 3).getValues();
  data.forEach(row => {
    const aptName = row[0] ? row[0].toString().trim() : "";
    const start = parseGermanDate(row[1]);
    const end = parseGermanDate(row[2]);
    if (aptName && start && end) {
      if (!bookings[aptName]) bookings[aptName] = [];
      bookings[aptName].push({ start, end });
    }
  });
  return bookings;
}

function checkCalendarOverlap(aptName, reqStart, reqEnd, calBookings) {
  if (!aptName) return false;
  const cleanSearchName = String(aptName).trim().toLowerCase();
  let aptBookings = [];
  
  Object.keys(calBookings).forEach(calKey => {
    if (calKey.trim().toLowerCase() === cleanSearchName) {
      aptBookings = aptBookings.concat(calBookings[calKey]);
    }
  });

  for (let b of aptBookings) {
    if (reqStart.getTime() < b.end.getTime() && reqEnd.getTime() > b.start.getTime()) {
      return true;
    }
  }
  return false;
}

// ==========================================
// 4. BUCHUNG SPEICHERN MIT AUDIT-LOG (§ 286 ZPO)
// ==========================================
function submitNewBooking(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetSites = ss.getSheetByName("sites") || ss.insertSheet("sites");
    const now = new Date();
    const formattedDate = Utilities.formatDate(now, "Europe/Berlin", "dd.MM.yyyy HH:mm:ss");
    
    const bookingTimestamp = payload.timestamp || now.toISOString();
    const bookingId = payload.bookingId || ("BK-SITE-" + Math.floor(100000 + Math.random() * 900000));
    const fullBillingAddress = `${payload.street || ''}, ${payload.zip || ''} ${payload.city || ''}, ${payload.country || ''}`;

    // Saubere CSV-Audit-Zeile
    var csvAuditLine = [
      bookingTimestamp,
      bookingId,
      payload.source || "GitHub_Pages_Native",
      payload.company || "Privat",
      payload.email || "k.A.",
      payload.phone || "k.A.",
      payload.clientIp || "CLIENT_IP",
      payload.userAgent || "Browser",
      (payload.aptTitle || "") + " (" + (payload.aptCode || "") + ")",
      (payload.startDate || "") + " bis " + (payload.endDate || ""),
      payload.agbAccepted || "AGB_AND_STORNO_ACCEPTED_TRUE",
      payload.agbVersion || "AGB_VERSION_2026_01",
      payload.actionBtn || "BUTTON_CLICK_VERBINDLICH_BUCHEN"
    ].join(";");

    // Fügt die Buchung in die lokale Tabelle "sites" ein
    sheetSites.appendRow([
      formattedDate,                     // Spalte A: Datum
      "Direktbuchung Website",          // Spalte B: Art
      payload.aptTitle || "",           // Spalte C: Titel
      payload.aptCode || "",            // Spalte D: Kürzel
      payload.company || "",            // Spalte E: Firmenname
      payload.vatId || "",              // Spalte F: USt-ID
      payload.email || "",              // Spalte G: E-Mail
      payload.phone || "",              // Spalte H: Tel
      payload.startDate || "",          // Spalte I: Start
      payload.endDate || "",            // Spalte J: Ende
      payload.guestsCount || "",        // Spalte K: Personen
      fullBillingAddress,               // Spalte L: Adresse
      payload.refNumber || "",          // Spalte M: Zusatz
      "Gebucht",                        // Spalte N: Status
      JSON.stringify(payload),          // Spalte O: JSON
      csvAuditLine                      // Spalte P (16): Audit_Log_CSV
    ]);

    // E-Mail-Audit-Trail senden
    var emailSubject = "AUDIT-LOG BEWEIS: Neue Buchung " + bookingId + " (" + (payload.aptTitle || "") + ")";
    try {
      var emailBody = "Eine neue rechtsverbindliche Buchung wurde getätigt.\n\n" +
        "--- BEWEIS-PROTOKOLL (§ 286 ZPO / § 312j BGB) ---\n" +
        "Buchungs-ID: " + bookingId + "\n" +
        "Zeitstempel (UTC): " + bookingTimestamp + "\n" +
        "IP-Adresse: " + (payload.clientIp || "k.A.") + "\n" +
        "User-Agent: " + (payload.userAgent || "k.A.") + "\n" +
        "AGB & Storno akzeptiert: JA\n\n" +
        "CSV-String:\n" + csvAuditLine;
      GmailApp.sendEmail(ADMIN_EMAIL, emailSubject, emailBody);
    } catch(mailErr) {
      console.error("E-Mail-Versand fehlgeschlagen: " + mailErr.toString());
    }

    // Module aus der 2. Datei aufrufen
    try {
      appendToReSheet(payload);
    } catch(e) {
      console.error("Fehler beim Aufruf von appendToReSheet: " + e.toString());
    }

    try {
      applyBookingLabel(emailSubject);
    } catch(e) {
      console.error("Fehler beim Aufruf von applyBookingLabel: " + e.toString());
    }

    return { success: true, message: "Vielen Dank! Ihre Buchung wurde erfolgreich erfasst." };
  } catch (err) {
    return { success: false, message: "Fehler beim Buchen: " + err.toString() };
  }
}

// ==========================================
// 5. HELPER FUNKTIONEN
// ==========================================
function parseGalleryUrls(rawGalleryStr) {
  let urls = [];
  if (!rawGalleryStr) return [];
  try {
    if (rawGalleryStr.startsWith("[") || rawGalleryStr.startsWith("{")) {
      let parsed = JSON.parse(rawGalleryStr);
      if (Array.isArray(parsed)) parsed.forEach(item => {
        if (item.src) urls.push(cleanWixUrl(item.src));
        else if (typeof item === 'string') urls.push(cleanWixUrl(item));
      });
    } else {
      rawGalleryStr.split(",").forEach(s => {
        let clean = s.trim().replace(/^\[\vert{}\]$/g, '').replace(/^"\vert{}"$/g, '');
        if (clean.startsWith("http")) urls.push(clean);
      });
    }
  } catch(e) {
    rawGalleryStr.split(",").forEach(s => {
      let clean = s.trim().replace(/^\[\vert{}\]$/g, '').replace(/^"\vert{}"$/g, '');
      if (clean.startsWith("http")) urls.push(clean);
    });
  }
  return urls;
}

function cleanWixUrl(wixSrc) {
  if (wixSrc.startsWith("wix:image://v1/")) return "https://static.wixstatic.com/media/" + wixSrc.replace("wix:image://v1/", "").split("/")[0];
  return wixSrc;
}

function parseGermanDate(dateVal) {
  if (!dateVal) return null;
  if (dateVal instanceof Date) { const d = new Date(dateVal.getTime()); d.setHours(12, 0, 0, 0); return d; }
  const str = dateVal.toString().trim();
  if (str.includes("-")) { const p = str.split("-"); return new Date(p[0], p[1] - 1, p[2], 12, 0, 0); }
  if (str.includes(".")) { const p = str.split("."); return new Date(p[2], p[1] - 1, p[0], 12, 0, 0); }
  return null;
}

function formatDateGerman(dateObj) {
  return Utilities.formatDate(dateObj, "Europe/Berlin", "dd.MM.yyyy");
}