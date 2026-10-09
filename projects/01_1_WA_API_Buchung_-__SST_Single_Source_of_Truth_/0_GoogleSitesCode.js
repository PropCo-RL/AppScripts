function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Gästeportal')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// MASTER-CONFIG
const MASTER_EMAIL = "info@l8street.com";
const MASTER_OTP = "b14e1dd4-5d5f"; 
const MASTER_PW = "b14e1dd4-5d5f"; 

// Hilfsfunktion: Lädt alle Apartment-Adressen aus Blatt "d"
function getApartmentDetailsMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetD = ss.getSheetByName("d");
  const aptMap = {};
  if (!sheetD) return aptMap;
  
  const dataD = sheetD.getDataRange().getValues();
  for (let i = 1; i < dataD.length; i++) {
    let linkStr = String(dataD[i][3] || "").trim();
    let aptCode = linkStr.includes("/a/") ? linkStr.split("/a/")[1].replace(/\/$/, "").trim() : linkStr;
    
    if (aptCode) {
      const name = dataD[i][2] || ""; 
      const strasse = dataD[i][6] || ""; 
      const nr = dataD[i][7] || ""; 
      const plz = dataD[i][8] || ""; 
      const stadt = dataD[i][9] || ""; 
      const etage = dataD[i][10] || ""; 
      const seite = dataD[i][11] || ""; 
      
      let fullAddress = strasse + " " + nr + ", " + plz + " " + stadt;
      let details = [];
      if (etage) details.push(etage + ". OG");
      if (seite) details.push(seite);
      if (details.length > 0) fullAddress += " (" + details.join(", ") + ")";
      
      aptMap[aptCode.toLowerCase()] = { name: name, address: fullAddress };
    }
  }
  return aptMap;
}

// 1. E-Mail prüfen & OTP generieren
function requestOTP(email) {
  const cleanEmail = String(email).trim().toLowerCase();
  if (cleanEmail === MASTER_EMAIL) {
    return { success: true, isMaster: true };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let found = false;
  
  ["RE", "VL"].forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      let email1 = String(data[i][8] || "").trim().toLowerCase();  // Spalte I
      let email2 = String(data[i][21] || "").trim().toLowerCase(); // Spalte V
      if (email1 === cleanEmail || email2.includes(cleanEmail)) found = true;
    }
  });

  if (found) {
    const otp = Math.floor(100000 + Math.random() * 900000).toString();
    const userProperties = PropertiesService.getScriptProperties();
    userProperties.setProperty("OTP_" + cleanEmail, otp);
    
    try {
      GmailApp.sendEmail(MASTER_EMAIL, "TEST-LOG: OTP für " + cleanEmail, "Der Code lautet: " + otp);
    } catch(e) {}
    
    return { success: true, isMaster: false };
  }
  return { success: false, message: "Diese E-Mail-Adresse wurde in den Buchungen nicht gefunden." };
}

// 2A. Login via Passwort
function loginWithPassword(email, password) {
  const cleanEmail = String(email).trim().toLowerCase();
  const cleanPw = String(password).trim();

  if (cleanEmail === MASTER_EMAIL && cleanPw === MASTER_PW) {
    return { found: true, isMasterAdmin: true, customerList: getAllCustomerEmails() };
  }

  const userProperties = PropertiesService.getScriptProperties();
  const storedPw = userProperties.getProperty("PW_" + cleanEmail);

  if (storedPw && storedPw === cleanPw) {
    return loadBookingsForEmail(cleanEmail);
  }

  return { found: false, message: "Ungültige E-Mail oder falsches Passwort." };
}

// 2B. Login via OTP
function verifyOTP(email, otp, targetCustomerEmail) {
  const cleanEmail = String(email).trim().toLowerCase();
  
  if (cleanEmail === MASTER_EMAIL) {
    if (String(otp).trim() !== MASTER_OTP) return { found: false, message: "Ungültiger Master-Code." };
    if (targetCustomerEmail) return loadBookingsForEmail(targetCustomerEmail.trim().toLowerCase());
    return { found: true, isMasterAdmin: true, customerList: getAllCustomerEmails() };
  }
  
  const userProperties = PropertiesService.getScriptProperties();
  const storedOTP = userProperties.getProperty("OTP_" + cleanEmail);
  if (!storedOTP || storedOTP !== String(otp).trim()) return { found: false, message: "Ungültiger oder abgelaufener Code." };
  
  userProperties.deleteProperty("OTP_" + cleanEmail);
  return loadBookingsForEmail(cleanEmail);
}

// Buchungs- & Profil-Daten laden (gruppiert nach Spalte S)
function loadBookingsForEmail(cleanEmail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aptMap = getApartmentDetailsMap();
  let bookingGroups = {};
  let customerProfile = { email: cleanEmail, name: "", firma: "", telefon: "" };

  ["RE", "VL"].forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      let email1 = String(data[i][8] || "").trim().toLowerCase();  // Spalte I
      let email2 = String(data[i][21] || "").trim().toLowerCase(); // Spalte V
      let status = String(data[i][31] || "").trim();               // Spalte AF
      
      if ((email1 === cleanEmail || email2.includes(cleanEmail)) && status !== "Storniert") {
        if (!customerProfile.name) {
          customerProfile.telefon = String(data[i][0] || "").trim(); // Spalte A
          customerProfile.firma = String(data[i][1] || "").trim();   // Spalte B
          customerProfile.name = String(data[i][2] || "").trim();    // Spalte C
        }
        
        let rawBookingNum = String(data[i][18] || "").trim(); // Spalte S
        let bookingKey = rawBookingNum || (sheetName + "_" + (i + 1));
        let rawAptCode = String(data[i][13] || "").trim();   // Spalte N
        let aptInfo = aptMap[rawAptCode.toLowerCase()] || { name: rawAptCode, address: "Adresse wird verarbeitet" };

        if (!bookingGroups[bookingKey]) {
          bookingGroups[bookingKey] = {
            id: sheetName + "_" + (i + 1),
            bookingNumber: rawBookingNum || "Buchung " + (i + 1),
            sheetName: sheetName,
            rowIndex: i + 1,
            firma: data[i][1],
            name: data[i][2],
            wohnungKuerzel: rawAptCode,
            wohnungName: aptInfo.name,
            wohnungAdresse: aptInfo.address,
            anreise: formatDate(data[i][10]),
            anreiseRaw: data[i][10],
            abreise: formatDate(data[i][11]),
            abreiseRaw: data[i][11],
            rhythmus: data[i][12] || "Wöchentlich",
            status: status || "Aktiv",
            invoices: []
          };
        }

        let rechnungId = String(data[i][19] || "").trim(); // Spalte T
        let reStart = formatDate(data[i][27]);             // Spalte AB
        let reEnde = formatDate(data[i][28]);              // Spalte AC
        
        if (rechnungId || reStart) {
          bookingGroups[bookingKey].invoices.push({
            id: rechnungId || ("RE-" + (i + 1)),
            zeitraum: (reStart && reEnde) ? (reStart + " bis " + reEnde) : "Gesamter Zeitraum",
            status: status || "Aktiv"
          });
        }
        
        if (data[i][11]) {
          bookingGroups[bookingKey].abreise = formatDate(data[i][11]);
          bookingGroups[bookingKey].abreiseRaw = data[i][11];
        }
      }
    }
  });

  let bookingsList = Object.values(bookingGroups);

  if (bookingsList.length > 0) {
    return { found: true, email: cleanEmail, profile: customerProfile, gast: customerProfile.name || customerProfile.firma, bookings: bookingsList };
  } else {
    return { found: false, message: "Keine aktiven Buchungen für " + cleanEmail + " gefunden." };
  }
}

// Profil update (Nichts in RE/VL überschreiben -> Passwort speichern & Log in Tab 'sites')
function updateProfileData(email, newName, newPhone, newPassword) {
  const cleanEmail = String(email).trim().toLowerCase();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (newPassword && newPassword.trim() !== "") {
    const userProperties = PropertiesService.getScriptProperties();
    userProperties.setProperty("PW_" + cleanEmail, newPassword.trim());
  }

  let sheetSites = ss.getSheetByName("sites");
  if (sheetSites) {
    sheetSites.appendRow([new Date(), cleanEmail, "-", "Profil-Update", "Neuer Name: " + newName + " | Neue Tel: " + newPhone]);
  }

  try {
    GmailApp.sendEmail(MASTER_EMAIL, "GÄSTEPORTAL: Profil aktualisiert", "Gast " + cleanEmail + " hat Kontaktdaten/Passwort angepasst.");
  } catch(e) {}

  return { success: true, message: "Profil & Einstellungen wurden erfolgreich gespeichert!" };
}

function getAllCustomerEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let emails = new Set();
  ["RE", "VL"].forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][8]) emails.add(String(data[i][8]).trim().toLowerCase());
    }
  });
  return Array.from(emails);
}

// 3. HAUPTLOGIK FÜR AKTIONEN (Sicheres Kopieren / Anfügen)
function submitGuestAction(email, bookingId, actionType, payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cleanEmail = String(email).trim().toLowerCase();
  
  // A) VERLÄNGERUNG -> Kopiert letzte Zeile & fügt sie am Ende von VL an
  if (actionType === 'extend') {
    const sheetVL = ss.getSheetByName("VL");
    const sheetRE = ss.getSheetByName("RE");
    
    let lastMatchingRow = null;

    [sheetVL, sheetRE].forEach(sheet => {
      if (lastMatchingRow || !sheet) return;
      const data = sheet.getDataRange().getValues();
      for (let i = data.length - 1; i >= 1; i--) {
        let e1 = String(data[i][8] || "").trim().toLowerCase();  
        let e2 = String(data[i][21] || "").trim().toLowerCase(); 
        if (e1 === cleanEmail || e2.includes(cleanEmail)) {
          lastMatchingRow = data[i];
          break;
        }
      }
    });

    if (!lastMatchingRow) {
      return { success: false, message: "Ursprüngliche Buchungszeile konnte nicht gefunden werden." };
    }

    let newRow = [...lastMatchingRow];

    // Spalten T bis X leeren (Index 19 bis 23)
    for (let col = 19; col <= 23; col++) {
      newRow[col] = "";
    }

    // Spalte L (Index 11) & Spalte M (Index 12) setzen
    const newEndDate = parseGermanDate(payload.endDate);
    newRow[11] = newEndDate; 
    newRow[12] = payload.rhythm;

    // Spalte AC (Index 28) wandert nach Spalte AB (Index 27)
    let oldAC = lastMatchingRow[28];
    let newAB_Date = oldAC ? new Date(oldAC) : new Date();
    newRow[27] = newAB_Date; 

    // Neue Spalte AC berechnen (AB + Rhythmus, max. bis Enddatum L)
    let rhythmDays = 7;
    if (payload.rhythm.includes("14")) rhythmDays = 14;
    if (payload.rhythm.includes("30") || payload.rhythm.includes("Monat")) rhythmDays = 30;

    let calculatedAC_Date = new Date(newAB_Date.getTime() + (rhythmDays * 24 * 60 * 60 * 1000));
    if (calculatedAC_Date > newEndDate) {
      calculatedAC_Date = newEndDate;
    }
    newRow[28] = calculatedAC_Date; 

    // Spalte AF (Index 31) Status mit Aktionscode
    const actionCode = "EXT-" + Math.floor(100000 + Math.random() * 900000);
    newRow[31] = "Google Sites Verlängerung [" + actionCode + "]";

    sheetVL.appendRow(newRow);

    try {
      GmailApp.sendEmail(
        MASTER_EMAIL,
        "GÄSTEPORTAL: Neue Verlängerungszeile in VL angelegt",
        "Gast: " + cleanEmail + "\nAktions-Code: " + actionCode + "\nNeues Enddatum: " + payload.endDate + "\nRhythmus: " + payload.rhythm
      );
    } catch(e) {}

    return { success: true, message: "Verlängerungsanfrage gebucht! (Ref-Code: " + actionCode + ")" };
  }

  // B) WOHNUNGSGEBERBESTÄTIGUNG -> Neue Zeile in Tab 'wgb'
  else if (actionType === 'wgb') {
    const sheetWGB = ss.getSheetByName("wgb");
    if (!sheetWGB) return { success: false, message: "Tab 'wgb' nicht gefunden." };

    let bookingID = "-";
    let apartmentCode = "-";
    
    ["RE", "VL"].forEach(sheet => {
      if (apartmentCode !== "-") return;
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        let e1 = String(data[i][8] || "").trim().toLowerCase();
        if (e1 === cleanEmail) {
          bookingID = String(data[i][18] || "-").trim();   
          apartmentCode = String(data[i][13] || "-").trim(); 
          break;
        }
      }
    });

    const moveInDate = parseGermanDate(payload.moveInDate);
    const moveOutDate = new Date(moveInDate.getTime());
    moveOutDate.setFullYear(moveOutDate.getFullYear() + 1);

    sheetWGB.appendRow([
      bookingID,
      apartmentCode,
      payload.names,
      moveInDate,
      moveOutDate,
      cleanEmail
    ]);

    try {
      GmailApp.sendEmail(MASTER_EMAIL, "GÄSTEPORTAL: WGB Beantragt", "Gast: " + cleanEmail + "\nPersonen: " + payload.names + "\nEinzug: " + payload.moveInDate);
    } catch(e) {}

    return { success: true, message: "Wohnungsgeberbestätigung wurde erfolgreich beantragt!" };
  }

  // C) STORNO & RECHNUNGSANSCHRIFT -> Neue Zeile in Tab 'sites'
  else {
    const sheetSites = ss.getSheetByName("sites");
    if (!sheetSites) return { success: false, message: "Tab 'sites' nicht gefunden." };

    let detailsText = "";
    if (actionType === 'cancel') {
      detailsText = "Typ: " + payload.type + " | Datum: " + payload.date;
    } else if (actionType === 'address') {
      detailsText = "Firma: " + payload.company + " | Adresse: " + payload.address;
    }

    sheetSites.appendRow([
      new Date(),
      cleanEmail,
      bookingId,
      actionType === 'cancel' ? "Stornierungsanfrage" : "Zweite Rechnungsanschrift",
      detailsText
    ]);

    try {
      GmailApp.sendEmail(MASTER_EMAIL, "GÄSTEPORTAL: Neue Anfrage (" + actionType + ")", "Gast: " + cleanEmail + "\nDetails:\n" + detailsText);
    } catch(e) {}

    return { success: true, message: "Anfrage wurde erfolgreich übermittelt!" };
  }
}

function parseGermanDate(dateStr) {
  if (!dateStr) return new Date();
  if (dateStr.includes("-")) { 
    const p = dateStr.split("-");
    return new Date(p[0], p[1] - 1, p[2]);
  } else if (dateStr.includes(".")) { 
    const p = dateStr.split(".");
    return new Date(p[2], p[1] - 1, p[0]);
  }
  return new Date(dateStr);
}

function formatDate(dateVal) {
  if (dateVal instanceof Date) {
    return Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "dd.MM.yyyy");
  }
  return dateVal;
}