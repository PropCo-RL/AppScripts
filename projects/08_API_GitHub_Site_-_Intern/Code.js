// SPALTEN-INDEXE:
// Spalte I (Index 8)  = E-Mail
// Spalte B (Index 1)  = Firma / Gast
// Spalte N (Index 13) = Apartment Kürzel (z.B. cs1, kf2)
// Spalte K (Index 10) = Anreise
// Spalte L (Index 11) = Abreise
// Spalte S (Index 18) = Lodgify Buchungs-ID
// Spalte T (Index 19) = Lexoffice Rechnungs-ID
// Spalte AF (Index 31)= Status (Bezahlt / Offen / Storniert)
// Spalte AK (Index 36)= Rechnungs-PDF Link (Google Drive)

function doGet(e) {
  const action = e.parameter.action;
  
  if (action === "requestOTP") {
    return jsonResponse(requestOTP(e.parameter.email));
  }
  
  if (action === "verifyOTP") {
    return jsonResponse(verifyOTP(e.parameter.email, e.parameter.otp, e.parameter.filterGuest));
  }

  return jsonResponse({ success: false, message: "Ungültige Aktion" });
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === "submitGuestAction") {
      const res = submitGuestAction(data.email, data.actionType, data.payload);
      return jsonResponse(res);
    }
  } catch (err) {
    return jsonResponse({ success: false, message: err.toString() });
  }
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getApartmentDetailsMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetD = ss.getSheetByName("d");
  if (!sheetD) return {};

  const lastRow = sheetD.getLastRow();
  if (lastRow < 2) return {};

  const data = sheetD.getRange(2, 1, lastRow - 1, 19).getValues();
  const map = {};

  data.forEach(row => {
    const name = String(row[2] || "").trim();
    const link = String(row[3] || "").trim();
    const strasse = String(row[7] || "").trim();
    const nr = String(row[8] || "").trim();
    const plz = String(row[9] || "").trim();
    const stadt = String(row[10] || "").trim();
    
    const etage = String(row[11] || "").trim();
    const seite = String(row[12] || "").trim();
    const checkinInfo = String(row[18] || "").trim();

    if (!link) return;
    const match = link.match(/\/a\/([^\/\?]+)/i);
    if (!match || !match[1]) return;

    const code = match[1].toLowerCase().trim();
    const address = [strasse + " " + nr, plz + " " + stadt].filter(x => x.trim().length > 1).join(", ");

    map[code] = {
      title: name || ("Apartment " + code.toUpperCase()),
      address: address || "Adresse auf Anfrage",
      etage: etage || "-",
      seite: seite || "-",
      checkinInfo: checkinInfo || "Check-in Informationen erhalten Sie vor Anreise."
    };
  });

  return map;
}

function getTargetEmail(inputVal) {
  const cleanInput = String(inputVal || "").trim().toLowerCase();
  if (!cleanInput) return null;

  if (cleanInput === "info@l8street.com") {
    return { email: "info@l8street.com", gast: "L8 Street Test-Admin", isAdmin: true };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToSearch = ["RE", "VL"];

  for (let s = 0; s < sheetsToSearch.length; s++) {
    const sheet = ss.getSheetByName(sheetsToSearch[s]);
    if (!sheet) continue;
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) continue;

    // Nur bis Spalte T einlesen für maximalen Speed beim Mail-Check
    const data = sheet.getRange(2, 1, lastRow - 1, 20).getValues();

    for (let i = 0; i < data.length; i++) {
      const rowEmail      = String(data[i][8] || "").trim().toLowerCase();
      const rowLodgifyId  = String(data[i][18] || "").trim().toLowerCase();
      const rowRechnungId = String(data[i][19] || "").trim().toLowerCase();

      if (rowEmail === cleanInput || rowLodgifyId === cleanInput || rowRechnungId === cleanInput) {
        return { email: data[i][8] || cleanInput, gast: data[i][1] || "Gast", isAdmin: false };
      }
    }
  }
  return null;
}

function requestOTP(inputVal) {
  const target = getTargetEmail(inputVal);
  if (!target || !target.email) {
    return { success: false, message: "Keine Buchung mit dieser E-Mail, Buchungs- oder Rechnungs-ID gefunden." };
  }

  const otp = Math.floor(100000 + Math.random() * 900000).toString();
  PropertiesService.getScriptProperties().setProperty("OTP_" + target.email.toLowerCase(), otp);

  GmailApp.sendEmail(
    target.email,
    "Ihr Sicherheitscode für das L8 Street Portal",
    "Guten Tag " + target.gast + ",\n\nIhr Sicherheitscode lautet: " + otp,
    {
      htmlBody: `
        <div style="font-family: Arial, sans-serif; color: #333; max-width: 500px; padding: 20px; border: 1px solid #e5e7eb; border-radius: 8px;">
          <h3 style="color: #111827; margin-top: 0;">L8 Street Kundenportal</h3>
          <p>Guten Tag <strong>${target.gast}</strong>,</p>
          <p>Ihr Sicherheitscode für die Anmeldung im Gästeportal lautet:</p>
          <div style="background: #f3f4f6; padding: 12px; font-size: 24px; font-weight: bold; letter-spacing: 4px; text-align: center; border-radius: 6px; margin: 15px 0;">
            ${otp}
          </div>
          <p style="font-size: 0.9em; color: #6b7280;">Dieser Code ist für Ihren aktuellen Login gültig.</p>
          <hr style="border: 0; border-top: 1px solid #e5e7eb; margin: 20px 0;">
          <p style="font-size: 0.85em; color: #9ca3af; margin: 0;">L8 Street – Monteurwohnungen & Apartments<br><a href="https://me.l8street.com" style="color: #2563eb;">me.l8street.com</a></p>
        </div>
      `
    }
  );

  return { success: true };
}

function verifyOTP(inputVal, otp, filterGuest) {
  const target = getTargetEmail(inputVal);
  if (!target) return { found: false, message: "Buchung nicht gefunden." };

  const storedOTP = PropertiesService.getScriptProperties().getProperty("OTP_" + target.email.toLowerCase());

  if (storedOTP && storedOTP === String(otp).trim()) {
    PropertiesService.getScriptProperties().deleteProperty("OTP_" + target.email.toLowerCase());
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetsToSearch = ["RE", "VL"];
    const aptMap = getApartmentDetailsMap();
    
    const cleanEmail = target.email.toLowerCase();
    const cleanFilter = String(filterGuest || "").trim().toLowerCase();

    const groupedBookings = {};

    for (let s = 0; s < sheetsToSearch.length; s++) {
      const sheet = ss.getSheetByName(sheetsToSearch[s]);
      if (!sheet) continue;
      
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) continue;

      const data = sheet.getRange(2, 1, lastRow - 1, 37).getValues();

      for (let i = 0; i < data.length; i++) {
        const rowEmail      = String(data[i][8] || "").trim().toLowerCase();
        const rowLodgifyId  = String(data[i][18] || "").trim();
        const rowRechnungId = String(data[i][19] || "").trim();

        let isMatch = false;

        if (cleanEmail === "info@l8street.com") {
          if (cleanFilter) {
            isMatch = (rowEmail.includes(cleanFilter) || rowLodgifyId.toLowerCase().includes(cleanFilter) || rowRechnungId.toLowerCase().includes(cleanFilter));
          } else {
            isMatch = true;
          }
        } else {
          isMatch = (rowEmail === cleanEmail);
        }

        if (isMatch) {
          const key = rowLodgifyId || ("SINGLE_" + s + "_" + i);
          const aptCode = String(data[i][13] || "").toLowerCase().trim();
          
          // Sichere Fallbacks gegen undefined!
          const aptDetails = aptMap[aptCode] || { 
            title: aptCode ? ("Apartment " + aptCode.toUpperCase()) : "Wohnung", 
            address: "Adresse auf Anfrage", 
            etage: "-", 
            seite: "-", 
            checkinInfo: "Code wird nach Zahlungsprüfung angezeigt." 
          };

          const isBookingCom = rowLodgifyId.toUpperCase().startsWith("B");
          const statusBezahlt = String(data[i][31] || "").toLowerCase().includes("bezahlt");

          if (!groupedBookings[key]) {
            groupedBookings[key] = {
              lodgifyId: rowLodgifyId || "-",
              isBookingCom: isBookingCom,
              gast: data[i][1] || "Gast",
              wohnungTitle: aptDetails.title,
              wohnungAddress: aptDetails.address,
              etage: aptDetails.etage,
              seite: aptDetails.seite,
              checkinInfo: aptDetails.checkinInfo,
              aptCode: aptCode,
              anreise: formatDate(data[i][10]),
              anreiseRaw: data[i][10],
              abreise: formatDate(data[i][11]),
              abreiseRaw: data[i][11],
              status: data[i][31] || "Aktiv",
              isPaid: statusBezahlt,
              invoices: []
            };
          } else {
            groupedBookings[key].abreise = formatDate(data[i][11]);
            groupedBookings[key].abreiseRaw = data[i][11];
            if (statusBezahlt) groupedBookings[key].isPaid = true;
          }

          // Nur eintragen wenn vorhanden, sonst "-" als Fallback
          groupedBookings[key].invoices.push({
            rechnungId: rowRechnungId || "-",
            pdfUrl: data[i][36] || ""
          });
        }
      }
    }

    return {
      found: true,
      email: target.email,
      gast: target.gast,
      isAdmin: (cleanEmail === "info@l8street.com"),
      bookings: Object.values(groupedBookings)
    };
  } else {
    return { found: false, message: "Ungültiger oder abgelaufener Sicherheitscode." };
  }
}

function submitGuestAction(email, actionType, payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let targetSheet = ss.getSheetByName("sites") || ss.getSheetByName("wgb");

  if (!targetSheet) {
    targetSheet = ss.insertSheet("Portal_Aktionen");
  }

  targetSheet.appendRow([
    new Date(),
    email,
    actionType,
    JSON.stringify(payload)
  ]);

  return { success: true, message: "Ihre Anfrage wurde erfolgreich übermittelt!" };
}

function formatDate(dateVal) {
  if (dateVal instanceof Date) {
    return Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "dd.MM.yyyy");
  }
  return dateVal || "-";
}