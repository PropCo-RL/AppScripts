// =================================================================
// MONTEURWOHNUNGEN - GEMINI MATCHING & VERTRIEBS-VERSAND
// =================================================================

function processApartmentRequestsWithGemini() {
  const SHEET1_ID = '1VB253Cq_hr4YfFrWDXNG-pgv8qOOKAzb7jN_KYciEHw';
  const SHEET2_ID = '1Aq9doxAkEw435bRCpXafayRwdNeaP9LVMU1YwWG7z70';
  
  const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('Gemini-Key');
  
  let s1, s2_m, s2_cal;
  try {
    s1 = SpreadsheetApp.openById(SHEET1_ID).getSheetByName('DMZ');
    s2_m = SpreadsheetApp.openById(SHEET2_ID).getSheetByName('M');
    s2_cal = SpreadsheetApp.openById(SHEET2_ID).getSheetByName('cal');
  } catch (e) {
    Logger.log("FEHLER beim Öffnen der Tabellen: " + e.toString());
    return;
  }

  if (!s1 || !s2_m || !s2_cal) {
    Logger.log("FEHLER: Ein oder mehrere Tabellenblätter nicht gefunden.");
    return;
  }

  const dmzData = s1.getDataRange().getValues();
  const mData = s2_m.getDataRange().getValues();
  const calData = s2_cal.getDataRange().getValues();

  Logger.log("--- GEMINI MATCHING SKRIPT START ---");

  // 1. Alle gebuchten Zeiträume aus "cal" einlesen
  let bookings = {};
  for (let i = 1; i < calData.length; i++) {
    let aptName = String(calData[i][0]).trim();
    let start = parseIsoDate(calData[i][1]);
    let end = parseIsoDate(calData[i][2]);
    if (aptName && start !== null && end !== null) {
      if (!bookings[aptName]) bookings[aptName] = [];
      bookings[aptName].push({ start: start, end: end });
    }
  }

  // 2. Stammdaten aus "M" einlesen (inkl. Ausschluss-Filter für Spalte AJ)
  let apartments = [];
  for (let i = 1; i < mData.length; i++) {
    let excludeFlag = String(mData[i][35] || "").trim().toLowerCase(); // Spalte AJ (Index 35)

    // WICHTIG: Wenn in Spalte AJ ein "x" steht, wird dieses Apartment komplett ignoriert!
    if (excludeFlag === "x") {
      continue;
    }

    let aptName = String(mData[i][0]).trim();
    let slug = String(mData[i][6] || "").trim();         // Spalte G
    let city = String(mData[i][9] || "").trim();         // Spalte J
    let zip = String(mData[i][12] || "").trim();         // Spalte M
    let street = String(mData[i][13] || "").trim();      // Spalte N
    let bedrooms = parseInt(mData[i][14]) || 0;         // Spalte O
    let beds = parseInt(mData[i][15]) || 0;             // Spalte P
    let pricePerNight = parseFloat(mData[i][17]) || 0;  // Spalte R
    let minPersons = parseInt(mData[i][34]) || 1;       // Spalte AI (Index 34)

    if (aptName) {
      apartments.push({
        name: aptName,
        slug: slug,
        city: city,
        zip: zip,
        street: street,
        fullAddress: `${street}, ${zip} ${city}`,
        bedrooms: bedrooms,
        beds: beds,
        pricePerNight: pricePerNight,
        minPersons: minPersons
      });
    }
  }

  // 3. Anfragen verarbeiten
  for (let i = 1; i < dmzData.length; i++) {
    const rowNum = i + 1;
    const currentStatus = String(dmzData[i][17] || "").trim(); // Spalte R (Status)

    // Nur ungestempelte Zeilen ausführen
    if (currentStatus !== "") {
      continue;
    }

    let rawLead = {
      portal: dmzData[i][0],
      kundenNummer: dmzData[i][1],
      gastName: dmzData[i][2],
      telefonnummer: dmzData[i][3],
      zeitraum: dmzData[i][4],
      personen: dmzData[i][5],
      unterkunftTitel: dmzData[i][6],
      link: dmzData[i][7],
      nachricht: dmzData[i][8],
      unterkunftAdresse: dmzData[i][9],
      email: dmzData[i][10],
      erfasstAm: dmzData[i][11]
    };

    if (!rawLead.zeitraum || !rawLead.unterkunftAdresse) {
      s1.getRange(rowNum, 17).setValue("Fehler: Unvollständige Anfrage");
      s1.getRange(rowNum, 18).setValue("Übersprungen");
      continue;
    }

    Logger.log("Verarbeite Zeile " + rowNum + " für " + rawLead.gastName);

    // A. KI-Analyse des Freitextes
    let parsedRequest = extractDataWithGemini(rawLead, GEMINI_API_KEY);
    
    // B. Zeiträume prüfen
    let dates = parseGermanDateRange(parsedRequest.correctedZeitraum || rawLead.zeitraum);
    if (!dates) {
      s1.getRange(rowNum, 17).setValue("Fehler: Datum nicht lesbar");
      s1.getRange(rowNum, 18).setValue("Fehler");
      continue;
    }

    let nightCount = Math.round((dates.end - dates.start) / (1000 * 60 * 60 * 24));
    if (nightCount < 5) {
      s1.getRange(rowNum, 17).setValue("Unter Mindestaufenthalt (" + nightCount + " Nächte)");
      s1.getRange(rowNum, 18).setValue("Abgelehnt");
      continue;
    }

    let reqPersonen = parseInt(parsedRequest.correctedPersonen) || parseInt(rawLead.personen) || 1;

    // C1. ERSTE WAHL: Wohnungen, die Kapazität + Mindestpersonen erreichbar erfüllen
    let candidates = apartments.filter(apt => {
      if (apt.beds < reqPersonen || reqPersonen < apt.minPersons) return false;

      // Belegungs-Check
      if (bookings[apt.name]) {
        for (let b of bookings[apt.name]) {
          if (dates.start < b.end && dates.end > b.start) {
            return false;
          }
        }
      }
      return true;
    });

    let isAlternativeDate = false;
    let nextAvailableDateStr = "";

    // C2. FALLBACK 1: Wunschdatum belegt? Suche nächstes freies Datum für aktivierte Wohnungen
    if (candidates.length === 0) {
      let altCandidates = apartments.filter(apt => apt.beds >= reqPersonen && reqPersonen >= apt.minPersons);
      
      for (let apt of altCandidates) {
        if (bookings[apt.name]) {
          let latestEnd = dates.start;
          for (let b of bookings[apt.name]) {
            if (dates.start < b.end && dates.end > b.start) {
              if (b.end > latestEnd) latestEnd = b.end;
            }
          }
          if (latestEnd > dates.start) {
            candidates.push(apt);
            isAlternativeDate = true;
            nextAvailableDateStr = formatDateUtc(latestEnd);
            break; 
          }
        }
      }
    }

    // D. Gemini entscheidet über Standort & Fahrzeit
    let bestMatch = selectBestApartmentWithGemini(parsedRequest, candidates.length > 0 ? candidates : apartments, GEMINI_API_KEY);

    if (candidates.length > 0 && bestMatch && bestMatch.selectedAptName && bestMatch.estimatedDriveTimeMinutes <= 40) {
      let matchedApt = candidates.find(c => c.name === bestMatch.selectedAptName) || candidates[0];
      let totalPrice = matchedApt.pricePerNight * nightCount;
      let cleanSlug = matchedApt.slug.startsWith('/') ? matchedApt.slug : '/' + matchedApt.slug;
      let link = "https://a.l8street.com" + cleanSlug;
      if (!link.endsWith('/')) link += '/';

      let resultText = "Empfehlung: " + matchedApt.name + "\n" +
                       "Link: " + link + "\n" +
                       "Preis: " + totalPrice + " € (" + nightCount + " Nächte x " + matchedApt.pricePerNight + " €)\n" +
                       "Fahrzeit: ca. " + bestMatch.estimatedDriveTimeMinutes + " Min." +
                       (isAlternativeDate ? "\n(Alternativ-Datum ab " + nextAvailableDateStr + ")" : "");

      s1.getRange(rowNum, 17).setValue(resultText);

      // E-Mail versenden
      sendSalesEmail({
        req: parsedRequest,
        apt: matchedApt,
        totalPrice: totalPrice,
        nights: nightCount,
        link: link,
        driveTimeMinutes: bestMatch.estimatedDriveTimeMinutes,
        isAlternativeDate: isAlternativeDate,
        nextAvailableDateStr: nextAvailableDateStr
      });

      s1.getRange(rowNum, 18).setValue("Versendet");

    } else {
      // C3. FALLBACK 2: Gar kein konkretes Apartment unter 40 Min frei -> Stadt-Link senden
      let nearestCity = (bestMatch && bestMatch.nearestCityName) ? bestMatch.nearestCityName.toLowerCase().replace(/[^a-z0-9äöüß-]/g, "") : "kaiserslautern";
      let cityLink = "https://a.l8street.com/" + nearestCity + "/";

      s1.getRange(rowNum, 17).setValue("Fallback Stadt-Link: " + cityLink);

      sendCityFallbackEmail(parsedRequest, nearestCity, cityLink);
      s1.getRange(rowNum, 18).setValue("Versendet (Stadt-Link)");
    }
  }

  Logger.log("--- GEMINI MATCHING SKRIPT BEENDET ---");
}

// =================================================================
// GEMINI INTERACTIONS API
// =================================================================

function extractDataWithGemini(rawLead, apiKey) {
  const promptText = `
Du bist ein Assistent für die Analyse von Anfragen für Monteurwohnungen.
Analysiere die folgende Kundenanfrage im JSON-Format.
Achte BESONDERS auf das Feld "nachricht". Falls der Kunde dort abweichende Angaben bezüglich Personenanzahl oder Zeitraum macht, korrigiere die Werte!

Eingangs-JSON:
"""${JSON.stringify(rawLead)}"""

ANTWORTE AUSSCHLIESSLICH ALS EIN EINZELNES JSON-OBJEKT IN FOLGENDEM FORMAT:
{
  "correctedPersonen": 3,
  "correctedZeitraum": "11.10.2026 - 23.10.2026",
  "destinationAddress": "Straße, PLZ Stadt",
  "gastName": "Name",
  "email": "E-Mail-Adresse"
}`;

  const result = executeGeminiInteraction(promptText, apiKey);

  return result || {
    correctedPersonen: rawLead.personen,
    correctedZeitraum: rawLead.zeitraum,
    destinationAddress: rawLead.unterkunftAdresse,
    gastName: rawLead.gastName,
    email: rawLead.email
  };
}

function selectBestApartmentWithGemini(parsedRequest, candidates, apiKey) {
  const promptText = `
Du bist ein Experte für Routen- und Fahrzeitplanung für Monteurunterkünfte.
Wähle aus den folgenden Kandidaten-Wohnungen die Option mit der KÜRZESTEN AUTOFARZEIT zur Zieladresse aus.

Zieladresse des Kunden: "${parsedRequest.destinationAddress}"
Benötigte Personen: ${parsedRequest.correctedPersonen}

FAHRZEIT-REGELN:
1. Wähle bevorzugt eine Unterkunft mit einer Fahrzeit von MAXIMAL 30 MINUTEN.
2. Wenn KEINE Unterkunft unter 30 Minuten erreichbar ist, wähle eine Unterkunft mit bis zu MAXIMAL 40 MINUTEN.
3. Gib zusätzlich den Namen der nächstgelegenen größeren Stadt/Region ("nearestCityName") für einen Fallback-Link an.

Verfügbare Kandidaten:
"""${JSON.stringify(candidates)}"""

ANTWORTE AUSSCHLIESSLICH ALS EIN EINZELNES JSON-OBJEKT IN FOLGENDEM FORMAT:
{
  "selectedAptName": "Exakter Name der gewählten Wohnung",
  "estimatedDriveTimeMinutes": 22,
  "nearestCityName": "kaiserslautern",
  "reasoning": "Kurze Begründung mit Angabe der geschätzten Fahrzeit"
}`;

  const result = executeGeminiInteraction(promptText, apiKey);

  return result || { 
    selectedAptName: candidates[0] ? candidates[0].name : "", 
    estimatedDriveTimeMinutes: 25, 
    nearestCityName: "kaiserslautern",
    reasoning: "Standardauswahl" 
  };
}

function executeGeminiInteraction(inputText, apiKey) {
  if (!apiKey) {
    Logger.log("FEHLER: Kein Gemini API Key in den Script Properties gefunden ('Gemini-Key').");
    return null;
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/interactions`;
  
  const payload = {
    "model": "gemini-3.8-flash",
    "input": inputText,
    "store": false
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: { "x-goog-api-key": apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) {
      Logger.log("Gemini API Fehler Code " + response.getResponseCode() + ": " + response.getContentText());
      return null;
    }
    const json = JSON.parse(response.getContentText());
    
    let rawText = "";
    if (json.output_text) {
      rawText = json.output_text.trim();
    } else if (json.steps && json.steps.length > 0) {
      const lastStep = json.steps[json.steps.length - 1];
      if (lastStep.content && lastStep.content[0] && lastStep.content[0].text) {
        rawText = lastStep.content[0].text.trim();
      }
    }

    const jsonMatch = rawText.match(/\[[\s\S]*\]|\{[\s\S]*\}/);
    if (jsonMatch) rawText = jsonMatch[0];
    return safeParseJSON(rawText);

  } catch (e) {
    Logger.log("Fehler bei Gemini API Aufruf: " + e.toString());
  }
  return null;
}

function safeParseJSON(rawText) {
  try { 
    return JSON.parse(rawText); 
  } catch (e) {
    try {
      let cleanText = rawText
        .replace(/[\u0000-\u001F\u007F-\u009F]/g, "")
        .replace(/\r?\n/g, " ")
        .replace(/",\s*}/g, '"}')
        .replace(/",\s*]/g, '"]');
      return JSON.parse(cleanText);
    } catch (e2) { 
      return null; 
    }
  }
}

// =================================================================
// VERTRIEBS-EMAILS AN INFO@L8STREET.COM WITH FOOTER & WHATSAPP
// =================================================================

function sendSalesEmail(data) {
  const targetEmail = "info@l8street.com";
  const subject = "Ihr Angebot für Monteurwohnung - " + data.apt.name + " (" + data.req.correctedZeitraum + ")";
  
  let dateTextNotice = "";
  if (data.isAlternativeDate) {
    dateTextNotice = "Hinweis zum Wunschdatum: Zu Ihrem genauen Wunschstart ist die Wohnung derzeit noch belegt. Ab dem " + data.nextAvailableDateStr + " ist diese Wohnung wieder vollständig für Sie verfügbar!\n\n";
  }

  const body = "Guten Tag " + data.req.gastName + ",\n\n" +
               "vielen Dank für Ihre Anfrage. Gerne bieten wir Ihnen für Ihre Mitarbeiter unsere passende Monteurwohnung an:\n\n" +
               dateTextNotice +
               "🏠 UNTERKUNFT: " + data.apt.name + "\n" +
               "📍 ADRESSE: " + data.apt.street + ", " + data.apt.zip + " " + data.apt.city + "\n" +
               "🚗 FAHRZEIT ZUM EINSATZORT: ca. " + data.driveTimeMinutes + " Minuten\n" +
               "🛏️ AUSSTATTUNG: " + data.apt.beds + " Einzelbetten in " + data.apt.bedrooms + " Schlafzimmer(n)\n\n" +
               "💶 IHRE KONDITIONEN:\n" +
               "- Zeitraum: " + data.req.correctedZeitraum + " (" + data.nights + " Nächte)\n" +
               "- Belegung: " + data.req.correctedPersonen + " Personen\n" +
               "- Preis pro Nacht: " + data.apt.pricePerNight.toFixed(2) + " € zzgl. USt.\n" +
               "- GESAMTPREIS: " + data.totalPrice.toFixed(2) + " € zzgl. USt.\n\n" +
               "📸 BILDER & DETAILS ZUR WOHNUNG:\n" +
               "Unter folgendem Link können Sie sich die Wohnung direkt ansehen:\n" +
               data.link + "\n\n" +
               "⚡ VERBINDLICHE BUCHUNG:\n" +
               "Soll das Angebot für Sie reserviert werden? Antworten Sie einfach kurz auf diese E-Mail oder schreiben Sie uns per WhatsApp.\n\n" +
               "--------------------------------------------------\n" +
               "Mit freundlichen Grüßen\n" +
               "Ihr Vertriebsteam von L8 Street\n\n" +
               "L8 Street GmbH\n" +
               "E-Mail: support@L8Street.com\n" +
               "WhatsApp / Mobil: +49 176 8480 1295\n" +
               "Web: www.L8Street.com\n\n" +
               "Amtsgericht Mannheim: HRB 728552\n" +
               "USt-ID: DE314603427\n" +
               "AGB: www.L8Street.com/agb\n" +
               "Datenschutz: www.L8Street.com/datenschutz\n" +
               "--------------------------------------------------";

  GmailApp.sendEmail(targetEmail, subject, body);
  addAngebotLabel(subject);
}

function sendCityFallbackEmail(req, city, cityLink) {
  const targetEmail = "info@l8street.com";
  const subject = "Übersicht verfügbarer Monteurwohnungen für " + req.destinationAddress;
  
  const body = "Guten Tag " + req.gastName + ",\n\n" +
               "vielen Dank für Ihre Anfrage.\n\n" +
               "Für Ihren genauen Einsatzort ist zum angefragten Zeitraum leider kein einzelnes Apartment direkt frei. Wir haben jedoch mehrere verfügbare Monteurwohnungen in der nahegelegenen Region für Sie bereitstehen.\n\n" +
               "Unter folgendem Link finden Sie alle verfügbaren Apartments in der Region sowie Fotos und Ausstattungsdetails:\n" +
               cityLink + "\n\n" +
               "Bitte lassen Sie uns wissen, welche der Unterkünfte für Ihr Team infrage kommt, damit wir das Angebot umgehend für Sie fixieren können.\n\n" +
               "--------------------------------------------------\n" +
               "Mit freundlichen Grüßen\n" +
               "Ihr Vertriebsteam von L8 Street\n\n" +
               "L8 Street GmbH\n" +
               "E-Mail: support@L8Street.com\n" +
               "WhatsApp / Mobil: +49 176 8480 1295\n" +
               "Web: www.L8Street.com\n\n" +
               "Amtsgericht Mannheim: HRB 728552\n" +
               "USt-ID: DE314603427\n" +
               "AGB: www.L8Street.com/agb\n" +
               "Datenschutz: www.L8Street.com/datenschutz\n" +
               "--------------------------------------------------";

  GmailApp.sendEmail(targetEmail, subject, body);
  addAngebotLabel(subject);
}

function addAngebotLabel(subject) {
  try {
    let label = GmailApp.getUserLabelByName("Angebot");
    if (!label) {
      label = GmailApp.createLabel("Angebot");
    }
    let threads = GmailApp.search("subject:\"" + subject + "\"");
    if (threads.length > 0) {
      threads[0].addLabel(label);
    }
  } catch (e) {
    Logger.log("Fehler beim Hinzufügen des Gmail-Labels: " + e.toString());
  }
}

// =================================================================
// HILFSFUNKTIONEN DATUM
// =================================================================

function parseGermanDateRange(dateStr) {
  try {
    let parts = String(dateStr).split('-');
    if (parts.length !== 2) return null;
    let d1 = parts[0].trim().split('.');
    let d2 = parts[1].trim().split('.');
    let start = Date.UTC(d1[2], d1[1] - 1, d1[0]);
    let end = Date.UTC(d2[2], d2[1] - 1, d2[0]);
    return { start: start, end: end };
  } catch (e) {
    return null;
  }
}

function parseIsoDate(isoStr) {
  try {
    if (!isoStr) return null;
    if (isoStr instanceof Date) {
      return Date.UTC(isoStr.getFullYear(), isoStr.getMonth(), isoStr.getDate());
    }
    let parts = String(isoStr).trim().split('-');
    if (parts.length !== 3) return null;
    return Date.UTC(parts[0], parts[1] - 1, parts[2]);
  } catch (e) {
    return null;
  }
}

function formatDateUtc(timestamp) {
  let d = new Date(timestamp);
  let day = ("0" + d.getUTCDate()).slice(-2);
  let month = ("0" + (d.getUTCMonth() + 1)).slice(-2);
  let year = d.getUTCFullYear();
  return `${day}.${month}.${year}`;
}