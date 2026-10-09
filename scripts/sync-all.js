const fs = require('fs');
const path = require('path');
const { google } = require('googleapis');

async function syncAllScripts() {
  try {
    // 1. Service Account JSON aus den GitHub Secrets laden
    const saCredentialsJson = process.env.GAPPSCRIPT_SA_CREDENTIALS;
    if (!saCredentialsJson) {
      throw new Error("Secret GAPPSCRIPT_SA_CREDENTIALS fehlt oder ist leer!");
    }

    const credentials = JSON.parse(saCredentialsJson);

    // 2. Domain-Wide Delegation für Ihren Haupt-Benutzer initialisieren
    const USER_TO_IMPERSONATE = process.env.WORKSPACE_USER_EMAIL || 'support@l8street.com';

    const auth = new google.auth.JWT({
      email: credentials.client_email,
      key: credentials.private_key,
      scopes: [
        'https://www.googleapis.com/auth/drive.readonly',
        'https://www.googleapis.com/auth/script.projects.readonly'
      ],
      subject: USER_TO_IMPERSONATE // Ausgabe als Ihr Workspace-Benutzer
    });

    const drive = google.drive({ version: 'v3', auth });
    const scriptApi = google.script({ version: 'v1', auth });

    console.log(`Suche Apps Script Projekte & Sheets für Benutzer: ${USER_TO_IMPERSONATE}...`);

    // 3. Suche nach freistehenden Scripts UND Google Sheets
    const driveRes = await drive.files.list({
      q: "(mimeType='application/vnd.google-apps.script' or mimeType='application/vnd.google-apps.spreadsheet') and trashed=false",
      fields: 'files(id, name, mimeType)',
      pageSize: 1000,
    });

    const filesFound = driveRes.data.files || [];
    console.log(`${filesFound.length} Dateien (Scripts/Sheets) gefunden. Starte Verarbeitungs-Check...`);

    const projectsDir = path.join(__dirname, '../projects');
    if (!fs.existsSync(projectsDir)) {
      fs.mkdirSync(projectsDir, { recursive: true });
    }

    let downloadedCount = 0;

    for (const file of filesFound) {
      let scriptId = file.id;

      try {
        // Versuch, Script-Inhalt von der Google Apps Script API zu laden
        const content = await scriptApi.projects.getContent({ scriptId: scriptId });
        const scriptFiles = content.data.files || [];

        if (scriptFiles.length === 0) continue;

        const folderName = file.name.replace(/[^a-zA-Z0-9_-]/g, '_');
        const folderPath = path.join(projectsDir, folderName);

        console.log(`Lade herunter: ${file.name} (${file.id})...`);

        if (!fs.existsSync(folderPath)) {
          fs.mkdirSync(folderPath, { recursive: true });
        }

        // .clasp.json erstellen
        const claspJson = { scriptId: file.id, rootDir: "." };
        fs.writeFileSync(path.join(folderPath, '.clasp.json'), JSON.stringify(claspJson, null, 2));

        for (const sFile of scriptFiles) {
          let ext = '.js';
          if (sFile.type === 'HTML') ext = '.html';
          if (sFile.type === 'JSON') ext = '.json';

          const filename = (sFile.name === 'appsscript' && ext === '.json') 
            ? 'appsscript.json' 
            : `${sFile.name}${ext}`;

          const filePath = path.join(folderPath, filename);
          fs.writeFileSync(filePath, sFile.source || '');
        }

        downloadedCount++;
      } catch (err) {
        // Stille Ignorierung bei reinen Sheets ohne Code oder fehlenden Berechtigungen
      }
    }

    console.log(`Sync abgeschlossen! ${downloadedCount} Apps Script Projekte erfolgreich heruntergeladen.`);

  } catch (error) {
    console.error("Kritischer Fehler:", error.message || error);
    process.exit(1);
  }
}

syncAllScripts();
