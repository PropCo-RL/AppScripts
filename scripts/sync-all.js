const fs = require('fs');
const path = require('path');
const { google } = require('googleapis');

async function syncAllScripts() {
  try {
    // 1. Service Account Credentials aus den Umgebungsvariablen laden
    const saCredentialsJson = process.env.GAPPSCRIPT_SA_CREDENTIALS;
    if (!saCredentialsJson) {
      throw new Error("Secret GAPPSCRIPT_SA_CREDENTIALS fehlt oder ist leer!");
    }

    const credentials = JSON.parse(saCredentialsJson);

    // 2. Authentifizierung für Google Drive & Apps Script API initialisieren
    const auth = new google.auth.GoogleAuth({
      credentials,
      scopes: [
        'https://www.googleapis.com/auth/drive.readonly',
        'https://www.googleapis.com/auth/script.projects.readonly'
      ],
    });

    const drive = google.drive({ version: 'v3', auth });
    const scriptApi = google.script({ version: 'v1', auth });

    console.log("Suche alle Apps Script Projekte in Google Drive...");

    // 3. Alle Apps Script Dateien auflisten
    const driveRes = await drive.files.list({
      q: "mimeType='application/vnd.google-apps.script' and trashed=false",
      fields: 'files(id, name)',
      pageSize: 1000,
    });

    const scripts = driveRes.data.files || [];
    console.log(`${scripts.length} Skripte gefunden. Starte Download...`);

    const projectsDir = path.join(__dirname, '../projects');
    if (!fs.existsSync(projectsDir)) {
      fs.mkdirSync(projectsDir, { recursive: true });
    }

    // 4. Jedes Skript herunterladen
    for (const script of scripts) {
      const folderName = script.name.replace(/[^a-zA-Z0-9_-]/g, '_');
      const folderPath = path.join(projectsDir, folderName);

      console.log(`Lade herunter: ${script.name} (${script.id})...`);

      if (!fs.existsSync(folderPath)) {
        fs.mkdirSync(folderPath, { recursive: true });
      }

      // .clasp.json anlegen, damit Sie später bei Bedarf via Clasp pushen können
      const claspJson = { scriptId: script.id, rootDir: "." };
      fs.writeFileSync(path.join(folderPath, '.clasp.json'), JSON.stringify(claspJson, null, 2));

      try {
        // Inhalt des Apps Script Projekts via Google Apps Script API abrufen
        const content = await scriptApi.projects.getContent({ scriptId: script.id });
        const files = content.data.files || [];

        for (const file of files) {
          let ext = '.js';
          if (file.type === 'HTML') ext = '.html';
          if (file.type === 'JSON') ext = '.json';

          // appsscript.json benötigt keine doppelte Endung
          const filename = file.name === 'appsscript' && ext === '.json' 
            ? 'appsscript.json' 
            : `${file.name}${ext}`;

          const filePath = path.join(folderPath, filename);
          fs.writeFileSync(filePath, file.source || '');
        }
      } catch (err) {
        console.error(`Fehler beim Herunterladen von ${script.name}:`, err.message);
      }
    }

    console.log("Sync erfolgreich abgeschlossen!");

  } catch (error) {
    console.error("Kritischer Fehler:", error.message || error);
    process.exit(1);
  }
}

syncAllScripts();
