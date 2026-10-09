const fs = require('fs');
const path = require('path');
const { google } = require('googleapis');

async function syncAllScripts() {
  try {
    const saCredentialsJson = process.env.GAPPSCRIPT_SA_CREDENTIALS;
    if (!saCredentialsJson) {
      throw new Error("Secret GAPPSCRIPT_SA_CREDENTIALS fehlt oder ist leer!");
    }

    const credentials = JSON.parse(saCredentialsJson);

    console.log(`Starte Sync via direkter Service Account Freigabe...`);

    // Direkte Auth OHNE "subject:" (Impersonation)
    const auth = new google.auth.JWT({
      email: credentials.client_email,
      key: credentials.private_key,
      scopes: [
        'https://www.googleapis.com/auth/drive.readonly',
        'https://www.googleapis.com/auth/script.projects.readonly'
      ]
    });

    const drive = google.drive({ version: 'v3', auth });
    const scriptApi = google.script({ version: 'v1', auth });

    console.log("Suche alle freigegebenen Apps Script Projekte & Sheets...");

    const driveRes = await drive.files.list({
      q: "(mimeType='application/vnd.google-apps.script' or mimeType='application/vnd.google-apps.spreadsheet') and trashed=false",
      fields: 'files(id, name, mimeType)',
      pageSize: 1000,
    });

    const filesFound = driveRes.data.files || [];
    console.log(`${filesFound.length} Dateien (Scripts/Sheets) gefunden. Starte Download...`);

    const projectsDir = path.join(__dirname, '../projects');
    if (!fs.existsSync(projectsDir)) {
      fs.mkdirSync(projectsDir, { recursive: true });
    }

    let downloadedCount = 0;

    for (const file of filesFound) {
      let scriptId = file.id;

      try {
        const content = await scriptApi.projects.getContent({ scriptId: scriptId });
        const scriptFiles = content.data.files || [];

        if (scriptFiles.length === 0) continue;

        const folderName = file.name.replace(/[^a-zA-Z0-9_-]/g, '_');
        const folderPath = path.join(projectsDir, folderName);

        console.log(`Lade herunter: ${file.name} (${file.id})...`);

        if (!fs.existsSync(folderPath)) {
          fs.mkdirSync(folderPath, { recursive: true });
        }

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
        // Ignoriert Sheets ohne Skript-Inhalt
      }
    }

    console.log(`Sync abgeschlossen! ${downloadedCount} Apps Script Projekte erfolgreich heruntergeladen.`);

  } catch (error) {
    console.error("Kritischer Fehler:", error.message || error);
    process.exit(1);
  }
}

syncAllScripts();
