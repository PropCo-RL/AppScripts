const fs = require('fs');
const { execSync } = require('child_process');
const { google } = require('googleapis');

async function syncScripts() {
  try {
    const configPath = `${process.env.HOME}/.clasprc.json`;
    if (!fs.existsSync(configPath)) {
      throw new Error(`Konnte .clasprc.json nicht unter ${configPath} finden.`);
    }

    const claspConfig = JSON.parse(fs.readFileSync(configPath, 'utf8'));

    // Verschiedene Formate von .clasprc.json abfangen
    const token = claspConfig.token || claspConfig;
    const clientSettings = claspConfig.oauth2ClientSettings || {};
    
    const clientId = clientSettings.clientId || clientSettings.client_id || process.env.CLIENT_ID;
    const clientSecret = clientSettings.clientSecret || clientSettings.client_secret || process.env.CLIENT_SECRET;

    const oAuth2Client = new google.auth.OAuth2(clientId, clientSecret);
    oAuth2Client.setCredentials(token);

    const drive = google.drive({ version: 'v3', auth: oAuth2Client });

    console.log("Suche nach allen Apps Script Dateien in Google Drive...");

    const response = await drive.files.list({
      q: "mimeType='application/vnd.google-apps.script' and trashed=false",
      fields: 'files(id, name)',
      pageSize: 1000
    });

    const scripts = response.data.files || [];
    console.log(`${scripts.length} Skripte gefunden. Starte Sync...`);

    if (!fs.existsSync('projects')) fs.mkdirSync('projects');

    for (const script of scripts) {
      const folderName = script.name.replace(/[^a-zA-Z0-9_-]/g, '_');
      const folderPath = `projects/${folderName}`;

      console.log(`Verarbeite: ${script.name} (${script.id})`);

      if (!fs.existsSync(folderPath)) {
        fs.mkdirSync(folderPath, { recursive: true });
      }

      const claspJson = { scriptId: script.id, rootDir: "." };
      fs.writeFileSync(`${folderPath}/.clasp.json`, JSON.stringify(claspJson, null, 2));

      try {
        execSync('clasp pull', { cwd: folderPath, stdio: 'ignore' });
      } catch (e) {
        console.error(`Fehler beim Pull von ${script.name}. Überspringe...`);
      }
    }

    console.log("Sync erfolgreich abgeschlossen!");

  } catch (error) {
    console.error("Kritischer Fehler:", error.message || error);
    process.exit(1);
  }
}

syncScripts();
