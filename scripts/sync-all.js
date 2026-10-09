const fs = require('fs');
const { execSync } = require('child_process');
const { google } = require('googleapis');

async function syncScripts() {
  try {
    // Lese die Clasp Credentials (die durch GitHub Secrets erstellt wurden)
    const claspConfig = JSON.parse(fs.readFileSync(`${process.env.HOME}/.clasprc.json`, 'utf8'));
    
    // Auth Client für Google Drive aufbauen
    const oAuth2Client = new google.auth.OAuth2(
      claspConfig.oauth2ClientSettings.clientId,
      claspConfig.oauth2ClientSettings.clientSecret
    );
    oAuth2Client.setCredentials(claspConfig.token);
    
    const drive = google.drive({ version: 'v3', auth: oAuth2Client });

    console.log("Suche nach allen Apps Script Dateien...");
    
    // Suche alle Apps Scripts in Ihrem Account
    const response = await drive.files.list({
      q: "mimeType='application/vnd.google-apps.script' and trashed=false",
      fields: 'files(id, name)',
      pageSize: 1000 // Falls Sie mehr als 1000 haben, müsste man paginieren
    });

    const scripts = response.data.files;
    console.log(`${scripts.length} Skripte gefunden. Starte Sync...`);

    // Erstelle den Hauptordner, falls er nicht existiert
    if (!fs.existsSync('projects')) fs.mkdirSync('projects');

    for (const script of scripts) {
      // Bereinige den Skriptnamen für Ordnernamen (keine Leerzeichen/Sonderzeichen)
      const folderName = script.name.replace(/[^a-zA-Z0-9_-]/g, '_');
      const folderPath = `projects/${folderName}`;

      console.log(`Verarbeite: ${script.name} (${script.id})`);

      if (!fs.existsSync(folderPath)) {
        fs.mkdirSync(folderPath);
      }

      // Erstelle .clasp.json für dieses spezifische Skript
      const claspJson = { scriptId: script.id, rootDir: "." };
      fs.writeFileSync(`${folderPath}/.clasp.json`, JSON.stringify(claspJson, null, 2));

      // Führe Clasp Pull in diesem Ordner aus
      try {
        execSync('clasp pull', { cwd: folderPath, stdio: 'ignore' });
      } catch (e) {
        console.error(`Fehler beim Pull von ${script.name}. Überspringe...`);
      }
    }
    
    console.log("Sync abgeschlossen.");

  } catch (error) {
    console.error("Kritischer Fehler:", error);
    process.exit(1);
  }
}

syncScripts();
