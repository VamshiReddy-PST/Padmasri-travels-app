/**
 * Fleet Supervisor App - MongoDB Atlas backup script.
 * ---------------------------------------------------------------------
 * MongoDB Atlas's free M0 tier does not include automatic cloud backups
 * (that's a paid-tier feature). This script is the free alternative: it
 * connects straight to the database with the `mongodb` driver (already a
 * dependency of the main app) and dumps everything to a dated JSON file
 * on disk, then deletes anything older than `retentionDays` (default 30).
 *
 * Run manually:   node backup/backup.js
 * Run on a schedule: see the Cowork scheduled task set up alongside this
 * script - it runs this file automatically and prunes old backups.
 *
 * Reads backup/config.json for the connection string, so the connection
 * string (which contains your database password) never needs to be typed
 * into a command or committed to git - backup/config.json is already
 * listed in .gitignore.
 * ---------------------------------------------------------------------
 */
const fs = require("fs");
const path = require("path");
const { MongoClient } = require("mongodb");

const ROOT = __dirname;
const CONFIG_FILE = path.join(ROOT, "config.json");
const HISTORY_DIR = path.join(ROOT, "history");

function loadConfig() {
  if (!fs.existsSync(CONFIG_FILE)) {
    console.error(
      "backup/config.json not found. Create it with: " +
        '{ "mongodbUri": "mongodb+srv://...", "retentionDays": 30 }'
    );
    process.exit(1);
  }
  return JSON.parse(fs.readFileSync(CONFIG_FILE, "utf8"));
}

function timestampSlug() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}`;
}

async function run() {
  const config = loadConfig();
  if (!config.mongodbUri) {
    console.error("config.json is missing mongodbUri.");
    process.exit(1);
  }
  const retentionDays = Number(config.retentionDays) || 30;

  if (!fs.existsSync(HISTORY_DIR)) fs.mkdirSync(HISTORY_DIR, { recursive: true });

  const client = new MongoClient(config.mongodbUri);
  console.log("Connecting to MongoDB Atlas...");
  await client.connect();
  const db = client.db("fleet_supervisor_app");

  const appData = await db.collection("appdata").findOne({ _id: "main" });
  const photos = await db.collection("photos").find({}).toArray();
  await client.close();

  if (!appData) {
    console.error("No app data found in the database - nothing to back up.");
    process.exit(1);
  }

  const slug = timestampSlug();
  const dataFile = path.join(HISTORY_DIR, `data_${slug}.json`);
  const photosFile = path.join(HISTORY_DIR, `photos_${slug}.json`);

  fs.writeFileSync(dataFile, JSON.stringify(appData, null, 2));
  fs.writeFileSync(photosFile, JSON.stringify(photos, null, 2));

  console.log(
    `Backup saved: ${path.basename(dataFile)} (${(fs.statSync(dataFile).size / 1024).toFixed(1)} KB) ` +
      `+ ${path.basename(photosFile)} (${photos.length} photos, ${(fs.statSync(photosFile).size / 1024).toFixed(1)} KB)`
  );

  // Prune anything older than retentionDays.
  const cutoff = Date.now() - retentionDays * 24 * 60 * 60 * 1000;
  let removed = 0;
  for (const file of fs.readdirSync(HISTORY_DIR)) {
    const full = path.join(HISTORY_DIR, file);
    const stat = fs.statSync(full);
    if (stat.mtimeMs < cutoff) {
      fs.unlinkSync(full);
      removed++;
    }
  }
  console.log(`Pruned ${removed} backup file(s) older than ${retentionDays} days.`);
  console.log("Backup complete.");
}

run().catch((err) => {
  console.error("Backup failed:", err);
  process.exit(1);
});
