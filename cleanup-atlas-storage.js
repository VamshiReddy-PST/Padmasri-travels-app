/**
 * cleanup-atlas-storage.js
 *
 * Frees up space on the Padmasri Travels MongoDB Atlas cluster when it's
 * hit the 512MB free-tier (M0) quota and writes are blocked.
 *
 * WHY THIS IS A SEPARATE SCRIPT YOU RUN YOURSELF:
 * Claude never sees or handles your database password. This script reads
 * your connection string from an environment variable on YOUR machine, not
 * from anything typed into chat.
 *
 * WHAT'S IN THIS DATABASE (confirmed by reading server.js):
 *   - appdata      : ONE document holding the entire app state (vehicles,
 *                     drivers, trips, expenses, users, up to 5000 audit-log
 *                     entries, etc). Small per-write, but every save()
 *                     rewrites the whole thing.
 *   - backups      : a FULL clone of that same appdata document, taken once
 *                     a day, kept for 30 days. This is very likely the
 *                     single biggest source of growth - 30 near-duplicate
 *                     copies of the entire database.
 *   - locationHistory: one row per GPS ping (~every 10s per vehicle). Grows
 *                     fast; the app already prunes anything older than 30
 *                     days on its own, but that prune is itself a WRITE, so
 *                     it silently stops working once the cluster is
 *                     write-blocked - the two problems reinforce each other.
 *   - photos       : base64 document/photo scans (driver docs, vehicle
 *                     docs, fuel receipts, repair invoices). No automatic
 *                     retention. This script only REPORTS its size - it
 *                     never deletes from here automatically, since these
 *                     are real business documents you may still need.
 *
 * WHAT THIS SCRIPT DELETES (in this order):
 *   1. Old daily backups beyond a smaller retention window you choose
 *      (default: keep the most recent 5, instead of 30 - a backup is a
 *      safety net, not a document you look at, so this is the safest cut).
 *   2. locationHistory older than a cutoff you choose (default: keep last
 *      7 days instead of 30).
 *   3. auditLog entries inside the appdata document, beyond the most
 *      recent 2000 (already capped at 5000 by the app; this trims further).
 *
 * Atlas allows deletes even while write-blocking is active for new
 * inserts - that's the whole reason its own error message offers "free up
 * storage by deleting data" as a fix. If a delete still gets rejected here,
 * the cluster is blocked too hard even for that - at that point upgrading
 * the Atlas cluster tier (M0 -> M10+) is the only remaining lever, since M0
 * has no "add storage" option.
 *
 * HOW TO RUN:
 *   1. Get your connection string (the same MONGODB_URI you set on Render)
 *      from Render's dashboard -> your service -> Environment, or from
 *      the Atlas dashboard -> Database -> Connect.
 *   2. In a terminal, on your own computer:
 *
 *        cd /path/to/PadmasriTravels_Supervisor_app
 *        npm install mongodb        (if not already installed)
 *        MONGODB_URI="your-connection-string-here" node cleanup-atlas-storage.js
 *
 *      (On Windows PowerShell:
 *        $env:MONGODB_URI="your-connection-string-here"; node cleanup-atlas-storage.js )
 *
 *   3. It first PRINTS A REPORT of every collection's size - read that
 *      before doing anything else, so you know what's actually eating
 *      your 512MB. By default it's a dry run (nothing deleted). Re-run
 *      with --confirm to actually delete:
 *
 *        MONGODB_URI="..." node cleanup-atlas-storage.js --confirm
 *
 *   4. Optional flags:
 *        --keep-days=3        location history: keep only last N days (default 7)
 *        --keep-backups=2     backups: keep only most recent N (default 5)
 *
 * After this frees enough space, Atlas unblocks writes automatically,
 * usually within a few minutes - no app restart needed.
 */

const { MongoClient } = require("mongodb");

const MONGODB_URI = process.env.MONGODB_URI;
const CONFIRM = process.argv.includes("--confirm");
const keepDaysArg = process.argv.find((a) => a.startsWith("--keep-days="));
const KEEP_DAYS = keepDaysArg ? Number(keepDaysArg.split("=")[1]) : 7;
const keepBackupsArg = process.argv.find((a) => a.startsWith("--keep-backups="));
const KEEP_BACKUPS = keepBackupsArg ? Number(keepBackupsArg.split("=")[1]) : 5;
const AUDIT_KEEP = 2000;

function fmtBytes(n) {
  if (n == null || Number.isNaN(n)) return "unknown";
  const units = ["B", "KB", "MB", "GB"];
  let i = 0;
  while (n >= 1024 && i < units.length - 1) {
    n /= 1024;
    i++;
  }
  return `${n.toFixed(1)} ${units[i]}`;
}

async function main() {
  if (!MONGODB_URI) {
    console.error('ERROR: Set MONGODB_URI first, e.g.\n  MONGODB_URI="..." node cleanup-atlas-storage.js');
    process.exit(1);
  }
  if (!Number.isFinite(KEEP_DAYS) || KEEP_DAYS < 0) {
    console.error("ERROR: --keep-days must be a non-negative number.");
    process.exit(1);
  }
  if (!Number.isFinite(KEEP_BACKUPS) || KEEP_BACKUPS < 1) {
    console.error("ERROR: --keep-backups must be at least 1.");
    process.exit(1);
  }

  console.log(`Mode: ${CONFIRM ? "LIVE (will delete)" : "DRY RUN (preview only - add --confirm to actually delete)"}`);
  console.log(`Backups: keeping most recent ${KEEP_BACKUPS}.`);
  console.log(`Location history: keeping last ${KEEP_DAYS} day(s).`);
  console.log(`Audit log: keeping most recent ${AUDIT_KEEP} entries.\n`);

  const client = new MongoClient(MONGODB_URI);
  await client.connect();
  const db = client.db("fleet_supervisor_app");
  console.log(`Connected to database: ${db.databaseName}\n`);

  try {
    // ---- Report: size per collection, so you can see what's actually big ----
    console.log("--- Collection sizes (before) ---");
    const collections = await db.listCollections().toArray();
    for (const { name } of collections) {
      const stats = await db.command({ collStats: name }).catch(() => null);
      const count = stats ? stats.count : await db.collection(name).estimatedDocumentCount().catch(() => "?");
      console.log(`  ${name}: ${count} doc(s), ${fmtBytes(stats ? stats.size : null)} (storage: ${fmtBytes(stats ? stats.storageSize : null)})`);
    }
    console.log("");

    // ---- 1. backups (usually the biggest win) ----
    const backupsCol = db.collection("backups");
    const allBackups = await backupsCol
      .find({}, { projection: { _id: 1, takenAt: 1 } })
      .sort({ takenAt: -1 })
      .toArray()
      .catch(() => []);
    const backupsToDelete = allBackups.slice(KEEP_BACKUPS);
    console.log(`backups: ${allBackups.length} total, ${backupsToDelete.length} would be deleted (keeping newest ${KEEP_BACKUPS}).`);
    if (CONFIRM && backupsToDelete.length) {
      const ids = backupsToDelete.map((d) => d._id);
      const result = await backupsCol.deleteMany({ _id: { $in: ids } });
      console.log(`  -> Deleted ${result.deletedCount} backup(s).`);
    }
    console.log("");

    // ---- 2. locationHistory ----
    const locCol = db.collection("locationHistory");
    const cutoffIso = new Date(Date.now() - KEEP_DAYS * 24 * 60 * 60 * 1000).toISOString();
    const locCountToDelete = await locCol.countDocuments({ ts: { $lt: cutoffIso } }).catch((err) => {
      console.error("Could not count locationHistory:", err.message);
      return 0;
    });
    console.log(`locationHistory: ${locCountToDelete} point(s) older than ${cutoffIso} would be deleted.`);
    if (CONFIRM && locCountToDelete > 0) {
      const result = await locCol.deleteMany({ ts: { $lt: cutoffIso } });
      console.log(`  -> Deleted ${result.deletedCount} point(s).`);
    }
    console.log("");

    // ---- 3. auditLog inside the single appdata document ----
    const appDataCol = db.collection("appdata");
    const doc = await appDataCol.findOne({ _id: "main" }, { projection: { auditLog: 1 } }).catch(() => null);
    if (doc && Array.isArray(doc.auditLog) && doc.auditLog.length > AUDIT_KEEP) {
      console.log(`auditLog: ${doc.auditLog.length} entries, would trim to most recent ${AUDIT_KEEP}.`);
      if (CONFIRM) {
        const trimmed = doc.auditLog.slice(0, AUDIT_KEEP); // app unshifts newest to the front
        await appDataCol.updateOne({ _id: "main" }, { $set: { auditLog: trimmed } });
        console.log(`  -> Trimmed to ${trimmed.length} entries.`);
      }
    } else if (doc) {
      console.log(`auditLog: ${doc.auditLog ? doc.auditLog.length : 0} entries - already under the ${AUDIT_KEEP} threshold, nothing to trim.`);
    }
    console.log("");

    if (CONFIRM) {
      console.log("--- Collection sizes (after) ---");
      for (const { name } of collections) {
        const stats = await db.command({ collStats: name }).catch(() => null);
        console.log(`  ${name}: ${fmtBytes(stats ? stats.size : null)} (storage: ${fmtBytes(stats ? stats.storageSize : null)})`);
      }
      console.log("\nDone. If Atlas was write-blocked, it should unblock automatically within a few minutes");
      console.log("once free disk space is back above its threshold - no app restart needed.");
    } else {
      console.log("This was a dry run - nothing was deleted. Re-run with --confirm to actually free space:");
      console.log(`  MONGODB_URI="..." node cleanup-atlas-storage.js --confirm --keep-days=${KEEP_DAYS} --keep-backups=${KEEP_BACKUPS}`);
    }
  } finally {
    await client.close();
  }
}

main().catch((err) => {
  console.error("\nCleanup failed:", err.message);
  console.error("If this mentions writes/deletes being blocked or disk space, the cluster is blocked too hard");
  console.error("even for deletes - at that point upgrading the Atlas cluster tier (M0 -> M10+) is the only");
  console.error("remaining fix, since M0 has no 'add storage' option.");
  process.exit(1);
});
