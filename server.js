/**
 * Fleet Supervisor App - shared backend
 * ---------------------------------------------------------------------
 * A small Express server that gives every supervisor, area supervisor,
 * HR, operations manager, data-team member and the owner a single
 * shared login + shared dataset - which is what makes this different
 * from the earlier click-through prototype (which only stored data in
 * one person's browser).
 *
 * STORAGE: if a MONGODB_URI environment variable is set, all app data
 * (and uploaded photos) are stored in MongoDB - this is what survives
 * Render's free-tier redeploys/restarts, which wipe local disk. See
 * DEPLOY.md for how to create a free MongoDB Atlas cluster and set
 * MONGODB_URI in Render.
 *
 * If MONGODB_URI is NOT set, the app falls back to a local JSON file
 * (./data/data.json) and local disk (./uploads) - handy for testing on
 * your own machine, but this data WILL be lost on every Render redeploy.
 * ---------------------------------------------------------------------
 */
const express = require("express");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const { MongoClient } = require("mongodb");
const XLSX = require("xlsx");
const PDFDocument = require("pdfkit");

const DATA_DIR = path.join(__dirname, "data");
const DATA_FILE = path.join(DATA_DIR, "data.json");
const SEED_FILE = path.join(DATA_DIR, "seed.json");
const UPLOADS_DIR = path.join(__dirname, "uploads");
const PORT = process.env.PORT || 3000;

const MONGODB_URI = process.env.MONGODB_URI || "";
const USE_MONGO = !!MONGODB_URI;

let db; // the whole app dataset, loaded into memory - all route handlers read/write this exactly as before
let mongoClient = null;
let appDataCol = null;
let photosCol = null;
let backupsCol = null;
const BACKUP_RETENTION_DAYS = 30;
const BACKUP_INTERVAL_MS = 24 * 60 * 60 * 1000; // once a day

async function initStorage() {
  if (USE_MONGO) {
    mongoClient = new MongoClient(MONGODB_URI);
    await mongoClient.connect();
    const mdb = mongoClient.db("fleet_supervisor_app");
    appDataCol = mdb.collection("appdata");
    photosCol = mdb.collection("photos");
    backupsCol = mdb.collection("backups");
    const existing = await appDataCol.findOne({ _id: "main" });
    if (existing) {
      delete existing._id;
      db = existing;
      backfillDefaults();
      console.log("Storage: connected to MongoDB, loaded existing data. Data will survive redeploys/restarts.");
    } else {
      db = JSON.parse(fs.readFileSync(SEED_FILE, "utf8"));
      await persistNow();
      console.log("Storage: connected to MongoDB, no existing data found - seeded with demo data.");
    }
    await runBackup("startup");
    setInterval(() => runBackup("daily"), BACKUP_INTERVAL_MS);
  } else {
    if (!fs.existsSync(UPLOADS_DIR)) fs.mkdirSync(UPLOADS_DIR, { recursive: true });
    if (!fs.existsSync(DATA_FILE)) fs.copyFileSync(SEED_FILE, DATA_FILE);
    db = JSON.parse(fs.readFileSync(DATA_FILE, "utf8"));
    backfillDefaults();
    console.warn(
      "Storage: MONGODB_URI not set - using local file storage. This is fine for local testing, but on Render's " +
        "free tier this data (and uploaded photos) will be WIPED on every redeploy/restart. Set MONGODB_URI to fix " +
        "this permanently - see DEPLOY.md."
    );
  }
}

// Fills in fields that didn't exist in older data (e.g. an existing Atlas
// database created before this update) so upgrades never crash on
// `undefined` - new installs already have these from seed.json.
function backfillDefaults() {
  if (!db.odometerLogs) db.odometerLogs = {};
  if (!Array.isArray(db.backupDownloads)) db.backupDownloads = [];
  if (!Array.isArray(db.attendanceBackupDownloads)) db.attendanceBackupDownloads = [];
  if (!db.payrollProfiles) db.payrollProfiles = {};
  if (!Array.isArray(db.payrollAdhocEntries)) db.payrollAdhocEntries = [];
  (db.vehicles || []).forEach((v) => {
    if (v.driverAssignedDate === undefined) v.driverAssignedDate = null;
    if (v.make === undefined) v.make = "";
    if (v.model === undefined) v.model = "";
    if (v.engineNo === undefined) v.engineNo = "";
    if (v.chassisNo === undefined) v.chassisNo = "";
    if (v.rcDate === undefined) v.rcDate = "";
    if (v.seatingCapacity === undefined) v.seatingCapacity = null;
    if (v.routeId === undefined) v.routeId = null;
    // Subvendor (Driver-cum-Owner) linkage - a vehicle owned/operated by an
    // external subvendor we contract with, with a fixed billing contract
    // (see the SUBVENDORS section below). null/null for the vast majority
    // of the fleet, which is company-owned.
    if (v.subvendorId === undefined) v.subvendorId = null;
    if (v.billing === undefined) v.billing = null;
    // Existing fleet is all vans/cabs run on Diesel - safe defaults for
    // records created before vehicle class/fuel type existed as fields.
    if (v.vehicleType === undefined) v.vehicleType = "Cab";
    if (v.fuelType === undefined) v.fuelType = "Diesel";
    if (v.docs && v.docs.Permit) {
      if (v.docs.Permit.isStatePermit === undefined) v.docs.Permit.isStatePermit = false;
      if (v.docs.Permit.districtCount === undefined) v.docs.Permit.districtCount = 0;
      if (v.docs.Permit.districtNames === undefined) v.docs.Permit.districtNames = "";
    }
    if (v.docs) {
      // TemporaryPermit and BorderTax were added after some vehicles already
      // existed - give any vehicle missing them a blank entry so the office
      // can fill them in and drivers can see the field at all.
      if (!v.docs.TemporaryPermit) v.docs.TemporaryPermit = { number: "", expiry: "", updatedAt: null, copyUrl: null };
      if (!v.docs.BorderTax) v.docs.BorderTax = { number: "", expiry: "", updatedAt: null, copyUrl: null };
      DOC_COPY_TYPES.forEach((docType) => {
        const doc = v.docs[docType];
        if (!doc) return;
        if (doc.updatedAt === undefined) doc.updatedAt = null;
        if (doc.copyUrl === undefined) doc.copyUrl = null;
      });
      // Older seed data set RC.expiry by hand - re-derive it from rcDate so
      // it always reflects the 15-year rule going forward.
      if (v.docs.RC && v.rcDate) v.docs.RC.expiry = computeRcExpiry(v.rcDate);
    }
  });
  (db.clients || []).forEach((c) => {
    if (c.routeCap === undefined) c.routeCap = 0;
    if (!Array.isArray(c.routes)) c.routes = [];
  });
  (db.users || []).forEach((u) => {
    if (u.phone === undefined) u.phone = "";
    if (u.dateOfJoining === undefined) u.dateOfJoining = "";
    if (u.email === undefined) u.email = "";
    if (!Array.isArray(u.passwordHistory)) u.passwordHistory = [];
    // Migrating off the old 4-digit PIN login: anyone who doesn't already
    // have a real password hash gets a one-time temporary password derived
    // from their old PIN, and is forced to set their own real password the
    // next time they log in - nobody gets locked out by this upgrade.
    if (!u.passwordHash) {
      const tempPassword = `Padmasri@${u.pin || "0000"}`;
      u.passwordHash = hashPassword(tempPassword);
      u.mustChangePassword = true;
    }
    if (u.mustChangePassword === undefined) u.mustChangePassword = false;
    delete u.pin;
  });
  // The Owner's email is a one-time, explicitly-specified migration - it
  // only lives in seed.json otherwise, which never touches an already-
  // existing database (only a brand new one). Without this, an existing
  // production Owner account would be stuck with a blank email forever and
  // unable to log in with it.
  const ownerAccount = (db.users || []).find((u) => u.role === "owner" && !u.email);
  if (ownerAccount) ownerAccount.email = "vamshi.reddy@padmasritravels.in";
  (db.drivers || []).forEach((d) => {
    if (d.phone === undefined) d.phone = "";
    if (d.licenseNumber === undefined) d.licenseNumber = "";
    if (d.aadharNumber === undefined) d.aadharNumber = "";
    if (d.esiNumber === undefined) d.esiNumber = "";
    if (d.pfNumber === undefined) d.pfNumber = "";
    if (d.uanNumber === undefined) d.uanNumber = "";
    if (d.esiCertificateUrl === undefined) d.esiCertificateUrl = null;
    if (d.pfCertificateUrl === undefined) d.pfCertificateUrl = null;
    // Documents tab in the driver app (view/download only, the office fills
    // these in via the Staff People screen): PAN, License copy, Aadhar
    // copy, and bank details + a cancelled cheque copy.
    if (d.panNumber === undefined) d.panNumber = "";
    if (d.panCopyUrl === undefined) d.panCopyUrl = null;
    if (d.licenseCopyUrl === undefined) d.licenseCopyUrl = null;
    if (d.aadharCopyUrl === undefined) d.aadharCopyUrl = null;
    if (d.bankAccountNumber === undefined) d.bankAccountNumber = "";
    if (d.bankIfsc === undefined) d.bankIfsc = "";
    if (d.bankAccountHolderName === undefined) d.bankAccountHolderName = "";
    if (d.bankChequeUrl === undefined) d.bankChequeUrl = null;
    // Health record, driver training certificate, and police verification
    // certificate - same view/download-only pattern as the docs above, each
    // with an optional reference number, the date it was done/issued, and a
    // copy upload (all set by HR/Ops Manager/Owner/Data Team via Staff).
    if (d.healthRecordNumber === undefined) d.healthRecordNumber = "";
    if (d.healthRecordDate === undefined) d.healthRecordDate = "";
    if (d.healthRecordCopyUrl === undefined) d.healthRecordCopyUrl = null;
    if (d.trainingCertNumber === undefined) d.trainingCertNumber = "";
    if (d.trainingCertDate === undefined) d.trainingCertDate = "";
    if (d.trainingCertCopyUrl === undefined) d.trainingCertCopyUrl = null;
    if (d.policeVerificationNumber === undefined) d.policeVerificationNumber = "";
    if (d.policeVerificationDate === undefined) d.policeVerificationDate = "";
    if (d.policeVerificationCopyUrl === undefined) d.policeVerificationCopyUrl = null;
    if (d.dateOfJoining === undefined) d.dateOfJoining = "";
    if (d.drivingLevel === undefined) d.drivingLevel = "";
    if (d.performanceScore === undefined) d.performanceScore = null;
    // Driver login is deliberately dead simple - mobile number + a 4-digit
    // PIN, defaulting to 1234 for everyone and never forced to change (per
    // the Owner's explicit instruction - drivers aren't expected to know
    // what a "strong password" even means). This is NOT the same security
    // bar as staff's hashed/complex passwords - it's an intentional
    // trade-off for a low-friction, low-literacy login.
    if (d.pin === undefined) d.pin = "1234";
  });
  if (!Array.isArray(db.driverShifts)) db.driverShifts = [];
  (db.driverShifts || []).forEach((s) => {
    // Odometer verification: who actually typed in each reading (the
    // driver themselves, normally - or a Site Supervisor/Data Team member
    // doing a PIN-verified manual entry when the driver couldn't), and
    // whether a Site Supervisor (or Data Team, only when the vehicle has no
    // Supervisor) has since checked the photo and approved it. Once
    // verified===true there is deliberately no endpoint anywhere that can
    // change that reading again - approval is a one-way door.
    if (s.odometerOpenEnteredBy === undefined) s.odometerOpenEnteredBy = s.odometerOpen != null ? "driver" : null;
    if (s.odometerOpenVerified === undefined) s.odometerOpenVerified = false;
    if (s.odometerOpenVerifiedBy === undefined) s.odometerOpenVerifiedBy = null;
    if (s.odometerOpenVerifiedByName === undefined) s.odometerOpenVerifiedByName = null;
    if (s.odometerOpenVerifiedAt === undefined) s.odometerOpenVerifiedAt = null;
    if (s.odometerCloseEnteredBy === undefined) s.odometerCloseEnteredBy = s.odometerClose != null ? "driver" : null;
    if (s.odometerCloseVerified === undefined) s.odometerCloseVerified = false;
    if (s.odometerCloseVerifiedBy === undefined) s.odometerCloseVerifiedBy = null;
    if (s.odometerCloseVerifiedByName === undefined) s.odometerCloseVerifiedByName = null;
    if (s.odometerCloseVerifiedAt === undefined) s.odometerCloseVerifiedAt = null;
    // A single dynamic 6-digit code the driver can generate and read out to
    // whoever is standing in for them - short-lived, single-use, cleared
    // once consumed by a manual entry (or once it expires).
    if (s.manualEntryPin === undefined) s.manualEntryPin = null;
    if (s.manualEntryPinExpiresAt === undefined) s.manualEntryPinExpiresAt = null;
  });
  if (!Array.isArray(db.trips)) db.trips = [];
  if (!Array.isArray(db.driverLocationLogs)) db.driverLocationLogs = [];
  if (!db.payrollApprovals) db.payrollApprovals = {};
  (db.trips || []).forEach((t) => {
    // pickupPoint/dropPoint are what the driver app's Current Trip tab
    // actually displays now - older trips (from before this existed) only
    // had a single free-text "description", which becomes the trip notes.
    if (t.pickupPoint === undefined) t.pickupPoint = "";
    if (t.dropPoint === undefined) t.dropPoint = "";
    if (t.notes === undefined) t.notes = t.description || "";
    if (t.startedAt === undefined) t.startedAt = null;
    if (t.completedAt === undefined) t.completedAt = null;
  });
  if (!Array.isArray(db.subvendors)) db.subvendors = [];
  if (!Array.isArray(db.subvendorPayments)) db.subvendorPayments = [];
  if (!Array.isArray(db.subvendorAdvances)) db.subvendorAdvances = [];
  if (!Array.isArray(db.subvendorIncomeEntries)) db.subvendorIncomeEntries = [];
  (db.subvendors || []).forEach((sv) => {
    if (sv.contactPerson === undefined) sv.contactPerson = "";
    if (sv.email === undefined) sv.email = "";
    if (sv.phone === undefined) sv.phone = "";
    if (sv.gstNumber === undefined) sv.gstNumber = "";
    if (sv.address === undefined) sv.address = "";
    if (!Array.isArray(sv.passwordHistory)) sv.passwordHistory = [];
    if (sv.mustChangePassword === undefined) sv.mustChangePassword = false;
    if (sv.active === undefined) sv.active = true;
  });
  (db.expenses || []).forEach((e) => {
    // Vehicle number is now mandatory unless the request is flagged as
    // Office Expenditure - older requests either had a vehicle already or
    // didn't, so infer the flag from whichever is true.
    if (e.isOfficeExpenditure === undefined) e.isOfficeExpenditure = !e.vehicleId;
    if (e.paymentStage === undefined) e.paymentStage = e.status === "approved" ? "approved" : null;
    if (e.utrNumber === undefined) e.utrNumber = null;
    if (e.paymentUpdatedAt === undefined) e.paymentUpdatedAt = null;
    if (e.siteId === undefined) {
      const v = e.vehicleId ? (db.vehicles || []).find((veh) => veh.id === e.vehicleId) : null;
      e.siteId = v ? v.siteId || null : null;
    }
  });
}

// MongoDB Atlas's free (M0) tier doesn't include automatic cloud backups -
// that's a paid-tier feature. This is the free substitute: a rolling 30-day
// history of full-data snapshots kept in their own "backups" collection in
// the same cluster, protecting against accidental deletions/bad edits made
// through the app itself (not against Atlas infrastructure failure, which
// Atlas already guards against on its own). Runs once at startup and then
// once a day for as long as the server process stays up.
async function runBackup(reason) {
  if (!USE_MONGO || !backupsCol) return;
  try {
    const clone = JSON.parse(JSON.stringify(db));
    delete clone._id;
    const takenAt = nowIso();
    await backupsCol.insertOne({ _id: uid("bk"), takenAt, reason: reason || "scheduled", data: clone });
    const cutoffIso = new Date(Date.now() - BACKUP_RETENTION_DAYS * 24 * 60 * 60 * 1000).toISOString();
    const pruned = await backupsCol.deleteMany({ takenAt: { $lt: cutoffIso } });
    console.log(`Backup taken (${reason || "scheduled"}) at ${takenAt}. Pruned ${pruned.deletedCount || 0} backup(s) older than ${BACKUP_RETENTION_DAYS} days.`);
  } catch (err) {
    console.error("Backup failed (app keeps running normally):", err);
  }
}

async function persistNow() {
  if (USE_MONGO) {
    const clone = JSON.parse(JSON.stringify(db));
    clone._id = "main";
    await appDataCol.replaceOne({ _id: "main" }, clone, { upsert: true });
  } else {
    fs.writeFileSync(DATA_FILE, JSON.stringify(db, null, 2));
  }
}

// Every route awaits this after mutating `db`, so the HTTP response only
// goes out once the change is actually durable - important for an audit
// trail where "it said saved" needs to actually mean saved.
async function save() {
  try {
    await persistNow();
  } catch (err) {
    console.error("Failed to persist data:", err);
    throw err;
  }
}

function uid(prefix) {
  return prefix + "_" + crypto.randomBytes(6).toString("hex");
}
function nowIso() {
  return new Date().toISOString();
}
function todayStr() {
  return new Date().toISOString().slice(0, 10);
}
// RC Date is the registration date, not an expiry - RC expiry is always 15
// years from registration, computed here rather than entered by hand.
function computeRcExpiry(rcDate) {
  if (!rcDate) return "";
  const d = new Date(rcDate);
  if (isNaN(d.getTime())) return "";
  d.setFullYear(d.getFullYear() + 15);
  return d.toISOString().slice(0, 10);
}
// These document types get a photo/scan of the physical document attached
// (Insurance does not, per the Owner's spec). TemporaryPermit and BorderTax
// cover whatever an RTA checkpoint outside the vehicle's home state might
// ask for, alongside the original 5 - drivers view/download all of these
// read-only in the driver app once the office has filled them in here.
const DOC_COPY_TYPES = ["RC", "Permit", "Fitness", "Tax", "PUC", "TemporaryPermit", "BorderTax"];

// ---------- PASSWORDS ----------
// Every login is now email (or mobile number, if no email on file) plus a
// real password - replacing the old 4-digit PIN entirely. Passwords are
// hashed with scrypt (salted, stored as "salt:hash" hex) - never kept or
// exported in plaintext anywhere, including the full-data backup.
const PASSWORD_MIN_LENGTH = 9; // "more than 8 letters"
const PASSWORD_HISTORY_LIMIT = 5;
function hashPassword(password, salt) {
  const useSalt = salt || crypto.randomBytes(16).toString("hex");
  const hash = crypto.scryptSync(String(password), useSalt, 64).toString("hex");
  return `${useSalt}:${hash}`;
}
function verifyPassword(password, stored) {
  if (!password || !stored || typeof stored !== "string" || !stored.includes(":")) return false;
  const [salt, hash] = stored.split(":");
  let candidate;
  try {
    candidate = crypto.scryptSync(String(password), salt, 64).toString("hex");
  } catch (err) {
    return false;
  }
  const a = Buffer.from(candidate, "hex");
  const b = Buffer.from(hash, "hex");
  if (a.length !== b.length) return false;
  return crypto.timingSafeEqual(a, b);
}
// Strong-password policy: more than 8 characters, and a mix of letters,
// numbers and symbols - not just length.
function passwordPolicyError(password) {
  if (!password || String(password).length < PASSWORD_MIN_LENGTH) {
    return `Password must be more than 8 characters (at least ${PASSWORD_MIN_LENGTH}).`;
  }
  if (!/[a-zA-Z]/.test(password)) return "Password must include at least one letter.";
  if (!/[0-9]/.test(password)) return "Password must include at least one number.";
  if (!/[^a-zA-Z0-9]/.test(password)) return "Password must include at least one symbol (e.g. ! @ # $ %).";
  return null;
}
// Blocks reusing the current password OR any of the last few - checked by
// re-hashing the candidate against each stored salt, since hashes are
// salted and can't be compared directly.
function passwordReused(password, user) {
  const history = [user.passwordHash, ...(user.passwordHistory || [])].filter(Boolean);
  return history.some((stored) => verifyPassword(password, stored));
}

// Vehicle class + fuel type - the fleet only has Diesel buses (no Petrol/
// CNG/EV/Hybrid buses), so a Bus is always forced to Diesel. Cabs can be
// any of the 5 fuel types. Used to break mileage reporting out by category
// instead of one misleading blended fleet-wide average.
const VEHICLE_TYPES = ["Bus", "Cab"];
const FUEL_TYPES = ["Diesel", "Petrol", "CNG", "EV", "Hybrid"];
function validateVehicleTypeFuel(vehicleType, fuelType) {
  if (!VEHICLE_TYPES.includes(vehicleType)) return `Vehicle Type must be one of: ${VEHICLE_TYPES.join(", ")}.`;
  if (!FUEL_TYPES.includes(fuelType)) return `Fuel Type must be one of: ${FUEL_TYPES.join(", ")}.`;
  if (vehicleType === "Bus" && fuelType !== "Diesel") return "Buses can only be Diesel - there are no Petrol/CNG/EV/Hybrid buses in this fleet.";
  return null;
}

function roleLabelServer(role) {
  return (
    {
      site_supervisor: "Site Supervisor",
      area_supervisor: "Area Supervisor",
      ops_manager: "Operations Manager",
      data_team: "Data Team",
      owner: "Owner",
      hr: "HR",
      bookings: "Bookings Department",
    }[role] || role
  );
}

async function audit(user, action, detail) {
  db.auditLog.unshift({
    id: uid("a"),
    ts: nowIso(),
    userId: user ? user.id : "system",
    userName: user ? user.name : "System",
    action,
    detail,
  });
  if (db.auditLog.length > 5000) db.auditLog.length = 5000;
  await save();
}

// Stores a data: URL image either in MongoDB (photos collection) or on
// local disk, depending on storage mode, and returns the URL the frontend
// should use in an <img src="..."> to display it.
async function savePhoto(dataUrl, tag) {
  if (!dataUrl || typeof dataUrl !== "string" || !dataUrl.startsWith("data:")) return null;
  // Accepts images (camera/gallery captures throughout the app) and PDFs
  // (scanned certificates like a driver's ESI/PF proof are often PDFs) -
  // same storage path either way.
  const match = dataUrl.match(/^data:(image\/[\w+.-]+|application\/pdf);base64,(.+)$/);
  if (!match) return null;
  const contentType = match[1];
  const base64Data = match[2];

  if (USE_MONGO) {
    const id = uid("img");
    await photosCol.insertOne({ _id: id, contentType, data: base64Data, tag: tag || "img", createdAt: nowIso() });
    return "/api/photo/" + id;
  }
  const ext = contentType === "application/pdf" ? "pdf" : contentType.split("/")[1] || "jpg";
  const filename = `${tag || "img"}_${Date.now()}_${crypto.randomBytes(4).toString("hex")}.${ext}`;
  fs.writeFileSync(path.join(UPLOADS_DIR, filename), Buffer.from(base64Data, "base64"));
  return "/uploads/" + filename;
}

// distance in meters between two lat/lng points (Haversine)
function distanceMeters(lat1, lng1, lat2, lng2) {
  if ([lat1, lng1, lat2, lng2].some((v) => v === undefined || v === null || Number.isNaN(Number(v)))) return null;
  const R = 6371000;
  const toRad = (d) => (d * Math.PI) / 180;
  const dLat = toRad(lat2 - lat1);
  const dLng = toRad(lng2 - lng1);
  const a =
    Math.sin(dLat / 2) ** 2 +
    Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLng / 2) ** 2;
  return Math.round(R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a)));
}
const GEOFENCE_METERS = 400; // configurable "close enough to site" radius

// ---------- app ----------
const app = express();
app.use(express.json({ limit: "25mb" }));
if (!USE_MONGO) app.use("/uploads", express.static(UPLOADS_DIR));
app.use(express.static(path.join(__dirname, "public")));

app.get("/api/photo/:id", async (req, res) => {
  if (!USE_MONGO) return res.status(404).end();
  try {
    const photo = await photosCol.findOne({ _id: req.params.id });
    if (!photo) return res.status(404).end();
    res.setHeader("Content-Type", photo.contentType);
    res.setHeader("Cache-Control", "public, max-age=31536000, immutable");
    res.send(Buffer.from(photo.data, "base64"));
  } catch (err) {
    console.error("Failed to load photo:", err);
    res.status(500).end();
  }
});

// ---------- auth ----------
const sessions = new Map(); // token -> userId

function publicUser(u) {
  if (!u) return null;
  const { passwordHash, passwordHistory, ...rest } = u;
  return rest;
}
// A driver's `pin` is a real login credential for the driver app now (not
// just a cosmetic field), so it's stripped out of anything the staff app
// shows - same idea as publicUser() hiding passwordHash above.
function publicDriverForStaff(d) {
  if (!d) return null;
  const { pin, ...rest } = d;
  return rest;
}

// A handful of routes stay reachable even while a user is mid-forced-
// password-change (mustChangePassword=true) - just enough for the frontend
// to know who they are and let them actually set the new password. Every
// other endpoint is blocked until that's done.
const ALLOWED_DURING_FORCED_PASSWORD_CHANGE = ["/api/meta", "/api/change-password", "/api/logout"];
function requireAuth(req, res, next) {
  const header = req.headers.authorization || "";
  const token = header.startsWith("Bearer ") ? header.slice(7) : null;
  const userId = token && sessions.get(token);
  const user = userId && db.users.find((u) => u.id === userId && u.active !== false);
  if (!user) return res.status(401).json({ error: "Not signed in. Please log in again." });
  if (user.mustChangePassword && !ALLOWED_DURING_FORCED_PASSWORD_CHANGE.includes(req.path)) {
    return res.status(403).json({ error: "You must set a new password before continuing.", mustChangePassword: true });
  }
  req.user = user;
  next();
}
function requireRole(...roles) {
  return (req, res, next) => {
    if (!roles.includes(req.user.role)) {
      return res.status(403).json({ error: `This action needs one of these roles: ${roles.join(", ")}.` });
    }
    next();
  };
}

// Wraps an async route handler so a thrown/rejected error becomes a clean
// 500 response instead of crashing the process or hanging the request.
function h(fn) {
  return (req, res) => {
    Promise.resolve(fn(req, res)).catch((err) => {
      console.error(err);
      if (!res.headersSent) res.status(500).json({ error: "Something went wrong saving that - please try again." });
    });
  };
}

// ---------- AUTH ROUTES ----------
// Everyone signs in with their email address (or, if they don't have one on
// file, their mobile number) plus their password - there's no more "pick
// your name from a list" picker, both because it doesn't scale to a real
// login and because it quietly exposed every active user's name/role to
// anyone who loaded the login page, unauthenticated.
app.post(
  "/api/login",
  h(async (req, res) => {
    const { identifier, password } = req.body || {};
    if (!identifier || !password) {
      return res.status(400).json({ error: "Email/mobile number and password are required." });
    }
    const idNorm = String(identifier).trim().toLowerCase();
    const phoneNorm = String(identifier).trim();
    const user = db.users.find(
      (u) =>
        u.active !== false &&
        ((u.email && u.email.toLowerCase() === idNorm) || (u.phone && u.phone === phoneNorm))
    );
    if (!user || !verifyPassword(password, user.passwordHash)) {
      return res.status(401).json({ error: "Incorrect email/mobile number or password." });
    }
    const token = crypto.randomBytes(24).toString("hex");
    sessions.set(token, user.id);
    await audit(user, "login", `${user.name} (${user.role}) logged in`);
    res.json({ token, user: publicUser(user) });
  })
);

app.post(
  "/api/logout",
  requireAuth,
  h(async (req, res) => {
    const header = req.headers.authorization || "";
    const token = header.slice(7);
    sessions.delete(token);
    await audit(req.user, "logout", `${req.user.name} logged out`);
    res.json({ ok: true });
  })
);

// Self-service password change - this is the one and only escape hatch
// from a forced mustChangePassword lock (first login, or after an Owner
// reset), and also how anyone can voluntarily change their own password
// at any time. Always requires re-entering the current password, even
// during a forced change, since that's simple proof they are who they
// say they are.
app.post(
  "/api/change-password",
  requireAuth,
  h(async (req, res) => {
    const { currentPassword, newPassword } = req.body || {};
    if (!currentPassword || !newPassword) {
      return res.status(400).json({ error: "Current and new password are required." });
    }
    if (!verifyPassword(currentPassword, req.user.passwordHash)) {
      return res.status(400).json({ error: "Current password is incorrect." });
    }
    const policyError = passwordPolicyError(newPassword);
    if (policyError) return res.status(400).json({ error: policyError });
    if (passwordReused(newPassword, req.user)) {
      return res.status(400).json({ error: "New password cannot be the same as your current or a recent previous password." });
    }
    req.user.passwordHistory = [req.user.passwordHash, ...(req.user.passwordHistory || [])].slice(0, PASSWORD_HISTORY_LIMIT);
    req.user.passwordHash = hashPassword(newPassword);
    req.user.mustChangePassword = false;
    await audit(req.user, "change_password", `${req.user.name} changed their own password`);
    res.json(publicUser(req.user));
  })
);

// ---------- META (clients, sites, drivers, users) ----------
app.get("/api/meta", requireAuth, (req, res) => {
  res.json({
    clients: db.clients,
    sites: db.sites,
    drivers: db.drivers.map(publicDriverForStaff),
    users: db.users.map(publicUser),
    me: publicUser(req.user),
  });
});

// ---------- USERS (People admin) ----------
const ADMIN_ROLES = ["ops_manager", "owner", "data_team"];
const VALID_ROLES = ["site_supervisor", "area_supervisor", "ops_manager", "data_team", "owner", "hr", "bookings"];

// People-management permission tiers:
// - Owner can create/edit anyone, including other Owner/Ops Manager/HR accounts.
// - HR is the primary staff onboarder: can create/edit Site Supervisor, Area
//   Supervisor, Operations Manager, Data Team and Bookings Department
//   accounts (everyone except Owner and other HR accounts).
// - Ops Manager can create/edit everyone EXCEPT Owner, Ops Manager, and HR
//   accounts (so they can't touch their own tier or promote themselves).
// - Data Team can view the People screen (needed for corrections/audit
//   context) but cannot create or edit anyone - enforced by simply not being
//   in PEOPLE_EDITOR_ROLES below.
// - Site/Area Supervisors have no People access at all (no route, no tab).
const PEOPLE_EDITOR_ROLES = ["ops_manager", "owner", "hr"];
const PEOPLE_VIEWER_ROLES = ["ops_manager", "owner", "hr", "data_team"];
const ROLE_MANAGEMENT_ALLOWED = {
  hr: ["site_supervisor", "area_supervisor", "ops_manager", "data_team", "bookings"],
  ops_manager: ["site_supervisor", "area_supervisor", "data_team", "bookings"],
};
function canManageUserRole(actorRole, targetRole) {
  if (actorRole === "owner") return true;
  const allowed = ROLE_MANAGEMENT_ALLOWED[actorRole];
  return !!allowed && allowed.includes(targetRole);
}

// Login identifiers (email, or phone as a fallback) must be unique across
// active accounts - otherwise login can't tell two people apart. Checked
// on both create and edit, excluding the user being edited themselves.
function findLoginIdentifierClash(email, phone, excludeUserId) {
  const emailNorm = email ? String(email).trim().toLowerCase() : "";
  const phoneNorm = phone ? String(phone).trim() : "";
  if (!emailNorm && !phoneNorm) return null;
  return db.users.find(
    (u) =>
      u.id !== excludeUserId &&
      ((emailNorm && u.email && u.email.toLowerCase() === emailNorm) || (phoneNorm && u.phone && u.phone === phoneNorm))
  );
}

app.post(
  "/api/users",
  requireAuth,
  requireRole(...PEOPLE_EDITOR_ROLES),
  h(async (req, res) => {
    const { name, role, password, siteId, supervises, phone, dateOfJoining, email } = req.body || {};
    if (!name || !VALID_ROLES.includes(role) || !password) {
      return res.status(400).json({ error: "name, role and password are required." });
    }
    if (!canManageUserRole(req.user.role, role)) {
      return res.status(403).json({ error: `Only the Owner can create ${roleLabelServer(role)} accounts.` });
    }
    const emailNorm = email ? String(email).trim().toLowerCase() : "";
    const phoneNorm = phone ? String(phone).trim() : "";
    if (!emailNorm && !phoneNorm) {
      return res.status(400).json({ error: "Either an email address or a mobile number is required to log in." });
    }
    if (findLoginIdentifierClash(emailNorm, phoneNorm, null)) {
      return res.status(400).json({ error: "That email or mobile number is already in use by another account." });
    }
    const policyError = passwordPolicyError(password);
    if (policyError) return res.status(400).json({ error: policyError });
    const user = {
      id: uid("u"),
      name,
      role,
      email: emailNorm,
      phone: phoneNorm,
      passwordHash: hashPassword(password),
      passwordHistory: [],
      // New accounts always have to set their own password on first login -
      // whoever created the account only ever sets a temporary one.
      mustChangePassword: true,
      active: true,
      siteId: siteId || null,
      supervises: Array.isArray(supervises) ? supervises : [],
      dateOfJoining: dateOfJoining || "",
    };
    db.users.push(user);
    await audit(req.user, "create_user", `${req.user.name} added ${name} as ${role}`);
    res.json(publicUser(user));
  })
);

app.patch(
  "/api/users/:id",
  requireAuth,
  requireRole(...PEOPLE_EDITOR_ROLES),
  h(async (req, res) => {
    const user = db.users.find((u) => u.id === req.params.id);
    if (!user) return res.status(404).json({ error: "User not found." });
    if (!canManageUserRole(req.user.role, user.role)) {
      return res.status(403).json({ error: `Only the Owner can edit ${roleLabelServer(user.role)} accounts.` });
    }
    const before = { name: user.name, role: user.role, siteId: user.siteId, supervises: user.supervises, active: user.active };

    // Whitelisted fields only - password is deliberately NOT settable here
    // at all, even by the Owner. Password changes only ever happen through
    // /api/change-password (self-service) or /api/users/:id/reset-password
    // (Owner-only), so there's exactly one audited path for each.
    const { name, role, siteId, supervises, active, phone, dateOfJoining, email } = req.body || {};
    if (name !== undefined && String(name).trim()) user.name = String(name).trim();
    if (role !== undefined) {
      if (!VALID_ROLES.includes(role)) return res.status(400).json({ error: "Unknown role." });
      // Prevent a non-Owner from promoting someone INTO a locked tier either.
      if (!canManageUserRole(req.user.role, role)) {
        return res.status(403).json({ error: `Only the Owner can set someone to ${roleLabelServer(role)}.` });
      }
      user.role = role;
    }
    if (email !== undefined || phone !== undefined) {
      const nextEmail = email !== undefined ? email : user.email;
      const nextPhone = phone !== undefined ? phone : user.phone;
      if (!nextEmail && !nextPhone) {
        return res.status(400).json({ error: "Either an email address or a mobile number is required to log in." });
      }
      if (findLoginIdentifierClash(nextEmail, nextPhone, user.id)) {
        return res.status(400).json({ error: "That email or mobile number is already in use by another account." });
      }
      if (email !== undefined) user.email = email ? String(email).trim().toLowerCase() : "";
      if (phone !== undefined) user.phone = phone ? String(phone).trim() : "";
    }
    if (siteId !== undefined) user.siteId = siteId || null;
    if (supervises !== undefined) user.supervises = Array.isArray(supervises) ? supervises : user.supervises;
    if (dateOfJoining !== undefined) user.dateOfJoining = dateOfJoining || "";
    if (active !== undefined) user.active = !!active;

    const after = { name: user.name, role: user.role, siteId: user.siteId, supervises: user.supervises, active: user.active };
    await audit(req.user, "update_user", `${req.user.name} updated ${user.name} (${user.id}): ${JSON.stringify(before)} -> ${JSON.stringify(after)}`);
    res.json(publicUser(user));
  })
);

// Owner-only password reset - gives the Owner full control over every
// staff member's password. Unlike a self-service change, this doesn't
// require knowing the old password, but it does force the target to set
// their own brand-new password the next time they log in.
app.patch(
  "/api/users/:id/reset-password",
  requireAuth,
  requireRole("owner"),
  h(async (req, res) => {
    const user = db.users.find((u) => u.id === req.params.id);
    if (!user) return res.status(404).json({ error: "User not found." });
    const { newPassword } = req.body || {};
    const policyError = passwordPolicyError(newPassword);
    if (policyError) return res.status(400).json({ error: policyError });
    if (passwordReused(newPassword, user)) {
      return res.status(400).json({ error: "New password cannot match this person's current or a recent previous password." });
    }
    user.passwordHistory = [user.passwordHash, ...(user.passwordHistory || [])].filter(Boolean).slice(0, PASSWORD_HISTORY_LIMIT);
    user.passwordHash = hashPassword(newPassword);
    user.mustChangePassword = true;
    await audit(req.user, "reset_password", `${req.user.name} reset the password for ${user.name} (${user.id}) - they must set a new one at next login`);
    res.json(publicUser(user));
  })
);

// ---------- CLIENTS ----------
app.post(
  "/api/clients",
  requireAuth,
  requireRole(...ADMIN_ROLES),
  h(async (req, res) => {
    const { name } = req.body || {};
    if (!name) return res.status(400).json({ error: "name is required." });
    const client = { id: uid("c"), name, routeCap: 0, routes: [] };
    db.clients.push(client);
    await audit(req.user, "create_client", `${req.user.name} added client ${name}`);
    res.json(client);
  })
);

// Route cap - how many routes this client is allowed to have entered below.
// Owner-only by design: Operations Manager fills in the actual routes, but
// only the Owner controls how many they're allowed to add.
app.patch(
  "/api/clients/:id/route-cap",
  requireAuth,
  requireRole("owner"),
  h(async (req, res) => {
    const client = db.clients.find((c) => c.id === req.params.id);
    if (!client) return res.status(404).json({ error: "Client not found." });
    const { routeCap } = req.body || {};
    const cap = Number(routeCap);
    if (!Number.isFinite(cap) || cap < 0) return res.status(400).json({ error: "Route cap must be a non-negative number." });
    const before = client.routeCap;
    client.routeCap = cap;
    await audit(req.user, "set_route_cap", `${req.user.name} set route cap for ${client.name}: ${before} -> ${cap}`);
    res.json(client);
  })
);

// Individual routes (starting point name + route number, e.g. "Alwal" / "6A")
// - Owner and Operations Manager only, and never more than the client's
// route cap. Data Team and Site/Area Supervisors can see this (it's already
// in /api/meta) but have no route to write to it.
const ROUTE_EDITOR_ROLES = ["owner", "ops_manager"];
app.post(
  "/api/clients/:id/routes",
  requireAuth,
  requireRole(...ROUTE_EDITOR_ROLES),
  h(async (req, res) => {
    const client = db.clients.find((c) => c.id === req.params.id);
    if (!client) return res.status(404).json({ error: "Client not found." });
    const { name, routeNumber } = req.body || {};
    if (!name || !String(name).trim()) return res.status(400).json({ error: "Route name (starting point) is required." });
    if (client.routes.length >= client.routeCap) {
      return res.status(400).json({
        error: `${client.name} is already at its route cap (${client.routeCap}). Ask the Owner to raise it before adding another route.`,
      });
    }
    const route = { id: uid("rt"), name: String(name).trim(), routeNumber: routeNumber || "" };
    client.routes.push(route);
    await audit(req.user, "add_route", `${req.user.name} added route "${route.name}" (${route.routeNumber || "no number"}) to ${client.name}`);
    res.json(client);
  })
);

app.patch(
  "/api/clients/:id/routes/:routeId",
  requireAuth,
  requireRole(...ROUTE_EDITOR_ROLES),
  h(async (req, res) => {
    const client = db.clients.find((c) => c.id === req.params.id);
    if (!client) return res.status(404).json({ error: "Client not found." });
    const route = client.routes.find((r) => r.id === req.params.routeId);
    if (!route) return res.status(404).json({ error: "Route not found." });
    const { name, routeNumber } = req.body || {};
    if (name !== undefined && String(name).trim()) route.name = String(name).trim();
    if (routeNumber !== undefined) route.routeNumber = routeNumber;
    await audit(req.user, "update_route", `${req.user.name} updated route "${route.name}" on ${client.name}`);
    res.json(client);
  })
);

app.delete(
  "/api/clients/:id/routes/:routeId",
  requireAuth,
  requireRole(...ROUTE_EDITOR_ROLES),
  h(async (req, res) => {
    const client = db.clients.find((c) => c.id === req.params.id);
    if (!client) return res.status(404).json({ error: "Client not found." });
    const idx = client.routes.findIndex((r) => r.id === req.params.routeId);
    if (idx === -1) return res.status(404).json({ error: "Route not found." });
    const [removed] = client.routes.splice(idx, 1);
    await audit(req.user, "remove_route", `${req.user.name} removed route "${removed.name}" from ${client.name}`);
    res.json(client);
  })
);

// ---------- SITES ----------
app.post(
  "/api/sites",
  requireAuth,
  requireRole(...ADMIN_ROLES),
  h(async (req, res) => {
    const { name, lat, lng, areaSupervisorId } = req.body || {};
    if (!name) return res.status(400).json({ error: "name is required." });
    const site = { id: uid("s"), name, lat: lat != null ? Number(lat) : null, lng: lng != null ? Number(lng) : null, areaSupervisorId: areaSupervisorId || null };
    db.sites.push(site);
    await audit(req.user, "create_site", `${req.user.name} added site ${name}`);
    res.json(site);
  })
);

// ---------- DRIVERS ----------
// HR is the primary driver onboarder (per the People/Staff permission
// model above), alongside Ops Manager, Owner and Data Team who could
// already add drivers from the Operations screen's quick-add form.
const DRIVER_EDITOR_ROLES = ["owner", "ops_manager", "hr", "data_team"];
const DRIVING_LEVELS = ["Trainee", "Standard", "Senior", "Expert"];
app.post(
  "/api/drivers",
  requireAuth,
  requireRole(...DRIVER_EDITOR_ROLES),
  h(async (req, res) => {
    const { name, phone, licenseNumber, aadharNumber, dateOfJoining, drivingLevel } = req.body || {};
    if (!name) return res.status(400).json({ error: "name is required." });
    if (drivingLevel && !DRIVING_LEVELS.includes(drivingLevel)) {
      return res.status(400).json({ error: `Driving level must be one of: ${DRIVING_LEVELS.join(", ")}.` });
    }
    const driver = {
      id: uid("d"),
      name,
      phone: phone || "",
      licenseNumber: licenseNumber || "",
      aadharNumber: aadharNumber || "",
      esiNumber: "",
      pfNumber: "",
      uanNumber: "",
      esiCertificateUrl: null,
      pfCertificateUrl: null,
      healthRecordNumber: "",
      healthRecordDate: "",
      healthRecordCopyUrl: null,
      trainingCertNumber: "",
      trainingCertDate: "",
      trainingCertCopyUrl: null,
      policeVerificationNumber: "",
      policeVerificationDate: "",
      policeVerificationCopyUrl: null,
      dateOfJoining: dateOfJoining || "",
      drivingLevel: drivingLevel || "",
      performanceScore: null,
      // Driver app login - mobile number + this PIN, defaulting to 1234
      // and never forced to change (see backfillDefaults for why).
      pin: "1234",
    };
    db.drivers.push(driver);
    await audit(req.user, "create_driver", `${req.user.name} added driver ${name}`);
    res.json(publicDriverForStaff(driver));
  })
);

// Full driver-profile editor - HR's onboarding fields (Aadhar/ESI/PF/UAN,
// certificate copies, date of joining, driving level) plus the basic
// name/phone/license already settable at creation. Performance score is
// intentionally NOT editable here - it's reserved for a future client-
// ratings feature to populate, so it stays null ("Not yet rated") until
// then rather than letting anyone hand-set it.
app.patch(
  "/api/drivers/:id",
  requireAuth,
  requireRole(...DRIVER_EDITOR_ROLES),
  h(async (req, res) => {
    const driver = db.drivers.find((d) => d.id === req.params.id);
    if (!driver) return res.status(404).json({ error: "Driver not found." });
    const {
      name, phone, licenseNumber, aadharNumber, esiNumber, pfNumber, uanNumber,
      dateOfJoining, drivingLevel, esiCertificate, pfCertificate,
      panNumber, bankAccountNumber, bankIfsc, bankAccountHolderName,
      licenseCopy, aadharCopy, panCopy, bankCheque,
      healthRecordNumber, healthRecordDate, healthRecordCopy,
      trainingCertNumber, trainingCertDate, trainingCertCopy,
      policeVerificationNumber, policeVerificationDate, policeVerificationCopy,
    } = req.body || {};
    if (drivingLevel !== undefined && drivingLevel && !DRIVING_LEVELS.includes(drivingLevel)) {
      return res.status(400).json({ error: `Driving level must be one of: ${DRIVING_LEVELS.join(", ")}.` });
    }
    let changed = false;
    if (name !== undefined && String(name).trim()) { driver.name = String(name).trim(); changed = true; }
    if (phone !== undefined) { driver.phone = phone || ""; changed = true; }
    if (licenseNumber !== undefined) { driver.licenseNumber = licenseNumber || ""; changed = true; }
    if (aadharNumber !== undefined) { driver.aadharNumber = aadharNumber || ""; changed = true; }
    if (esiNumber !== undefined) { driver.esiNumber = esiNumber || ""; changed = true; }
    if (pfNumber !== undefined) { driver.pfNumber = pfNumber || ""; changed = true; }
    if (uanNumber !== undefined) { driver.uanNumber = uanNumber || ""; changed = true; }
    if (dateOfJoining !== undefined) { driver.dateOfJoining = dateOfJoining || ""; changed = true; }
    if (drivingLevel !== undefined) { driver.drivingLevel = drivingLevel || ""; changed = true; }
    // Documents tab in the driver app - PAN, bank account details, and a
    // cancelled cheque copy, alongside the License/Aadhar copies below.
    // View/download only for the driver; only HR/Ops Manager/Owner/Data
    // Team can actually set these, same as everything else on this route.
    if (panNumber !== undefined) { driver.panNumber = panNumber || ""; changed = true; }
    if (bankAccountNumber !== undefined) { driver.bankAccountNumber = bankAccountNumber || ""; changed = true; }
    if (bankIfsc !== undefined) { driver.bankIfsc = bankIfsc || ""; changed = true; }
    if (bankAccountHolderName !== undefined) { driver.bankAccountHolderName = bankAccountHolderName || ""; changed = true; }
    if (esiCertificate) {
      const url = await savePhoto(esiCertificate, "driver_esi_" + driver.id);
      if (url) { driver.esiCertificateUrl = url; changed = true; }
    }
    if (pfCertificate) {
      const url = await savePhoto(pfCertificate, "driver_pf_" + driver.id);
      if (url) { driver.pfCertificateUrl = url; changed = true; }
    }
    if (licenseCopy) {
      const url = await savePhoto(licenseCopy, "driver_license_" + driver.id);
      if (url) { driver.licenseCopyUrl = url; changed = true; }
    }
    if (aadharCopy) {
      const url = await savePhoto(aadharCopy, "driver_aadhar_" + driver.id);
      if (url) { driver.aadharCopyUrl = url; changed = true; }
    }
    if (panCopy) {
      const url = await savePhoto(panCopy, "driver_pan_" + driver.id);
      if (url) { driver.panCopyUrl = url; changed = true; }
    }
    if (bankCheque) {
      const url = await savePhoto(bankCheque, "driver_cheque_" + driver.id);
      if (url) { driver.bankChequeUrl = url; changed = true; }
    }
    // Health record, driver training certificate, police verification -
    // same number/date/copy-upload pattern as the docs above.
    if (healthRecordNumber !== undefined) { driver.healthRecordNumber = healthRecordNumber || ""; changed = true; }
    if (healthRecordDate !== undefined) { driver.healthRecordDate = healthRecordDate || ""; changed = true; }
    if (healthRecordCopy) {
      const url = await savePhoto(healthRecordCopy, "driver_health_" + driver.id);
      if (url) { driver.healthRecordCopyUrl = url; changed = true; }
    }
    if (trainingCertNumber !== undefined) { driver.trainingCertNumber = trainingCertNumber || ""; changed = true; }
    if (trainingCertDate !== undefined) { driver.trainingCertDate = trainingCertDate || ""; changed = true; }
    if (trainingCertCopy) {
      const url = await savePhoto(trainingCertCopy, "driver_training_" + driver.id);
      if (url) { driver.trainingCertCopyUrl = url; changed = true; }
    }
    if (policeVerificationNumber !== undefined) { driver.policeVerificationNumber = policeVerificationNumber || ""; changed = true; }
    if (policeVerificationDate !== undefined) { driver.policeVerificationDate = policeVerificationDate || ""; changed = true; }
    if (policeVerificationCopy) {
      const url = await savePhoto(policeVerificationCopy, "driver_police_" + driver.id);
      if (url) { driver.policeVerificationCopyUrl = url; changed = true; }
    }
    if (changed) {
      await audit(req.user, "update_driver", `${req.user.name} updated driver ${driver.name} (${driver.id})`);
    }
    res.json(publicDriverForStaff(driver));
  })
);

// Safety valve if a driver forgets their PIN or it's been fumbled -
// resets it straight back to the default (1234), same as a brand new
// driver. HR/Owner only.
app.patch(
  "/api/drivers/:id/reset-pin",
  requireAuth,
  requireRole("owner", "hr"),
  h(async (req, res) => {
    const driver = db.drivers.find((d) => d.id === req.params.id);
    if (!driver) return res.status(404).json({ error: "Driver not found." });
    driver.pin = "1234";
    await audit(req.user, "reset_driver_pin", `${req.user.name} reset the app PIN for driver ${driver.name} (${driver.id}) back to the default`);
    res.json(publicDriverForStaff(driver));
  })
);

// Per-driver mileage/efficiency summary for the People tab's Drivers card -
// averages every fuel-fill day recorded against whichever vehicle(s) that
// driver is/was currently assigned to (from their driverAssignedDate
// onward), skipping any day a temporary driver covered instead so a
// driver's numbers only ever reflect their own driving.
function computeDriverMileageSummaries() {
  const stats = {};
  db.drivers.forEach((d) => { stats[d.id] = { fillCount: 0, totalMileage: 0, belowStandardCount: 0, lastFillDate: null }; });
  Object.values(db.records).forEach((r) => {
    if (!r.fuel || !(r.fuel.litres > 0)) return;
    const vehicle = db.vehicles.find((v) => v.id === r.vehicleId);
    if (!vehicle || !vehicle.driverId || !stats[vehicle.driverId]) return;
    const tempArranged = r.attendance && r.attendance.tempDriver && r.attendance.tempDriver.arranged;
    if (tempArranged) return;
    if (vehicle.driverAssignedDate && r.date < vehicle.driverAssignedDate) return;
    const s = stats[vehicle.driverId];
    s.fillCount += 1;
    s.totalMileage += r.fuel.mileage;
    if (r.fuel.belowStandard) s.belowStandardCount += 1;
    if (!s.lastFillDate || r.date > s.lastFillDate) s.lastFillDate = r.date;
  });
  return db.drivers.map((d) => {
    const s = stats[d.id];
    return {
      driverId: d.id,
      fillCount: s.fillCount,
      avgMileage: s.fillCount ? Math.round((s.totalMileage / s.fillCount) * 10) / 10 : null,
      belowStandardCount: s.belowStandardCount,
      lastFillDate: s.lastFillDate,
    };
  });
}
app.get(
  "/api/drivers/mileage-summary",
  requireAuth,
  requireRole(...PEOPLE_VIEWER_ROLES),
  (req, res) => {
    res.json(computeDriverMileageSummaries());
  }
);

// ---------- VEHICLES ----------
function visibleVehiclesFor(user) {
  if (user.role === "site_supervisor") return db.vehicles.filter((v) => v.supervisorId === user.id);
  if (user.role === "area_supervisor") {
    const mySupervisors = new Set(user.supervises || []);
    return db.vehicles.filter((v) => mySupervisors.has(v.supervisorId));
  }
  return db.vehicles;
}
// A route entry's "name" field is what the route catalog UI labels
// "Starting Point" - this just looks it up for a given vehicle's routeId.
function vehicleRouteInfo(v) {
  if (!v.routeId) return { routeNumber: v.route || "", startingPoint: "" };
  const client = db.clients.find((c) => c.id === v.clientId);
  const routeEntry = client ? (client.routes || []).find((r) => r.id === v.routeId) : null;
  if (!routeEntry) return { routeNumber: v.route || "", startingPoint: "" };
  return { routeNumber: routeEntry.routeNumber || "", startingPoint: routeEntry.name || "" };
}

app.get("/api/vehicles", requireAuth, (req, res) => {
  res.json(visibleVehiclesFor(req.user));
});

app.post(
  "/api/vehicles",
  requireAuth,
  requireRole(...ADMIN_ROLES),
  h(async (req, res) => {
    const { reg, route, usage, standardMileage, make, model, engineNo, chassisNo, rcDate, seatingCapacity, vehicleType, fuelType } = req.body || {};
    // A vehicle record without its full onboarding data is not usable for
    // compliance/reporting purposes, so every field below is mandatory at
    // creation time - there is no "add now, fill in details later" path.
    const missing = [];
    if (!reg || !String(reg).trim()) missing.push("Registration number");
    if (!make || !String(make).trim()) missing.push("Make");
    if (!model || !String(model).trim()) missing.push("Model");
    if (!engineNo || !String(engineNo).trim()) missing.push("Engine No");
    if (!chassisNo || !String(chassisNo).trim()) missing.push("Chassis No");
    if (!rcDate || !String(rcDate).trim()) missing.push("RC Date");
    if (seatingCapacity === undefined || seatingCapacity === null || seatingCapacity === "" || !(Number(seatingCapacity) > 0)) {
      missing.push("Seating Capacity");
    }
    if (!usage || !String(usage).trim()) missing.push("Usage");
    if (!vehicleType || !String(vehicleType).trim()) missing.push("Vehicle Type");
    if (!fuelType || !String(fuelType).trim()) missing.push("Fuel Type");
    if (missing.length) {
      return res.status(400).json({
        error: `Cannot add vehicle - missing required field(s): ${missing.join(", ")}.`,
      });
    }
    const typeFuelError = validateVehicleTypeFuel(vehicleType, fuelType);
    if (typeFuelError) return res.status(400).json({ error: typeFuelError });
    const vehicle = {
      id: uid("v"),
      reg,
      route: route || "",
      routeId: null,
      usage: usage || "fixed_route",
      clientId: null,
      siteId: null,
      supervisorId: null,
      driverId: null,
      driverAssignedDate: null,
      make: make || "",
      model: model || "",
      engineNo: engineNo || "",
      chassisNo: chassisNo || "",
      rcDate: rcDate || "",
      seatingCapacity: seatingCapacity != null && seatingCapacity !== "" ? Number(seatingCapacity) : null,
      vehicleType,
      fuelType,
      standardMileage: Number(standardMileage) || 4.0,
      lastOdometer: null,
      docs: {
        RC: { number: "", expiry: computeRcExpiry(rcDate), updatedAt: null, copyUrl: null },
        Permit: { number: "", expiry: "", isStatePermit: false, districtCount: 0, districtNames: "", updatedAt: null, copyUrl: null },
        Insurance: { number: "", expiry: "" },
        Fitness: { number: "", expiry: "", updatedAt: null, copyUrl: null },
        Tax: { number: "", expiry: "", updatedAt: null, copyUrl: null },
        PUC: { number: "", expiry: "", updatedAt: null, copyUrl: null },
        TemporaryPermit: { number: "", expiry: "", updatedAt: null, copyUrl: null },
        BorderTax: { number: "", expiry: "", updatedAt: null, copyUrl: null },
      },
    };
    db.vehicles.push(vehicle);
    await audit(req.user, "create_vehicle", `${req.user.name} added vehicle ${reg}${make||model? ` (${make||''} ${model||''})`.trim() : ""}`);
    res.json(vehicle);
  })
);

// Onboarding details - Make/Model/Engine No/Chassis No/RC Date/Seating
// Capacity. Separate from /assign (site/driver/client) and /docs (document
// numbers & expiries) since it's a distinct step in bringing on a vehicle,
// normally done by Data Team.
app.patch(
  "/api/vehicles/:id/details",
  requireAuth,
  requireRole(...ADMIN_ROLES),
  h(async (req, res) => {
    const vehicle = db.vehicles.find((v) => v.id === req.params.id);
    if (!vehicle) return res.status(404).json({ error: "Vehicle not found." });
    const before = { make: vehicle.make, model: vehicle.model, engineNo: vehicle.engineNo, chassisNo: vehicle.chassisNo, rcDate: vehicle.rcDate, seatingCapacity: vehicle.seatingCapacity, vehicleType: vehicle.vehicleType, fuelType: vehicle.fuelType };
    const { make, model, engineNo, chassisNo, rcDate, seatingCapacity, vehicleType, fuelType } = req.body || {};
    // Vehicle Type / Fuel Type are validated together (Bus => Diesel only),
    // using whichever value is being kept if only one of the two changed.
    if (vehicleType !== undefined || fuelType !== undefined) {
      const resultType = vehicleType !== undefined ? vehicleType : vehicle.vehicleType;
      const resultFuel = fuelType !== undefined ? fuelType : vehicle.fuelType;
      const typeFuelError = validateVehicleTypeFuel(resultType, resultFuel);
      if (typeFuelError) return res.status(400).json({ error: typeFuelError });
      vehicle.vehicleType = resultType;
      vehicle.fuelType = resultFuel;
    }
    if (make !== undefined) {
      if (!String(make).trim()) return res.status(400).json({ error: "Make cannot be blank." });
      vehicle.make = make;
    }
    if (model !== undefined) {
      if (!String(model).trim()) return res.status(400).json({ error: "Model cannot be blank." });
      vehicle.model = model;
    }
    if (engineNo !== undefined) {
      if (!String(engineNo).trim()) return res.status(400).json({ error: "Engine No cannot be blank." });
      vehicle.engineNo = engineNo;
    }
    if (chassisNo !== undefined) {
      if (!String(chassisNo).trim()) return res.status(400).json({ error: "Chassis No cannot be blank." });
      vehicle.chassisNo = chassisNo;
    }
    if (rcDate !== undefined) {
      if (!String(rcDate).trim()) return res.status(400).json({ error: "RC Date cannot be blank." });
      vehicle.rcDate = rcDate;
      // RC Date is the registration date, not an expiry - RC expiry is
      // always derived as 15 years from registration, recomputed here
      // whenever the registration date changes.
      vehicle.docs.RC.expiry = computeRcExpiry(vehicle.rcDate);
      vehicle.docs.RC.updatedAt = nowIso();
    }
    if (seatingCapacity !== undefined) {
      if (seatingCapacity === "" || seatingCapacity == null || !(Number(seatingCapacity) > 0)) {
        return res.status(400).json({ error: "Seating Capacity must be a positive number." });
      }
      vehicle.seatingCapacity = Number(seatingCapacity);
    }
    const after = { make: vehicle.make, model: vehicle.model, engineNo: vehicle.engineNo, chassisNo: vehicle.chassisNo, rcDate: vehicle.rcDate, seatingCapacity: vehicle.seatingCapacity, vehicleType: vehicle.vehicleType, fuelType: vehicle.fuelType };
    await audit(req.user, "update_vehicle_details", `${req.user.name} updated onboarding details for ${vehicle.reg}: ${JSON.stringify(before)} -> ${JSON.stringify(after)}`);
    res.json(vehicle);
  })
);

// Assignment - Ops Manager / Owner / Data Team can assign a vehicle (and
// its driver) to a site + supervisor + client.
app.patch(
  "/api/vehicles/:id/assign",
  requireAuth,
  requireRole(...ADMIN_ROLES),
  h(async (req, res) => {
    const vehicle = db.vehicles.find((v) => v.id === req.params.id);
    if (!vehicle) return res.status(404).json({ error: "Vehicle not found." });
    const before = { siteId: vehicle.siteId, supervisorId: vehicle.supervisorId, driverId: vehicle.driverId, clientId: vehicle.clientId, routeId: vehicle.routeId, driverAssignedDate: vehicle.driverAssignedDate };
    const { siteId, supervisorId, driverId, clientId, routeId, usage, standardMileage, driverAssignedDate } = req.body || {};
    if (siteId !== undefined) vehicle.siteId = siteId || null;
    if (supervisorId !== undefined) vehicle.supervisorId = supervisorId || null;
    if (driverId !== undefined) {
      const driverChanged = vehicle.driverId !== (driverId || null);
      vehicle.driverId = driverId || null;
      // A new driver being put on this vehicle - stamp the date it happened
      // (today, unless the caller explicitly sent a date in the same request,
      // e.g. to backdate a correction) so "Driver since" has something to show.
      if (driverChanged) vehicle.driverAssignedDate = driverAssignedDate || (vehicle.driverId ? todayStr() : null);
    } else if (driverAssignedDate !== undefined) {
      // Manual edit of the date only, without changing the driver.
      vehicle.driverAssignedDate = driverAssignedDate || null;
    }
    if (clientId !== undefined) {
      const clientChanged = vehicle.clientId !== (clientId || null);
      vehicle.clientId = clientId || null;
      // A route belongs to a single client's catalog - switching the
      // client invalidates whatever route was picked before, unless the
      // caller also sent a new routeId in this same request (handled below).
      if (clientChanged && routeId === undefined) {
        vehicle.routeId = null;
        vehicle.route = "";
      }
    }
    if (routeId !== undefined) {
      if (!routeId) {
        vehicle.routeId = null;
        vehicle.route = "";
      } else {
        const client = db.clients.find((c) => c.id === vehicle.clientId);
        const routeEntry = client ? (client.routes || []).find((r) => r.id === routeId) : null;
        if (!routeEntry) {
          return res.status(400).json({ error: "Select a company for this vehicle first, then choose one of its Route IDs." });
        }
        vehicle.routeId = routeEntry.id;
        vehicle.route = routeEntry.routeNumber || "";
      }
    }
    if (usage !== undefined) vehicle.usage = usage;
    if (standardMileage !== undefined) vehicle.standardMileage = Number(standardMileage) || vehicle.standardMileage;
    await audit(
      req.user,
      "assign_vehicle",
      `${req.user.name} reassigned ${vehicle.reg}: ${JSON.stringify(before)} -> ${JSON.stringify({ siteId: vehicle.siteId, supervisorId: vehicle.supervisorId, driverId: vehicle.driverId, clientId: vehicle.clientId, routeId: vehicle.routeId, driverAssignedDate: vehicle.driverAssignedDate })}`
    );
    res.json(vehicle);
  })
);

// Usage (Fixed / Client on demand / Booking) is deliberately editable by a
// wider set of roles than the rest of the vehicle assignment - Site
// Supervisors, Area Supervisors and the Bookings Department all need to be
// able to flip this without going through Operations Manager. Supervisors
// are still scoped to vehicles they can already see elsewhere in the app.
const USAGE_EDITOR_ROLES = [...ADMIN_ROLES, "site_supervisor", "area_supervisor", "bookings"];
const VALID_USAGE_VALUES = ["fixed_route", "client_on_demand", "booking"];
app.patch(
  "/api/vehicles/:id/usage",
  requireAuth,
  requireRole(...USAGE_EDITOR_ROLES),
  h(async (req, res) => {
    const vehicle = db.vehicles.find((v) => v.id === req.params.id);
    if (!vehicle) return res.status(404).json({ error: "Vehicle not found." });
    if (["site_supervisor", "area_supervisor"].includes(req.user.role)) {
      const canSee = visibleVehiclesFor(req.user).some((v) => v.id === vehicle.id);
      if (!canSee) return res.status(403).json({ error: "You can only change the usage of vehicles assigned to you." });
    }
    const { usage } = req.body || {};
    if (!VALID_USAGE_VALUES.includes(usage)) {
      return res.status(400).json({ error: "Usage must be Fixed, Client on demand, or Booking." });
    }
    const before = vehicle.usage;
    vehicle.usage = usage;
    await audit(req.user, "update_vehicle_usage", `${req.user.name} changed ${vehicle.reg} usage: ${before} -> ${usage}`);
    res.json(vehicle);
  })
);

// Lets the currently-logged-in user re-confirm their own password
// mid-session - used as a "step-up" confirmation (not a separate login)
// before letting the Owner edit Assignments data or confirm a data
// restore, so an accidental tap can't silently change a vehicle/route/
// supervisor or wipe the fleet's data. Never exposes the stored password
// hash to the client.
app.post(
  "/api/verify-password",
  requireAuth,
  h(async (req, res) => {
    const { password } = req.body || {};
    // Deliberately always a 200 (never 401) - the frontend's generic api()
    // helper treats any 401 as "session expired" and force-logs-out, which
    // would be wrong here: an incorrect password re-entry is an expected,
    // normal outcome, not an invalid/expired session.
    if (!verifyPassword(password, req.user.passwordHash)) {
      return res.json({ ok: false, error: "Incorrect password." });
    }
    res.json({ ok: true });
  })
);

app.patch(
  "/api/vehicles/:id/docs",
  requireAuth,
  h(async (req, res) => {
    const vehicle = db.vehicles.find((v) => v.id === req.params.id);
    if (!vehicle) return res.status(404).json({ error: "Vehicle not found." });
    const isOwnerSupervisor = req.user.role === "site_supervisor" && vehicle.supervisorId === req.user.id;
    if (!isOwnerSupervisor && !ADMIN_ROLES.includes(req.user.role)) {
      return res.status(403).json({ error: "Not allowed to edit this vehicle's documents." });
    }
    const { docType, number, expiry, isStatePermit, districtCount, districtNames, copy } = req.body || {};
    if (!vehicle.docs[docType]) return res.status(400).json({ error: "Unknown document type." });
    // Merge onto the existing doc object rather than replacing it wholesale,
    // so setting one field (e.g. just the expiry, from the daily checklist)
    // never wipes out Permit's district details or vice versa.
    const doc = vehicle.docs[docType];
    let changed = false;

    if (docType === "RC" && expiry !== undefined) {
      return res.status(400).json({
        error: "RC expiry is calculated automatically as 15 years from the RC Date - edit RC Date in Vehicle Details instead.",
      });
    }

    if (number !== undefined) {
      doc.number = number;
      changed = true;
    }
    if (expiry !== undefined) {
      doc.expiry = expiry;
      changed = true;
    }
    if (docType === "Permit") {
      // A permit is either a State Permit or a district permit, never both -
      // checking State Permit clears any district details, and a district
      // permit must name 1 or 2 districts (use State Permit for anything
      // wider than that).
      const resultStatePermit = isStatePermit !== undefined ? !!isStatePermit : doc.isStatePermit;
      if (resultStatePermit) {
        if (isStatePermit !== undefined || districtCount !== undefined || districtNames !== undefined) {
          doc.isStatePermit = true;
          doc.districtCount = 0;
          doc.districtNames = "";
          changed = true;
        }
      } else {
        if (isStatePermit !== undefined) {
          doc.isStatePermit = false;
          changed = true;
        }
        if (districtCount !== undefined) {
          const dc = Number(districtCount);
          if (!Number.isFinite(dc) || dc < 1 || dc > 2) {
            return res.status(400).json({
              error: "Number of districts must be 1 or 2 for a district permit - use State Permit instead if it covers more.",
            });
          }
          doc.districtCount = dc;
          changed = true;
        }
        if (districtNames !== undefined) {
          doc.districtNames = districtNames || "";
          changed = true;
        }
      }
    }
    if (DOC_COPY_TYPES.includes(docType) && copy) {
      const url = await savePhoto(copy, "doc_" + docType);
      if (url) {
        doc.copyUrl = url;
        changed = true;
      }
    }
    if (changed) doc.updatedAt = nowIso();

    await audit(req.user, "update_doc", `${req.user.name} updated ${docType} on ${vehicle.reg}`);
    res.json(vehicle);
  })
);

// ---------- RECORDS (daily checklist / breathalyzer / attendance / fuel) ----------
function recordKey(vehicleId, date) {
  return vehicleId + "_" + date;
}

app.get("/api/records", requireAuth, (req, res) => {
  const date = req.query.date || todayStr();
  const visibleIds = new Set(visibleVehiclesFor(req.user).map((v) => v.id));
  const out = {};
  Object.values(db.records).forEach((r) => {
    if (r.date === date && visibleIds.has(r.vehicleId)) out[r.vehicleId] = r;
  });
  res.json(out);
});

// Every vehicle fuelled on a given day, plus that day's total fuel spend
// broken out by station brand (IOCL/BPCL/HPCL/Other) and by payment mode
// (Cash/UPI/Card/Other) - Site Supervisor, Area Supervisor, Operations
// Manager and Owner all use this to get a one-glance picture of the day's
// fuel filling, scoped to whichever vehicles they can already see elsewhere.
// Shared by the on-screen overview and the CSV/XLSX download below, so both
// always show exactly the same numbers.
const FUEL_OVERVIEW_ROLES = ["site_supervisor", "area_supervisor", "ops_manager", "owner"];
function computeFuelOverviewRows(user, date) {
  const visibleIds = new Set(visibleVehiclesFor(user).map((v) => v.id));
  const rows = [];
  Object.values(db.records).forEach((r) => {
    if (r.date !== date || !visibleIds.has(r.vehicleId)) return;
    if (!r.fuel || !(r.fuel.litres > 0)) return;
    const vehicle = db.vehicles.find((v) => v.id === r.vehicleId);
    rows.push({
      vehicleId: r.vehicleId,
      reg: vehicle ? vehicle.reg : r.vehicleId,
      mileage: r.fuel.mileage,
      standardMileage: vehicle ? vehicle.standardMileage : null,
      belowStandard: r.fuel.belowStandard,
      litres: r.fuel.litres,
      distance: r.fuel.distance,
      totalCost: r.fuel.totalCost,
      station: r.fuel.station || "Unspecified",
      paymentMode: r.fuel.paymentMode || "Unspecified",
    });
  });
  rows.sort((a, b) => a.mileage - b.mileage);
  return rows;
}
app.get("/api/fuel-overview", requireAuth, requireRole(...FUEL_OVERVIEW_ROLES), (req, res) => {
  const date = req.query.date || todayStr();
  const rows = computeFuelOverviewRows(req.user, date);
  const byStation = {};
  const byPaymentMode = {};
  let totalLitres = 0;
  let totalCost = 0;
  rows.forEach((r) => {
    byStation[r.station] = (byStation[r.station] || 0) + r.totalCost;
    byPaymentMode[r.paymentMode] = (byPaymentMode[r.paymentMode] || 0) + r.totalCost;
    totalLitres += r.litres;
    totalCost += r.totalCost;
  });
  res.json({ date, rows, byStation, byPaymentMode, totalLitres: Math.round(totalLitres * 10) / 10, totalCost, fillCount: rows.length });
});

// Same vehicle+driver-for-the-day picture as the Assignments date view,
// server-side, so the download always matches what's on screen (including
// showing a temporary driver in place of the regular one, when arranged).
function computeVehiclesRunningRows(user, date) {
  const visibleVehicles = visibleVehiclesFor(user);
  const visibleIds = new Set(visibleVehicles.map((v) => v.id));
  const recordsForDate = {};
  Object.values(db.records).forEach((r) => {
    if (r.date === date && visibleIds.has(r.vehicleId)) recordsForDate[r.vehicleId] = r;
  });
  return visibleVehicles.map((v) => {
    const r = recordsForDate[v.id];
    const temp = r && r.attendance && r.attendance.tempDriver && r.attendance.tempDriver.arranged ? r.attendance.tempDriver : null;
    const site = v.siteId ? db.sites.find((s) => s.id === v.siteId) : null;
    const regularDriver = v.driverId ? db.drivers.find((d) => d.id === v.driverId) : null;
    const supervisorUser = v.supervisorId ? db.users.find((u) => u.id === v.supervisorId) : null;
    const routeInfo = vehicleRouteInfo(v);
    return {
      reg: v.reg,
      company: v.clientId ? clientNameServer(v.clientId) : "Internal / Fixed Route",
      areaSupervisor: site && site.areaSupervisorId ? userNameServer(site.areaSupervisorId) : "",
      siteSupervisor: v.supervisorId ? userNameServer(v.supervisorId) : "",
      siteSupervisorMobile: supervisorUser ? supervisorUser.phone || "" : "",
      driver: temp ? temp.name : (r && r.attendance && r.attendance.driver) || (regularDriver ? regularDriver.name : ""),
      driverMobile: temp ? temp.phone : regularDriver ? regularDriver.phone || "" : "",
      route: routeInfo.routeNumber || (site ? site.name : ""),
      routeStartingPoint: routeInfo.startingPoint,
      isTemporary: !!temp ? "YES" : "",
      tempAmount: temp ? temp.amount : "",
    };
  });
}
app.get("/api/reports/vehicles-running", requireAuth, requireRole(...FUEL_OVERVIEW_ROLES), (req, res) => {
  const date = req.query.date || todayStr();
  const rows = computeVehiclesRunningRows(req.user, date);
  const columns = [
    { label: "Vehicle", key: "reg" },
    { label: "Company", key: "company" },
    { label: "Area Supervisor", key: "areaSupervisor" },
    { label: "Site Supervisor", key: "siteSupervisor" },
    { label: "Site Supervisor Mobile", key: "siteSupervisorMobile" },
    { label: "Driver", key: "driver" },
    { label: "Driver Mobile", key: "driverMobile" },
    { label: "Route ID", key: "route" },
    { label: "Route Starting Point", key: "routeStartingPoint" },
    { label: "Temporary Driver", key: "isTemporary" },
    { label: "Temp Driver Amount", key: "tempAmount" },
  ];
  sendDownload(res, rows, columns, `vehicles_running_${date}`, req.query.format);
});
app.get("/api/reports/fuel-overview", requireAuth, requireRole(...FUEL_OVERVIEW_ROLES), (req, res) => {
  const date = req.query.date || todayStr();
  const rows = computeFuelOverviewRows(req.user, date);
  const columns = [
    { label: "Vehicle", key: "reg" },
    { label: "Mileage (km/l)", key: "mileage" },
    { label: "Standard Mileage", key: "standardMileage" },
    { label: "Below Standard", key: (r) => (r.belowStandard ? "YES" : "") },
    { label: "Litres", key: "litres" },
    { label: "Distance (km)", key: "distance" },
    { label: "Cost", key: "totalCost" },
    { label: "Station", key: "station" },
    { label: "Payment Mode", key: "paymentMode" },
  ];
  sendDownload(res, rows, columns, `fuel_overview_${date}`, req.query.format);
});

app.post(
  "/api/records",
  requireAuth,
  requireRole("site_supervisor"),
  h(async (req, res) => {
    const body = req.body || {};
    const vehicle = db.vehicles.find((v) => v.id === body.vehicleId);
    if (!vehicle) return res.status(400).json({ error: "Unknown vehicle." });
    if (vehicle.supervisorId !== req.user.id) {
      return res.status(403).json({ error: "This vehicle is not allotted to you." });
    }
    const date = body.date || todayStr();

    // photos: expect { tag: [dataUrl, ...] }
    const photoTags = {};
    if (body.photoTags) {
      for (const [tag, arr] of Object.entries(body.photoTags)) {
        const list = Array.isArray(arr) ? arr : [];
        photoTags[tag] = [];
        for (const src of list) {
          if (typeof src === "string" && src.startsWith("data:")) {
            photoTags[tag].push((await savePhoto(src, tag)) || src);
          } else {
            photoTags[tag].push(src); // already a saved URL
          }
        }
      }
    }

    // server-authoritative mileage calculation
    const f = body.fuel || {};
    const previousOdometer = vehicle.lastOdometer != null ? vehicle.lastOdometer : Number(f.odometer) || 0;
    const odometer = Number(f.odometer) || previousOdometer;
    const litres = Number(f.litres) || 0;
    const fuelPrice = Number(f.fuelPrice) || 0;
    const distance = Math.max(odometer - previousOdometer, 0);
    const mileage = litres > 0 ? Math.round((distance / litres) * 10) / 10 : 0;
    const belowStandard = litres > 0 && distance > 0 && mileage < vehicle.standardMileage;

    // A receipt photo is mandatory any time fuel was actually filled - not
    // required on days with no fuel entry (litres left at 0).
    const fuelReceiptUrls = photoTags.fuelReceipt || [];
    if (litres > 0 && fuelReceiptUrls.length === 0) {
      return res.status(400).json({ error: "A photo of the fuel receipt is required whenever fuel is filled." });
    }

    // Temporary driver: if the supervisor recorded a paid replacement for an
    // absent driver, a payment-proof photo is mandatory, and submitting the
    // checklist automatically raises an expense request for that amount
    // under this supervisor's name (same "pending, needs Ops approval" flow
    // as any other expense) - noted as covering for the regular driver, so
    // whoever runs payroll knows to deduct it from that driver's pay. The
    // app itself doesn't have a payroll/salary module, so that deduction is
    // a manual step outside this app - this just gives a clear paper trail.
    const attendance = body.attendance || {};
    const tempDriver = attendance.tempDriver || {};
    const tempDriverPaymentUrls = photoTags.tempDriverPayment || [];
    const tempAmount = Number(tempDriver.amount) || 0;
    let autoExpense = null;
    if (attendance.status === "absent" && tempDriver.arranged && tempAmount > 0) {
      if (tempDriverPaymentUrls.length === 0) {
        return res.status(400).json({ error: "A photo of the payment (UPI screenshot) is required when paying a temporary driver." });
      }
      if (!tempDriver.name || !String(tempDriver.name).trim()) {
        return res.status(400).json({ error: "The temporary driver's name is required." });
      }
      const regularDriver = db.drivers.find((d) => d.id === vehicle.driverId);
      autoExpense = {
        id: uid("e"),
        userId: req.user.id,
        userName: req.user.name,
        vehicleId: vehicle.id,
        category: "Temporary Driver",
        amount: tempAmount,
        description:
          `Temporary driver ${tempDriver.name}${tempDriver.phone ? " (" + tempDriver.phone + ")" : ""}` +
          `${tempDriver.upiId ? ", UPI: " + tempDriver.upiId : ""} covered for ${regularDriver ? regularDriver.name : "the regular driver"} ` +
          `on ${vehicle.reg} (${date}) - to be deducted from ${regularDriver ? regularDriver.name : "the regular driver"}'s pay.`,
        billUrl: tempDriverPaymentUrls[0],
        status: "pending",
        decidedBy: null,
        comment: "",
        submittedAt: nowIso(),
        autoGenerated: true,
        coveredForDriverId: vehicle.driverId || null,
        coveredForDriverName: regularDriver ? regularDriver.name : null,
        tempDriver: { name: tempDriver.name, phone: tempDriver.phone || "", upiId: tempDriver.upiId || "" },
      };
      db.expenses.unshift(autoExpense);
      await audit(
        req.user,
        "expense_request",
        `${req.user.name} requested ₹${autoExpense.amount} (Temporary Driver, auto-created) - covering for ${autoExpense.coveredForDriverName || "unknown driver"} on ${vehicle.reg}`
      );
    }

    vehicle.lastOdometer = odometer;

    const record = {
      vehicleId: vehicle.id,
      date,
      safety: body.safety || {},
      condition: body.condition || {},
      photoTags,
      breathalyzer: body.breathalyzer || {},
      attendance: Object.assign({}, attendance, {
        tempDriver: tempDriver.arranged
          ? { arranged: true, name: tempDriver.name || "", phone: tempDriver.phone || "", upiId: tempDriver.upiId || "", amount: tempAmount, expenseId: autoExpense ? autoExpense.id : null }
          : { arranged: false },
      }),
      fuel: {
        previousOdometer,
        odometer,
        fuelPrice,
        litres,
        fillLevel: f.fillLevel || "other",
        station: litres > 0 ? f.station || "" : "",
        paymentMode: litres > 0 ? f.paymentMode || "" : "",
        distance,
        mileage,
        totalCost: Math.round(litres * fuelPrice),
        belowStandard,
        receiptUrls: fuelReceiptUrls,
      },
      verification: { status: "pending", verifiedBy: null, comment: "" },
      submittedBy: req.user.id,
      submittedAt: nowIso(),
    };
    db.records[recordKey(vehicle.id, date)] = record;
    await audit(req.user, "submit_record", `${req.user.name} submitted daily record for ${vehicle.reg} (${date})`);
    res.json(Object.assign({}, record, { autoExpense }));
  })
);

app.patch(
  "/api/records/:vehicleId/:date/verify",
  requireAuth,
  requireRole("area_supervisor", ...ADMIN_ROLES),
  h(async (req, res) => {
    const key = recordKey(req.params.vehicleId, req.params.date);
    const record = db.records[key];
    if (!record) return res.status(404).json({ error: "Record not found." });
    const { status, comment } = req.body || {};
    if (!["approved", "rejected"].includes(status)) return res.status(400).json({ error: "status must be approved or rejected." });
    record.verification = { status, verifiedBy: req.user.id, comment: comment || "" };
    await audit(req.user, status === "approved" ? "approve" : "reject", `${req.user.name} ${status} ${req.params.vehicleId} (${req.params.date})${comment ? ": " + comment : ""}`);
    res.json(record);
  })
);

app.patch(
  "/api/records/:vehicleId/:date/correct",
  requireAuth,
  requireRole("data_team", "owner"),
  h(async (req, res) => {
    const key = recordKey(req.params.vehicleId, req.params.date);
    const record = db.records[key];
    if (!record) return res.status(404).json({ error: "Record not found." });
    const { path: fieldPath, value } = req.body || {};
    if (!fieldPath) return res.status(400).json({ error: "path is required, e.g. fuel.odometer" });
    const parts = fieldPath.split(".");
    let obj = record;
    for (let i = 0; i < parts.length - 1; i++) obj = obj[parts[i]];
    const before = obj[parts[parts.length - 1]];
    obj[parts[parts.length - 1]] = value;
    await audit(req.user, "correction", `${req.user.name} corrected ${fieldPath} on ${req.params.vehicleId} (${req.params.date}): ${before} -> ${value}`);
    res.json(record);
  })
);

// ---------- DAY-START / DAY-CLOSE ODOMETER (separate from the fuel entry) ----------
// Site supervisors log two odometer readings per vehicle per day - one when
// the vehicle starts its day, one when it closes - independent of whether
// fuel was filled that day. These are saved immediately (like clock in/out),
// not held in a draft.
app.get("/api/odometer", requireAuth, (req, res) => {
  const date = req.query.date || todayStr();
  const visibleIds = new Set(visibleVehiclesFor(req.user).map((v) => v.id));
  const out = {};
  Object.values(db.odometerLogs).forEach((o) => {
    if (o.date === date && visibleIds.has(o.vehicleId)) out[o.vehicleId] = o;
  });
  res.json(out);
});

app.post(
  "/api/odometer/:vehicleId/:type(start|close)",
  requireAuth,
  requireRole("site_supervisor"),
  h(async (req, res) => {
    const vehicle = db.vehicles.find((v) => v.id === req.params.vehicleId);
    if (!vehicle) return res.status(400).json({ error: "Unknown vehicle." });
    if (vehicle.supervisorId !== req.user.id) {
      return res.status(403).json({ error: "This vehicle is not allotted to you." });
    }
    const reading = Number((req.body || {}).value);
    if (!Number.isFinite(reading) || reading < 0) {
      return res.status(400).json({ error: "Enter a valid odometer reading." });
    }
    const date = todayStr();
    const key = recordKey(vehicle.id, date);
    const entry = db.odometerLogs[key] || { vehicleId: vehicle.id, date, dayStart: null, dayClose: null };
    if (req.params.type === "start") {
      entry.dayStart = { value: reading, ts: nowIso(), by: req.user.id };
    } else {
      if (entry.dayStart && reading < entry.dayStart.value) {
        return res.status(400).json({
          error: `Day-close reading (${reading} km) can't be less than the day-start reading (${entry.dayStart.value} km).`,
        });
      }
      entry.dayClose = { value: reading, ts: nowIso(), by: req.user.id };
    }
    db.odometerLogs[key] = entry;
    await audit(
      req.user,
      req.params.type === "start" ? "odometer_day_start" : "odometer_day_close",
      `${req.user.name} logged ${req.params.type === "start" ? "day-start" : "day-close"} odometer ${reading} km for ${vehicle.reg}`
    );
    res.json(entry);
  })
);

// ---------- LEAVES ----------
app.get("/api/leaves", requireAuth, (req, res) => res.json(db.leaves));

app.post(
  "/api/leaves",
  requireAuth,
  h(async (req, res) => {
    const { driverId, driverName, type, start, end } = req.body || {};
    if (!driverName || !start || !end) return res.status(400).json({ error: "driverName, start, end are required." });
    const leave = { id: uid("l"), driverId: driverId || null, driver: driverName, type: type || "Planned", start, end, status: "pending", requestedBy: req.user.id };
    db.leaves.push(leave);
    await audit(req.user, "leave_request", `${req.user.name} requested leave for ${driverName} (${start} to ${end})`);
    res.json(leave);
  })
);

app.patch(
  "/api/leaves/:id",
  requireAuth,
  requireRole("ops_manager", "owner"),
  h(async (req, res) => {
    const leave = db.leaves.find((l) => l.id === req.params.id);
    if (!leave) return res.status(404).json({ error: "Leave not found." });
    const { status } = req.body || {};
    if (!["approved", "rejected"].includes(status)) return res.status(400).json({ error: "status must be approved or rejected." });
    leave.status = status;
    await audit(req.user, "leave_decision", `${req.user.name} set leave ${leave.id} (${leave.driver}) to ${status}`);
    res.json(leave);
  })
);

// ---------- EXPENSES (supervisor reimbursement requests) ----------
// Every expense needs a photo of the bill - no exceptions - and is approved
// by the Operations Manager before it's marked payable. This app only
// tracks and approves the request; it never moves money itself.
function visibleExpensesFor(user) {
  if (user.role === "site_supervisor") return db.expenses.filter((e) => e.userId === user.id);
  if (user.role === "area_supervisor") {
    const mine = new Set(user.supervises || []);
    return db.expenses.filter((e) => mine.has(e.userId));
  }
  return db.expenses; // ops_manager, owner, data_team, hr(read-only via audit) see all
}

app.get("/api/expenses", requireAuth, (req, res) => {
  res.json(visibleExpensesFor(req.user));
});

const EXPENSE_DECIDE_ROLES = ["ops_manager", "owner"];

app.post(
  "/api/expenses",
  requireAuth,
  h(async (req, res) => {
    const { amount, category, description, vehicleId, isOfficeExpenditure, bill } = req.body || {};
    if (!amount || Number(amount) <= 0) return res.status(400).json({ error: "A valid amount is required." });
    if (!bill) return res.status(400).json({ error: "A photo of the bill/receipt is required for every expense request." });
    const officeExpenditure = !!isOfficeExpenditure;
    if (!vehicleId && !officeExpenditure) {
      return res.status(400).json({ error: "Select the vehicle this expense is for, or mark it as Office Expenditure." });
    }
    const billUrl = await savePhoto(bill, "bill");
    if (!billUrl) return res.status(400).json({ error: "Could not read the bill photo - please try again." });
    // Every vehicle-tagged expense is snapshotted with the vehicle's Site at
    // the moment it's filed - not looked up live later - so cost-by-site
    // reporting stays accurate even if the vehicle moves sites afterward.
    // Office Expenditure has no vehicle/site of its own; it gets divided
    // across every site when computing total operational cost per site (see
    // /api/reports/site-cost-analysis).
    const expenseVehicle = !officeExpenditure && vehicleId ? db.vehicles.find((v) => v.id === vehicleId) : null;
    const expense = {
      id: uid("e"),
      userId: req.user.id,
      userName: req.user.name,
      vehicleId: officeExpenditure ? null : vehicleId,
      siteId: expenseVehicle ? expenseVehicle.siteId || null : null,
      isOfficeExpenditure: officeExpenditure,
      category: category || "Other",
      amount: Number(amount),
      description: description || "",
      billUrl,
      status: "pending",
      decidedBy: null,
      comment: "",
      paymentStage: null, // set once approved: "approved" -> "in_process" -> "completed"
      utrNumber: null,
      paymentUpdatedAt: null,
      submittedAt: nowIso(),
    };
    db.expenses.unshift(expense);
    await audit(req.user, "expense_request", `${req.user.name} requested ₹${expense.amount} (${expense.category})`);
    res.json(expense);
  })
);

app.patch(
  "/api/expenses/:id/decide",
  requireAuth,
  requireRole(...EXPENSE_DECIDE_ROLES),
  h(async (req, res) => {
    const expense = db.expenses.find((e) => e.id === req.params.id);
    if (!expense) return res.status(404).json({ error: "Expense not found." });
    const { status, comment } = req.body || {};
    if (!["approved", "rejected"].includes(status)) return res.status(400).json({ error: "status must be approved or rejected." });
    if (status === "rejected" && !(comment || "").trim()) return res.status(400).json({ error: "A reason is required when rejecting an expense." });
    expense.status = status;
    expense.decidedBy = req.user.id;
    expense.comment = comment || "";
    expense.paymentStage = status === "approved" ? "approved" : null;
    await audit(req.user, "expense_decision", `${req.user.name} ${status} ₹${expense.amount} expense from ${expense.userName}${comment ? ": " + comment : ""}`);
    res.json(expense);
  })
);

app.patch(
  "/api/expenses/:id/payment-status",
  requireAuth,
  requireRole(...EXPENSE_DECIDE_ROLES),
  h(async (req, res) => {
    const expense = db.expenses.find((e) => e.id === req.params.id);
    if (!expense) return res.status(404).json({ error: "Expense not found." });
    if (expense.status !== "approved") return res.status(400).json({ error: "Only approved expenses can have their payment status updated." });
    const { paymentStage, utrNumber } = req.body || {};
    if (!["in_process", "completed"].includes(paymentStage)) return res.status(400).json({ error: "paymentStage must be in_process or completed." });
    if (paymentStage === "completed" && !(utrNumber || "").trim()) return res.status(400).json({ error: "A UTR Number is required to mark this payment as completed." });
    expense.paymentStage = paymentStage;
    if (paymentStage === "completed") expense.utrNumber = utrNumber.trim();
    expense.paymentUpdatedAt = nowIso();
    await audit(req.user, "expense_payment_status", `${req.user.name} marked ₹${expense.amount} expense from ${expense.userName} as ${paymentStage}${paymentStage === "completed" ? " (UTR " + expense.utrNumber + ")" : ""}`);
    res.json(expense);
  })
);

// ---------- SUPERVISOR ATTENDANCE (selfie + geolocation) ----------
function siteFor(user) {
  return db.sites.find((s) => s.id === user.siteId) || null;
}

app.post(
  "/api/attendance/:type(clockin|clockout)",
  requireAuth,
  h(async (req, res) => {
    const { selfie, lat, lng } = req.body || {};
    const site = siteFor(req.user);
    const selfieUrl = await savePhoto(selfie, "selfie_" + req.params.type);
    const distance = site ? distanceMeters(site.lat, site.lng, lat, lng) : null;
    const withinGeofence = distance != null ? distance <= GEOFENCE_METERS : null;
    const entry = {
      id: uid("att"),
      userId: req.user.id,
      userName: req.user.name,
      role: req.user.role,
      date: todayStr(),
      type: req.params.type,
      ts: nowIso(),
      selfieUrl,
      lat: lat != null ? Number(lat) : null,
      lng: lng != null ? Number(lng) : null,
      siteId: site ? site.id : null,
      siteName: site ? site.name : null,
      distanceMeters: distance,
      withinGeofence,
    };
    db.supervisorAttendance.unshift(entry);
    await audit(
      req.user,
      req.params.type,
      `${req.user.name} ${req.params.type === "clockin" ? "clocked in" : "clocked out"}${site ? " at " + site.name : ""}${distance != null ? ` (${distance}m from site, ${withinGeofence ? "within" : "OUTSIDE"} geofence)` : " (no site location on file)"}`
    );
    res.json(entry);
  })
);

app.get("/api/attendance", requireAuth, requireRole("hr", "owner", "ops_manager", "data_team"), (req, res) => {
  const { date, userId, from, to } = req.query;
  let rows = db.supervisorAttendance;
  if (from || to) {
    const f = from || to;
    const t = to || from;
    rows = rows.filter((r) => r.date >= f && r.date <= t);
  } else if (date) {
    rows = rows.filter((r) => r.date === date);
  }
  if (userId) rows = rows.filter((r) => r.userId === userId);
  res.json(rows);
});

// own attendance history - any logged-in user can see their own
app.get("/api/attendance/me", requireAuth, (req, res) => {
  res.json(db.supervisorAttendance.filter((r) => r.userId === req.user.id).slice(0, 60));
});

// Driver attendance (present/absent + uniform compliance) for a date range -
// unlike supervisor clock-in/out (its own log), a driver's daily attendance
// only lives inside that day's vehicle checklist record, so this walks
// every record in range and pulls it out into its own rows.
function computeDriverAttendanceRows(from, to) {
  const rows = [];
  Object.values(db.records).forEach((r) => {
    if (from && r.date < from) return;
    if (to && r.date > to) return;
    if (!r.attendance || !r.attendance.driver) return;
    const vehicle = db.vehicles.find((v) => v.id === r.vehicleId);
    rows.push({
      date: r.date,
      driver: r.attendance.driver,
      vehicleReg: vehicle ? vehicle.reg : r.vehicleId,
      status: r.attendance.status || "",
      uniform: r.attendance.uniform || "",
    });
  });
  rows.sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : 0));
  return rows;
}
app.get("/api/driver-attendance", requireAuth, requireRole("hr", "owner", "ops_manager", "data_team"), (req, res) => {
  const { date, from, to } = req.query;
  let f = from;
  let t = to;
  if (!f && !t && date) {
    f = date;
    t = date;
  }
  res.json(computeDriverAttendanceRows(f, t));
});

// ---------- ATTENDANCE DOWNLOADS + BACKUP (HR, Owner only) ----------
// HR is the one who actually needs the last-month attendance picture (for
// payroll/discipline review), so unlike the general Reports tab (Owner/Ops
// Manager/Area Supervisor), these downloads are scoped to HR and Owner only.
const ATTENDANCE_DOWNLOAD_ROLES = ["hr", "owner"];
app.get("/api/reports/attendance-supervisors", requireAuth, requireRole(...ATTENDANCE_DOWNLOAD_ROLES), (req, res) => {
  const { from, to, format } = req.query;
  let rows = db.supervisorAttendance;
  if (from || to) {
    const f = from || to;
    const t = to || from;
    rows = rows.filter((r) => r.date >= f && r.date <= t);
  }
  const columns = [
    { label: "Date", key: "date" },
    { label: "Name", key: "userName" },
    { label: "Role", key: (a) => roleLabelServer(a.role) },
    { label: "Type", key: (a) => (a.type === "clockin" ? "Clock In" : "Clock Out") },
    { label: "Time", key: (a) => new Date(a.ts).toLocaleTimeString() },
    { label: "Site", key: (a) => a.siteName || "" },
    { label: "Distance From Site (m)", key: "distanceMeters" },
    { label: "Within Geofence", key: (a) => (a.withinGeofence == null ? "n/a" : a.withinGeofence ? "Yes" : "No") },
  ];
  sendDownload(res, rows, columns, `supervisor_attendance_${from || "all"}_${to || "all"}`, format);
});
app.get("/api/reports/attendance-drivers", requireAuth, requireRole(...ATTENDANCE_DOWNLOAD_ROLES), (req, res) => {
  const { from, to, format } = req.query;
  const rows = computeDriverAttendanceRows(from, to);
  const columns = [
    { label: "Date", key: "date" },
    { label: "Driver", key: "driver" },
    { label: "Vehicle", key: "vehicleReg" },
    { label: "Attendance", key: "status" },
    { label: "Uniform", key: "uniform" },
  ];
  sendDownload(res, rows, columns, `driver_attendance_${from || "all"}_${to || "all"}`, format);
});

// Attendance Backup - separate from the general full-data Backup tab
// (Owner/Ops Manager). HR needs to pull a complete archive of every
// supervisor clock-in/out and driver attendance record roughly every 15
// days and keep it outside the app - every download is logged (who, when)
// so it's visible whether this is actually happening on schedule, the same
// way the full-data backup already tracks itself.
const ATTENDANCE_BACKUP_ROLES = ["hr", "owner"];
app.get(
  "/api/attendance-backup/full",
  requireAuth,
  requireRole(...ATTENDANCE_BACKUP_ROLES),
  h(async (req, res) => {
    const takenAt = nowIso();
    db.attendanceBackupDownloads = db.attendanceBackupDownloads || [];
    db.attendanceBackupDownloads.unshift({ id: uid("attbkdl"), userId: req.user.id, userName: req.user.name, role: req.user.role, takenAt });
    if (db.attendanceBackupDownloads.length > 200) db.attendanceBackupDownloads.length = 200;
    await audit(req.user, "download_attendance_backup", `${req.user.name} downloaded a full attendance backup`);

    const payload = {
      exportedAt: takenAt,
      supervisorAttendance: db.supervisorAttendance,
      driverAttendance: computeDriverAttendanceRows(null, null),
    };
    const filename = `padmasri_attendance_backup_${takenAt.slice(0, 10)}.json`;
    res.setHeader("Content-Type", "application/json");
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.send(JSON.stringify(payload, null, 2));
  })
);
app.get(
  "/api/attendance-backup/history",
  requireAuth,
  requireRole(...ATTENDANCE_BACKUP_ROLES),
  (req, res) => {
    res.json((db.attendanceBackupDownloads || []).slice(0, 100));
  }
);

// ---------- PAYROLL ----------
// Two pieces per person (a "user" = any staff role, or a "driver"):
//   1. A recurring profile (base salary + components like ESI/PF/Gratuity,
//      percent-of-base or a fixed amount, each toggleable) - set up ONCE by
//      HR and from then on applies automatically to every month's payslip
//      with no re-entry needed.
//   2. Ad-hoc entries tied to a specific month (traffic challans, fuel
//      theft deductions, vehicle cleanliness penalties, other penalties,
//      or a one-off bonus) - added by HR as they happen.
// A payslip for a given month is always computed fresh from these two
// pieces, never hand-entered or manually "generated" - that's what makes
// it automatic going forward.
function payrollKey(personType, personId) {
  return `${personType}:${personId}`;
}
function personExists(personType, personId) {
  if (personType === "user") return db.users.some((u) => u.id === personId);
  if (personType === "driver") return db.drivers.some((d) => d.id === personId);
  return false;
}
function personLabel(personType, personId) {
  if (personType === "user") {
    const u = db.users.find((x) => x.id === personId);
    return u ? { name: u.name, subLabel: roleLabelServer(u.role) } : { name: personId, subLabel: "" };
  }
  const d = db.drivers.find((x) => x.id === personId);
  return d ? { name: d.name, subLabel: "Driver" } : { name: personId, subLabel: "" };
}
function getPayrollProfile(personType, personId) {
  return db.payrollProfiles[payrollKey(personType, personId)] || null;
}
// Drivers a Site Supervisor can see payroll for - whichever vehicle(s) are
// currently assigned to them, same scoping rule used everywhere else
// (visibleVehiclesFor), narrowed down to just the driver on each one.
function visibleDriverIdsFor(user) {
  if (user.role === "site_supervisor") {
    return new Set(visibleVehiclesFor(user).map((v) => v.driverId).filter(Boolean));
  }
  return new Set(db.drivers.map((d) => d.id));
}
const PAYROLL_EDIT_ROLES = ["hr"];
const PAYROLL_VIEW_ALL_ROLES = ["hr", "owner"]; // full staff + driver visibility
const PAYROLL_VIEW_ALL_DRIVERS_ROLES = ["hr", "owner", "ops_manager"]; // every driver, no staff
const PAYROLL_VIEW_OWN_DRIVERS_ROLES = ["hr", "owner", "ops_manager", "site_supervisor"]; // site_supervisor narrowed further below
function canViewPayroll(user, personType, personId) {
  if (personType === "user" && personId === user.id) return true; // always allowed to see your own
  if (PAYROLL_VIEW_ALL_ROLES.includes(user.role)) return true;
  if (personType !== "driver") return false; // nobody else gets another staff member's payroll
  if (PAYROLL_VIEW_ALL_DRIVERS_ROLES.includes(user.role)) return true;
  if (user.role === "site_supervisor") return visibleDriverIdsFor(user).has(personId);
  return false;
}
function canEditPayroll(user) {
  return PAYROLL_EDIT_ROLES.includes(user.role);
}
// Computes a month's payslip purely from the recurring profile + whichever
// ad-hoc entries were logged for that exact month - nothing is stored per
// month, so changing the profile automatically reshapes every future
// month without anyone having to re-run or regenerate anything.
function computePayslip(personType, personId, month) {
  const profile = getPayrollProfile(personType, personId);
  const base = profile ? Number(profile.baseSalary) || 0 : 0;
  const earnings = [{ label: "Base Salary", amount: base }];
  const deductions = [];
  (profile && Array.isArray(profile.components) ? profile.components : []).forEach((c) => {
    if (!c.enabled) return;
    const amt = c.mode === "percent" ? Math.round(base * (Number(c.value) || 0)) / 100 : Number(c.value) || 0;
    const rounded = Math.round(amt);
    if (c.type === "earning") earnings.push({ label: c.label, amount: rounded });
    else deductions.push({ label: c.label, amount: rounded, source: "recurring" });
  });
  db.payrollAdhocEntries
    .filter((e) => e.personType === personType && e.personId === personId && e.month === month)
    .forEach((e) => {
      if (e.type === "earning") earnings.push({ label: e.label, amount: e.amount, id: e.id, category: e.category });
      else deductions.push({ label: e.label, amount: e.amount, source: "adhoc", id: e.id, category: e.category });
    });
  const grossPay = earnings.reduce((s, x) => s + x.amount, 0);
  const totalDeductions = deductions.reduce((s, x) => s + x.amount, 0);
  const netPay = grossPay - totalDeductions;
  return { personType, personId, month, hasProfile: !!profile, earnings, deductions, grossPay, totalDeductions, netPay };
}

// People this role is allowed to browse payroll for, beyond their own -
// used to populate the management/view list. Empty for roles that only
// ever see their own payslip (Area Supervisor, Data Team, Bookings).
app.get("/api/payroll/people", requireAuth, (req, res) => {
  const user = req.user;
  const people = [];
  if (PAYROLL_VIEW_ALL_ROLES.includes(user.role)) {
    db.users.forEach((u) => people.push({ personType: "user", personId: u.id, name: u.name, subLabel: roleLabelServer(u.role) }));
    db.drivers.forEach((d) => people.push({ personType: "driver", personId: d.id, name: d.name, subLabel: "Driver" }));
  } else if (PAYROLL_VIEW_ALL_DRIVERS_ROLES.includes(user.role)) {
    db.drivers.forEach((d) => people.push({ personType: "driver", personId: d.id, name: d.name, subLabel: "Driver" }));
  } else if (user.role === "site_supervisor") {
    const ids = visibleDriverIdsFor(user);
    db.drivers.filter((d) => ids.has(d.id)).forEach((d) => people.push({ personType: "driver", personId: d.id, name: d.name, subLabel: "Driver" }));
  }
  res.json(people);
});

app.get("/api/payroll/profile/:personType/:personId", requireAuth, (req, res) => {
  const { personType, personId } = req.params;
  if (!["user", "driver"].includes(personType) || !personExists(personType, personId)) {
    return res.status(404).json({ error: "Person not found." });
  }
  if (!canViewPayroll(req.user, personType, personId)) {
    return res.status(403).json({ error: "You don't have access to this person's payroll." });
  }
  res.json(getPayrollProfile(personType, personId));
});

app.patch(
  "/api/payroll/profile/:personType/:personId",
  requireAuth,
  requireRole(...PAYROLL_EDIT_ROLES),
  h(async (req, res) => {
    const { personType, personId } = req.params;
    if (!["user", "driver"].includes(personType) || !personExists(personType, personId)) {
      return res.status(404).json({ error: "Person not found." });
    }
    const { baseSalary, components } = req.body || {};
    const base = Number(baseSalary);
    if (!Number.isFinite(base) || base < 0) return res.status(400).json({ error: "Base salary must be a non-negative number." });
    if (!Array.isArray(components)) return res.status(400).json({ error: "components must be an array." });
    for (const c of components) {
      if (!c.label || !String(c.label).trim()) return res.status(400).json({ error: "Every component needs a label." });
      if (!["earning", "deduction"].includes(c.type)) return res.status(400).json({ error: "Component type must be earning or deduction." });
      if (!["percent", "fixed"].includes(c.mode)) return res.status(400).json({ error: "Component mode must be percent or fixed." });
      if (!Number.isFinite(Number(c.value)) || Number(c.value) < 0) return res.status(400).json({ error: `Component "${c.label}" needs a non-negative value.` });
    }
    const key = payrollKey(personType, personId);
    db.payrollProfiles[key] = {
      personType,
      personId,
      baseSalary: base,
      components: components.map((c) => ({
        label: String(c.label).trim(),
        type: c.type,
        mode: c.mode,
        value: Number(c.value),
        enabled: c.enabled !== false,
      })),
      updatedBy: req.user.id,
      updatedAt: nowIso(),
    };
    const { name } = personLabel(personType, personId);
    await audit(req.user, "set_payroll_profile", `${req.user.name} set the payroll formula for ${name} (base ₹${base}, ${components.length} component(s))`);
    res.json(db.payrollProfiles[key]);
  })
);

const PAYROLL_ADHOC_CATEGORIES = [
  { key: "traffic_challan", label: "Traffic Challan", type: "deduction" },
  { key: "fuel_theft", label: "Fuel Deduction (Theft)", type: "deduction" },
  { key: "cleanliness_penalty", label: "Vehicle Cleanliness Penalty", type: "deduction" },
  { key: "missed_odometer_entry", label: "Missed Odometer Entry Penalty", type: "deduction" },
  { key: "other_penalty", label: "Other Penalty", type: "deduction" },
  { key: "bonus", label: "Bonus / Incentive", type: "earning" },
  { key: "other_earning", label: "Other Earning", type: "earning" },
];
app.post(
  "/api/payroll/adhoc",
  requireAuth,
  requireRole(...PAYROLL_EDIT_ROLES),
  h(async (req, res) => {
    const { personType, personId, month, category, amount, note } = req.body || {};
    if (!["user", "driver"].includes(personType) || !personExists(personType, personId)) {
      return res.status(404).json({ error: "Person not found." });
    }
    if (!/^\d{4}-\d{2}$/.test(month || "")) return res.status(400).json({ error: "month must be in YYYY-MM format." });
    const cat = PAYROLL_ADHOC_CATEGORIES.find((c) => c.key === category);
    if (!cat) return res.status(400).json({ error: `category must be one of: ${PAYROLL_ADHOC_CATEGORIES.map((c) => c.key).join(", ")}.` });
    const amt = Number(amount);
    if (!Number.isFinite(amt) || amt <= 0) return res.status(400).json({ error: "amount must be a positive number." });
    const entry = {
      id: uid("payadh"),
      personType,
      personId,
      month,
      category: cat.key,
      label: note && String(note).trim() ? `${cat.label} - ${String(note).trim()}` : cat.label,
      type: cat.type,
      amount: Math.round(amt),
      addedBy: req.user.id,
      addedByName: req.user.name,
      addedAt: nowIso(),
    };
    db.payrollAdhocEntries.unshift(entry);
    const { name } = personLabel(personType, personId);
    await audit(req.user, "add_payroll_adhoc", `${req.user.name} added ${cat.label} of ₹${entry.amount} for ${name} (${month})`);
    res.json(entry);
  })
);

app.delete(
  "/api/payroll/adhoc/:id",
  requireAuth,
  requireRole(...PAYROLL_EDIT_ROLES),
  h(async (req, res) => {
    const idx = db.payrollAdhocEntries.findIndex((e) => e.id === req.params.id);
    if (idx === -1) return res.status(404).json({ error: "Entry not found." });
    const [removed] = db.payrollAdhocEntries.splice(idx, 1);
    const { name } = personLabel(removed.personType, removed.personId);
    await audit(req.user, "remove_payroll_adhoc", `${req.user.name} removed ${removed.label} of ₹${removed.amount} for ${name} (${removed.month})`);
    res.json({ ok: true });
  })
);

app.get("/api/payroll/payslip/:personType/:personId", requireAuth, (req, res) => {
  const { personType, personId } = req.params;
  if (!["user", "driver"].includes(personType) || !personExists(personType, personId)) {
    return res.status(404).json({ error: "Person not found." });
  }
  if (!canViewPayroll(req.user, personType, personId)) {
    return res.status(403).json({ error: "You don't have access to this person's payroll." });
  }
  const month = /^\d{4}-\d{2}$/.test(req.query.month || "") ? req.query.month : nowIso().slice(0, 7);
  const payslip = computePayslip(personType, personId, month);
  const { name, subLabel } = personLabel(personType, personId);
  res.json(Object.assign({ name, subLabel }, payslip));
});

// Payroll register for a month - every person this role can see, with
// their computed net pay, so HR/Owner/Ops Manager/Site Supervisor get an
// at-a-glance table instead of opening each person one at a time.
app.get("/api/payroll/summary", requireAuth, (req, res) => {
  const user = req.user;
  const month = /^\d{4}-\d{2}$/.test(req.query.month || "") ? req.query.month : nowIso().slice(0, 7);
  let people = [];
  if (PAYROLL_VIEW_ALL_ROLES.includes(user.role)) {
    people = [
      ...db.users.map((u) => ({ personType: "user", personId: u.id })),
      ...db.drivers.map((d) => ({ personType: "driver", personId: d.id })),
    ];
  } else if (PAYROLL_VIEW_ALL_DRIVERS_ROLES.includes(user.role)) {
    people = db.drivers.map((d) => ({ personType: "driver", personId: d.id }));
  } else if (user.role === "site_supervisor") {
    const ids = visibleDriverIdsFor(user);
    people = db.drivers.filter((d) => ids.has(d.id)).map((d) => ({ personType: "driver", personId: d.id }));
  }
  const rows = people.map(({ personType, personId }) => {
    const payslip = computePayslip(personType, personId, month);
    const { name, subLabel } = personLabel(personType, personId);
    return { personType, personId, name, subLabel, hasProfile: payslip.hasProfile, grossPay: payslip.grossPay, totalDeductions: payslip.totalDeductions, netPay: payslip.netPay };
  });
  res.json({ month, rows });
});

app.get("/api/payroll/payslip/:personType/:personId/download", requireAuth, (req, res) => {
  const { personType, personId } = req.params;
  if (!["user", "driver"].includes(personType) || !personExists(personType, personId)) {
    return res.status(404).json({ error: "Person not found." });
  }
  if (!canViewPayroll(req.user, personType, personId)) {
    return res.status(403).json({ error: "You don't have access to this person's payroll." });
  }
  const month = /^\d{4}-\d{2}$/.test(req.query.month || "") ? req.query.month : nowIso().slice(0, 7);
  const payslip = computePayslip(personType, personId, month);
  const { name, subLabel } = personLabel(personType, personId);
  const filenameBase = `payslip_${name.replace(/\s+/g, "_")}_${month}`;
  if ((req.query.format || "").toLowerCase() === "pdf") {
    return sendPayslipPdf(res, { name, subLabel, month, payslip }, filenameBase);
  }
  const rows = [
    ...payslip.earnings.map((e) => ({ item: e.label, type: "Earning", amount: e.amount })),
    ...payslip.deductions.map((d) => ({ item: d.label, type: "Deduction", amount: -d.amount })),
    { item: "NET PAY", type: "", amount: payslip.netPay },
  ];
  const columns = [
    { label: "Item", key: "item" },
    { label: "Type", key: "type" },
    { label: "Amount (₹)", key: "amount" },
  ];
  sendDownload(res, rows, columns, filenameBase, req.query.format);
});

// HR "approves" a month's payslip for a person, which snapshots the
// computed numbers at that moment. Drivers only ever see this approved
// snapshot in the driver app (never the live-computed figures staff see
// when viewing their own payslip), so a driver's downloadable payslip
// can't change out from under them after HR has signed off on it.
function payrollApprovalKey(personType, personId, month) {
  return `${personType}:${personId}:${month}`;
}
app.post(
  "/api/payroll/approve",
  requireAuth,
  requireRole("hr"),
  h(async (req, res) => {
    const { personType, personId, month } = req.body || {};
    if (!["user", "driver"].includes(personType) || !personExists(personType, personId)) {
      return res.status(400).json({ error: "Person not found." });
    }
    if (!/^\d{4}-\d{2}$/.test(month || "")) return res.status(400).json({ error: "A valid month (YYYY-MM) is required." });
    const payslip = computePayslip(personType, personId, month);
    const { name, subLabel } = personLabel(personType, personId);
    const approval = {
      ...payslip,
      name,
      subLabel,
      approvedBy: req.user.id,
      approvedByName: req.user.name,
      approvedAt: nowIso(),
    };
    db.payrollApprovals[payrollApprovalKey(personType, personId, month)] = approval;
    await audit(req.user, "approve_payslip", `${req.user.name} approved the ${month} payslip for ${name}`);
    res.json(approval);
  })
);
app.get("/api/payroll/approval/:personType/:personId", requireAuth, (req, res) => {
  const { personType, personId } = req.params;
  if (!["user", "driver"].includes(personType) || !personExists(personType, personId)) {
    return res.status(404).json({ error: "Person not found." });
  }
  if (!canViewPayroll(req.user, personType, personId)) {
    return res.status(403).json({ error: "You don't have access to this person's payroll." });
  }
  const month = /^\d{4}-\d{2}$/.test(req.query.month || "") ? req.query.month : nowIso().slice(0, 7);
  res.json(db.payrollApprovals[payrollApprovalKey(personType, personId, month)] || null);
});

// ---------- TRIPS ----------
// Trips are allocated to a driver by whichever Site Supervisor runs that
// driver's vehicle (Owner can also allocate, as an administrative
// override, same as Owner's override on vehicle Assignments elsewhere).
const TRIP_ASSIGN_ROLES = ["site_supervisor", "owner", "ops_manager", "data_team", "bookings"];
function vehiclesForTripAssignment(user) {
  // What vehicles (and therefore drivers) this user is allowed to allocate
  // trips for - identical scoping rule to Assignments/visibleVehiclesFor.
  return visibleVehiclesFor(user);
}
app.get(
  "/api/trips",
  requireAuth,
  h(async (req, res) => {
    let rows;
    if (req.user.role === "site_supervisor") {
      const vehicleIds = new Set(vehiclesForTripAssignment(req.user).map((v) => v.id));
      rows = db.trips.filter((t) => vehicleIds.has(t.vehicleId));
    } else if (["owner", "ops_manager", "hr", "area_supervisor", "data_team", "bookings"].includes(req.user.role)) {
      rows = db.trips.slice();
    } else {
      rows = [];
    }
    rows = rows
      .slice()
      .sort((a, b) => (b.date + (b.scheduledTime || "")).localeCompare(a.date + (a.scheduledTime || "")));
    res.json(rows);
  })
);
app.post(
  "/api/trips",
  requireAuth,
  requireRole(...TRIP_ASSIGN_ROLES),
  h(async (req, res) => {
    const { vehicleId, driverId, date, scheduledTime, pickupPoint, dropPoint, notes } = req.body || {};
    const vehicle = db.vehicles.find((v) => v.id === vehicleId);
    if (!vehicle) return res.status(400).json({ error: "Vehicle not found." });
    if (req.user.role === "site_supervisor" && vehicle.supervisorId !== req.user.id) {
      return res.status(403).json({ error: "You can only allocate trips for your own vehicles." });
    }
    const useDriverId = driverId || vehicle.driverId;
    if (!useDriverId || !db.drivers.some((d) => d.id === useDriverId)) {
      return res.status(400).json({ error: "This vehicle doesn't have a driver assigned to allocate a trip to." });
    }
    if (!date) return res.status(400).json({ error: "A trip date is required." });
    const trip = {
      id: uid("trip"),
      vehicleId,
      driverId: useDriverId,
      date,
      scheduledTime: scheduledTime || "",
      pickupPoint: pickupPoint || "",
      dropPoint: dropPoint || "",
      notes: notes || "",
      status: "upcoming",
      createdBy: req.user.id,
      createdByName: req.user.name,
      createdAt: nowIso(),
      updatedAt: nowIso(),
      startedAt: null,
      completedAt: null,
    };
    db.trips.push(trip);
    await audit(req.user, "create_trip", `${req.user.name} allocated a trip on ${date} to driver ${useDriverId} (${vehicle.reg})`);
    res.json(trip);
  })
);
app.patch(
  "/api/trips/:id",
  requireAuth,
  requireRole(...TRIP_ASSIGN_ROLES),
  h(async (req, res) => {
    const trip = db.trips.find((t) => t.id === req.params.id);
    if (!trip) return res.status(404).json({ error: "Trip not found." });
    const vehicle = db.vehicles.find((v) => v.id === trip.vehicleId);
    if (req.user.role === "site_supervisor" && (!vehicle || vehicle.supervisorId !== req.user.id)) {
      return res.status(403).json({ error: "You can only manage trips for your own vehicles." });
    }
    const { date, scheduledTime, pickupPoint, dropPoint, notes, status } = req.body || {};
    if (date !== undefined) trip.date = date;
    if (scheduledTime !== undefined) trip.scheduledTime = scheduledTime;
    if (pickupPoint !== undefined) trip.pickupPoint = pickupPoint;
    if (dropPoint !== undefined) trip.dropPoint = dropPoint;
    if (notes !== undefined) trip.notes = notes;
    if (status !== undefined) {
      if (!["upcoming", "in_progress", "completed", "cancelled"].includes(status)) {
        return res.status(400).json({ error: "Invalid trip status." });
      }
      trip.status = status;
      if (status === "in_progress" && !trip.startedAt) trip.startedAt = nowIso();
      if (status === "completed" && !trip.completedAt) trip.completedAt = nowIso();
      if (status === "completed") accrueTripIncomeIfNeeded(trip);
    }
    trip.updatedAt = nowIso();
    await audit(req.user, "update_trip", `${req.user.name} updated trip ${trip.id}`);
    res.json(trip);
  })
);
app.delete(
  "/api/trips/:id",
  requireAuth,
  requireRole(...TRIP_ASSIGN_ROLES),
  h(async (req, res) => {
    const trip = db.trips.find((t) => t.id === req.params.id);
    if (!trip) return res.status(404).json({ error: "Trip not found." });
    const vehicle = db.vehicles.find((v) => v.id === trip.vehicleId);
    if (req.user.role === "site_supervisor" && (!vehicle || vehicle.supervisorId !== req.user.id)) {
      return res.status(403).json({ error: "You can only manage trips for your own vehicles." });
    }
    db.trips = db.trips.filter((t) => t.id !== trip.id);
    await audit(req.user, "delete_trip", `${req.user.name} removed trip ${trip.id}`);
    res.json({ ok: true });
  })
);

// ---------- SUBVENDORS (Driver-cum-Owner) ----------
// A Subvendor is an external vehicle owner/operator we contract with -
// distinct from a "driver" (our employee, paid a salary) and distinct
// from a "client" (who we provide transport service to). A vehicle can
// be linked to at most one Subvendor, with a fixed billing contract (per
// trip / per km / per package) that stays constant for the life of the
// contract and only changes via an explicit revision (e.g. a fuel price
// variation) - never silently overwritten, so a full history is kept.
// Subvendors get their own completely separate login (email/mobile +
// password, same mechanics as staff) so each can see only their own
// vehicles, trip counts, income earned running for us, and the payments
// (Diesel/Advance/Repair/Other) and interest we've recorded against them.
const SUBVENDOR_MANAGE_ROLES = ["owner", "ops_manager"];
// Recording a payment (Diesel/Advance/Repair/Other) TO a subvendor is a
// day-to-day site expense, not an account-management action - Site
// Supervisors file these for whichever of their own subvendor vehicles are
// running at their site, same scoping as trips/assignments/expenses
// (vehicle.supervisorId === their own id). Everything else about a
// Subvendor (creating the account, editing contact details, resetting the
// password, setting the billing contract, viewing the full income summary)
// stays owner/ops_manager only.
const SUBVENDOR_PAYMENT_ROLES = [...SUBVENDOR_MANAGE_ROLES, "site_supervisor"];
const SUBVENDOR_ADVANCE_INTEREST_RATE_MONTHLY = 0.36 / 12; // 36% per annum, simple interest, accrued monthly

function publicSubvendor(sv) {
  if (!sv) return null;
  const { passwordHash, passwordHistory, ...rest } = sv;
  return rest;
}
function monthKey(d) {
  return (d instanceof Date ? d : new Date(d)).toISOString().slice(0, 7); // "YYYY-MM"
}
function monthsElapsed(fromKey, toKey) {
  const [fy, fm] = fromKey.split("-").map(Number);
  const [ty, tm] = toKey.split("-").map(Number);
  return (ty - fy) * 12 + (tm - fm);
}
function nextMonthKey(key) {
  const [y, m] = key.split("-").map(Number);
  return monthKey(new Date(Date.UTC(y, m, 1))); // m is 1-indexed already, so this lands on the next month
}
// Lazily catches up simple monthly interest (3%/month on the outstanding
// principal only - never compounded) on one advance. Same "check on read,
// no cron job needed" pattern as expireStaleShift() elsewhere in this
// file - safe to call as often as needed, a no-op once caught up.
function accrueAdvanceInterest(advance) {
  if (advance.status !== "outstanding" || !(advance.outstandingPrincipal > 0)) return;
  const nowKey = monthKey(nowIso());
  const months = monthsElapsed(advance.lastInterestAccrualMonth, nowKey);
  if (months <= 0) return;
  const interestAdded = Math.round(advance.outstandingPrincipal * SUBVENDOR_ADVANCE_INTEREST_RATE_MONTHLY * months * 100) / 100;
  advance.interestAccrued = Math.round((advance.interestAccrued + interestAdded) * 100) / 100;
  advance.lastInterestAccrualMonth = nowKey;
}
// Lazily catches up a "per package" vehicle's flat monthly income - one
// income entry per whole calendar month elapsed since the contract's
// effective date (inclusive of the start month itself), dated the 1st of
// that month.
function accruePackageIncome(vehicle) {
  if (!vehicle.subvendorId || !vehicle.billing || vehicle.billing.mode !== "per_package") return;
  const b = vehicle.billing;
  const startKey = monthKey(b.effectiveFrom || nowIso());
  if (!b.lastPackageAccrualMonth) {
    const [y, m] = startKey.split("-").map(Number);
    b.lastPackageAccrualMonth = monthKey(new Date(Date.UTC(y, m - 2, 1))); // the month BEFORE the start month
  }
  const nowKey = monthKey(nowIso());
  while (true) {
    const due = nextMonthKey(b.lastPackageAccrualMonth);
    if (due > nowKey) break;
    db.subvendorIncomeEntries.push({
      id: uid("svinc"),
      subvendorId: vehicle.subvendorId,
      vehicleId: vehicle.id,
      source: "package",
      refId: null,
      amount: b.rate,
      date: due + "-01",
      note: `Monthly package charge - ${due}`,
      createdAt: nowIso(),
    });
    b.lastPackageAccrualMonth = due;
  }
}
// Event-driven (not lazy) - called right when a trip is marked completed.
function accrueTripIncomeIfNeeded(trip) {
  const vehicle = db.vehicles.find((v) => v.id === trip.vehicleId);
  if (!vehicle || !vehicle.subvendorId || !vehicle.billing || vehicle.billing.mode !== "per_trip") return;
  if (db.subvendorIncomeEntries.some((e) => e.source === "trip" && e.refId === trip.id)) return; // already accrued
  db.subvendorIncomeEntries.push({
    id: uid("svinc"),
    subvendorId: vehicle.subvendorId,
    vehicleId: vehicle.id,
    source: "trip",
    refId: trip.id,
    amount: vehicle.billing.rate,
    date: todayStr(),
    note: `Trip on ${trip.date}${trip.pickupPoint ? " (" + trip.pickupPoint + (trip.dropPoint ? " -> " + trip.dropPoint : "") + ")" : ""}`,
    createdAt: nowIso(),
  });
}
// Event-driven - called right when a driver's closing odometer reading is
// recorded (end of that day's distance for whichever vehicle they drive).
function accrueKmIncomeIfNeeded(shift) {
  const vehicle = db.vehicles.find((v) => v.driverId === shift.driverId);
  if (!vehicle || !vehicle.subvendorId || !vehicle.billing || vehicle.billing.mode !== "per_km") return;
  if (typeof shift.odometerOpen !== "number" || typeof shift.odometerClose !== "number") return;
  const distanceKm = shift.odometerClose - shift.odometerOpen;
  if (!(distanceKm > 0)) return;
  if (db.subvendorIncomeEntries.some((e) => e.source === "km" && e.refId === shift.id)) return; // already accrued
  db.subvendorIncomeEntries.push({
    id: uid("svinc"),
    subvendorId: vehicle.subvendorId,
    vehicleId: vehicle.id,
    source: "km",
    refId: shift.id,
    amount: Math.round(vehicle.billing.rate * distanceKm * 100) / 100,
    date: (shift.odometerCloseAt || nowIso()).slice(0, 10),
    note: `${distanceKm} km on ${(shift.loginAt || "").slice(0, 10)}`,
    createdAt: nowIso(),
  });
}
// Everything a Subvendor (or the staff managing them) needs to see: their
// vehicles, trip counts, total income earned, payments received broken
// down by type, and outstanding advances with interest. Shared by both
// the Subvendor's own dashboard and the staff-facing summary endpoint.
function subvendorSummary(subvendorId) {
  const vehiclesForSv = db.vehicles.filter((v) => v.subvendorId === subvendorId);
  vehiclesForSv.forEach((v) => accruePackageIncome(v));
  const advancesForSv = db.subvendorAdvances.filter((a) => a.subvendorId === subvendorId);
  advancesForSv.forEach((a) => accrueAdvanceInterest(a));
  const vehicleIds = new Set(vehiclesForSv.map((v) => v.id));
  const tripsForSv = db.trips.filter((t) => vehicleIds.has(t.vehicleId));
  const completedTrips = tripsForSv.filter((t) => t.status === "completed");
  const incomeForSv = db.subvendorIncomeEntries.filter((e) => e.subvendorId === subvendorId);
  const totalIncome = Math.round(incomeForSv.reduce((s, e) => s + e.amount, 0) * 100) / 100;
  const paymentsForSv = db.subvendorPayments.filter((p) => p.subvendorId === subvendorId);
  const paymentsByType = {};
  paymentsForSv.forEach((p) => {
    paymentsByType[p.type] = Math.round(((paymentsByType[p.type] || 0) + p.amount) * 100) / 100;
  });
  const totalPaid = Math.round(paymentsForSv.reduce((s, p) => s + p.amount, 0) * 100) / 100;
  const totalOutstandingPrincipal =
    Math.round(advancesForSv.filter((a) => a.status === "outstanding").reduce((s, a) => s + a.outstandingPrincipal, 0) * 100) / 100;
  const totalInterestAccrued = Math.round(advancesForSv.reduce((s, a) => s + a.interestAccrued, 0) * 100) / 100;
  return {
    vehicles: vehiclesForSv.map((v) => ({
      id: v.id,
      reg: v.reg,
      vehicleType: v.vehicleType,
      fuelType: v.fuelType,
      clientId: v.clientId,
      clientName: v.clientId ? clientNameServer(v.clientId) : "",
      route: vehicleRouteInfo(v).routeNumber,
      routeStartingPoint: vehicleRouteInfo(v).startingPoint,
      driverId: v.driverId,
      driverName: v.driverId ? driverNameServer(v.driverId) : "",
      billing: v.billing,
    })),
    totalTrips: tripsForSv.length,
    completedTrips: completedTrips.length,
    totalIncome,
    incomeEntries: incomeForSv.slice().sort((a, b) => b.date.localeCompare(a.date)).slice(0, 200),
    payments: paymentsForSv.slice().sort((a, b) => b.date.localeCompare(a.date)),
    paymentsByType,
    totalPaid,
    advances: advancesForSv.slice().sort((a, b) => b.date.localeCompare(a.date)),
    totalOutstandingPrincipal,
    totalInterestAccrued,
    // What we still owe them: income earned, minus what we've already paid
    // out (fuel/advances/repairs/other), minus interest they owe us back on
    // any outstanding advance.
    netPayableToSubvendor: Math.round((totalIncome - totalPaid - totalInterestAccrued) * 100) / 100,
  };
}

app.get("/api/subvendors", requireAuth, requireRole(...SUBVENDOR_PAYMENT_ROLES), (req, res) => {
  let list = db.subvendors;
  if (req.user.role === "site_supervisor") {
    // Scoped down to just the subvendors running under this supervisor's
    // own vehicles - not the whole company's subvendor roster - since this
    // is only here so they can pick who to file a payment against.
    const mySubvendorIds = new Set(visibleVehiclesFor(req.user).filter((v) => v.subvendorId).map((v) => v.subvendorId));
    list = list.filter((s) => mySubvendorIds.has(s.id));
  }
  res.json(list.map(publicSubvendor));
});

app.post(
  "/api/subvendors",
  requireAuth,
  requireRole(...SUBVENDOR_MANAGE_ROLES),
  h(async (req, res) => {
    const { name, contactPerson, email, phone, gstNumber, address, password } = req.body || {};
    if (!name || !String(name).trim()) return res.status(400).json({ error: "Subvendor name is required." });
    const emailNorm = email ? String(email).trim().toLowerCase() : "";
    const phoneNorm = phone ? String(phone).trim() : "";
    if (!emailNorm && !phoneNorm) return res.status(400).json({ error: "Either an email address or a mobile number is required to log in." });
    const clash = db.subvendors.find((s) => (emailNorm && s.email === emailNorm) || (phoneNorm && s.phone === phoneNorm));
    if (clash) return res.status(400).json({ error: "That email or mobile number is already in use by another subvendor." });
    if (!password) return res.status(400).json({ error: "A temporary password is required." });
    const policyError = passwordPolicyError(password);
    if (policyError) return res.status(400).json({ error: policyError });
    const sv = {
      id: uid("sv"),
      name,
      contactPerson: contactPerson || "",
      email: emailNorm,
      phone: phoneNorm,
      gstNumber: gstNumber || "",
      address: address || "",
      passwordHash: hashPassword(password),
      passwordHistory: [],
      mustChangePassword: true,
      active: true,
      createdAt: nowIso(),
    };
    db.subvendors.push(sv);
    await audit(req.user, "create_subvendor", `${req.user.name} added subvendor ${name}`);
    res.json(publicSubvendor(sv));
  })
);

app.patch(
  "/api/subvendors/:id",
  requireAuth,
  requireRole(...SUBVENDOR_MANAGE_ROLES),
  h(async (req, res) => {
    const sv = db.subvendors.find((s) => s.id === req.params.id);
    if (!sv) return res.status(404).json({ error: "Subvendor not found." });
    const { name, contactPerson, email, phone, gstNumber, address, active } = req.body || {};
    if (name !== undefined) sv.name = name;
    if (contactPerson !== undefined) sv.contactPerson = contactPerson;
    if (email !== undefined) sv.email = String(email).trim().toLowerCase();
    if (phone !== undefined) sv.phone = String(phone).trim();
    if (gstNumber !== undefined) sv.gstNumber = gstNumber;
    if (address !== undefined) sv.address = address;
    if (active !== undefined) sv.active = !!active;
    await audit(req.user, "update_subvendor", `${req.user.name} updated subvendor ${sv.name}`);
    res.json(publicSubvendor(sv));
  })
);

app.patch(
  "/api/subvendors/:id/reset-password",
  requireAuth,
  requireRole("owner"),
  h(async (req, res) => {
    const sv = db.subvendors.find((s) => s.id === req.params.id);
    if (!sv) return res.status(404).json({ error: "Subvendor not found." });
    const { newPassword } = req.body || {};
    const policyError = passwordPolicyError(newPassword);
    if (policyError) return res.status(400).json({ error: policyError });
    sv.passwordHistory = [sv.passwordHash, ...(sv.passwordHistory || [])].filter(Boolean).slice(0, PASSWORD_HISTORY_LIMIT);
    sv.passwordHash = hashPassword(newPassword);
    sv.mustChangePassword = true;
    await audit(req.user, "reset_subvendor_password", `${req.user.name} reset the password for subvendor ${sv.name} - they must set a new one at next login`);
    res.json(publicSubvendor(sv));
  })
);

app.get(
  "/api/subvendors/:id/summary",
  requireAuth,
  requireRole(...SUBVENDOR_MANAGE_ROLES),
  h(async (req, res) => {
    const sv = db.subvendors.find((s) => s.id === req.params.id);
    if (!sv) return res.status(404).json({ error: "Subvendor not found." });
    res.json({ subvendor: publicSubvendor(sv), ...subvendorSummary(sv.id) });
  })
);

// Link (or unlink) a vehicle to a Subvendor and set/revise its billing
// contract. Switching mode or rate on an already-billed vehicle doesn't
// overwrite history - the old contract is pushed into billing.revisions
// first, so a full record of every rate change (e.g. a fuel price
// variation) is kept.
app.patch(
  "/api/vehicles/:id/subvendor",
  requireAuth,
  requireRole(...SUBVENDOR_MANAGE_ROLES),
  h(async (req, res) => {
    const vehicle = db.vehicles.find((v) => v.id === req.params.id);
    if (!vehicle) return res.status(404).json({ error: "Vehicle not found." });
    const { subvendorId, billingMode, billingRate, reason } = req.body || {};
    if (subvendorId === null || subvendorId === "") {
      vehicle.subvendorId = null;
      vehicle.billing = null;
      await audit(req.user, "unlink_subvendor", `${req.user.name} unlinked vehicle ${vehicle.reg} from its subvendor`);
      return res.json(vehicle);
    }
    if (subvendorId !== undefined) {
      const sv = db.subvendors.find((s) => s.id === subvendorId && s.active !== false);
      if (!sv) return res.status(400).json({ error: "Subvendor not found." });
      vehicle.subvendorId = subvendorId;
    }
    if (!vehicle.subvendorId) return res.status(400).json({ error: "Select a Subvendor before setting a billing contract." });
    if (billingMode !== undefined || billingRate !== undefined) {
      const mode = billingMode !== undefined ? billingMode : vehicle.billing && vehicle.billing.mode;
      const rate = billingRate !== undefined ? Number(billingRate) : vehicle.billing && vehicle.billing.rate;
      if (!["per_trip", "per_km", "per_package"].includes(mode)) {
        return res.status(400).json({ error: "Billing mode must be per_trip, per_km, or per_package." });
      }
      if (!(rate > 0)) return res.status(400).json({ error: "Billing rate must be a positive number." });
      const revisions = (vehicle.billing && vehicle.billing.revisions) || [];
      const changed = vehicle.billing && (vehicle.billing.mode !== mode || vehicle.billing.rate !== rate);
      if (vehicle.billing && changed) {
        revisions.push({
          mode: vehicle.billing.mode,
          rate: vehicle.billing.rate,
          effectiveFrom: vehicle.billing.effectiveFrom,
          endedAt: nowIso(),
          changedBy: req.user.id,
          changedByName: req.user.name,
          changedAt: nowIso(),
          reason: reason || "",
        });
      }
      vehicle.billing = {
        mode,
        rate,
        effectiveFrom: vehicle.billing && !changed ? vehicle.billing.effectiveFrom : nowIso(),
        lastPackageAccrualMonth: vehicle.billing && !changed ? vehicle.billing.lastPackageAccrualMonth : null,
        revisions,
      };
    }
    await audit(req.user, "set_subvendor_billing", `${req.user.name} set billing for ${vehicle.reg}: ${JSON.stringify(vehicle.billing)}`);
    res.json(vehicle);
  })
);

// A payment WE make TO a Subvendor - Diesel / Advance / Repair Costs /
// Other. An "Advance" additionally opens an interest-bearing balance
// (simple interest, 3%/month = 36% p.a., accrued monthly on whatever
// principal remains outstanding) that has to be settled separately.
app.post(
  "/api/subvendors/:id/payments",
  requireAuth,
  requireRole(...SUBVENDOR_PAYMENT_ROLES),
  h(async (req, res) => {
    const sv = db.subvendors.find((s) => s.id === req.params.id);
    if (!sv) return res.status(404).json({ error: "Subvendor not found." });
    const { vehicleId, type, amount, date, note } = req.body || {};
    const validTypes = ["Diesel", "Advance", "Repair", "Other"];
    if (!validTypes.includes(type)) return res.status(400).json({ error: "Payment type must be one of: " + validTypes.join(", ") + "." });
    if (!amount || Number(amount) <= 0) return res.status(400).json({ error: "A valid amount is required." });
    // Site Supervisors only file expenses for their own subvendor vehicles -
    // same scoping as trips/assignments (vehicle.supervisorId === them).
    if (req.user.role === "site_supervisor") {
      if (!vehicleId) return res.status(400).json({ error: "Select which of your vehicles this expense is for." });
      const vehicle = db.vehicles.find((v) => v.id === vehicleId);
      if (!vehicle || vehicle.supervisorId !== req.user.id) {
        return res.status(403).json({ error: "You can only record payments for your own vehicles." });
      }
      if (vehicle.subvendorId !== sv.id) {
        return res.status(400).json({ error: "That vehicle isn't linked to this subvendor." });
      }
    }
    const payment = {
      id: uid("svpay"),
      subvendorId: sv.id,
      vehicleId: vehicleId || null,
      type,
      amount: Number(amount),
      date: date || todayStr(),
      note: note || "",
      createdBy: req.user.id,
      createdByName: req.user.name,
      createdAt: nowIso(),
    };
    db.subvendorPayments.push(payment);
    let advance = null;
    if (type === "Advance") {
      advance = {
        id: uid("svadv"),
        subvendorId: sv.id,
        vehicleId: vehicleId || null,
        paymentId: payment.id,
        principal: payment.amount,
        outstandingPrincipal: payment.amount,
        interestAccrued: 0,
        status: "outstanding",
        date: payment.date,
        lastInterestAccrualMonth: monthKey(payment.date),
        note: note || "",
        createdAt: nowIso(),
      };
      db.subvendorAdvances.push(advance);
    }
    await audit(req.user, "subvendor_payment", `${req.user.name} recorded a ${type} payment of ₹${payment.amount} to subvendor ${sv.name}`);
    res.json({ payment, advance });
  })
);

// Lightweight payments list - lets a Site Supervisor see (and double-check)
// what they've filed for their own vehicles without exposing the full
// subvendor summary (company-wide income entries, other sites' payments,
// advance/interest details) that owner/ops_manager get via /summary.
app.get(
  "/api/subvendors/:id/payments",
  requireAuth,
  requireRole(...SUBVENDOR_PAYMENT_ROLES),
  h(async (req, res) => {
    const sv = db.subvendors.find((s) => s.id === req.params.id);
    if (!sv) return res.status(404).json({ error: "Subvendor not found." });
    let payments = db.subvendorPayments.filter((p) => p.subvendorId === sv.id);
    if (req.user.role === "site_supervisor") {
      const myVehicleIds = new Set(visibleVehiclesFor(req.user).map((v) => v.id));
      payments = payments.filter((p) => p.vehicleId && myVehicleIds.has(p.vehicleId));
    }
    payments = payments.slice().sort((a, b) => b.createdAt.localeCompare(a.createdAt));
    res.json(payments);
  })
);

// Settle (fully or partially) an outstanding advance - e.g. deducted from
// a future payment, or repaid directly.
app.patch(
  "/api/subvendors/:id/advances/:advanceId/settle",
  requireAuth,
  requireRole(...SUBVENDOR_MANAGE_ROLES),
  h(async (req, res) => {
    const advance = db.subvendorAdvances.find((a) => a.id === req.params.advanceId && a.subvendorId === req.params.id);
    if (!advance) return res.status(404).json({ error: "Advance not found." });
    accrueAdvanceInterest(advance);
    const { amount } = req.body || {};
    const settleAmount = amount !== undefined ? Number(amount) : advance.outstandingPrincipal;
    if (!(settleAmount > 0)) return res.status(400).json({ error: "A valid settlement amount is required." });
    if (settleAmount > advance.outstandingPrincipal) {
      return res.status(400).json({ error: `Cannot settle more than the outstanding principal (₹${advance.outstandingPrincipal}).` });
    }
    advance.outstandingPrincipal = Math.round((advance.outstandingPrincipal - settleAmount) * 100) / 100;
    if (advance.outstandingPrincipal <= 0) {
      advance.status = "settled";
      advance.settledAt = nowIso();
    }
    await audit(req.user, "settle_subvendor_advance", `${req.user.name} settled ₹${settleAmount} against an advance for subvendor ${advance.subvendorId}`);
    res.json(advance);
  })
);

// ---------- SUBVENDOR AUTH (their own separate login) ----------
// Mirrors staff login mechanics (email/mobile + hashed password, forced
// change on first login) but is entirely its own session map/table -
// Subvendors are not staff `db.users` and never appear in People/Staff
// management, payroll, or the audit-triggered user list.
const subvendorSessions = new Map(); // token -> subvendorId
const SUBVENDOR_ALLOWED_DURING_FORCED_PASSWORD_CHANGE = ["/api/subvendor-auth/me", "/api/subvendor-auth/change-password", "/api/subvendor-auth/logout"];
function subvendorAuth(req, res, next) {
  const header = req.headers.authorization || "";
  const token = header.startsWith("Bearer ") ? header.slice(7) : null;
  const subvendorId = token && subvendorSessions.get(token);
  const sv = subvendorId && db.subvendors.find((s) => s.id === subvendorId && s.active !== false);
  if (!sv) return res.status(401).json({ error: "Not signed in. Please log in again." });
  if (sv.mustChangePassword && !SUBVENDOR_ALLOWED_DURING_FORCED_PASSWORD_CHANGE.includes(req.path)) {
    return res.status(403).json({ error: "You must set a new password before continuing.", mustChangePassword: true });
  }
  req.subvendor = sv;
  req.subvendorToken = token;
  next();
}

app.post(
  "/api/subvendor-auth/login",
  h(async (req, res) => {
    const { identifier, password } = req.body || {};
    if (!identifier || !password) return res.status(400).json({ error: "Email/mobile number and password are required." });
    const idNorm = String(identifier).trim().toLowerCase();
    const phoneNorm = String(identifier).trim();
    const sv = db.subvendors.find((s) => s.active !== false && ((s.email && s.email === idNorm) || (s.phone && s.phone === phoneNorm)));
    if (!sv || !verifyPassword(password, sv.passwordHash)) {
      return res.status(401).json({ error: "Incorrect email/mobile number or password." });
    }
    const token = crypto.randomBytes(24).toString("hex");
    subvendorSessions.set(token, sv.id);
    await audit({ id: sv.id, name: sv.name }, "subvendor_login", `Subvendor ${sv.name} logged in`);
    res.json({ token, subvendor: publicSubvendor(sv) });
  })
);

app.get("/api/subvendor-auth/me", subvendorAuth, (req, res) => {
  res.json(publicSubvendor(req.subvendor));
});

app.post(
  "/api/subvendor-auth/change-password",
  subvendorAuth,
  h(async (req, res) => {
    const { currentPassword, newPassword } = req.body || {};
    if (!currentPassword || !newPassword) return res.status(400).json({ error: "Current and new password are required." });
    if (!verifyPassword(currentPassword, req.subvendor.passwordHash)) return res.status(400).json({ error: "Current password is incorrect." });
    const policyError = passwordPolicyError(newPassword);
    if (policyError) return res.status(400).json({ error: policyError });
    if (passwordReused(newPassword, req.subvendor)) {
      return res.status(400).json({ error: "New password cannot be the same as your current or a recent previous password." });
    }
    req.subvendor.passwordHistory = [req.subvendor.passwordHash, ...(req.subvendor.passwordHistory || [])].slice(0, PASSWORD_HISTORY_LIMIT);
    req.subvendor.passwordHash = hashPassword(newPassword);
    req.subvendor.mustChangePassword = false;
    await audit(req.subvendor, "subvendor_change_password", `${req.subvendor.name} changed their own password`);
    res.json(publicSubvendor(req.subvendor));
  })
);

app.post(
  "/api/subvendor-auth/logout",
  subvendorAuth,
  h(async (req, res) => {
    subvendorSessions.delete(req.subvendorToken);
    await audit(req.subvendor, "subvendor_logout", `${req.subvendor.name} logged out`);
    res.json({ ok: true });
  })
);

app.get(
  "/api/subvendor-auth/dashboard",
  subvendorAuth,
  h(async (req, res) => {
    res.json(subvendorSummary(req.subvendor.id));
  })
);

// ---------- DRIVER AUTH (mobile app) ----------
// A completely separate, deliberately simple login for drivers - mobile
// number + a 4-digit PIN (default 1234, never forced to change, see
// backfillDefaults for why). This never touches the staff `sessions` map
// or `db.users` at all; drivers get their own session map and their own
// req.driver, parallel to (but independent from) requireAuth/req.user.
const driverSessions = new Map(); // token -> driverId
const DRIVER_ODOMETER_WINDOW_MS = 10 * 60 * 1000; // 10 minutes to enter the opening odometer reading or the shift auto-closes
const MANUAL_ENTRY_PIN_WINDOW_MS = 10 * 60 * 1000; // 10 minutes to hand the code to a Supervisor before it expires
const MISSED_ODOMETER_ENTRY_PENALTY = 100; // ₹, charged to the driver each time a Supervisor has to manually enter a reading for them
const DRIVER_LOCATION_LOG_CAP = 50000; // oldest pings drop off so this file doesn't grow forever

function publicDriverAuth(d) {
  if (!d) return null;
  const { pin, ...rest } = d;
  return rest;
}
// The most recent shift for this driver that hasn't been fully closed out -
// i.e. what "currently logged in" means for a driver.
function currentDriverShift(driverId) {
  for (let i = db.driverShifts.length - 1; i >= 0; i--) {
    const s = db.driverShifts[i];
    if (s.driverId === driverId && s.status !== "closed") return s;
  }
  return null;
}
// Lazily expires a shift that's been sitting in "awaiting_odometer" for
// more than 10 minutes - checked on every relevant request rather than via
// a server-side timer, same pattern as the staff forced-password-change
// gate elsewhere in this file.
function expireStaleShift(shift) {
  if (shift && shift.status === "awaiting_odometer") {
    const ageMs = Date.now() - new Date(shift.loginAt).getTime();
    if (ageMs > DRIVER_ODOMETER_WINDOW_MS) {
      shift.status = "closed";
      shift.autoClosedReason = "Logged out automatically - the opening odometer reading wasn't entered within 10 minutes of login.";
      shift.closedAt = nowIso();
      return true;
    }
  }
  return false;
}
function driverAuth(req, res, next) {
  const header = req.headers.authorization || "";
  const token = header.startsWith("Bearer ") ? header.slice(7) : null;
  const driverId = token && driverSessions.get(token);
  const driver = driverId && db.drivers.find((d) => d.id === driverId && d.active !== false);
  if (!driver) return res.status(401).json({ error: "Not signed in. Please log in again." });
  req.driver = driver;
  req.driverToken = token;
  next();
}

app.post(
  "/api/driver-auth/login",
  h(async (req, res) => {
    const { phone, pin, lat, lng } = req.body || {};
    if (!phone || !pin) return res.status(400).json({ error: "Mobile number and PIN are required." });
    const phoneNorm = String(phone).trim();
    const driver = db.drivers.find((d) => d.active !== false && d.phone && d.phone === phoneNorm);
    if (!driver || String(driver.pin || "1234") !== String(pin).trim()) {
      return res.status(401).json({ error: "Incorrect mobile number or PIN." });
    }
    let shift = currentDriverShift(driver.id);
    if (shift) expireStaleShift(shift);
    if (!shift || shift.status === "closed") {
      shift = {
        id: uid("shift"),
        driverId: driver.id,
        loginAt: nowIso(),
        loginLat: typeof lat === "number" ? lat : null,
        loginLng: typeof lng === "number" ? lng : null,
        status: "awaiting_odometer",
        odometerOpen: null,
        odometerOpenPhoto: null,
        odometerOpenAt: null,
        odometerClose: null,
        odometerClosePhoto: null,
        odometerCloseAt: null,
        logoutAt: null,
        logoutLat: null,
        logoutLng: null,
        autoClosedReason: null,
        odometerOpenEnteredBy: null,
        odometerOpenVerified: false,
        odometerOpenVerifiedBy: null,
        odometerOpenVerifiedByName: null,
        odometerOpenVerifiedAt: null,
        odometerCloseEnteredBy: null,
        odometerCloseVerified: false,
        odometerCloseVerifiedBy: null,
        odometerCloseVerifiedByName: null,
        odometerCloseVerifiedAt: null,
        manualEntryPin: null,
        manualEntryPinExpiresAt: null,
      };
      db.driverShifts.push(shift);
    }
    const token = uid("dtok");
    driverSessions.set(token, driver.id);
    await audit({ id: driver.id, name: driver.name }, "driver_login", `Driver ${driver.name} logged in via the driver app`);
    res.json({ token, driver: publicDriverAuth(driver), shift });
  })
);

app.get(
  "/api/driver-auth/me",
  driverAuth,
  h(async (req, res) => {
    let shift = currentDriverShift(req.driver.id);
    const expired = expireStaleShift(shift);
    if (expired) await save();
    res.json({ driver: publicDriverAuth(req.driver), shift: shift && shift.status !== "closed" ? shift : null });
  })
);

app.post(
  "/api/driver-auth/odometer-open",
  driverAuth,
  h(async (req, res) => {
    const shift = currentDriverShift(req.driver.id);
    if (!shift) return res.status(400).json({ error: "No active login found. Please log in again." });
    if (expireStaleShift(shift)) {
      await save();
      return res.status(400).json({ error: shift.autoClosedReason });
    }
    if (shift.status !== "awaiting_odometer") {
      return res.status(400).json({ error: "The opening odometer reading has already been recorded for this shift." });
    }
    const { odometer, photo, lat, lng } = req.body || {};
    if (odometer === undefined || odometer === null || isNaN(Number(odometer))) {
      return res.status(400).json({ error: "Opening odometer reading is required." });
    }
    if (!photo) return res.status(400).json({ error: "A photo of the opening odometer reading is required." });
    const photoUrl = await savePhoto(photo, "driver_odo_open_" + req.driver.id);
    if (!photoUrl) return res.status(400).json({ error: "That photo couldn't be saved - please try again." });
    shift.odometerOpen = Number(odometer);
    shift.odometerOpenPhoto = photoUrl;
    shift.odometerOpenAt = nowIso();
    shift.odometerOpenLat = typeof lat === "number" ? lat : null;
    shift.odometerOpenLng = typeof lng === "number" ? lng : null;
    shift.odometerOpenEnteredBy = "driver";
    shift.status = "on_trip";
    await audit({ id: req.driver.id, name: req.driver.name }, "driver_odometer_open", `Driver ${req.driver.name} recorded opening odometer ${shift.odometerOpen}`);
    res.json(shift);
  })
);

// Adds everything the driver app's trip screens display beyond the bare
// trip record: the vehicle's registration, its client/site (so the driver
// knows who the run is for), its route, and the Site Supervisor who owns
// that vehicle - name + mobile number, so the driver has someone to call.
function enrichTripForDriver(trip) {
  const vehicle = db.vehicles.find((v) => v.id === trip.vehicleId);
  const client = vehicle && vehicle.clientId ? db.clients.find((c) => c.id === vehicle.clientId) : null;
  const site = vehicle && vehicle.siteId ? db.sites.find((s) => s.id === vehicle.siteId) : null;
  const supervisor = vehicle && vehicle.supervisorId ? db.users.find((u) => u.id === vehicle.supervisorId) : null;
  return Object.assign({}, trip, {
    vehicleReg: vehicle ? vehicle.reg : "",
    clientName: client ? client.name : "",
    siteName: site ? site.name : "",
    route: vehicle ? vehicle.route || "" : "",
    supervisorName: supervisor ? supervisor.name : "",
    supervisorPhone: supervisor ? supervisor.phone || "" : "",
  });
}

// The whole month's schedule for the "Upcoming Trips" tab - defaults to the
// current month if none given. Still excludes cancelled trips (nothing
// useful for a driver to do with those).
app.get(
  "/api/driver-auth/trips",
  driverAuth,
  h(async (req, res) => {
    const month = /^\d{4}-\d{2}$/.test(req.query.month || "") ? req.query.month : nowIso().slice(0, 7);
    const rows = db.trips
      .filter((t) => t.driverId === req.driver.id && t.status !== "cancelled" && t.date.slice(0, 7) === month)
      .sort((a, b) => (a.date + (a.scheduledTime || "")).localeCompare(b.date + (b.scheduledTime || "")))
      .map(enrichTripForDriver);
    res.json({
      month,
      upcoming: rows.filter((t) => t.status === "upcoming" || t.status === "in_progress"),
      completed: rows.filter((t) => t.status === "completed"),
    });
  })
);

// The single trip currently in progress (if any) - what the "Current Trip"
// tab is built around. Not month-scoped, since "what am I doing right now"
// shouldn't disappear just because the driver is browsing a different
// month in the Upcoming Trips tab.
app.get(
  "/api/driver-auth/trips/current",
  driverAuth,
  h(async (req, res) => {
    const trip = db.trips.find((t) => t.driverId === req.driver.id && t.status === "in_progress") || null;
    res.json(trip ? enrichTripForDriver(trip) : null);
  })
);

app.post(
  "/api/driver-auth/trips/:id/status",
  driverAuth,
  h(async (req, res) => {
    const trip = db.trips.find((t) => t.id === req.params.id && t.driverId === req.driver.id);
    if (!trip) return res.status(404).json({ error: "Trip not found." });
    const { status } = req.body || {};
    if (!["in_progress", "completed"].includes(status)) return res.status(400).json({ error: "Invalid trip status." });
    if (status === "in_progress") {
      const alreadyRunning = db.trips.find((t) => t.driverId === req.driver.id && t.status === "in_progress" && t.id !== trip.id);
      if (alreadyRunning) {
        return res.status(400).json({ error: "You already have a trip in progress - complete or finish that one first." });
      }
      trip.startedAt = nowIso();
    }
    if (status === "completed") {
      trip.completedAt = nowIso();
      accrueTripIncomeIfNeeded(trip);
    }
    trip.status = status;
    trip.updatedAt = nowIso();
    await audit({ id: req.driver.id, name: req.driver.name }, "driver_trip_status", `Driver ${req.driver.name} marked trip ${trip.id} as ${status}`);
    res.json(trip);
  })
);

app.post(
  "/api/driver-auth/odometer-close",
  driverAuth,
  h(async (req, res) => {
    const shift = currentDriverShift(req.driver.id);
    if (!shift) return res.status(400).json({ error: "No active login found. Please log in again." });
    if (shift.status !== "on_trip") {
      return res.status(400).json({ error: "Nothing to close yet - make sure you've entered your opening odometer reading first." });
    }
    const { odometer, photo, lat, lng } = req.body || {};
    if (odometer === undefined || odometer === null || isNaN(Number(odometer))) {
      return res.status(400).json({ error: "Closing odometer reading is required." });
    }
    if (Number(odometer) < Number(shift.odometerOpen)) {
      return res.status(400).json({ error: "Closing odometer reading can't be less than the opening reading." });
    }
    if (!photo) return res.status(400).json({ error: "A photo of the closing odometer reading is required." });
    const photoUrl = await savePhoto(photo, "driver_odo_close_" + req.driver.id);
    if (!photoUrl) return res.status(400).json({ error: "That photo couldn't be saved - please try again." });
    shift.odometerClose = Number(odometer);
    shift.odometerClosePhoto = photoUrl;
    shift.odometerCloseAt = nowIso();
    shift.odometerCloseLat = typeof lat === "number" ? lat : null;
    shift.odometerCloseLng = typeof lng === "number" ? lng : null;
    shift.odometerCloseEnteredBy = "driver";
    shift.status = "ready_to_logout";
    accrueKmIncomeIfNeeded(shift);
    await audit({ id: req.driver.id, name: req.driver.name }, "driver_odometer_close", `Driver ${req.driver.name} recorded closing odometer ${shift.odometerClose}`);
    res.json(shift);
  })
);

// If the driver genuinely can't enter a reading themselves (phone trouble,
// injury, app issue, etc.), their Supervisor can key it in for them - but
// only after the driver reads out this fresh 6-digit code, so the
// Supervisor can't just make up a number unchallenged. Dynamic and
// short-lived (10 minutes) rather than reusing the driver's static login
// PIN, and single-use - consumed the moment a manual entry succeeds.
app.post(
  "/api/driver-auth/manual-entry-code",
  driverAuth,
  h(async (req, res) => {
    const shift = currentDriverShift(req.driver.id);
    if (!shift || shift.status === "closed") {
      return res.status(400).json({ error: "No active shift found. Please log in again." });
    }
    const code = String(Math.floor(100000 + Math.random() * 900000));
    shift.manualEntryPin = code;
    shift.manualEntryPinExpiresAt = new Date(Date.now() + MANUAL_ENTRY_PIN_WINDOW_MS).toISOString();
    await audit({ id: req.driver.id, name: req.driver.name }, "driver_manual_entry_code", `Driver ${req.driver.name} generated a verification code for a Supervisor to manually enter an odometer reading`);
    res.json({ code, expiresAt: shift.manualEntryPinExpiresAt });
  })
);

app.post(
  "/api/driver-auth/logout",
  driverAuth,
  h(async (req, res) => {
    const shift = currentDriverShift(req.driver.id);
    if (!shift || shift.status !== "ready_to_logout") {
      return res.status(400).json({ error: "You must enter your closing odometer reading (with a photo) before you can log out." });
    }
    const { lat, lng } = req.body || {};
    shift.status = "closed";
    shift.logoutAt = nowIso();
    shift.logoutLat = typeof lat === "number" ? lat : null;
    shift.logoutLng = typeof lng === "number" ? lng : null;
    driverSessions.delete(req.driverToken);
    await audit({ id: req.driver.id, name: req.driver.name }, "driver_logout", `Driver ${req.driver.name} logged out of the driver app`);
    res.json({ ok: true });
  })
);

app.post(
  "/api/driver-auth/location-ping",
  driverAuth,
  h(async (req, res) => {
    const shift = currentDriverShift(req.driver.id);
    if (!shift || shift.status === "closed") return res.status(400).json({ error: "Not on an active shift." });
    const { lat, lng } = req.body || {};
    if (typeof lat !== "number" || typeof lng !== "number") return res.status(400).json({ error: "lat/lng are required." });
    db.driverLocationLogs.push({ id: uid("loc"), driverId: req.driver.id, shiftId: shift.id, lat, lng, ts: nowIso() });
    if (db.driverLocationLogs.length > DRIVER_LOCATION_LOG_CAP) {
      db.driverLocationLogs.splice(0, db.driverLocationLogs.length - DRIVER_LOCATION_LOG_CAP);
    }
    await save();
    res.json({ ok: true });
  })
);

// Drivers only ever see the HR-approved snapshot for a month, never the
// live-computed figures staff can see for themselves - so what they
// download can't silently change after HR has signed off on it.
app.get(
  "/api/driver-auth/payslip",
  driverAuth,
  h(async (req, res) => {
    const month = /^\d{4}-\d{2}$/.test(req.query.month || "") ? req.query.month : nowIso().slice(0, 7);
    const approval = db.payrollApprovals[payrollApprovalKey("driver", req.driver.id, month)] || null;
    res.json({ month, approved: !!approval, payslip: approval });
  })
);
app.get(
  "/api/driver-auth/payslip/download",
  driverAuth,
  h(async (req, res) => {
    const month = /^\d{4}-\d{2}$/.test(req.query.month || "") ? req.query.month : nowIso().slice(0, 7);
    const approval = db.payrollApprovals[payrollApprovalKey("driver", req.driver.id, month)];
    if (!approval) return res.status(404).json({ error: "This month's payslip hasn't been approved by HR yet." });
    const filenameBase = `payslip_${req.driver.name.replace(/\s+/g, "_")}_${month}`;
    if ((req.query.format || "").toLowerCase() === "pdf") {
      return sendPayslipPdf(res, { name: req.driver.name, subLabel: "Driver", month, payslip: approval }, filenameBase);
    }
    const rows = [
      ...approval.earnings.map((e) => ({ item: e.label, type: "Earning", amount: e.amount })),
      ...approval.deductions.map((d) => ({ item: d.label, type: "Deduction", amount: -d.amount })),
      { item: "NET PAY", type: "", amount: approval.netPay },
    ];
    const columns = [
      { label: "Item", key: "item" },
      { label: "Type", key: "type" },
      { label: "Amount (₹)", key: "amount" },
    ];
    sendDownload(res, rows, columns, filenameBase, req.query.format);
  })
);

// A driver requesting leave against a specific upcoming trip (or just a
// date) lands in the exact same db.leaves list/approval flow the Ops
// Manager already uses for staff-submitted leave requests - nothing new
// to build on that side, it just shows up in their existing queue with
// type "Requested by Driver" so it's obviously distinguishable.
app.post(
  "/api/driver-auth/leave-request",
  driverAuth,
  h(async (req, res) => {
    const { date, reason, tripId } = req.body || {};
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date || "")) return res.status(400).json({ error: "A valid date is required." });
    if (tripId && !db.trips.some((t) => t.id === tripId && t.driverId === req.driver.id)) {
      return res.status(400).json({ error: "Trip not found." });
    }
    const leave = {
      id: uid("l"),
      driverId: req.driver.id,
      driver: req.driver.name,
      type: "Requested by Driver",
      start: date,
      end: date,
      status: "pending",
      requestedBy: null,
      tripId: tripId || null,
      reason: reason || "",
    };
    db.leaves.push(leave);
    await audit({ id: req.driver.id, name: req.driver.name }, "driver_leave_request", `Driver ${req.driver.name} requested leave for ${date}${reason ? ` (${reason})` : ""}`);
    res.json(leave);
  })
);
app.get(
  "/api/driver-auth/leave-requests",
  driverAuth,
  h(async (req, res) => {
    const rows = db.leaves
      .filter((l) => l.driverId === req.driver.id)
      .slice()
      .sort((a, b) => b.start.localeCompare(a.start));
    res.json(rows);
  })
);

// Documents tab: the driver's own personal documents (filled in by HR via
// the Staff screen) plus whichever vehicle they're currently mapped to's
// RTA-facing documents - both strictly view/download here, no editing.
app.get(
  "/api/driver-auth/documents",
  driverAuth,
  h(async (req, res) => {
    const d = req.driver;
    const personal = {
      licenseNumber: d.licenseNumber || "",
      licenseCopyUrl: d.licenseCopyUrl || null,
      aadharNumber: d.aadharNumber || "",
      aadharCopyUrl: d.aadharCopyUrl || null,
      panNumber: d.panNumber || "",
      panCopyUrl: d.panCopyUrl || null,
      bankAccountNumber: d.bankAccountNumber || "",
      bankIfsc: d.bankIfsc || "",
      bankAccountHolderName: d.bankAccountHolderName || "",
      bankChequeUrl: d.bankChequeUrl || null,
      uanNumber: d.uanNumber || "",
      pfCertificateUrl: d.pfCertificateUrl || null,
      esiNumber: d.esiNumber || "",
      esiCertificateUrl: d.esiCertificateUrl || null,
      healthRecordNumber: d.healthRecordNumber || "",
      healthRecordDate: d.healthRecordDate || "",
      healthRecordCopyUrl: d.healthRecordCopyUrl || null,
      trainingCertNumber: d.trainingCertNumber || "",
      trainingCertDate: d.trainingCertDate || "",
      trainingCertCopyUrl: d.trainingCertCopyUrl || null,
      policeVerificationNumber: d.policeVerificationNumber || "",
      policeVerificationDate: d.policeVerificationDate || "",
      policeVerificationCopyUrl: d.policeVerificationCopyUrl || null,
    };
    const vehicle = db.vehicles.find((v) => v.driverId === req.driver.id);
    const vehicleDocs = vehicle
      ? {
          reg: vehicle.reg,
          docs: DOC_COPY_TYPES.reduce((acc, docType) => {
            const doc = vehicle.docs[docType];
            if (doc) acc[docType] = { number: doc.number || "", expiry: doc.expiry || "", copyUrl: doc.copyUrl || null };
            return acc;
          }, {}),
        }
      : null;
    res.json({ personal, vehicle: vehicleDocs });
  })
);

// Attendance tab: the driver's own shift history, drawn straight from
// driverShifts (the same records that gate login/odometer/logout) rather
// than a separate attendance log - login-to-logout, with distance driven
// if both odometer readings were taken.
app.get(
  "/api/driver-auth/attendance",
  driverAuth,
  h(async (req, res) => {
    const month = /^\d{4}-\d{2}$/.test(req.query.month || "") ? req.query.month : nowIso().slice(0, 7);
    const rows = db.driverShifts
      .filter((s) => s.driverId === req.driver.id && s.loginAt && s.loginAt.slice(0, 7) === month)
      .map((s) => {
        const hours =
          s.logoutAt && s.loginAt ? Math.round(((new Date(s.logoutAt) - new Date(s.loginAt)) / 3600000) * 10) / 10 : null;
        const distanceKm =
          typeof s.odometerClose === "number" && typeof s.odometerOpen === "number" ? s.odometerClose - s.odometerOpen : null;
        return {
          id: s.id,
          date: s.loginAt.slice(0, 10),
          loginAt: s.loginAt,
          logoutAt: s.logoutAt,
          status: s.status,
          odometerOpen: s.odometerOpen,
          odometerClose: s.odometerClose,
          distanceKm,
          hours,
          autoClosedReason: s.autoClosedReason || null,
        };
      })
      .sort((a, b) => b.loginAt.localeCompare(a.loginAt));
    res.json({ month, rows });
  })
);

// ---------- ODOMETER VERIFICATION (Site Supervisor / Data Team / Ops / Owner) ----------
// Every odometer reading a driver enters needs a human to actually look at
// the photo and confirm it's real before it's trusted - that's the
// Supervisor of whichever vehicle the driver is on. Vehicles don't always
// have a Supervisor assigned though, so Data Team is the fallback for
// those. Once a reading is verified, it's locked forever - there is no
// endpoint anywhere (this file included) that can change a verified
// reading, by design.
const ODOMETER_VERIFY_ROLES = ["site_supervisor", "data_team", "ops_manager", "owner"];
function shiftVehicleFor(shift) {
  return db.vehicles.find((v) => v.driverId === shift.driverId) || null;
}
// Owner/Ops Manager: everything. Site Supervisor: only shifts for drivers
// currently assigned to one of their own vehicles. Data Team: everything,
// but full detail (reading + photo) only for vehicles with no Supervisor -
// see publicDriverShiftRow for how the "status only" redaction works for
// the rest.
function driverShiftsVisibleTo(user) {
  return db.driverShifts
    .map((s) => ({ shift: s, vehicle: shiftVehicleFor(s) }))
    .filter(({ vehicle }) => {
      if (["owner", "ops_manager", "data_team"].includes(user.role)) return true;
      if (user.role === "site_supervisor") return vehicle && vehicle.supervisorId === user.id;
      return false;
    });
}
function odometerPointHtmlSafe(point) {
  return ["open", "close"].includes(point);
}
// Most recent GPS ping for a shift, falling back to whatever fixed lat/lng
// we do have (opening odometer reading, then login) so a live-tracking
// "Open in Maps" link still shows *something* even before the driver's
// phone has sent a single background ping yet.
function lastKnownLocationFor(shift) {
  const pings = (db.driverLocationLogs || []).filter((p) => p.shiftId === shift.id);
  if (pings.length) {
    const latest = pings.reduce((a, b) => ((a.ts || "") > (b.ts || "") ? a : b));
    return { lat: latest.lat, lng: latest.lng, ts: latest.ts, source: "live" };
  }
  if (shift.odometerCloseLat != null && shift.odometerCloseLng != null) {
    return { lat: shift.odometerCloseLat, lng: shift.odometerCloseLng, ts: shift.odometerCloseAt, source: "closing_reading" };
  }
  if (shift.odometerOpenLat != null && shift.odometerOpenLng != null) {
    return { lat: shift.odometerOpenLat, lng: shift.odometerOpenLng, ts: shift.odometerOpenAt, source: "opening_reading" };
  }
  if (shift.loginLat != null && shift.loginLng != null) {
    return { lat: shift.loginLat, lng: shift.loginLng, ts: shift.loginAt, source: "login" };
  }
  return null;
}
function publicDriverShiftRow(entry, user) {
  const { shift: s, vehicle } = entry;
  const driver = db.drivers.find((d) => d.id === s.driverId);
  const supervisor = vehicle && vehicle.supervisorId ? db.users.find((u) => u.id === vehicle.supervisorId) : null;
  const hasSupervisor = !!supervisor;
  // Data Team only gets the raw reading/photo for orphan (no-Supervisor)
  // vehicles - for everything else they see verification status only, per
  // the Owner's explicit call on scope.
  const fullAccess = user.role !== "data_team" || !hasSupervisor;
  function pointView(prefix) {
    const enteredAt = s[`odometer${prefix}At`];
    const hasReading = s[`odometer${prefix}`] != null;
    const base = {
      hasReading,
      enteredBy: s[`odometer${prefix}EnteredBy`] || null,
      enteredAt: enteredAt || null,
      verified: !!s[`odometer${prefix}Verified`],
      verifiedByName: s[`odometer${prefix}VerifiedByName`] || null,
      verifiedAt: s[`odometer${prefix}VerifiedAt`] || null,
    };
    if (!fullAccess) return base; // no odometer value or photo leaks to Data Team for a supervised vehicle
    return Object.assign(base, {
      odometer: s[`odometer${prefix}`],
      photoUrl: s[`odometer${prefix}Photo`] || null,
    });
  }
  return {
    id: s.id,
    driverId: s.driverId,
    driverName: driver ? driver.name : "Unknown driver",
    driverPhone: driver ? driver.phone : "",
    vehicleReg: vehicle ? vehicle.reg : null,
    vehicleId: vehicle ? vehicle.id : null,
    hasSupervisor,
    supervisorName: supervisor ? supervisor.name : null,
    status: s.status,
    date: s.loginAt ? s.loginAt.slice(0, 10) : null,
    loginAt: s.loginAt,
    open: pointView("Open"),
    close: pointView("Close"),
    lastLocation: lastKnownLocationFor(s),
    // Only meaningful to whoever can actually act on this shift - never
    // handed to a Data Team member looking at someone else's supervised
    // vehicle, since it's not their call to make a manual entry there.
    canAct: user.role === "owner" || user.role === "ops_manager" || (user.role === "site_supervisor" && vehicle && vehicle.supervisorId === user.id) || (user.role === "data_team" && !hasSupervisor),
  };
}
app.get(
  "/api/driver-shifts",
  requireAuth,
  requireRole(...ODOMETER_VERIFY_ROLES),
  h(async (req, res) => {
    const rows = driverShiftsVisibleTo(req.user)
      .map((e) => publicDriverShiftRow(e, req.user))
      .sort((a, b) => (b.loginAt || "").localeCompare(a.loginAt || ""));
    res.json(rows);
  })
);

// Approve a driver-entered reading after checking the photo. One-way -
// there's no un-verify endpoint. Can only verify a reading that's actually
// there, and only once.
app.patch(
  "/api/driver-shifts/:id/verify",
  requireAuth,
  requireRole(...ODOMETER_VERIFY_ROLES),
  h(async (req, res) => {
    const shift = db.driverShifts.find((s) => s.id === req.params.id);
    if (!shift) return res.status(404).json({ error: "Shift not found." });
    const vehicle = shiftVehicleFor(shift);
    const hasSupervisor = !!(vehicle && vehicle.supervisorId);
    const canAct =
      req.user.role === "owner" ||
      req.user.role === "ops_manager" ||
      (req.user.role === "site_supervisor" && vehicle && vehicle.supervisorId === req.user.id) ||
      (req.user.role === "data_team" && !hasSupervisor);
    if (!canAct) return res.status(403).json({ error: "You don't have access to verify this driver's readings." });
    const { point } = req.body || {};
    if (!odometerPointHtmlSafe(point)) return res.status(400).json({ error: "point must be 'open' or 'close'." });
    const prefix = point === "open" ? "Open" : "Close";
    if (shift[`odometer${prefix}`] == null) return res.status(400).json({ error: "There's no reading to verify yet." });
    if (shift[`odometer${prefix}Verified`]) return res.status(400).json({ error: "This reading is already verified." });
    shift[`odometer${prefix}Verified`] = true;
    shift[`odometer${prefix}VerifiedBy`] = req.user.id;
    shift[`odometer${prefix}VerifiedByName`] = req.user.name;
    shift[`odometer${prefix}VerifiedAt`] = nowIso();
    const driver = db.drivers.find((d) => d.id === shift.driverId);
    await audit(req.user, "verify_odometer", `${req.user.name} verified the ${point==="open"?"opening":"closing"} odometer reading for ${driver ? driver.name : shift.driverId} (${shift[`odometer${prefix}`]} km)`);
    res.json(publicDriverShiftRow({ shift, vehicle }, req.user));
  })
);

// A Supervisor (or Data Team, for a vehicle with no Supervisor) keying in a
// reading the driver couldn't enter themselves - gated behind the fresh
// 6-digit code the driver reads out to them, single-use. Auto-verified the
// moment it's entered (the code exchange already establishes that this
// really is the driver, standing there), and automatically raises a ₹100
// penalty against the driver for the missed entry.
app.post(
  "/api/driver-shifts/:id/manual-odometer",
  requireAuth,
  requireRole(...ODOMETER_VERIFY_ROLES),
  h(async (req, res) => {
    const shift = db.driverShifts.find((s) => s.id === req.params.id);
    if (!shift) return res.status(404).json({ error: "Shift not found." });
    const vehicle = shiftVehicleFor(shift);
    const hasSupervisor = !!(vehicle && vehicle.supervisorId);
    const canAct =
      req.user.role === "owner" ||
      req.user.role === "ops_manager" ||
      (req.user.role === "site_supervisor" && vehicle && vehicle.supervisorId === req.user.id) ||
      (req.user.role === "data_team" && !hasSupervisor);
    if (!canAct) return res.status(403).json({ error: "You don't have access to enter a reading for this driver." });
    if (shift.status === "closed") return res.status(400).json({ error: "This shift is already closed." });
    const { point, odometer, pin, photo } = req.body || {};
    if (!odometerPointHtmlSafe(point)) return res.status(400).json({ error: "point must be 'open' or 'close'." });
    const prefix = point === "open" ? "Open" : "Close";
    if (point === "open" && shift.status !== "awaiting_odometer") {
      return res.status(400).json({ error: "The opening reading has already been recorded for this shift." });
    }
    if (point === "close" && shift.status !== "on_trip") {
      return res.status(400).json({ error: "The closing reading can only be entered once the trip is underway and hasn't already been closed." });
    }
    if (odometer === undefined || odometer === null || isNaN(Number(odometer))) {
      return res.status(400).json({ error: "An odometer reading is required." });
    }
    if (point === "close" && Number(odometer) < Number(shift.odometerOpen)) {
      return res.status(400).json({ error: "Closing odometer reading can't be less than the opening reading." });
    }
    if (!pin || !String(pin).trim()) {
      return res.status(400).json({ error: "Enter the 6-digit code the driver read out to you." });
    }
    if (!shift.manualEntryPin || String(shift.manualEntryPin) !== String(pin).trim()) {
      return res.status(400).json({ error: "That code is incorrect. Ask the driver to check it, or generate a new one." });
    }
    if (!shift.manualEntryPinExpiresAt || new Date(shift.manualEntryPinExpiresAt).getTime() < Date.now()) {
      return res.status(400).json({ error: "That code has expired - ask the driver to generate a new one." });
    }
    let photoUrl = null;
    if (photo) {
      photoUrl = await savePhoto(photo, `driver_odo_${point}_manual_` + shift.driverId);
    }
    shift[`odometer${prefix}`] = Number(odometer);
    shift[`odometer${prefix}Photo`] = photoUrl;
    shift[`odometer${prefix}At`] = nowIso();
    shift[`odometer${prefix}EnteredBy`] = "supervisor";
    shift[`odometer${prefix}Verified`] = true;
    shift[`odometer${prefix}VerifiedBy`] = req.user.id;
    shift[`odometer${prefix}VerifiedByName`] = req.user.name;
    shift[`odometer${prefix}VerifiedAt`] = nowIso();
    if (point === "open") shift.status = "on_trip";
    if (point === "close") {
      shift.status = "ready_to_logout";
      accrueKmIncomeIfNeeded(shift);
    }
    // The code is single-use - consumed the moment it's successfully used,
    // whether for the opening or closing reading.
    shift.manualEntryPin = null;
    shift.manualEntryPinExpiresAt = null;

    // Automatic ₹100 penalty on the driver for the missed entry - not a
    // manually-typed HR adjustment, so it goes straight into
    // payrollAdhocEntries the same way the temp-driver auto-expense does.
    const month = nowIso().slice(0, 7);
    const penalty = {
      id: uid("payadh"),
      personType: "driver",
      personId: shift.driverId,
      month,
      category: "missed_odometer_entry",
      label: `Missed Odometer Entry (${point === "open" ? "Opening" : "Closing"} reading, ${vehicle ? vehicle.reg : "vehicle"} on ${shift.loginAt ? shift.loginAt.slice(0, 10) : "-"}) - entered by ${req.user.name}`,
      type: "deduction",
      amount: MISSED_ODOMETER_ENTRY_PENALTY,
      addedBy: req.user.id,
      addedByName: req.user.name,
      addedAt: nowIso(),
      autoGenerated: true,
    };
    db.payrollAdhocEntries.unshift(penalty);

    const driver = db.drivers.find((d) => d.id === shift.driverId);
    await audit(
      req.user,
      "manual_odometer_entry",
      `${req.user.name} manually entered the ${point === "open" ? "opening" : "closing"} odometer reading (${odometer} km) for ${driver ? driver.name : shift.driverId} after code verification - ₹${MISSED_ODOMETER_ENTRY_PENALTY} penalty applied`
    );
    res.json({ shift: publicDriverShiftRow({ shift, vehicle }, req.user), penalty });
  })
);

// ---------- OWNER DASHBOARD - fleet average mileage ----------
// Computed from every daily record ever submitted (not just today's), since
// a single day's fuel fills are too few to mean much - this is a proper
// running average per vehicle, worst-performing first, plus a fleet-wide
// figure. Records are never deleted, so this covers the vehicle's whole
// history in the app.
// A single blended "fleet average" hides which vehicle class/fuel type is
// actually underperforming (Buses only run Diesel; Cabs may be Diesel,
// Petrol, CNG, EV or Hybrid), so mileage is broken out into one category
// per Vehicle Type + Fuel Type combination that actually exists in the
// fleet. Within each category, only the worst-performing vehicles (lowest
// mileage first) are returned, so the dashboard stays a short, actionable
// list instead of a full per-vehicle table.
const MILEAGE_WORST_N = 5;
app.get("/api/dashboard/mileage-summary", requireAuth, requireRole("owner", "ops_manager"), (req, res) => {
  const perVehicle = {};
  Object.values(db.records).forEach((r) => {
    if (!r.fuel || !(r.fuel.litres > 0) || !(r.fuel.mileage > 0)) return;
    if (!perVehicle[r.vehicleId]) perVehicle[r.vehicleId] = { totalMileage: 0, fillCount: 0 };
    perVehicle[r.vehicleId].totalMileage += r.fuel.mileage;
    perVehicle[r.vehicleId].fillCount += 1;
  });
  const rows = Object.entries(perVehicle).map(([vehicleId, agg]) => {
    const vehicle = db.vehicles.find((v) => v.id === vehicleId);
    const avgMileage = Math.round((agg.totalMileage / agg.fillCount) * 10) / 10;
    return {
      vehicleId,
      reg: vehicle ? vehicle.reg : vehicleId,
      vehicleType: (vehicle && vehicle.vehicleType) || "Cab",
      fuelType: (vehicle && vehicle.fuelType) || "Diesel",
      avgMileage,
      standardMileage: vehicle ? vehicle.standardMileage : null,
      fillCount: agg.fillCount,
      belowStandard: vehicle ? avgMileage < vehicle.standardMileage : false,
    };
  });

  const byCategory = {};
  rows.forEach((r) => {
    const key = `${r.vehicleType}_${r.fuelType}`;
    if (!byCategory[key]) byCategory[key] = { vehicleType: r.vehicleType, fuelType: r.fuelType, rows: [] };
    byCategory[key].rows.push(r);
  });
  const categories = Object.values(byCategory)
    .map((cat) => {
      const sorted = cat.rows.slice().sort((a, b) => a.avgMileage - b.avgMileage);
      const categoryAverage = sorted.length
        ? Math.round((sorted.reduce((s, r) => s + r.avgMileage, 0) / sorted.length) * 10) / 10
        : 0;
      return {
        vehicleType: cat.vehicleType,
        fuelType: cat.fuelType,
        label: `${cat.vehicleType === "Bus" ? "Buses" : "Cabs"} - ${cat.fuelType}`,
        vehicleCount: sorted.length,
        categoryAverage,
        worst: sorted.slice(0, MILEAGE_WORST_N),
      };
    })
    .sort((a, b) => (a.vehicleType === b.vehicleType ? a.fuelType.localeCompare(b.fuelType) : a.vehicleType.localeCompare(b.vehicleType)));

  res.json({ categories });
});

// ---------- AUDIT ----------
// The Audit Log tab shows a date range (today only by default, or whichever
// from/to the caller picks) rather than the whole history at once - and the
// Dashboard's "recent activity" feed asks for the last hour specifically
// via sinceMinutes, independent of day boundaries.
app.get("/api/audit", requireAuth, requireRole("owner", "data_team", "hr", "ops_manager"), (req, res) => {
  const limit = Math.min(Number(req.query.limit) || 200, 1000);
  let rows = db.auditLog;
  if (req.query.sinceMinutes) {
    const cutoff = Date.now() - Number(req.query.sinceMinutes) * 60000;
    rows = rows.filter((a) => new Date(a.ts).getTime() >= cutoff);
  } else if (req.query.from || req.query.to) {
    const from = req.query.from || req.query.to;
    const to = req.query.to || req.query.from;
    rows = rows.filter((a) => {
      const d = (a.ts || "").slice(0, 10);
      return d >= from && d <= to;
    });
  } else {
    const date = req.query.date || todayStr();
    rows = rows.filter((a) => (a.ts || "").slice(0, 10) === date);
  }
  res.json(rows.slice(0, limit));
});

// ---------- REPORTS (Owner, Operations Manager, Area Supervisor) ----------
// Downloadable CSV/Excel exports so these three roles can pull the raw data
// out and analyze it however they like, outside the app. Each dataset is
// scoped to what that role can already see elsewhere in the app - an Area
// Supervisor only gets their own team's rows, never the whole fleet's.
function csvEscape(v) {
  if (v === null || v === undefined) return "";
  const s = String(v);
  return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}
function toCsv(rows, columns) {
  const header = columns.map((c) => csvEscape(c.label)).join(",");
  const lines = rows.map((r) => columns.map((c) => csvEscape(typeof c.key === "function" ? c.key(r) : r[c.key])).join(","));
  return [header, ...lines].join("\r\n");
}
function toXlsxBuffer(rows, columns, sheetName) {
  const data = rows.map((r) => {
    const obj = {};
    columns.forEach((c) => {
      obj[c.label] = typeof c.key === "function" ? c.key(r) : r[c.key];
    });
    return obj;
  });
  const ws = XLSX.utils.json_to_sheet(data.length ? data : [{}]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, (sheetName || "Sheet1").slice(0, 31));
  return XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
}
// Shared CSV/XLSX response writer for the one-off date-scoped downloads
// (Vehicles Running, Fuel Overview) that sit outside the main Reports
// dataset system - same format handling and headers as /api/reports/:dataset.
function sendDownload(res, rows, columns, filenameBase, format) {
  const fmt = (format || "csv").toLowerCase() === "xlsx" ? "xlsx" : "csv";
  const filename = `${filenameBase}.${fmt}`;
  if (fmt === "xlsx") {
    const buf = toXlsxBuffer(rows, columns, filenameBase.slice(0, 31));
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    return res.send(buf);
  }
  res.setHeader("Content-Type", "text/csv; charset=utf-8");
  res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
  res.send(toCsv(rows, columns));
}
// Every payslip - staff or driver - can be downloaded as a proper PDF
// payslip, not just CSV/Excel. Kept separate from sendDownload() (which
// stays CSV/XLSX-only and is shared by a bunch of unrelated reports) since
// a payslip is the one thing in this app that actually wants to look like
// a real, printable document.
function sendPayslipPdf(res, { name, subLabel, month, payslip }, filenameBase) {
  const doc = new PDFDocument({ size: "A4", margin: 50 });
  res.setHeader("Content-Type", "application/pdf");
  res.setHeader("Content-Disposition", `attachment; filename="${filenameBase}.pdf"`);
  doc.pipe(res);

  doc.fontSize(16).fillColor("#1f3864").text("Padmasri Travels", { align: "left" });
  doc.fontSize(11).fillColor("#666").text("Payslip", { align: "left" });
  doc.moveDown(0.6);
  doc.fontSize(10).fillColor("#1a1a1a");
  doc.text(`Name: ${name}`);
  if (subLabel) doc.text(`Role: ${subLabel}`);
  doc.text(`Month: ${month}`);
  doc.moveDown(0.8);

  const colX = { item: 50, type: 320, amount: 420 };
  const rowH = 20;
  let y = doc.y;
  doc.fontSize(10).fillColor("#fff");
  doc.rect(50, y, 495, rowH).fill("#2e74b5");
  doc.fillColor("#fff").text("Item", colX.item + 5, y + 5).text("Type", colX.type, y + 5).text("Amount (Rs.)", colX.amount, y + 5);
  y += rowH;

  const rows = [
    ...payslip.earnings.map((e) => ({ item: e.label, type: "Earning", amount: e.amount })),
    ...payslip.deductions.map((d) => ({ item: d.label, type: "Deduction", amount: -d.amount })),
  ];
  doc.fontSize(9.5);
  rows.forEach((r, i) => {
    doc.rect(50, y, 495, rowH).fill(i % 2 === 0 ? "#f7f9fc" : "#ffffff");
    doc.fillColor("#1a1a1a")
      .text(String(r.item), colX.item + 5, y + 5, { width: 260 })
      .text(String(r.type), colX.type, y + 5)
      .text(r.amount.toLocaleString("en-IN"), colX.amount, y + 5);
    y += rowH;
  });

  y += 4;
  doc.rect(50, y, 495, rowH + 4).fill("#1f3864");
  doc.fillColor("#fff").fontSize(11).text("NET PAY", colX.item + 5, y + 7).text(
    `Rs. ${payslip.netPay.toLocaleString("en-IN")}`,
    colX.amount,
    y + 7
  );

  doc.moveDown(3);
  doc.fontSize(8).fillColor("#9aa4b2").text("Generated by the Padmasri Travels Fleet Supervisor App.", 50, doc.y + 20);
  doc.end();
}
function vehicleReg(id) {
  const v = db.vehicles.find((x) => x.id === id);
  return v ? v.reg : id || "";
}
function paymentStageLabel(stage) {
  if (stage === "in_process") return "In Process";
  if (stage === "completed") return "Completed";
  if (stage === "approved") return "Approved";
  return "";
}
function driverNameServer(id) {
  const d = db.drivers.find((x) => x.id === id);
  return d ? d.name : "";
}
function userNameServer(id) {
  const u = db.users.find((x) => x.id === id);
  return u ? u.name : "";
}
function clientNameServer(id) {
  const c = db.clients.find((x) => x.id === id);
  return c ? c.name : "";
}
function siteNameServer(id) {
  const s = db.sites.find((x) => x.id === id);
  return s ? s.name : "";
}

const REPORT_ROLES = ["owner", "ops_manager", "area_supervisor"];

function buildReport(dataset, user, query) {
  const from = query.from || null;
  const to = query.to || null;
  const inRange = (dateStr) => {
    if (!dateStr) return true;
    if (from && dateStr < from) return false;
    if (to && dateStr > to) return false;
    return true;
  };

  if (dataset === "records") {
    const visibleIds = new Set(visibleVehiclesFor(user).map((v) => v.id));
    const rows = Object.values(db.records).filter((r) => visibleIds.has(r.vehicleId) && inRange(r.date));
    const columns = [
      { label: "Date", key: "date" },
      { label: "Vehicle", key: (r) => vehicleReg(r.vehicleId) },
      { label: "Attendance", key: (r) => (r.attendance || {}).status || "" },
      { label: "Uniform", key: (r) => (r.attendance || {}).uniform || "" },
      { label: "Breathalyzer Result", key: (r) => (r.breathalyzer || {}).result || "" },
      { label: "Breathalyzer Reading", key: (r) => (r.breathalyzer || {}).reading || "" },
      { label: "Odometer", key: (r) => (r.fuel || {}).odometer ?? "" },
      { label: "Distance (km)", key: (r) => (r.fuel || {}).distance ?? "" },
      { label: "Litres Filled", key: (r) => (r.fuel || {}).litres ?? "" },
      { label: "Mileage (km/l)", key: (r) => (r.fuel || {}).mileage ?? "" },
      { label: "Below Standard Mileage", key: (r) => ((r.fuel || {}).belowStandard ? "YES" : "") },
      { label: "Fuel Cost", key: (r) => (r.fuel || {}).totalCost ?? "" },
      { label: "Verification Status", key: (r) => (r.verification || {}).status || "" },
      { label: "Verification Comment", key: (r) => (r.verification || {}).comment || "" },
      { label: "Submitted At", key: "submittedAt" },
    ];
    return { rows, columns };
  }

  if (dataset === "expenses") {
    const rows = visibleExpensesFor(user).filter((e) => inRange((e.submittedAt || "").slice(0, 10)));
    const columns = [
      { label: "Date", key: (e) => (e.submittedAt || "").slice(0, 10) },
      { label: "Supervisor", key: "userName" },
      { label: "Category", key: "category" },
      { label: "Amount", key: "amount" },
      { label: "Description", key: "description" },
      { label: "Vehicle", key: (e) => (e.vehicleId ? vehicleReg(e.vehicleId) : e.isOfficeExpenditure ? "Office Expenditure" : "") },
      { label: "Status", key: (e) => (e.status === "rejected" ? "Rejected" : e.status === "approved" ? paymentStageLabel(e.paymentStage) : "Pending") },
      { label: "Rejection Reason / Note", key: "comment" },
      { label: "UTR Number", key: (e) => e.utrNumber || "" },
      { label: "Covered For Driver", key: (e) => e.coveredForDriverName || "" },
    ];
    return { rows, columns };
  }

  if (dataset === "attendance") {
    let rows = db.supervisorAttendance.filter((a) => inRange(a.date));
    if (user.role === "area_supervisor") {
      const mine = new Set(user.supervises || []);
      rows = rows.filter((a) => mine.has(a.userId));
    }
    const columns = [
      { label: "Date", key: "date" },
      { label: "Name", key: "userName" },
      { label: "Role", key: "role" },
      { label: "Type", key: (a) => (a.type === "clockin" ? "Clock In" : "Clock Out") },
      { label: "Time", key: (a) => new Date(a.ts).toLocaleTimeString() },
      { label: "Site", key: (a) => a.siteName || "" },
      { label: "Distance From Site (m)", key: "distanceMeters" },
      { label: "Within Geofence", key: (a) => (a.withinGeofence == null ? "n/a" : a.withinGeofence ? "Yes" : "No") },
    ];
    return { rows, columns };
  }

  if (dataset === "vehicles") {
    const rows = visibleVehiclesFor(user);
    const columns = [
      { label: "Registration", key: "reg" },
      { label: "Vehicle Type", key: "vehicleType" },
      { label: "Fuel Type", key: "fuelType" },
      { label: "Route", key: "route" },
      { label: "Usage", key: "usage" },
      { label: "Client", key: (v) => (v.clientId ? clientNameServer(v.clientId) : "") },
      { label: "Site", key: (v) => (v.siteId ? siteNameServer(v.siteId) : "") },
      { label: "Supervisor", key: (v) => (v.supervisorId ? userNameServer(v.supervisorId) : "") },
      { label: "Driver", key: (v) => (v.driverId ? driverNameServer(v.driverId) : "") },
      { label: "Driver Since", key: "driverAssignedDate" },
      { label: "Standard Mileage", key: "standardMileage" },
      { label: "Last Odometer", key: "lastOdometer" },
    ];
    return { rows, columns };
  }

  if (dataset === "leaves") {
    const rows = db.leaves.filter((l) => inRange(l.start));
    const columns = [
      { label: "Driver", key: "driver" },
      { label: "Type", key: "type" },
      { label: "Start", key: "start" },
      { label: "End", key: "end" },
      { label: "Status", key: "status" },
    ];
    return { rows, columns };
  }

  if (dataset === "audit") {
    if (!["owner", "ops_manager"].includes(user.role)) {
      const err = new Error("Only the Owner and Operations Manager can download the audit log.");
      err.status = 403;
      throw err;
    }
    const rows = db.auditLog.filter((a) => inRange((a.ts || "").slice(0, 10)));
    const columns = [
      { label: "Time", key: "ts" },
      { label: "User", key: "userName" },
      { label: "Action", key: "action" },
      { label: "Detail", key: "detail" },
    ];
    return { rows, columns };
  }

  if (dataset === "temp_drivers") {
    const rows = visibleExpensesFor(user)
      .filter((e) => e.category === "Temporary Driver")
      .filter((e) => inRange((e.submittedAt || "").slice(0, 10)));
    const columns = [
      { label: "Date", key: (e) => (e.submittedAt || "").slice(0, 10) },
      { label: "Temporary Driver", key: (e) => (e.tempDriver || {}).name || "" },
      { label: "Phone", key: (e) => (e.tempDriver || {}).phone || "" },
      { label: "UPI ID", key: (e) => (e.tempDriver || {}).upiId || "" },
      { label: "Amount", key: "amount" },
      { label: "Vehicle", key: (e) => (e.vehicleId ? vehicleReg(e.vehicleId) : "") },
      { label: "Route", key: (e) => { const v = db.vehicles.find((x) => x.id === e.vehicleId); return v ? v.route : ""; } },
      { label: "Company", key: (e) => { const v = db.vehicles.find((x) => x.id === e.vehicleId); return v && v.clientId ? clientNameServer(v.clientId) : "Internal / Fixed Route"; } },
      { label: "Covered For Driver", key: (e) => e.coveredForDriverName || "" },
      { label: "Requested By", key: "userName" },
      { label: "Status", key: "status" },
    ];
    return { rows, columns };
  }

  if (dataset === "doc_expiry") {
    const visibleIds = new Set(visibleVehiclesFor(user).map((v) => v.id));
    const rows = [];
    db.vehicles.filter((v) => visibleIds.has(v.id)).forEach((v) => {
      Object.entries(v.docs || {}).forEach(([docType, doc]) => {
        if (!doc.expiry && !doc.number) return;
        rows.push({ vehicle: v, docType, doc });
      });
    });
    const columns = [
      { label: "Vehicle", key: (r) => r.vehicle.reg },
      { label: "Document", key: "docType" },
      { label: "Number", key: (r) => r.doc.number || "" },
      { label: "Expiry Date", key: (r) => r.doc.expiry || "" },
      { label: "Days Left", key: (r) => (r.doc.expiry ? Math.round((new Date(r.doc.expiry) - new Date(todayStr())) / 86400000) : "") },
      { label: "State Permit", key: (r) => (r.docType === "Permit" ? (r.doc.isStatePermit ? "Yes" : "No") : "") },
      { label: "Districts", key: (r) => (r.docType === "Permit" ? r.doc.districtCount || 0 : "") },
      { label: "District Names", key: (r) => (r.docType === "Permit" ? r.doc.districtNames || "" : "") },
      { label: "Site", key: (r) => (r.vehicle.siteId ? siteNameServer(r.vehicle.siteId) : "") },
      { label: "Supervisor", key: (r) => (r.vehicle.supervisorId ? userNameServer(r.vehicle.supervisorId) : "") },
    ];
    return { rows, columns };
  }

  if (dataset === "mileage") {
    const visibleIds = new Set(visibleVehiclesFor(user).map((v) => v.id));
    const rows = Object.values(db.records).filter((r) => visibleIds.has(r.vehicleId) && r.fuel && r.fuel.litres > 0 && inRange(r.date));
    const columns = [
      { label: "Date", key: "date" },
      { label: "Vehicle", key: (r) => vehicleReg(r.vehicleId) },
      { label: "Vehicle Type", key: (r) => { const v = db.vehicles.find((x) => x.id === r.vehicleId); return v ? v.vehicleType : ""; } },
      { label: "Fuel Type", key: (r) => { const v = db.vehicles.find((x) => x.id === r.vehicleId); return v ? v.fuelType : ""; } },
      { label: "Previous Odometer", key: (r) => r.fuel.previousOdometer },
      { label: "Odometer", key: (r) => r.fuel.odometer },
      { label: "Distance (km)", key: (r) => r.fuel.distance },
      { label: "Litres Filled", key: (r) => r.fuel.litres },
      { label: "Mileage (km/l)", key: (r) => r.fuel.mileage },
      { label: "Standard Mileage", key: (r) => { const v = db.vehicles.find((x) => x.id === r.vehicleId); return v ? v.standardMileage : ""; } },
      { label: "Below Standard", key: (r) => (r.fuel.belowStandard ? "YES" : "") },
      { label: "Fuel Cost", key: (r) => r.fuel.totalCost },
    ];
    return { rows, columns };
  }

  if (dataset === "odometer") {
    const visibleIds = new Set(visibleVehiclesFor(user).map((v) => v.id));
    const rows = Object.values(db.odometerLogs).filter((o) => visibleIds.has(o.vehicleId) && inRange(o.date));
    const columns = [
      { label: "Date", key: "date" },
      { label: "Vehicle", key: (o) => vehicleReg(o.vehicleId) },
      { label: "Day Start (km)", key: (o) => (o.dayStart ? o.dayStart.value : "") },
      { label: "Day Start Time", key: (o) => (o.dayStart ? new Date(o.dayStart.ts).toLocaleTimeString() : "") },
      { label: "Day Close (km)", key: (o) => (o.dayClose ? o.dayClose.value : "") },
      { label: "Day Close Time", key: (o) => (o.dayClose ? new Date(o.dayClose.ts).toLocaleTimeString() : "") },
      { label: "Distance (km)", key: (o) => (o.dayStart && o.dayClose ? o.dayClose.value - o.dayStart.value : "") },
    ];
    return { rows, columns };
  }

  if (dataset === "monthly") {
    const visibleIds = new Set(visibleVehiclesFor(user).map((v) => v.id));
    const byVehicleMonth = {};
    Object.values(db.records).forEach((r) => {
      if (!visibleIds.has(r.vehicleId)) return;
      const month = (r.date || "").slice(0, 7); // YYYY-MM
      if (!month || !inRange(r.date)) return;
      const key = r.vehicleId + "_" + month;
      if (!byVehicleMonth[key]) {
        byVehicleMonth[key] = { vehicleId: r.vehicleId, month, daysSubmitted: 0, totalDistance: 0, totalLitres: 0, totalFuelCost: 0, totalMileage: 0, fillCount: 0, absentDays: 0 };
      }
      const agg = byVehicleMonth[key];
      agg.daysSubmitted += 1;
      if (r.fuel) {
        agg.totalDistance += r.fuel.distance || 0;
        agg.totalLitres += r.fuel.litres || 0;
        agg.totalFuelCost += r.fuel.totalCost || 0;
        if (r.fuel.litres > 0 && r.fuel.mileage > 0) {
          agg.totalMileage += r.fuel.mileage;
          agg.fillCount += 1;
        }
      }
      if (r.attendance && r.attendance.status === "absent") agg.absentDays += 1;
    });
    const rows = Object.values(byVehicleMonth);
    const columns = [
      { label: "Month", key: "month" },
      { label: "Vehicle", key: (a) => vehicleReg(a.vehicleId) },
      { label: "Days Submitted", key: "daysSubmitted" },
      { label: "Total Distance (km)", key: "totalDistance" },
      { label: "Total Litres", key: "totalLitres" },
      { label: "Total Fuel Cost", key: "totalFuelCost" },
      { label: "Average Mileage (km/l)", key: (a) => (a.fillCount ? Math.round((a.totalMileage / a.fillCount) * 10) / 10 : "") },
      { label: "Driver Absent Days", key: "absentDays" },
    ];
    return { rows, columns };
  }

  const err = new Error("Unknown report.");
  err.status = 400;
  throw err;
}

// ---------- Operational Cost Analysis (Vehicle-wise + Site-wise) ----------
// Every vehicle-tagged expense already carries the vehicle's registration
// (vehicleId) and a Site snapshot (siteId, taken at the moment it was
// filed - see POST /api/expenses). Office Expenditure has neither, so per
// the Owner's explicit instruction it's divided evenly across every Site
// when computing each Site's Total Operational Cost, rather than being
// left out of the picture entirely. Only approved expenses count as real
// spend - pending/rejected requests aren't money that's actually gone out
// yet.
app.get(
  "/api/reports/operational-cost-analysis",
  requireAuth,
  requireRole(...REPORT_ROLES),
  (req, res) => {
    const approved = db.expenses.filter((e) => e.status === "approved");
    const pending = db.expenses.filter((e) => e.status === "pending");
    const officeTotal = approved.filter((e) => e.isOfficeExpenditure).reduce((s, e) => s + e.amount, 0);
    const vehicleExpenses = approved.filter((e) => e.vehicleId);

    const byVehicle = {};
    vehicleExpenses.forEach((e) => {
      if (!byVehicle[e.vehicleId]) byVehicle[e.vehicleId] = { vehicleId: e.vehicleId, total: 0, count: 0, byCategory: {} };
      const agg = byVehicle[e.vehicleId];
      agg.total += e.amount;
      agg.count += 1;
      agg.byCategory[e.category] = (agg.byCategory[e.category] || 0) + e.amount;
    });
    const vehicleRows = Object.values(byVehicle)
      .map((v) => {
        const vehicle = db.vehicles.find((x) => x.id === v.vehicleId);
        return {
          vehicleId: v.vehicleId,
          reg: vehicle ? vehicle.reg : "Unknown vehicle",
          siteId: vehicle ? vehicle.siteId || null : null,
          siteName: vehicle && vehicle.siteId ? siteNameServer(vehicle.siteId) : "",
          clientName: vehicle && vehicle.clientId ? clientNameServer(vehicle.clientId) : vehicle ? "Internal / Fixed Route" : "",
          total: Math.round(v.total * 100) / 100,
          count: v.count,
          byCategory: v.byCategory,
        };
      })
      .sort((a, b) => b.total - a.total);

    const bySite = {};
    vehicleExpenses.forEach((e) => {
      const key = e.siteId || "__no_site__";
      if (!bySite[key]) bySite[key] = { siteId: e.siteId || null, direct: 0 };
      bySite[key].direct += e.amount;
    });
    const siteIds = db.sites.map((s) => s.id);
    const siteCount = siteIds.length || 1;
    const officeShare = Math.round((officeTotal / siteCount) * 100) / 100;
    siteIds.forEach((id) => {
      if (!bySite[id]) bySite[id] = { siteId: id, direct: 0 };
    });
    const siteRows = Object.values(bySite)
      .filter((s) => s.siteId)
      .map((s) => ({
        siteId: s.siteId,
        siteName: siteNameServer(s.siteId),
        directVehicleCost: Math.round(s.direct * 100) / 100,
        officeExpenditureShare: officeShare,
        totalOperationalCost: Math.round((s.direct + officeShare) * 100) / 100,
      }))
      .sort((a, b) => b.totalOperationalCost - a.totalOperationalCost);
    const noSiteVehicleCost = Math.round(((bySite["__no_site__"] && bySite["__no_site__"].direct) || 0) * 100) / 100;

    res.json({
      officeExpenditureTotal: Math.round(officeTotal * 100) / 100,
      siteCount,
      officeExpenditureSharePerSite: officeShare,
      vehicles: vehicleRows,
      sites: siteRows,
      noSiteVehicleCost,
      pendingTotal: Math.round(pending.reduce((s, e) => s + e.amount, 0) * 100) / 100,
      grandTotal: Math.round((vehicleExpenses.reduce((s, e) => s + e.amount, 0) + officeTotal) * 100) / 100,
    });
  }
);

app.get(
  "/api/reports/:dataset",
  requireAuth,
  requireRole(...REPORT_ROLES),
  h(async (req, res) => {
    let result;
    try {
      result = buildReport(req.params.dataset, req.user, req.query);
    } catch (err) {
      return res.status(err.status || 400).json({ error: err.message });
    }
    const format = (req.query.format || "csv").toLowerCase() === "xlsx" ? "xlsx" : "csv";
    const filename = `${req.params.dataset}_${todayStr()}.${format}`;
    if (format === "xlsx") {
      const buf = toXlsxBuffer(result.rows, result.columns, req.params.dataset);
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
      return res.send(buf);
    }
    res.setHeader("Content-Type", "text/csv; charset=utf-8");
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.send(toCsv(result.rows, result.columns));
  })
);

// ---------- BACKUPS (Owner only) ----------
// A rolling 30-day history of full-data snapshots, taken automatically once
// a day (see runBackup above). This is what "save this for a month" means in
// practice on Atlas's free tier - it protects against an accidental bad
// edit/deletion made through the app, restorable by the Owner without
// needing anyone's help.
app.get(
  "/api/admin/backups",
  requireAuth,
  requireRole("owner"),
  h(async (req, res) => {
    if (!USE_MONGO) return res.json([]);
    const list = await backupsCol
      .find({}, { projection: { _id: 1, takenAt: 1, reason: 1 } })
      .sort({ takenAt: -1 })
      .toArray();
    res.json(list);
  })
);

app.post(
  "/api/admin/backups/:id/restore",
  requireAuth,
  requireRole("owner"),
  h(async (req, res) => {
    if (!USE_MONGO) return res.status(400).json({ error: "Backups only apply when using MongoDB storage." });
    const snapshot = await backupsCol.findOne({ _id: req.params.id });
    if (!snapshot) return res.status(404).json({ error: "Backup not found." });
    db = JSON.parse(JSON.stringify(snapshot.data));
    backfillDefaults();
    await audit(req.user, "restore_backup", `${req.user.name} restored all app data from the backup taken ${snapshot.takenAt}`);
    res.json({ ok: true, restoredFrom: snapshot.takenAt });
  })
);

// ---------- MANUAL FULL DATA BACKUP DOWNLOAD (Owner, Operations Manager) ----------
// Separate from the automatic 30-day Mongo snapshots above (which only
// exist in Mongo mode and aren't downloadable) - this is a complete,
// downloadable export of every record in the app, meant to be taken
// roughly every 15 days and kept outside the app entirely as an
// independent copy. Every download is logged (who, when) so the Owner and
// Operations Manager can see whether the fleet's data is actually being
// backed up on schedule.
const BACKUP_DOWNLOAD_ROLES = ["owner", "ops_manager"];
app.get(
  "/api/backup/full",
  requireAuth,
  requireRole(...BACKUP_DOWNLOAD_ROLES),
  h(async (req, res) => {
    const takenAt = nowIso();
    db.backupDownloads = db.backupDownloads || [];
    db.backupDownloads.unshift({ id: uid("bkdl"), userId: req.user.id, userName: req.user.name, role: req.user.role, takenAt });
    if (db.backupDownloads.length > 200) db.backupDownloads.length = 200;
    await audit(req.user, "download_full_backup", `${req.user.name} downloaded a full data backup`);

    // Snapshot taken AFTER recording this download, so the exported file's
    // own embedded history includes the download that produced it.
    const clone = JSON.parse(JSON.stringify(db));
    delete clone._id;
    // Never include password hashes (or history) in an exportable file,
    // even though this is restricted to Owner/Ops Manager - a downloaded
    // file can end up anywhere (email, a shared drive, a personal laptop).
    if (Array.isArray(clone.users)) {
      clone.users.forEach((u) => {
        delete u.passwordHash;
        delete u.passwordHistory;
      });
    }
    if (Array.isArray(clone.subvendors)) {
      clone.subvendors.forEach((sv) => {
        delete sv.passwordHash;
        delete sv.passwordHistory;
      });
    }

    const filename = `padmasri_backup_${takenAt.slice(0, 10)}.json`;
    res.setHeader("Content-Type", "application/json");
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.send(JSON.stringify(clone, null, 2));
  })
);

app.get(
  "/api/backup/history",
  requireAuth,
  requireRole(...BACKUP_DOWNLOAD_ROLES),
  (req, res) => {
    res.json((db.backupDownloads || []).slice(0, 100));
  }
);

// ---------- fallback to the app shell ----------
app.get("*", (req, res, next) => {
  if (req.path.startsWith("/api/") || req.path.startsWith("/uploads/")) return next();
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

initStorage()
  .then(() => {
    app.listen(PORT, () => {
      console.log(`Fleet Supervisor App listening on port ${PORT} (storage: ${USE_MONGO ? "MongoDB" : "local file"})`);
    });
  })
  .catch((err) => {
    console.error("Failed to start - could not initialize storage:", err);
    process.exit(1);
  });
