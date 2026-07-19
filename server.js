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
  (db.vehicles || []).forEach((v) => {
    if (v.driverAssignedDate === undefined) v.driverAssignedDate = null;
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
  const match = dataUrl.match(/^data:(image\/[\w+.-]+);base64,(.+)$/);
  if (!match) return null;
  const contentType = match[1];
  const base64Data = match[2];

  if (USE_MONGO) {
    const id = uid("img");
    await photosCol.insertOne({ _id: id, contentType, data: base64Data, tag: tag || "img", createdAt: nowIso() });
    return "/api/photo/" + id;
  }
  const ext = contentType.split("/")[1] || "jpg";
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
  const { pin, ...rest } = u;
  return rest;
}

function requireAuth(req, res, next) {
  const header = req.headers.authorization || "";
  const token = header.startsWith("Bearer ") ? header.slice(7) : null;
  const userId = token && sessions.get(token);
  const user = userId && db.users.find((u) => u.id === userId && u.active !== false);
  if (!user) return res.status(401).json({ error: "Not signed in. Please log in again." });
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
app.get("/api/users/login-list", (req, res) => {
  res.json(
    db.users
      .filter((u) => u.active !== false)
      .map((u) => ({ id: u.id, name: u.name, role: u.role }))
  );
});

app.post(
  "/api/login",
  h(async (req, res) => {
    const { id, pin } = req.body || {};
    const user = db.users.find((u) => u.id === id && u.active !== false);
    if (!user || String(user.pin) !== String(pin)) {
      return res.status(401).json({ error: "Wrong user or PIN." });
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

// ---------- META (clients, sites, drivers, users) ----------
app.get("/api/meta", requireAuth, (req, res) => {
  res.json({
    clients: db.clients,
    sites: db.sites,
    drivers: db.drivers,
    users: db.users.map(publicUser),
    me: publicUser(req.user),
  });
});

// ---------- USERS (People admin) ----------
const ADMIN_ROLES = ["ops_manager", "owner", "data_team"];
const VALID_ROLES = ["site_supervisor", "area_supervisor", "ops_manager", "data_team", "owner", "hr"];

app.post(
  "/api/users",
  requireAuth,
  requireRole(...ADMIN_ROLES),
  h(async (req, res) => {
    const { name, role, pin, siteId, supervises } = req.body || {};
    if (!name || !VALID_ROLES.includes(role) || !pin) {
      return res.status(400).json({ error: "name, role and pin are required." });
    }
    const user = {
      id: uid("u"),
      name,
      role,
      pin: String(pin),
      active: true,
      siteId: siteId || null,
      supervises: Array.isArray(supervises) ? supervises : [],
    };
    db.users.push(user);
    await audit(req.user, "create_user", `${req.user.name} added ${name} as ${role}`);
    res.json(publicUser(user));
  })
);

app.patch(
  "/api/users/:id",
  requireAuth,
  requireRole(...ADMIN_ROLES),
  h(async (req, res) => {
    const user = db.users.find((u) => u.id === req.params.id);
    if (!user) return res.status(404).json({ error: "User not found." });
    const before = { name: user.name, role: user.role, siteId: user.siteId, supervises: user.supervises, active: user.active };

    // Whitelisted fields only - and PIN is left alone unless a new non-empty
    // one is actually provided, so an edit form with a blank PIN field can't
    // accidentally lock someone out.
    const { name, role, pin, siteId, supervises, active } = req.body || {};
    if (name !== undefined && String(name).trim()) user.name = String(name).trim();
    if (role !== undefined) {
      if (!VALID_ROLES.includes(role)) return res.status(400).json({ error: "Unknown role." });
      user.role = role;
    }
    if (pin !== undefined && String(pin).trim()) user.pin = String(pin).trim();
    if (siteId !== undefined) user.siteId = siteId || null;
    if (supervises !== undefined) user.supervises = Array.isArray(supervises) ? supervises : user.supervises;
    if (active !== undefined) user.active = !!active;

    const after = { name: user.name, role: user.role, siteId: user.siteId, supervises: user.supervises, active: user.active };
    await audit(req.user, "update_user", `${req.user.name} updated ${user.name} (${user.id}): ${JSON.stringify(before)} -> ${JSON.stringify(after)}`);
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
    const client = { id: uid("c"), name };
    db.clients.push(client);
    await audit(req.user, "create_client", `${req.user.name} added client ${name}`);
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
app.post(
  "/api/drivers",
  requireAuth,
  requireRole(...ADMIN_ROLES),
  h(async (req, res) => {
    const { name, phone, licenseNumber } = req.body || {};
    if (!name) return res.status(400).json({ error: "name is required." });
    const driver = { id: uid("d"), name, phone: phone || "", licenseNumber: licenseNumber || "" };
    db.drivers.push(driver);
    await audit(req.user, "create_driver", `${req.user.name} added driver ${name}`);
    res.json(driver);
  })
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

app.get("/api/vehicles", requireAuth, (req, res) => {
  res.json(visibleVehiclesFor(req.user));
});

app.post(
  "/api/vehicles",
  requireAuth,
  requireRole(...ADMIN_ROLES),
  h(async (req, res) => {
    const { reg, route, usage, standardMileage } = req.body || {};
    if (!reg) return res.status(400).json({ error: "Registration number is required." });
    const vehicle = {
      id: uid("v"),
      reg,
      route: route || "",
      usage: usage || "fixed_route",
      clientId: null,
      siteId: null,
      supervisorId: null,
      driverId: null,
      driverAssignedDate: null,
      standardMileage: Number(standardMileage) || 4.0,
      lastOdometer: null,
      docs: {
        RC: { number: "", expiry: "" },
        Permit: { number: "", expiry: "" },
        Insurance: { number: "", expiry: "" },
        Fitness: { number: "", expiry: "" },
        Tax: { number: "", expiry: "" },
        PUC: { number: "", expiry: "" },
      },
    };
    db.vehicles.push(vehicle);
    await audit(req.user, "create_vehicle", `${req.user.name} added vehicle ${reg}`);
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
    const before = { siteId: vehicle.siteId, supervisorId: vehicle.supervisorId, driverId: vehicle.driverId, clientId: vehicle.clientId, driverAssignedDate: vehicle.driverAssignedDate };
    const { siteId, supervisorId, driverId, clientId, usage, standardMileage, driverAssignedDate } = req.body || {};
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
    if (clientId !== undefined) vehicle.clientId = clientId || null;
    if (usage !== undefined) vehicle.usage = usage;
    if (standardMileage !== undefined) vehicle.standardMileage = Number(standardMileage) || vehicle.standardMileage;
    await audit(
      req.user,
      "assign_vehicle",
      `${req.user.name} reassigned ${vehicle.reg}: ${JSON.stringify(before)} -> ${JSON.stringify({ siteId: vehicle.siteId, supervisorId: vehicle.supervisorId, driverId: vehicle.driverId, clientId: vehicle.clientId, driverAssignedDate: vehicle.driverAssignedDate })}`
    );
    res.json(vehicle);
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
    const { docType, number, expiry } = req.body || {};
    if (!vehicle.docs[docType]) return res.status(400).json({ error: "Unknown document type." });
    vehicle.docs[docType] = { number: number ?? vehicle.docs[docType].number, expiry: expiry ?? vehicle.docs[docType].expiry };
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

    vehicle.lastOdometer = odometer;

    const record = {
      vehicleId: vehicle.id,
      date,
      safety: body.safety || {},
      condition: body.condition || {},
      photoTags,
      breathalyzer: body.breathalyzer || {},
      attendance: body.attendance || {},
      fuel: {
        previousOdometer,
        odometer,
        fuelPrice,
        litres,
        fillLevel: f.fillLevel || "other",
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
    res.json(record);
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

app.post(
  "/api/expenses",
  requireAuth,
  h(async (req, res) => {
    const { amount, category, description, vehicleId, bill } = req.body || {};
    if (!amount || Number(amount) <= 0) return res.status(400).json({ error: "A valid amount is required." });
    if (!bill) return res.status(400).json({ error: "A photo of the bill/receipt is required for every expense request." });
    const billUrl = await savePhoto(bill, "bill");
    if (!billUrl) return res.status(400).json({ error: "Could not read the bill photo - please try again." });
    const expense = {
      id: uid("e"),
      userId: req.user.id,
      userName: req.user.name,
      vehicleId: vehicleId || null,
      category: category || "Other",
      amount: Number(amount),
      description: description || "",
      billUrl,
      status: "pending",
      decidedBy: null,
      comment: "",
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
  requireRole("ops_manager", "owner"),
  h(async (req, res) => {
    const expense = db.expenses.find((e) => e.id === req.params.id);
    if (!expense) return res.status(404).json({ error: "Expense not found." });
    const { status, comment } = req.body || {};
    if (!["approved", "rejected"].includes(status)) return res.status(400).json({ error: "status must be approved or rejected." });
    expense.status = status;
    expense.decidedBy = req.user.id;
    expense.comment = comment || "";
    await audit(req.user, "expense_decision", `${req.user.name} ${status} ₹${expense.amount} expense from ${expense.userName}${comment ? ": " + comment : ""}`);
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
  const { date, userId } = req.query;
  let rows = db.supervisorAttendance;
  if (date) rows = rows.filter((r) => r.date === date);
  if (userId) rows = rows.filter((r) => r.userId === userId);
  res.json(rows);
});

// own attendance history - any logged-in user can see their own
app.get("/api/attendance/me", requireAuth, (req, res) => {
  res.json(db.supervisorAttendance.filter((r) => r.userId === req.user.id).slice(0, 60));
});

// ---------- AUDIT ----------
app.get("/api/audit", requireAuth, requireRole("owner", "data_team", "hr", "ops_manager"), (req, res) => {
  const limit = Math.min(Number(req.query.limit) || 100, 1000);
  res.json(db.auditLog.slice(0, limit));
});

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
