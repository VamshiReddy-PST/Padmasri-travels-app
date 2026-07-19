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
  (db.vehicles || []).forEach((v) => {
    if (v.driverAssignedDate === undefined) v.driverAssignedDate = null;
    if (v.make === undefined) v.make = "";
    if (v.model === undefined) v.model = "";
    if (v.engineNo === undefined) v.engineNo = "";
    if (v.chassisNo === undefined) v.chassisNo = "";
    if (v.rcDate === undefined) v.rcDate = "";
    if (v.seatingCapacity === undefined) v.seatingCapacity = null;
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
  });
  (db.drivers || []).forEach((d) => {
    if (d.phone === undefined) d.phone = "";
    if (d.licenseNumber === undefined) d.licenseNumber = "";
    if (d.aadharNumber === undefined) d.aadharNumber = "";
    if (d.esiNumber === undefined) d.esiNumber = "";
    if (d.pfNumber === undefined) d.pfNumber = "";
    if (d.uanNumber === undefined) d.uanNumber = "";
    if (d.esiCertificateUrl === undefined) d.esiCertificateUrl = null;
    if (d.pfCertificateUrl === undefined) d.pfCertificateUrl = null;
    if (d.dateOfJoining === undefined) d.dateOfJoining = "";
    if (d.drivingLevel === undefined) d.drivingLevel = "";
    if (d.performanceScore === undefined) d.performanceScore = null;
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
// These 5 document types get a photo/scan of the physical document attached
// (Insurance does not, per the Owner's spec).
const DOC_COPY_TYPES = ["RC", "Permit", "Fitness", "Tax", "PUC"];

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

app.post(
  "/api/users",
  requireAuth,
  requireRole(...PEOPLE_EDITOR_ROLES),
  h(async (req, res) => {
    const { name, role, pin, siteId, supervises, phone, dateOfJoining } = req.body || {};
    if (!name || !VALID_ROLES.includes(role) || !pin) {
      return res.status(400).json({ error: "name, role and pin are required." });
    }
    if (!canManageUserRole(req.user.role, role)) {
      return res.status(403).json({ error: `Only the Owner can create ${roleLabelServer(role)} accounts.` });
    }
    const user = {
      id: uid("u"),
      name,
      role,
      pin: String(pin),
      active: true,
      siteId: siteId || null,
      supervises: Array.isArray(supervises) ? supervises : [],
      phone: phone || "",
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

    // Whitelisted fields only - and PIN is left alone unless a new non-empty
    // one is actually provided, so an edit form with a blank PIN field can't
    // accidentally lock someone out.
    const { name, role, pin, siteId, supervises, active, phone, dateOfJoining } = req.body || {};
    if (name !== undefined && String(name).trim()) user.name = String(name).trim();
    if (role !== undefined) {
      if (!VALID_ROLES.includes(role)) return res.status(400).json({ error: "Unknown role." });
      // Prevent a non-Owner from promoting someone INTO a locked tier either.
      if (!canManageUserRole(req.user.role, role)) {
        return res.status(403).json({ error: `Only the Owner can set someone to ${roleLabelServer(role)}.` });
      }
      user.role = role;
    }
    if (pin !== undefined && String(pin).trim()) user.pin = String(pin).trim();
    if (siteId !== undefined) user.siteId = siteId || null;
    if (supervises !== undefined) user.supervises = Array.isArray(supervises) ? supervises : user.supervises;
    if (phone !== undefined) user.phone = phone || "";
    if (dateOfJoining !== undefined) user.dateOfJoining = dateOfJoining || "";
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
      dateOfJoining: dateOfJoining || "",
      drivingLevel: drivingLevel || "",
      performanceScore: null,
    };
    db.drivers.push(driver);
    await audit(req.user, "create_driver", `${req.user.name} added driver ${name}`);
    res.json(driver);
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
    if (esiCertificate) {
      const url = await savePhoto(esiCertificate, "driver_esi_" + driver.id);
      if (url) { driver.esiCertificateUrl = url; changed = true; }
    }
    if (pfCertificate) {
      const url = await savePhoto(pfCertificate, "driver_pf_" + driver.id);
      if (url) { driver.pfCertificateUrl = url; changed = true; }
    }
    if (changed) {
      await audit(req.user, "update_driver", `${req.user.name} updated driver ${driver.name} (${driver.id})`);
    }
    res.json(driver);
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

// Lets the currently-logged-in user re-confirm their own PIN mid-session -
// used as a "step-up" confirmation (not a separate login) before letting the
// Owner edit Assignments data, so an accidental tap can't silently change a
// vehicle/route/supervisor. Never exposes the stored PIN to the client.
app.post(
  "/api/verify-pin",
  requireAuth,
  h(async (req, res) => {
    const { pin } = req.body || {};
    // Deliberately always a 200 (never 401) - the frontend's generic api()
    // helper treats any 401 as "session expired" and force-logs-out, which
    // would be wrong here: an incorrect PIN re-entry is an expected, normal
    // outcome, not an invalid/expired session.
    if (!pin || String(pin) !== String(req.user.pin)) {
      return res.json({ ok: false, error: "Incorrect PIN." });
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
    return {
      reg: v.reg,
      company: v.clientId ? clientNameServer(v.clientId) : "Internal / Fixed Route",
      areaSupervisor: site && site.areaSupervisorId ? userNameServer(site.areaSupervisorId) : "",
      siteSupervisor: v.supervisorId ? userNameServer(v.supervisorId) : "",
      driver: temp ? temp.name : (r && r.attendance && r.attendance.driver) || (regularDriver ? regularDriver.name : ""),
      driverMobile: temp ? temp.phone : regularDriver ? regularDriver.phone || "" : "",
      route: v.route || (site ? site.name : ""),
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
    { label: "Driver", key: "driver" },
    { label: "Driver Mobile", key: "driverMobile" },
    { label: "Route", key: "route" },
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
function vehicleReg(id) {
  const v = db.vehicles.find((x) => x.id === id);
  return v ? v.reg : id || "";
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
      { label: "Vehicle", key: (e) => (e.vehicleId ? vehicleReg(e.vehicleId) : "") },
      { label: "Status", key: "status" },
      { label: "Decision Comment", key: "comment" },
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
    // Never include PINs in an exportable file, even though this is
    // restricted to Owner/Ops Manager - a downloaded file can end up
    // anywhere (email, a shared drive, a personal laptop).
    if (Array.isArray(clone.users)) clone.users.forEach((u) => delete u.pin);

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
