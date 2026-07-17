/**
 * Fleet Supervisor App - shared backend
 * ---------------------------------------------------------------------
 * A small, dependency-light Express server that gives every supervisor,
 * area supervisor, HR, operations manager, data-team member and the
 * owner a single shared login + shared dataset - which is what makes
 * this different from the earlier click-through prototype (which only
 * stored data in one person's browser).
 *
 * Data lives in ./data/data.json (created from ./data/seed.json on first
 * run). Photos/selfies are written to ./uploads and served statically.
 * This is intentionally simple (a JSON file, not a real database) so it
 * deploys in minutes on a free host - see DEPLOY.md. For your full 250+
 * vehicle fleet in daily production use, move this to a real managed
 * database (see the spec doc, Phase 4 - Hardening).
 * ---------------------------------------------------------------------
 */
const express = require("express");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

const DATA_DIR = path.join(__dirname, "data");
const DATA_FILE = path.join(DATA_DIR, "data.json");
const SEED_FILE = path.join(DATA_DIR, "seed.json");
const UPLOADS_DIR = path.join(__dirname, "uploads");
const PORT = process.env.PORT || 3000;

if (!fs.existsSync(UPLOADS_DIR)) fs.mkdirSync(UPLOADS_DIR, { recursive: true });
if (!fs.existsSync(DATA_FILE)) {
  fs.copyFileSync(SEED_FILE, DATA_FILE);
}

// ---------- tiny JSON "database" ----------
let db = JSON.parse(fs.readFileSync(DATA_FILE, "utf8"));
// Writes happen synchronously and immediately (no debounce) - this is a small
// JSON file, so the cost is trivial, and it means a save() call has actually
// hit disk before the HTTP response is sent. Don't change this to a debounced/
// async write without also handling process-exit flushing - free-tier hosts
// can kill the process on redeploy/restart with very little notice.
function save() {
  fs.writeFileSync(DATA_FILE, JSON.stringify(db, null, 2));
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

function audit(user, action, detail) {
  db.auditLog.unshift({
    id: uid("a"),
    ts: nowIso(),
    userId: user ? user.id : "system",
    userName: user ? user.name : "System",
    action,
    detail,
  });
  if (db.auditLog.length > 5000) db.auditLog.length = 5000;
  save();
}

function saveBase64Image(dataUrl, tag) {
  if (!dataUrl || typeof dataUrl !== "string" || !dataUrl.startsWith("data:")) return null;
  const match = dataUrl.match(/^data:(image\/\w+);base64,(.+)$/);
  if (!match) return null;
  const ext = match[1].split("/")[1] || "jpg";
  const filename = `${tag || "img"}_${Date.now()}_${crypto.randomBytes(4).toString("hex")}.${ext}`;
  fs.writeFileSync(path.join(UPLOADS_DIR, filename), Buffer.from(match[2], "base64"));
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
app.use("/uploads", express.static(UPLOADS_DIR));
app.use(express.static(path.join(__dirname, "public")));

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

// ---------- AUTH ROUTES ----------
app.get("/api/users/login-list", (req, res) => {
  res.json(
    db.users
      .filter((u) => u.active !== false)
      .map((u) => ({ id: u.id, name: u.name, role: u.role }))
  );
});

app.post("/api/login", (req, res) => {
  const { id, pin } = req.body || {};
  const user = db.users.find((u) => u.id === id && u.active !== false);
  if (!user || String(user.pin) !== String(pin)) {
    return res.status(401).json({ error: "Wrong user or PIN." });
  }
  const token = crypto.randomBytes(24).toString("hex");
  sessions.set(token, user.id);
  audit(user, "login", `${user.name} (${user.role}) logged in`);
  res.json({ token, user: publicUser(user) });
});

app.post("/api/logout", requireAuth, (req, res) => {
  const header = req.headers.authorization || "";
  const token = header.slice(7);
  sessions.delete(token);
  audit(req.user, "logout", `${req.user.name} logged out`);
  res.json({ ok: true });
});

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

app.post("/api/users", requireAuth, requireRole(...ADMIN_ROLES), (req, res) => {
  const { name, role, pin, siteId, supervises } = req.body || {};
  const validRoles = ["site_supervisor", "area_supervisor", "ops_manager", "data_team", "owner", "hr"];
  if (!name || !validRoles.includes(role) || !pin) {
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
  audit(req.user, "create_user", `${req.user.name} added ${name} as ${role}`);
  save();
  res.json(publicUser(user));
});

app.patch("/api/users/:id", requireAuth, requireRole(...ADMIN_ROLES), (req, res) => {
  const user = db.users.find((u) => u.id === req.params.id);
  if (!user) return res.status(404).json({ error: "User not found." });
  const before = { ...user };
  Object.assign(user, req.body, { id: user.id });
  audit(req.user, "update_user", `${req.user.name} updated ${user.name}: ${JSON.stringify(before)} -> ${JSON.stringify(user)}`);
  save();
  res.json(publicUser(user));
});

// ---------- CLIENTS ----------
app.post("/api/clients", requireAuth, requireRole(...ADMIN_ROLES), (req, res) => {
  const { name } = req.body || {};
  if (!name) return res.status(400).json({ error: "name is required." });
  const client = { id: uid("c"), name };
  db.clients.push(client);
  audit(req.user, "create_client", `${req.user.name} added client ${name}`);
  save();
  res.json(client);
});

// ---------- SITES ----------
app.post("/api/sites", requireAuth, requireRole(...ADMIN_ROLES), (req, res) => {
  const { name, lat, lng, areaSupervisorId } = req.body || {};
  if (!name) return res.status(400).json({ error: "name is required." });
  const site = { id: uid("s"), name, lat: lat != null ? Number(lat) : null, lng: lng != null ? Number(lng) : null, areaSupervisorId: areaSupervisorId || null };
  db.sites.push(site);
  audit(req.user, "create_site", `${req.user.name} added site ${name}`);
  save();
  res.json(site);
});

// ---------- DRIVERS ----------
app.post("/api/drivers", requireAuth, requireRole(...ADMIN_ROLES), (req, res) => {
  const { name, phone, licenseNumber } = req.body || {};
  if (!name) return res.status(400).json({ error: "name is required." });
  const driver = { id: uid("d"), name, phone: phone || "", licenseNumber: licenseNumber || "" };
  db.drivers.push(driver);
  audit(req.user, "create_driver", `${req.user.name} added driver ${name}`);
  save();
  res.json(driver);
});

// ---------- VEHICLES ----------
function visibleVehiclesFor(user) {
  if (user.role === "site_supervisor") return db.vehicles.filter((v) => v.supervisorId === user.id);
  if (user.role === "area_supervisor") {
    const mySupervisors = new Set((user.supervises || []));
    return db.vehicles.filter((v) => mySupervisors.has(v.supervisorId));
  }
  return db.vehicles;
}

app.get("/api/vehicles", requireAuth, (req, res) => {
  res.json(visibleVehiclesFor(req.user));
});

app.post("/api/vehicles", requireAuth, requireRole(...ADMIN_ROLES), (req, res) => {
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
  audit(req.user, "create_vehicle", `${req.user.name} added vehicle ${reg}`);
  save();
  res.json(vehicle);
});

// Assignment - the new capability: Ops Manager / Owner / Data Team can
// assign a vehicle (and its driver) to a site + supervisor + client.
app.patch("/api/vehicles/:id/assign", requireAuth, requireRole(...ADMIN_ROLES), (req, res) => {
  const vehicle = db.vehicles.find((v) => v.id === req.params.id);
  if (!vehicle) return res.status(404).json({ error: "Vehicle not found." });
  const before = { siteId: vehicle.siteId, supervisorId: vehicle.supervisorId, driverId: vehicle.driverId, clientId: vehicle.clientId };
  const { siteId, supervisorId, driverId, clientId, usage, standardMileage } = req.body || {};
  if (siteId !== undefined) vehicle.siteId = siteId || null;
  if (supervisorId !== undefined) vehicle.supervisorId = supervisorId || null;
  if (driverId !== undefined) vehicle.driverId = driverId || null;
  if (clientId !== undefined) vehicle.clientId = clientId || null;
  if (usage !== undefined) vehicle.usage = usage;
  if (standardMileage !== undefined) vehicle.standardMileage = Number(standardMileage) || vehicle.standardMileage;
  audit(
    req.user,
    "assign_vehicle",
    `${req.user.name} reassigned ${vehicle.reg}: ${JSON.stringify(before)} -> ${JSON.stringify({ siteId: vehicle.siteId, supervisorId: vehicle.supervisorId, driverId: vehicle.driverId, clientId: vehicle.clientId })}`
  );
  save();
  res.json(vehicle);
});

app.patch("/api/vehicles/:id/docs", requireAuth, (req, res) => {
  const vehicle = db.vehicles.find((v) => v.id === req.params.id);
  if (!vehicle) return res.status(404).json({ error: "Vehicle not found." });
  const isOwnerSupervisor = req.user.role === "site_supervisor" && vehicle.supervisorId === req.user.id;
  if (!isOwnerSupervisor && !ADMIN_ROLES.includes(req.user.role)) {
    return res.status(403).json({ error: "Not allowed to edit this vehicle's documents." });
  }
  const { docType, number, expiry } = req.body || {};
  if (!vehicle.docs[docType]) return res.status(400).json({ error: "Unknown document type." });
  vehicle.docs[docType] = { number: number ?? vehicle.docs[docType].number, expiry: expiry ?? vehicle.docs[docType].expiry };
  audit(req.user, "update_doc", `${req.user.name} updated ${docType} on ${vehicle.reg}`);
  save();
  res.json(vehicle);
});

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

app.post("/api/records", requireAuth, requireRole("site_supervisor"), (req, res) => {
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
    Object.entries(body.photoTags).forEach(([tag, arr]) => {
      photoTags[tag] = (Array.isArray(arr) ? arr : []).map((src) => {
        if (typeof src === "string" && src.startsWith("data:")) {
          return saveBase64Image(src, tag) || src;
        }
        return src; // already a saved URL
      });
    });
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
    },
    verification: { status: "pending", verifiedBy: null, comment: "" },
    submittedBy: req.user.id,
    submittedAt: nowIso(),
  };
  db.records[recordKey(vehicle.id, date)] = record;
  audit(req.user, "submit_record", `${req.user.name} submitted daily record for ${vehicle.reg} (${date})`);
  save();
  res.json(record);
});

app.patch("/api/records/:vehicleId/:date/verify", requireAuth, requireRole("area_supervisor", ...ADMIN_ROLES), (req, res) => {
  const key = recordKey(req.params.vehicleId, req.params.date);
  const record = db.records[key];
  if (!record) return res.status(404).json({ error: "Record not found." });
  const { status, comment } = req.body || {};
  if (!["approved", "rejected"].includes(status)) return res.status(400).json({ error: "status must be approved or rejected." });
  record.verification = { status, verifiedBy: req.user.id, comment: comment || "" };
  audit(req.user, status === "approved" ? "approve" : "reject", `${req.user.name} ${status} ${req.params.vehicleId} (${req.params.date})${comment ? ": " + comment : ""}`);
  save();
  res.json(record);
});

app.patch("/api/records/:vehicleId/:date/correct", requireAuth, requireRole("data_team", "owner"), (req, res) => {
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
  audit(req.user, "correction", `${req.user.name} corrected ${fieldPath} on ${req.params.vehicleId} (${req.params.date}): ${before} -> ${value}`);
  save();
  res.json(record);
});

// ---------- LEAVES ----------
app.get("/api/leaves", requireAuth, (req, res) => res.json(db.leaves));

app.post("/api/leaves", requireAuth, (req, res) => {
  const { driverId, driverName, type, start, end } = req.body || {};
  if (!driverName || !start || !end) return res.status(400).json({ error: "driverName, start, end are required." });
  const leave = { id: uid("l"), driverId: driverId || null, driver: driverName, type: type || "Planned", start, end, status: "pending", requestedBy: req.user.id };
  db.leaves.push(leave);
  audit(req.user, "leave_request", `${req.user.name} requested leave for ${driverName} (${start} to ${end})`);
  save();
  res.json(leave);
});

app.patch("/api/leaves/:id", requireAuth, requireRole("ops_manager", "owner"), (req, res) => {
  const leave = db.leaves.find((l) => l.id === req.params.id);
  if (!leave) return res.status(404).json({ error: "Leave not found." });
  const { status } = req.body || {};
  if (!["approved", "rejected"].includes(status)) return res.status(400).json({ error: "status must be approved or rejected." });
  leave.status = status;
  audit(req.user, "leave_decision", `${req.user.name} set leave ${leave.id} (${leave.driver}) to ${status}`);
  save();
  res.json(leave);
});

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

app.post("/api/expenses", requireAuth, (req, res) => {
  const { amount, category, description, vehicleId, bill } = req.body || {};
  if (!amount || Number(amount) <= 0) return res.status(400).json({ error: "A valid amount is required." });
  if (!bill) return res.status(400).json({ error: "A photo of the bill/receipt is required for every expense request." });
  const billUrl = saveBase64Image(bill, "bill");
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
  audit(req.user, "expense_request", `${req.user.name} requested ₹${expense.amount} (${expense.category})`);
  save();
  res.json(expense);
});

app.patch("/api/expenses/:id/decide", requireAuth, requireRole("ops_manager", "owner"), (req, res) => {
  const expense = db.expenses.find((e) => e.id === req.params.id);
  if (!expense) return res.status(404).json({ error: "Expense not found." });
  const { status, comment } = req.body || {};
  if (!["approved", "rejected"].includes(status)) return res.status(400).json({ error: "status must be approved or rejected." });
  expense.status = status;
  expense.decidedBy = req.user.id;
  expense.comment = comment || "";
  audit(req.user, "expense_decision", `${req.user.name} ${status} ₹${expense.amount} expense from ${expense.userName}${comment ? ": " + comment : ""}`);
  save();
  res.json(expense);
});

// ---------- SUPERVISOR ATTENDANCE (selfie + geolocation) ----------
function siteFor(user) {
  return db.sites.find((s) => s.id === user.siteId) || null;
}

app.post("/api/attendance/:type(clockin|clockout)", requireAuth, (req, res) => {
  const { selfie, lat, lng } = req.body || {};
  const site = siteFor(req.user);
  const selfieUrl = saveBase64Image(selfie, "selfie_" + req.params.type);
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
  audit(
    req.user,
    req.params.type,
    `${req.user.name} ${req.params.type === "clockin" ? "clocked in" : "clocked out"}${site ? " at " + site.name : ""}${distance != null ? ` (${distance}m from site, ${withinGeofence ? "within" : "OUTSIDE"} geofence)` : " (no site location on file)"}`
  );
  save();
  res.json(entry);
});

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

// ---------- fallback to the app shell ----------
app.get("*", (req, res, next) => {
  if (req.path.startsWith("/api/") || req.path.startsWith("/uploads/")) return next();
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

app.listen(PORT, () => {
  console.log(`Fleet Supervisor App listening on port ${PORT}`);
});
