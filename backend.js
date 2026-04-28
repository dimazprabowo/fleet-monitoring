/**
 * FLEET MONITORING SYSTEM — GOOGLE APPS SCRIPT BACKEND
 * Versi: 2.0 (Blogger-Ready, JSONP, CORS-safe)
 *
 * CARA DEPLOY:
 *   1. Buka https://script.google.com → New Project
 *   2. Paste seluruh isi file ini ke Code.gs, Save
 *   3. Pastikan project ini terikat ke sebuah Google Spreadsheet
 *      (Resources / Container: Google Sheets). Jika project standalone,
 *      buka spreadsheet tujuan lalu Extensions → Apps Script,
 *      kemudian paste kode ini di sana.
 *   4. Run setupDatabase() SEKALI untuk inisialisasi seluruh sheet.
 *   5. Deploy → New Deployment → Web App
 *        - Execute as: Me
 *        - Who has access: Anyone
 *   6. Salin URL "/exec", paste ke variabel GAS_URL di file frontend.
 *   7. Set IS_PREVIEW = false di file frontend.
 *
 * DEFAULT LOGIN ADMIN:
 *   username: admin
 *   password: admin123
 */

const SS = SpreadsheetApp.getActiveSpreadsheet();
const SH_VEHICLES = "vehicles";
const SH_BOOKINGS = "bookings";
const SH_USERS    = "users";
const SH_SESSIONS = "sessions";

const SESSION_TTL_HOURS = 24;

// ============================================================
// MAIN HANDLER — semua via doGet + JSONP supaya aman di Blogger
// ============================================================
function doGet(e) {
  let result = { success: false, msg: "Action tidak ditemukan." };

  try {
    const action = (e && e.parameter && e.parameter.action) || "";
    const p = (e && e.parameter) || {};

    switch (action) {
      // ---- Public ----
      case "ping":             result = { success: true, msg: "Fleet Monitoring API aktif." }; break;
      case "getPublicData":    result = handleGetPublicData(p.date); break;
      case "submitBooking":    result = handleSubmitBooking(p); break;
      case "checkStatus":      result = handleCheckStatus(p.q); break;

      // ---- Auth ----
      case "login":            result = handleLogin(p.username, p.password); break;
      case "validateSession":  result = handleValidateSession(p.token); break;
      case "logout":           result = handleLogout(p.token); break;

      // ---- Admin (butuh session valid) ----
      case "getMonitoring":    result = withAuth(p.token, () => handleGetMonitoring(p.date)); break;
      case "getBookingHistory":result = withAuth(p.token, () => handleGetBookingHistory(p)); break;
      case "getDashboardStats":result = withAuth(p.token, () => handleGetDashboardStats()); break;
      case "updateBooking":    result = withAuth(p.token, (u) => handleUpdateBooking(p, u)); break;
      case "deleteBooking":    result = withAuth(p.token, () => handleDeleteBooking(p.id)); break;
      case "getVehicles":      result = withAuth(p.token, () => ({ success: true, data: getSheetData(SH_VEHICLES) })); break;
      case "addVehicle":       result = withAuth(p.token, () => handleAddVehicle(p)); break;
      case "updateVehicle":    result = withAuth(p.token, () => handleUpdateVehicle(p)); break;
      case "deleteVehicle":    result = withAuth(p.token, () => handleDeleteVehicle(p.id)); break;

      default:
        result = { success: false, msg: "Action '" + action + "' tidak dikenal." };
    }
  } catch (err) {
    result = { success: false, msg: "Server error: " + err.toString() };
    console.error("doGet error:", err);
  }

  // Output JSONP jika ada callback, JSON biasa jika tidak
  const output = ContentService.createTextOutput();
  const jsonStr = JSON.stringify(result);
  const cb = e && e.parameter && e.parameter.callback;
  if (cb) {
    output.setMimeType(ContentService.MimeType.JAVASCRIPT);
    output.setContent(cb + "(" + jsonStr + ");");
  } else {
    output.setMimeType(ContentService.MimeType.JSON);
    output.setContent(jsonStr);
  }
  return output;
}

// ============================================================
// SETUP DATABASE — Run sekali dari editor GAS
// ============================================================
function setupDatabase() {
  // 1. vehicles
  let vSh = SS.getSheetByName(SH_VEHICLES);
  if (!vSh) {
    vSh = SS.insertSheet(SH_VEHICLES);
    vSh.appendRow(["id","nama","plat","kategori","status","created_at"]);

    // Seed data dari sheet contoh
    const seed = [
      // Operasional Harian
      ["EXPANDER","B 2094 UYN"],["AVANZA","B 2386 UYV"],["EXPANDER","B 2668 UYZ"],
      ["EXPANDER","B 2684 UYP"],["EXPANDER","B 1860 DKG"],["VELLOZ","B 1181 DKO"],
      ["AVANZA","B 1856 ROZ"],["RUSH","B 1137 VZB"],["AVANZA","B 1776 RYA"],
      ["AVANZA","B 1601 RZQ"],["AVANZA","B 2753 UYO"],["EXPANDER","B 1421 DKU"],
      ["EXPANDER","B 1574 DOD"],["EXPANDER","B 2189 UZD"],["XL7","B 2144 UYI"],
    ].map(r => [...r, "Operasional Harian"]);

    const proj = [
      ["TOYOTA INNOVA","BG 1896 NH"],["TOYOTA HILUX","B 9420 UBC"],["TOYOTA HILUX","B 9118 UBE"],
      ["EXPANDER","B 1080 DOF"],["TOYOTA INNOVA","DK 1992 ACG"],["MITSUBISHI TRITON","BG 8587 GD"],
      ["MITSUBISHI TRITON","BG 8340 CG"],["AVANZA","B 1576 RZE"],["INNOVA ZENIX NON HYBRID","B 2191 UYX"],
      ["INNOVA ZENIX","B 1893 DYE"],["HONDA BRV","B 1727 RYA"],["AVANZA","B 1069 RZC"],
    ].map(r => [...r, "Project"]);

    const struk = [
      ["INNOVA ZENIX","B 1083 VZH"],["HONDA BRV","B 1625 RZF"],["HONDA BRV","B 1231 RZF"],
      ["HONDA BRV","B 1290 RZD"],["HONDA HRV","B 1036 RYB"],["HONDA HRV","B 1907 DZC"],
      ["INNOVA ZENIX","B 1353 VZO"],
    ].map(r => [...r, "Operasional Struktural"]);

    [...seed, ...proj, ...struk].forEach((r, i) => {
      vSh.appendRow(["V" + Utilities.formatString("%03d", i + 1), r[0], r[1], r[2], "Tersedia", new Date()]);
    });
  }

  // 2. bookings
  if (!SS.getSheetByName(SH_BOOKINGS)) {
    const sh = SS.insertSheet(SH_BOOKINGS);
    sh.appendRow([
      "id","tanggal","vehicle_id","pic","divisi","tujuan","status",
      "driver","keterangan","created_at","updated_at","updated_by","durasi"
    ]);
  }
  ensureBookingsSchema();

  // 3. users
  let uSh = SS.getSheetByName(SH_USERS);
  if (!uSh) {
    uSh = SS.insertSheet(SH_USERS);
    uSh.appendRow(["id","username","password","name","role","created_at"]);
    uSh.appendRow(["U001","admin","admin123","Administrator","Admin", new Date()]);
  }

  // 4. sessions
  if (!SS.getSheetByName(SH_SESSIONS)) {
    const sh = SS.insertSheet(SH_SESSIONS);
    sh.appendRow(["token","user_id","user_name","user_role","expires_at","created_at"]);
  }

  // Format header
  [SH_VEHICLES, SH_BOOKINGS, SH_USERS, SH_SESSIONS].forEach(name => {
    const sh = SS.getSheetByName(name);
    if (sh) sh.getRange(1, 1, 1, sh.getLastColumn())
              .setFontWeight("bold").setBackground("#E8F0FE");
  });

  return "✅ Database siap. Login default: admin / admin123";
}

// ============================================================
// PUBLIC HANDLERS
// ============================================================
function handleGetPublicData(date) {
  const today = date || isoDate(new Date());
  const vehicles = getSheetData(SH_VEHICLES);
  const all = getSheetData(SH_BOOKINGS);
  const bookings = all.filter(b => b.tanggal === today);
  // Jadwal aktif: semua booking non-Rejected dari 30 hari lalu sd 180 hari ke depan
  const todayDate = new Date();
  const minStr = isoDate(new Date(todayDate.getTime() - 30 * 86400000));
  const maxStr = isoDate(new Date(todayDate.getTime() + 180 * 86400000));
  const schedule = all.filter(b =>
    b.status !== "Rejected" && b.tanggal >= minStr && b.tanggal <= maxStr
  );
  return { success: true, vehicles, bookings, schedule, date: today };
}

function handleSubmitBooking(p) {
  const required = ["pic","divisi","vehicle_id","tanggal","tujuan"];
  for (const f of required) {
    if (!p[f] || String(p[f]).trim() === "") {
      return { success: false, msg: "Field '" + f + "' wajib diisi." };
    }
  }

  const sh = SS.getSheetByName(SH_BOOKINGS);
  if (!sh) return { success: false, msg: "Sheet bookings tidak ditemukan. Jalankan setupDatabase()." };

  ensureBookingsSchema();

  // Cek status armada — hanya yang "Tersedia" yang boleh dibooking
  const vehicle = getSheetData(SH_VEHICLES).find(x => x.id === p.vehicle_id);
  if (!vehicle) return { success: false, msg: "Armada tidak ditemukan." };
  if (vehicle.status && vehicle.status !== "Tersedia") {
    return { success: false, msg: "Armada sedang " + vehicle.status + " dan tidak dapat dibooking." };
  }

  // Durasi (lama penggunaan, dalam hari). Default 1 = 1 hari pakai.
  const durasi = Math.max(1, parseInt(p.durasi) || 1);
  const newEnd = addDaysIso(p.tanggal, durasi - 1);

  // Cek bentrok range vs semua booking aktif (Pending/Approved) di vehicle yang sama
  const existing = getSheetData(SH_BOOKINGS).find(b => {
    if (b.vehicle_id !== p.vehicle_id) return false;
    if (b.status === "Rejected" || b.status === "Completed") return false;
    const bDur = Math.max(1, parseInt(b.durasi) || 1);
    const bEnd = addDaysIso(b.tanggal, bDur - 1);
    // Range overlap: [p.tanggal, newEnd] ∩ [b.tanggal, bEnd]
    return p.tanggal <= bEnd && newEnd >= b.tanggal;
  });
  if (existing) {
    const eDur = Math.max(1, parseInt(existing.durasi) || 1);
    const eEnd = addDaysIso(existing.tanggal, eDur - 1);
    return { success: false, msg: "Tanggal bentrok dengan booking " + existing.tanggal + " → " + eEnd + " (status: " + existing.status + ")." };
  }

  const id = "BK-" + new Date().getTime();
  const now = new Date();
  // Tulis pakai header-mapped supaya aman terhadap posisi kolom 'durasi'
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const row = new Array(headers.length).fill("");
  const fields = {
    id: id, tanggal: p.tanggal, vehicle_id: p.vehicle_id,
    pic: p.pic.trim(), divisi: p.divisi.trim(), tujuan: p.tujuan.trim(),
    status: "Pending", driver: "", keterangan: "",
    created_at: now, updated_at: now, updated_by: "", durasi: durasi
  };
  Object.keys(fields).forEach(k => {
    const c = headers.indexOf(k);
    if (c !== -1) row[c] = fields[k];
  });
  sh.appendRow(row);
  return { success: true, msg: "Booking berhasil diajukan.", id };
}

// Tambah kolom 'durasi' ke sheet bookings jika belum ada (migration)
function ensureBookingsSchema() {
  const sh = SS.getSheetByName(SH_BOOKINGS);
  if (!sh) return;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (headers.indexOf("durasi") === -1) {
    sh.getRange(1, sh.getLastColumn() + 1).setValue("durasi");
    // Backfill default 1 untuk row lama
    const n = sh.getLastRow() - 1;
    if (n > 0) {
      const col = sh.getLastColumn();
      sh.getRange(2, col, n, 1).setValues(Array.from({length: n}, () => [1]));
    }
    sh.getRange(1, sh.getLastColumn()).setFontWeight("bold").setBackground("#E8F0FE");
  }
}

// Tambah n hari ke ISO date string (yyyy-mm-dd), return ISO
function addDaysIso(iso, n) {
  const parts = String(iso).split("-").map(x => parseInt(x, 10));
  const d = new Date(parts[0], parts[1] - 1, parts[2]);
  d.setDate(d.getDate() + n);
  return isoDate(d);
}

function handleCheckStatus(q) {
  if (!q) return { success: false, msg: "Masukkan nama PIC." };
  const needle = String(q).toLowerCase().trim();
  const vehicles = getSheetData(SH_VEHICLES);
  const vMap = {};
  vehicles.forEach(v => vMap[v.id] = v);

  const data = getSheetData(SH_BOOKINGS)
    .filter(b => b.pic && b.pic.toLowerCase().includes(needle))
    .map(b => {
      const v = vMap[b.vehicle_id] || {};
      return { ...b, vehicle_nama: v.nama || "-", vehicle_plat: v.plat || "-", kategori: v.kategori || "-" };
    })
    .sort((a, b) => (b.tanggal || "").localeCompare(a.tanggal || ""));

  return { success: true, data };
}

// ============================================================
// ADMIN HANDLERS
// ============================================================
function handleGetMonitoring(date) {
  const target = date || isoDate(new Date());
  const vehicles = getSheetData(SH_VEHICLES);
  const all = getSheetData(SH_BOOKINGS);
  const bookings = all.filter(b => b.tanggal === target);
  // Schedule: semua booking non-Rejected dari 60 hari lalu sd 365 hari ke depan
  const targetDate = new Date(target);
  const minStr = isoDate(new Date(targetDate.getTime() - 60 * 86400000));
  const maxStr = isoDate(new Date(targetDate.getTime() + 365 * 86400000));
  const schedule = all.filter(b =>
    b.status !== "Rejected" && b.tanggal >= minStr && b.tanggal <= maxStr
  );
  return { success: true, vehicles, bookings, schedule, date: target };
}

function handleGetBookingHistory(p) {
  const vehicles = getSheetData(SH_VEHICLES);
  const vMap = {};
  vehicles.forEach(v => vMap[v.id] = v);

  let rows = getSheetData(SH_BOOKINGS).map(b => {
    const v = vMap[b.vehicle_id] || {};
    return { ...b, vehicle_nama: v.nama || "-", vehicle_plat: v.plat || "-", kategori: v.kategori || "-" };
  });

  if (p.from)     rows = rows.filter(r => (r.tanggal || "") >= p.from);
  if (p.to)       rows = rows.filter(r => (r.tanggal || "") <= p.to);
  if (p.status)   rows = rows.filter(r => r.status === p.status);
  if (p.kategori) rows = rows.filter(r => r.kategori === p.kategori);

  rows.sort((a, b) => (b.tanggal || "").localeCompare(a.tanggal || ""));
  return { success: true, data: rows };
}

function handleGetDashboardStats() {
  const today = isoDate(new Date());
  const vehicles = getSheetData(SH_VEHICLES);
  const bookings = getSheetData(SH_BOOKINGS);
  const todays = bookings.filter(b => b.tanggal === today);

  return {
    success: true,
    stats: {
      total_vehicles: vehicles.length,
      by_category: {
        "Operasional Harian":      vehicles.filter(v => v.kategori === "Operasional Harian").length,
        "Project":                 vehicles.filter(v => v.kategori === "Project").length,
        "Operasional Struktural":  vehicles.filter(v => v.kategori === "Operasional Struktural").length,
      },
      bookings_today:  todays.length,
      used_today:      todays.filter(b => b.status === "Approved").length,
      pending_today:   todays.filter(b => b.status === "Pending").length,
      standby_today:   vehicles.length - todays.filter(b => b.status === "Approved").length,
      total_bookings:  bookings.length,
      pending_total:   bookings.filter(b => b.status === "Pending").length,
    }
  };
}

function handleUpdateBooking(p, user) {
  if (!p.id) return { success: false, msg: "ID booking wajib ada." };
  const sh = SS.getSheetByName(SH_BOOKINGS);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf("id");

  ensureBookingsSchema();
  const updatable = ["status","driver","keterangan","tujuan","pic","divisi","tanggal","durasi"];

  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === p.id) {
      updatable.forEach(f => {
        if (p[f] !== undefined && p[f] !== null) {
          const c = headers.indexOf(f);
          if (c !== -1) sh.getRange(i + 1, c + 1).setValue(p[f]);
        }
      });
      sh.getRange(i + 1, headers.indexOf("updated_at") + 1).setValue(new Date());
      sh.getRange(i + 1, headers.indexOf("updated_by") + 1).setValue(user.name || "");
      return { success: true, msg: "Booking diperbarui." };
    }
  }
  return { success: false, msg: "Booking tidak ditemukan." };
}

function handleDeleteBooking(id) {
  if (!id) return { success: false, msg: "ID wajib ada." };
  const sh = SS.getSheetByName(SH_BOOKINGS);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) { sh.deleteRow(i + 1); return { success: true, msg: "Booking dihapus." }; }
  }
  return { success: false, msg: "Booking tidak ditemukan." };
}

function handleAddVehicle(p) {
  if (!p.nama || !p.plat || !p.kategori) {
    return { success: false, msg: "nama, plat, kategori wajib diisi." };
  }
  const sh = SS.getSheetByName(SH_VEHICLES);
  const id = "V" + new Date().getTime();
  sh.appendRow([id, String(p.nama).trim().toUpperCase(), String(p.plat).trim().toUpperCase(),
                p.kategori, "Tersedia", new Date()]);
  return { success: true, msg: "Armada ditambahkan.", id };
}

function handleUpdateVehicle(p) {
  if (!p.id) return { success: false, msg: "ID wajib ada." };
  const sh = SS.getSheetByName(SH_VEHICLES);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const fields = ["nama","plat","kategori","status"];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === p.id) {
      fields.forEach(f => {
        if (p[f] !== undefined && p[f] !== null) {
          const c = headers.indexOf(f);
          if (c !== -1) {
            const val = (f === "nama" || f === "plat") ? String(p[f]).trim().toUpperCase() : p[f];
            sh.getRange(i + 1, c + 1).setValue(val);
          }
        }
      });
      return { success: true, msg: "Armada diperbarui." };
    }
  }
  return { success: false, msg: "Armada tidak ditemukan." };
}

function handleDeleteVehicle(id) {
  if (!id) return { success: false, msg: "ID wajib ada." };
  const sh = SS.getSheetByName(SH_VEHICLES);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) { sh.deleteRow(i + 1); return { success: true, msg: "Armada dihapus." }; }
  }
  return { success: false, msg: "Armada tidak ditemukan." };
}

// ============================================================
// AUTH
// ============================================================
function handleLogin(username, password) {
  if (!username || !password) return { success: false, msg: "Username & password wajib diisi." };
  const sh = SS.getSheetByName(SH_USERS);
  if (!sh) return { success: false, msg: "Sheet users tidak ditemukan." };

  const data = sh.getDataRange().getValues();
  const h = data[0];
  const cU = h.indexOf("username"), cP = h.indexOf("password"),
        cN = h.indexOf("name"), cR = h.indexOf("role");

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][cU]) === username && String(data[i][cP]) === password) {
      const user = { id: data[i][0], username, name: data[i][cN], role: data[i][cR] };
      const token = Utilities.getUuid();
      const now = new Date();
      const exp = new Date(now.getTime() + SESSION_TTL_HOURS * 3600 * 1000);
      SS.getSheetByName(SH_SESSIONS).appendRow([token, user.id, user.name, user.role, exp, now]);
      return { success: true, user, token };
    }
  }
  return { success: false, msg: "Username atau password salah." };
}

function handleValidateSession(token) {
  const u = getUserByToken(token);
  if (!u) return { success: false, msg: "Session tidak valid / sudah expired." };
  return { success: true, user: u };
}

function handleLogout(token) {
  if (!token) return { success: false, msg: "Token kosong." };
  const sh = SS.getSheetByName(SH_SESSIONS);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === token) { sh.deleteRow(i + 1); return { success: true, msg: "Logout berhasil." }; }
  }
  return { success: false, msg: "Session tidak ditemukan." };
}

function getUserByToken(token) {
  if (!token) return null;
  const sh = SS.getSheetByName(SH_SESSIONS);
  if (!sh) return null;
  const data = sh.getDataRange().getValues();
  const h = data[0];
  const cT = h.indexOf("token"), cExp = h.indexOf("expires_at");
  const now = new Date();
  for (let i = 1; i < data.length; i++) {
    if (data[i][cT] === token) {
      const exp = data[i][cExp];
      if (exp instanceof Date && exp < now) { sh.deleteRow(i + 1); return null; }
      return {
        id:   data[i][h.indexOf("user_id")],
        name: data[i][h.indexOf("user_name")],
        role: data[i][h.indexOf("user_role")],
      };
    }
  }
  return null;
}

function withAuth(token, fn) {
  const user = getUserByToken(token);
  if (!user) return { success: false, msg: "Unauthorized. Silakan login ulang.", needLogin: true };
  return fn(user);
}

// ============================================================
// HELPERS
// ============================================================
function isoDate(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return y + "-" + m + "-" + day;
}

function getSheetData(sheetName) {
  const sh = SS.getSheetByName(sheetName);
  if (!sh) return [];
  const raw = sh.getDataRange().getValues();
  if (raw.length <= 1) return [];
  const headers = raw[0];
  return raw.slice(1).map(row => {
    const obj = {};
    headers.forEach((header, i) => {
      const v = row[i];
      if (v instanceof Date) {
        if (header === "tanggal") obj[header] = isoDate(v);
        else obj[header] = v.toISOString();
      } else {
        obj[header] = v;
      }
    });
    return obj;
  }).filter(o => o.id);
}
