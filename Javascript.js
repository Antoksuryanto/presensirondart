// =========================================================================
// ID Spreadsheet Google Anda.
// Ganti dengan ID spreadsheet tempat script ini berada.
// Cara dapatkan ID: buka spreadsheet, URL-nya seperti:
//   https://docs.google.com/spreadsheets/d/ID_SPREADSHEET_ANDA/edit
// ID-nya adalah bagian antara /d/ dan /edit
// =========================================================================
var SPREADSHEET_ID = "1G4MERfP-il--aPfqSLgndx3TWt9_0r1Kab-sKKectnDM7SseG4etK0_k";  // <-- GANTI INI

// Nama sheet untuk tiap kategori data (bisa diubah sesuai keinginan)
var SHEETS = {
  AKUN:       "AKUN_LOGIN",
  KELOMPOK:   "KELOMPOK_RONDA",
  REKAP:      "REKAP_PRESENSI",
  KEJADIAN:   "KEJADIAN",
  TAMU:       "TAMU",
  PENGATURAN: "PENGATURAN",
  RIWAYAT:    "RIWAYAT"
};

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({status: "success", message: "Koneksi API Siskamling Aktif"}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  // CORS tidak diperlukan untuk Apps Script Web App (fetch lintas origin sudah diizinkan)
  if (e.parameter && e.parameter._method === "OPTIONS") {
    return ContentService.createTextOutput("")
      .setMimeType(ContentService.MimeType.TEXT);
  }

  try {
    var postData = JSON.parse(e.postData.contents);
    var action = postData.action;
    var data = postData.data;
    var result;

    initializeDatabase();

    switch (action) {
      case "getInitialData":
        result = loadAllData();
        break;
      case "saveData":
        result = saveAllData(data);
        break;
      case "resetTodayAttendance":
        result = resetTodayAttendance();
        break;
      default:
        result = { status: "error", message: "Aksi tidak dikenali" };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// --- Spreadsheet Accessor (robust) ---
function getSpreadsheet() {
  // 1. Coba dapatkan active spreadsheet (untuk script terikat di spreadsheet)
  try {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) {
      return active;
    }
  } catch (ex) {}

  // 2. Fallback: buka via Spreadsheet ID eksplisit
  if (SPREADSHEET_ID && SPREADSHEET_ID !== "AKfycbx...ganttung_ID_spreadsheet_disini") {
    try {
      return SpreadsheetApp.openById(SPREADSHEET_ID);
    } catch (ex) {
      throw new Error("Gagal membuka spreadsheet dengan ID: " + ex.toString());
    }
  }

  throw new Error("Tidak dapat mengakses spreadsheet. Pastikan script terikat ke spreadsheet atau isi SPREADSHEET_ID yang valid.");
}

// --- Inisialisasi: buat semua sheet + header jika belum ada ---
function initializeDatabase() {
  var ss = getSpreadsheet();
  ensureSheet(ss, SHEETS.AKUN,       ["username", "password", "name", "role", "regu"]);
  ensureSheet(ss, SHEETS.KELOMPOK,   ["regu_key", "regu_name", "ketua", "member_name", "member_role"]);
  ensureSheet(ss, SHEETS.REKAP,      ["date", "regu_key", "regu_name", "member_name", "status", "time"]);
  ensureSheet(ss, SHEETS.KEJADIAN,   ["id", "reporter", "kategori", "lokasi", "deskripsi", "time"]);
  ensureSheet(ss, SHEETS.TAMU,       ["id", "plat", "jenis", "tujuan", "time"]);
  ensureSheet(ss, SHEETS.PENGATURAN, ["key", "value"]);
  ensureSheet(ss, SHEETS.RIWAYAT,    ["time", "type", "icon", "color", "title", "desc"]);
}

function ensureSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  }
  return sheet;
}

function clearSheet(sheet) {
  var last = sheet.getLastRow();
  if (last > 1) {
    sheet.deleteRows(2, last - 1);
  }
}

// =========================================================================
// LOAD: baca semua sheet -> gabung jadi satu objek database
// =========================================================================
function loadAllData() {
  var ss = getSpreadsheet();
  initializeDatabase();

  var usersDB = loadAkun(ss);
  var regus = loadKelompok(ss);
  var rekap = loadRekap(ss);
  var incidentReports = loadKejadian(ss);
  var guestList = loadTamu(ss);
  var geofenceSettings = loadPengaturan(ss);
  var appHistory = loadRiwayat(ss);

  // Terapkan presensi hari ini dari REKAP_PRESENSI ke anggota kelompok
  var todayKey = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  rekap.forEach(function (r) {
    if (r.date === todayKey && regus[r.regu_key]) {
      var member = regus[r.regu_key].members.find(function (m) { return m.name === r.member_name; });
      if (member) {
        member.status = r.status;
        member.time = r.time;
      }
    }
  });

  // Jika semua sheet kosong -> beri tahu frontend agar mengirim data awal
  var isEmpty = Object.keys(usersDB).length === 0 &&
                Object.keys(regus).length === 0 &&
                incidentReports.length === 0 &&
                guestList.length === 0 &&
                rekap.length === 0;
  if (isEmpty) {
    return { status: "empty", message: "Belum ada data tersimpan di awan." };
  }

  return {
    status: "success",
    data: {
      usersDB: usersDB,
      regus: regus,
      appHistory: appHistory,
      incidentReports: incidentReports,
      guestList: guestList,
      attendanceHistory: rekap,
      geofenceSettings: geofenceSettings
    }
  };
}

function loadAkun(ss) {
  var sheet = ss.getSheetByName(SHEETS.AKUN);
  var data = sheet.getDataRange().getValues();
  var users = {};
  for (var i = 1; i < data.length; i++) {
    var username = String(data[i][0]).trim();
    if (!username) continue;
    users[username] = {
      password: String(data[i][1]),
      name: String(data[i][2]),
      role: String(data[i][3]),
      regu: data[i][4] ? String(data[i][4]) : null
    };
  }
  return users;
}

function loadKelompok(ss) {
  var sheet = ss.getSheetByName(SHEETS.KELOMPOK);
  var data = sheet.getDataRange().getValues();
  var regus = {};
  for (var i = 1; i < data.length; i++) {
    var key = String(data[i][0]).trim();
    if (!key) continue;
    if (!regus[key]) {
      regus[key] = { name: String(data[i][1]), ketua: String(data[i][2]), members: [] };
    }
    var memberName = String(data[i][3]).trim();
    if (memberName) {
      regus[key].members.push({ name: memberName, role: String(data[i][4]), status: "none", time: null });
    }
  }
  return regus;
}

function loadRekap(ss) {
  var sheet = ss.getSheetByName(SHEETS.REKAP);
  var data = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    if (!String(data[i][0]).trim()) continue;
    rows.push({
      date: normalizeDateStr(data[i][0]),
      regu_key: String(data[i][1]),
      regu_name: String(data[i][2]),
      member_name: String(data[i][3]),
      status: String(data[i][4]),
      time: data[i][5] ? String(data[i][5]) : null
    });
  }
  return rows;
}

function loadKejadian(ss) {
  var sheet = ss.getSheetByName(SHEETS.KEJADIAN);
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    if (!String(data[i][0]).trim()) continue;
    list.push({
      id: Number(data[i][0]),
      reporter: String(data[i][1]),
      kategori: String(data[i][2]),
      lokasi: String(data[i][3]),
      deskripsi: String(data[i][4]),
      time: String(data[i][5])
    });
  }
  return list;
}

function loadTamu(ss) {
  var sheet = ss.getSheetByName(SHEETS.TAMU);
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    if (!String(data[i][0]).trim()) continue;
    list.push({
      id: Number(data[i][0]),
      plat: String(data[i][1]),
      jenis: String(data[i][2]),
      tujuan: String(data[i][3]),
      time: String(data[i][4])
    });
  }
  return list;
}

function loadPengaturan(ss) {
  var sheet = ss.getSheetByName(SHEETS.PENGATURAN);
  var data = sheet.getDataRange().getValues();
  var settings = {
    lat: -7.7820628,
    lng: 110.3030864,
    radius: 10,
    name: "Angkringan Pak Santo",
    mockMode: true,
    simulatedDistanceState: "inside"
  };
  for (var i = 1; i < data.length; i++) {
    var key = String(data[i][0]).trim();
    if (!key) continue;
    var val = data[i][1];
    if (key === "lat" || key === "lng" || key === "radius") {
      settings[key] = Number(val);
    } else if (key === "mockMode") {
      settings[key] = String(val) === "true";
    } else {
      settings[key] = String(val);
    }
  }
  return settings;
}

function loadRiwayat(ss) {
  var sheet = ss.getSheetByName(SHEETS.RIWAYAT);
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    if (!String(data[i][0]).trim()) continue;
    list.push({
      time: parseDateTime(String(data[i][0])),
      type: String(data[i][1]),
      icon: String(data[i][2]),
      color: String(data[i][3]),
      title: String(data[i][4]),
      desc: String(data[i][5])
    });
  }
  return list;
}

// Parse "yyyy-MM-dd HH:mm:ss" atau "yyyy-MM-dd" jadi Date (aman di Apps Script)
function parseDateTime(str) {
  var m = str.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?)?/);
  if (!m) return new Date(str);
  return new Date(
    Number(m[1]), Number(m[2]) - 1, Number(m[3]),
    m[4] ? Number(m[4]) : 0,
    m[5] ? Number(m[5]) : 0,
    m[6] ? Number(m[6]) : 0
  );
}

// Ubah nilai sel (bisa Date object atau string) jadi "yyyy-MM-dd"
function normalizeDateStr(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  var s = String(val).trim();
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[1] + "-" + m[2] + "-" + m[3];
  return s;
}

// Hapus semua baris presensi hari ini dari REKAP_PRESENSI (reset harian)
function resetTodayAttendance() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.REKAP);
  if (!sheet) {
    initializeDatabase();
    sheet = ss.getSheetByName(SHEETS.REKAP);
  }

  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  var data = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    if (!String(data[i][0]).trim()) continue;
    if (normalizeDateStr(data[i][0]) !== today) {
      rows.push(data[i]);
    }
  }

  var out = [["date", "regu_key", "regu_name", "member_name", "status", "time"]].concat(rows);
  clearSheet(sheet);
  if (out.length > 1) {
    sheet.getRange(1, 1, out.length, 6).setValues(out);
  }

  Logger.log("Presensi hari ini (" + today + ") direset.");
  return { status: "success", message: "Presensi hari ini direset!" };
}

// =========================================================================
// SAVE: tulis state aplikasi ke masing-masing sheet
// =========================================================================
function saveAllData(data) {
  var ss = getSpreadsheet();
  initializeDatabase();

  saveAkun(ss, data.usersDB || {});
  saveKelompok(ss, data.regus || {});
  saveRekap(ss, data.regus || {});
  saveRekapHistory(ss, data.attendanceHistory || []);
  saveKejadian(ss, data.incidentReports || []);
  saveTamu(ss, data.guestList || []);
  savePengaturan(ss, data.geofenceSettings || {});
  saveRiwayat(ss, data.appHistory || []);

  Logger.log("Data tersinkronisasi ke semua sheet.");
  return { status: "success", message: "Data tersinkronisasi ke semua sheet!" };
}

function saveAkun(ss, users) {
  var sheet = ss.getSheetByName(SHEETS.AKUN);
  clearSheet(sheet);
  var rows = [["username", "password", "name", "role", "regu"]];
  Object.keys(users).forEach(function (u) {
    var acc = users[u];
    rows.push([u, acc.password, acc.name, acc.role, acc.regu || ""]);
  });
  if (rows.length > 1) {
    sheet.getRange(1, 1, rows.length, 5).setValues(rows);
  }
}

function saveKelompok(ss, regus) {
  var sheet = ss.getSheetByName(SHEETS.KELOMPOK);
  clearSheet(sheet);
  var rows = [["regu_key", "regu_name", "ketua", "member_name", "member_role"]];
  Object.keys(regus).forEach(function (key) {
    var regu = regus[key];
    regu.members.forEach(function (m) {
      rows.push([key, regu.name, regu.ketua, m.name, m.role]);
    });
  });
  if (rows.length > 1) {
    sheet.getRange(1, 1, rows.length, 5).setValues(rows);
  }
}

function saveRekap(ss, regus) {
  var sheet = ss.getSheetByName(SHEETS.REKAP);
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  // Baca baris lama (termasuk injeksi manual user)
  var existing = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < existing.length; i++) {
    if (String(existing[i][0]).trim()) {
      // Normalisasi kolom tanggal: Date object -> "yyyy-MM-dd"
      existing[i][0] = normalizeDateStr(existing[i][0]);
      rows.push(existing[i]);
    }
  }

  // Upsert presensi hari ini dari state aplikasi (per tanggal + nama anggota)
  Object.keys(regus).forEach(function (key) {
    var regu = regus[key];
    regu.members.forEach(function (m) {
      if (!m.status || m.status === "none") return;
      var found = false;
      for (var r = 0; r < rows.length; r++) {
        if (String(rows[r][0]) === today && String(rows[r][1]) === key && String(rows[r][3]) === m.name) {
          rows[r][4] = m.status;
          rows[r][5] = m.time || "";
          found = true;
          break;
        }
      }
      if (!found) {
        rows.push([today, key, regu.name, m.name, m.status, m.time || ""]);
      }
    });
  });

  // Tulis ulang: header + semua baris (lama + baru hari ini)
  var out = [["date", "regu_key", "regu_name", "member_name", "status", "time"]].concat(rows);
  clearSheet(sheet);
  if (out.length > 1) {
    sheet.getRange(1, 1, out.length, 6).setValues(out);
  }
}

// Merge attendanceHistory dari app ke REKAP_PRESENSI (tambah baris yang belum ada)
function saveRekapHistory(ss, history) {
  if (!history || !history.length) return;
  var sheet = ss.getSheetByName(SHEETS.REKAP);

  var existing = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < existing.length; i++) {
    if (String(existing[i][0]).trim()) {
      existing[i][0] = normalizeDateStr(existing[i][0]);
      rows.push(existing[i]);
    }
  }

  history.forEach(function (h) {
    if (!h || !h.date || !h.reguKey || !h.memberName) return;
    var found = false;
    for (var r = 0; r < rows.length; r++) {
      if (String(rows[r][0]) === h.date && String(rows[r][1]) === h.reguKey && String(rows[r][3]) === h.memberName) {
        // Update status & time bila app punya data lebih baru
        rows[r][4] = h.status;
        rows[r][5] = h.time || "";
        found = true;
        break;
      }
    }
    if (!found) {
      rows.push([h.date, h.reguKey, h.reguName || "", h.memberName, h.status, h.time || ""]);
    }
  });

  var out = [["date", "regu_key", "regu_name", "member_name", "status", "time"]].concat(rows);
  clearSheet(sheet);
  if (out.length > 1) {
    sheet.getRange(1, 1, out.length, 6).setValues(out);
  }
}

function saveKejadian(ss, list) {
  var sheet = ss.getSheetByName(SHEETS.KEJADIAN);
  var existing = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < existing.length; i++) {
    if (String(existing[i][0]).trim()) {
      rows.push(existing[i]);
    }
  }

  // Upsert per id: baris injeksi manual (id tidak dikenal app) dipertahankan
  list.forEach(function (item) {
    var found = false;
    for (var r = 0; r < rows.length; r++) {
      if (Number(rows[r][0]) === Number(item.id)) {
        rows[r] = [item.id, item.reporter, item.kategori, item.lokasi, item.deskripsi, item.time];
        found = true;
        break;
      }
    }
    if (!found) {
      rows.push([item.id, item.reporter, item.kategori, item.lokasi, item.deskripsi, item.time]);
    }
  });

  var out = [["id", "reporter", "kategori", "lokasi", "deskripsi", "time"]].concat(rows);
  clearSheet(sheet);
  if (out.length > 1) {
    sheet.getRange(1, 1, out.length, 6).setValues(out);
  }
}

function saveTamu(ss, list) {
  var sheet = ss.getSheetByName(SHEETS.TAMU);
  var existing = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < existing.length; i++) {
    if (String(existing[i][0]).trim()) {
      rows.push(existing[i]);
    }
  }

  // Upsert per id: baris injeksi manual (id tidak dikenal app) dipertahankan
  list.forEach(function (item) {
    var found = false;
    for (var r = 0; r < rows.length; r++) {
      if (Number(rows[r][0]) === Number(item.id)) {
        rows[r] = [item.id, item.plat, item.jenis, item.tujuan, item.time];
        found = true;
        break;
      }
    }
    if (!found) {
      rows.push([item.id, item.plat, item.jenis, item.tujuan, item.time]);
    }
  });

  var out = [["id", "plat", "jenis", "tujuan", "time"]].concat(rows);
  clearSheet(sheet);
  if (out.length > 1) {
    sheet.getRange(1, 1, out.length, 5).setValues(out);
  }
}

function savePengaturan(ss, settings) {
  var sheet = ss.getSheetByName(SHEETS.PENGATURAN);
  clearSheet(sheet);
  var rows = [["key", "value"]];
  rows.push(["lat", settings.lat]);
  rows.push(["lng", settings.lng]);
  rows.push(["radius", settings.radius]);
  rows.push(["name", settings.name]);
  rows.push(["mockMode", settings.mockMode ? "true" : "false"]);
  rows.push(["simulatedDistanceState", settings.simulatedDistanceState]);
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
}

function saveRiwayat(ss, list) {
  var sheet = ss.getSheetByName(SHEETS.RIWAYAT);
  clearSheet(sheet);
  var rows = [["time", "type", "icon", "color", "title", "desc"]];
  list.forEach(function (item) {
    var t = item.time;
    if (t instanceof Date) {
      t = Utilities.formatDate(t, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    }
    rows.push([t, item.type, item.icon, item.color, item.title, item.desc]);
  });
  if (rows.length > 1) {
    sheet.getRange(1, 1, rows.length, 6).setValues(rows);
  }
}
