// =========================================================================
// ID Spreadsheet Google Anda.
// Ganti dengan ID spreadsheet tempat script ini berada.
// Cara dapatkan ID: buka spreadsheet, URL-nya seperti:
//   https://docs.google.com/spreadsheets/d/ID_SPREADSHEET_ANDA/edit
// ID-nya adalah bagian antara /d/ dan /edit
// =========================================================================
var SPREADSHEET_ID = "1G4MERfP-il--aPfqSLgndx3TWt9_0r1Kab-sKKectnDM7SseG4etK0_k";  // <-- GANTI INI

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

function initializeDatabase() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("SYSTEM_DB");
  if (!sheet) {
    sheet = ss.insertSheet("SYSTEM_DB");
    sheet.appendRow(["JSON_DATA"]);
    sheet.getRange(1, 1).setFontWeight("bold");
  }
}

function loadAllData() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("SYSTEM_DB");

  if (!sheet) {
    initializeDatabase();
    sheet = ss.getSheetByName("SYSTEM_DB");
  }

  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return { status: "empty", message: "Belum ada data tersimpan di awan." };
  }

  var jsonString = sheet.getRange(2, 1).getValue();

  if (!jsonString || jsonString.toString().trim() === "") {
    return { status: "empty", message: "Sel data masih kosong." };
  }

  try {
    var parsedData = JSON.parse(jsonString);
    return { status: "success", data: parsedData };
  } catch (e) {
    return { status: "error", message: "Gagal memproses data JSON: " + e.toString() };
  }
}

function saveAllData(data) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("SYSTEM_DB");

  if (!sheet) {
    initializeDatabase();
    sheet = ss.getSheetByName("SYSTEM_DB");
  }

  var jsonString = JSON.stringify(data);

  // Pastikan baris 2 ada sebelum menulis
  if (sheet.getLastRow() < 2) {
    sheet.insertRowBefore(2);
  }

  sheet.getRange(2, 1).setValue(jsonString);

  // Optional: log ke Apps Script log untuk debug
  Logger.log("Data saved to cloud, length: " + jsonString.length + " chars");

  return { status: "success", message: "Data tersinkronisasi!" };
}
