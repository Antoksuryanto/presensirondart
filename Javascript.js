function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({status: "success", message: "Koneksi API Siskamling Aktif"}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
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

function initializeDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("SYSTEM_DB");
  if (!sheet) {
    sheet = ss.insertSheet("SYSTEM_DB");
    sheet.appendRow(["JSON_DATA"]);
    sheet.getRange(1, 1).setFontWeight("bold");
  }
}

function loadAllData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("SYSTEM_DB");
  var lastRow = sheet.getLastRow();
  
  if (lastRow < 2) {
    return { status: "empty", message: "Belum ada data tersimpan di awan." };
  }
  
  var jsonString = sheet.getRange(2, 1).getValue();
  try {
    var parsedData = JSON.parse(jsonString);
    return { status: "success", data: parsedData };
  } catch (e) {
    return { status: "error", message: "Gagal memproses data JSON." };
  }
}

function saveAllData(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("SYSTEM_DB");
  var jsonString = JSON.stringify(data);
  
  sheet.getRange(2, 1).setValue(jsonString);
  return { status: "success", message: "Data tersinkronisasi!" };
}