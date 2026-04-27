/**
 * SISTEM MONITORING KENDARAAN HARIAN - BACKEND
 */

function doGet(e) {
  return HtmlService.createHtmlOutput("FleetMonitoring API is Active.");
}

function doPost(e) {
  const params = JSON.parse(e.postData.contents);
  const action = e.parameter.action;
  let result = { success: false };

  try {
    switch (action) {
      case "getPublicData":
        result = { 
          vehicles: getData("kendaraan"),
          bookings: getData("booking")
        };
        break;
      case "getAdminData":
        result = { 
          vehicles: getData("kendaraan"),
          bookings: getData("booking")
        };
        break;
      case "saveBooking":
        result = saveBooking(params);
        break;
      case "updateDetailBooking":
        result = updateDetailBooking(params);
        break;
      case "setup":
        result = setupDatabase();
        break;
    }
  } catch (err) {
    result.message = err.toString();
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function getDb() { return SpreadsheetApp.getActiveSpreadsheet(); }

function getData(sheetName) {
  const sheet = getDb().getSheetByName(sheetName);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values.shift();
  return values.map(row => {
    let obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function saveBooking(p) {
  const sheet = getDb().getSheetByName("booking");
  const headers = ["id", "timestamp", "nama", "divisi", "id_kendaraan", "tgl_mulai", "tujuan", "status_approval", "driver", "keterangan"];
  const newRow = headers.map(h => {
    if(h === "id") return "B" + new Date().getTime();
    if(h === "timestamp") return new Date();
    if(h === "status_approval") return "Pending";
    return p[h] || "";
  });
  sheet.appendRow(newRow);
  return { success: true };
}

function updateDetailBooking(p) {
  const sheet = getDb().getSheetByName("booking");
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] == p.id) {
      sheet.getRange(i + 1, headers.indexOf("driver") + 1).setValue(p.driver);
      sheet.getRange(i + 1, headers.indexOf("keterangan") + 1).setValue(p.keterangan);
      sheet.getRange(i + 1, headers.indexOf("status_approval") + 1).setValue(p.status);
      return { success: true };
    }
  }
  return { success: false };
}

function setupDatabase() {
  const ss = getDb();
  const tables = [
    { name: "kendaraan", headers: ["id", "nama", "plat", "kategori", "status"] },
    { name: "booking", headers: ["id", "timestamp", "nama", "divisi", "id_kendaraan", "tgl_mulai", "tujuan", "status_approval", "driver", "keterangan"] },
    { name: "users", headers: ["username", "password", "role"] }
  ];

  tables.forEach(t => {
    let sheet = ss.getSheetByName(t.name);
    if (!sheet) {
      sheet = ss.insertSheet(t.name);
      sheet.appendRow(t.headers);
    }
  });

  return { success: true };
}