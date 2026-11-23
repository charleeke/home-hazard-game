/**
 * Home Fall Risk Assessment Game — Google Apps Script back end
 *
 * Security change (Nov 2026):
 *  - The instructor PIN is no longer in the client code. It is stored in Script Properties
 *    (Project Settings > Script Properties > key EDITOR_PIN) and checked here, on the server.
 *  - Every write action (save / update / delete) must send the correct PIN; otherwise it is rejected.
 *  - After 10 wrong PIN attempts within 15 minutes, all PIN checks are locked for 15 minutes.
 *  This file contains no credentials and can be published in the repository.
 *
 * Note: doGet still returns hazard locations to the browser (client-side hit-testing),
 * as described in Section 3.2.4 of the manuscript.
 */
const SHEET_NAME = "Data";
const FOLDER_NAME = "PT_Safety_Images";
const MAX_FAILS = 10;          // wrong attempts allowed per window
const LOCK_SECONDS = 15 * 60;  // window / lockout length

function getTargetFolder() {
  var ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var ssFile = DriveApp.getFileById(ssId);
  var parents = ssFile.getParents();
  var parentFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  var folders = parentFolder.getFoldersByName(FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  else return parentFolder.createFolder(FOLDER_NAME);
}

function setupSystem() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, 5).setValues([["ID", "RoomType", "ImageURL", "HazardsJSON", "Timestamp"]]);
  }
  getTargetFolder();
  return sheet;
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Returns true only if the PIN matches EDITOR_PIN; counts failures and applies a lockout. */
function checkPin(pin) {
  var expected = PropertiesService.getScriptProperties().getProperty("EDITOR_PIN");
  if (!expected) throw new Error("EDITOR_PIN has not been set in Script Properties");
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get("pinFails") || 0);
  if (fails >= MAX_FAILS) throw new Error("Too many attempts. Please try again later.");
  if (typeof pin === "string" && pin === expected) return true;
  cache.put("pinFails", String(fails + 1), LOCK_SECONDS);
  return false;
}

function doGet(e) {
  try {
    var sheet = setupSystem();
    var data = sheet.getDataRange().getValues();
    var result = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        try {
          result.push({
            id: data[i][0].toString(),
            roomType: data[i][1],
            imageUrl: data[i][2],
            hazards: JSON.parse(data[i][3] || "[]")
          });
        } catch (err) {}
      }
    }
    return jsonOut(result);
  } catch (error) {
    return jsonOut({ error: error.message });
  }
}

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);

    // Login check from the authoring page
    if (payload.action === "verify") {
      return jsonOut({ ok: checkPin(payload.pin) });
    }

    // All other actions change data: the PIN is required
    if (!checkPin(payload.pin)) return jsonOut({ error: "Unauthorized" });

    var sheet = setupSystem();

    if (payload.action === "delete") {
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (data[i][0].toString() === payload.id) {
          sheet.deleteRow(i + 1);
          SpreadsheetApp.flush();
          break;
        }
      }
      return jsonOut({ status: "deleted" });
    }
    else if (payload.action === "save" || payload.action === "update") {
      var imageUrl = payload.imageUrl || "";

      if (payload.imageBase64) {
        var folder = getTargetFolder();
        var base64Data = payload.imageBase64.split(",")[1];
        var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), payload.mimeType, payload.fileName);
        var file = folder.createFile(blob);
        try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (sharingError) {}
        imageUrl = "https://drive.google.com/uc?id=" + file.getId();
      }

      var hazardsStr = JSON.stringify(payload.hazards);
      var timeStr = new Date().toLocaleString("th-TH");

      if (payload.action === "update" && payload.id) {
        var data2 = sheet.getDataRange().getValues();
        for (var j = 1; j < data2.length; j++) {
          if (data2[j][0].toString() === payload.id) {
            sheet.getRange(j + 1, 2, 1, 4).setValues([[payload.roomType, imageUrl, hazardsStr, timeStr]]);
            SpreadsheetApp.flush();
            return jsonOut({ status: "updated", imageUrl: imageUrl });
          }
        }
        return jsonOut({ error: "Record not found" });
      } else {
        var newId = new Date().getTime().toString();
        var columnA = sheet.getRange("A:A").getValues();
        var lastRow = 1;
        for (var k = 0; k < columnA.length; k++) { if (columnA[k][0] !== "") lastRow = k + 1; }
        sheet.getRange(lastRow + 1, 1, 1, 5).setValues([[newId, payload.roomType, imageUrl, hazardsStr, timeStr]]);
        SpreadsheetApp.flush();
        return jsonOut({ status: "saved", imageUrl: imageUrl });
      }
    }
    return jsonOut({ error: "Unknown action" });
  } catch (error) {
    return jsonOut({ error: "Error: " + error.message });
  }
}
