// Apps Script Web App backing the "협력업체 업무가이드 작성 도구" app.
// Bound to the master Google Sheet. Because Apps Script always runs with
// the owner's own permissions ("Execute as: Me"), this lets the actual
// spreadsheet stay fully private (sharing: Restricted / owner only) while
// still serving its checklist rows out to the public web app:
//
//   doGet  -> read-only: returns the sheet's rows as JSON, each tagged with
//             its actual sheet row number so edits can target it precisely.
//             No auth needed; this is the same data the app's users are
//             meant to see.
//   doPost -> write: gated by a SECRET_TOKEN script property so random
//             callers of this URL can't spam the sheet.
//             - default (no "action"): appends a new row.
//             - action: "update": overwrites an existing row (by its row
//               number) with new category/content/days/remarks.
//
// Sheet columns: A 구분, B 내용, C 표준조치요구일, D 비고.
//
// Setup:
// 1. Paste this file into the bound Apps Script project's Code.gs, save.
// 2. Project Settings > Script Properties > Add script property:
//    key "SECRET_TOKEN", value = any string you pick. The real value never
//    appears in this source file, which is tracked in the public repo.
// 3. Deploy > New deployment > Web app (Execute as: Me, Who has access:
//    Anyone). Copy the resulting .../exec URL.
// 4. Set the Google Sheet's own sharing to "제한됨" (Restricted, only you)
//    - the web app above still works regardless, since it runs as you.
// 5. Give the .../exec URL and the SECRET_TOKEN value to the frontend's
//    SHEET_WEBAPP_URL / SHEET_WEBAPP_TOKEN constants in index.html.
//
// To push future edits: after Claude updates this file, copy its contents
// into the Apps Script editor, save, then Deploy > Manage deployments >
// edit (pencil) icon > New version > Deploy. The web app URL stays the same.

function doGet(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var values = sheet.getDataRange().getValues();
    var items = [];

    for (var i = 1; i < values.length; i++) { // skip header row
      var row = values[i];
      var category = (row[0] || "").toString().trim();
      var content = (row[1] || "").toString().trim();
      if (!category && !content) continue;

      items.push({
        row: i + 1,
        category: category,
        content: content,
        days: normalizeDays(row[2]),
        remarks: (row[3] || "").toString().trim()
      });
    }

    return jsonOutput({ ok: true, items: items });
  } catch (err) {
    return jsonOutput({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    var secretToken = PropertiesService.getScriptProperties().getProperty("SECRET_TOKEN");
    var data = JSON.parse(e.postData.contents);

    if (!secretToken || data.token !== secretToken) {
      return jsonOutput({ ok: false, error: "unauthorized" });
    }
    if (!data.content) {
      return jsonOutput({ ok: false, error: "content required" });
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var daysValue = data.days === null || data.days === undefined ? "" : data.days;
    var remarksValue = data.remarks || "";

    if (data.action === "update") {
      var rowNum = parseInt(data.row, 10);
      var lastRow = sheet.getLastRow();
      if (!rowNum || rowNum < 2 || rowNum > lastRow) {
        return jsonOutput({ ok: false, error: "invalid row" });
      }
      sheet.getRange(rowNum, 1, 1, 4).setValues([[data.category || "", data.content, daysValue, remarksValue]]);
      return jsonOutput({ ok: true, row: rowNum });
    }

    sheet.appendRow([data.category || "", data.content, daysValue, remarksValue]);
    return jsonOutput({ ok: true, row: sheet.getLastRow() });
  } catch (err) {
    return jsonOutput({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function normalizeDays(rawDays) {
  if (rawDays === "" || rawDays === null || rawDays === undefined) return "";
  if (typeof rawDays === "number") return rawDays;
  return String(rawDays).trim();
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
