// Apps Script Web App backing the "협력업체 업무가이드 작성 도구" app.
// Bound to the master Google Sheet. Because Apps Script always runs with
// the owner's own permissions ("Execute as: Me"), this lets the actual
// spreadsheet stay fully private (sharing: Restricted / owner only) while
// still serving its checklist rows out to the public web app:
//
//   doGet  -> read-only: returns the sheet's rows as JSON (no auth needed;
//             this is the same data the app's users are meant to see).
//   doPost -> write: appends a row, gated by a SECRET_TOKEN script
//             property so random callers of this URL can't spam the sheet.
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

      var rawDays = row[2];
      var days;
      if (rawDays === "" || rawDays === null || rawDays === undefined) {
        days = "";
      } else if (typeof rawDays === "number") {
        days = rawDays;
      } else {
        days = String(rawDays).trim();
      }

      items.push({ category: category, content: content, days: days });
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
    sheet.appendRow([
      data.category || "",
      data.content,
      data.days === null || data.days === undefined ? "" : data.days
    ]);

    return jsonOutput({ ok: true });
  } catch (err) {
    return jsonOutput({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
