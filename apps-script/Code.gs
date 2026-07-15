// Apps Script Web App backing the "협력업체 업무가이드 작성 도구" app's
// optional Google Sheet write-through. Appends a row to the sheet's first
// (leftmost) tab whenever the frontend saves a new local guide item.
//
// Setup: paste this file's contents into the bound Apps Script project's
// Code.gs, change SECRET_TOKEN below, save, then Deploy > New deployment >
// Web app (Execute as: Me, Who has access: Anyone). Copy the resulting
// .../exec URL and give it plus SECRET_TOKEN to the frontend's
// SHEET_WEBAPP_URL / SHEET_WEBAPP_TOKEN constants in index.html.
//
// To push future edits: after Claude updates this file, copy its contents
// into the Apps Script editor, save, then Deploy > Manage deployments >
// edit (pencil) icon > New version > Deploy. The web app URL stays the same.

var SECRET_TOKEN = "원하는_비밀문자열로_변경하세요";

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    var data = JSON.parse(e.postData.contents);

    if (data.token !== SECRET_TOKEN) {
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
