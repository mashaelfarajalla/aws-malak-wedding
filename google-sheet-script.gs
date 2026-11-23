/**
 * سكربت استقبال رسائل التهاني — أوس وملك
 * يُلصق في: Google Sheets ← Extensions (الإضافات) ← Apps Script
 * ثم Deploy ← New deployment ← Web app
 *   Execute as: Me
 *   Who has access: Anyone
 * وانسخ رابط الـ Web app (ينتهي بـ /exec) وضعه في SHEET_URL داخل index.html
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["التاريخ والوقت", "الاسم", "الرسالة"]);
      sheet.setRightToLeft(true);
      sheet.getRange(1, 1, 1, 3).setFontWeight("bold");
    }
    var p = (e && e.parameter) || {};
    var clean = function (v, max) {
      v = String(v || "").trim().slice(0, max);
      // منع تنفيذ الصيغ داخل الجدول
      return /^[=+\-@]/.test(v) ? "'" + v : v;
    };
    var msg = clean(p.message, 2000);
    if (!msg) return ContentService.createTextOutput("empty");
    sheet.appendRow([new Date(), clean(p.name, 80), msg]);
    return ContentService.createTextOutput("ok");
  } finally {
    lock.releaseLock();
  }
}

/** يرجّع الرسائل للصفحة (الأحدث أولاً) */
function doGet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var last = sheet.getLastRow();
  var out = [];
  if (last > 1) {
    var rows = sheet.getRange(2, 2, last - 1, 2).getDisplayValues(); // الاسم، الرسالة
    for (var i = rows.length - 1; i >= 0 && out.length < 200; i--) {
      var name = String(rows[i][0]).replace(/^'/, "");
      var msg = String(rows[i][1]).replace(/^'/, "");
      if (msg) out.push({ name: name, message: msg });
    }
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}
