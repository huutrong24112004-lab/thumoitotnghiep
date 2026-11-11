/**
 * GOOGLE SHEETS RSVP RECEIVER — THƯ MỜI TỐT NGHIỆP HỮU TRỌNG
 * Gắn script này với Google Sheet nhận phản hồi.
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    const headers = ["Thời gian nhận", "Họ và tên", "Xác nhận tham dự", "Lời nhắn", "Sự kiện"];
    if (sheet.getLastRow() === 0) { sheet.appendRow(headers); sheet.setFrozenRows(1); }
    const p = (e && e.parameter) ? e.parameter : {};
    const safe = value => { let s = String(value || "").trim(); if (/^[=+\\-@]/.test(s)) s = "'" + s; return s; };
    sheet.appendRow([new Date(), safe(p.name), safe(p.attendance), safe(p.message), safe(p.event || "Lễ tốt nghiệp Nguyễn Hữu Trọng - IUH 2026")]);
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } finally { lock.releaseLock(); }
}
