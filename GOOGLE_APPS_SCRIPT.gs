/**
 * PHẢN HỒI KHÁCH MỜI — THƯ MỜI TỐT NGHIỆP HỮU TRỌNG
 * Spreadsheet ID: 1Xkex0TlNz480bcOwYq5qqHYqDY_Rtv-ShoxRfZfJ8iE
 *
 * Bảng "Danh sách khách mời", hàng tiêu đề 4:
 * G = Phản hồi; H = Có tham dự?; K = Ghi chú.
 *
 * Người có tên trùng trong danh sách: cập nhật dòng hiện có.
 * Người chưa có trong danh sách: thêm một dòng mới ở cuối.
 */

const SPREADSHEET_ID = '1Xkex0TlNz480bcOwYq5qqHYqDY_Rtv-ShoxRfZfJ8iE';
const SHEET_NAME = 'Danh sách khách mời';
const HEADER_ROW = 4;

function doGet() {
  return ContentService
    .createTextOutput('RSVP endpoint is ready.')
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('Không tìm thấy sheet: ' + SHEET_NAME);

    const p = (e && e.parameter) ? e.parameter : {};
    const guestName = clean_(p.name);
    const attendance = clean_(p.attendance);
    const message = clean_(p.message);
    if (!guestName) throw new Error('Thiếu họ và tên khách mời.');

    const attendanceMap = {
      'Chắc chắn đến': { response: 'Sẽ tham dự', presence: 'Có' },
      'Sẽ cố gắng đến': { response: 'Đang cân nhắc', presence: 'Chưa rõ' },
      'Rất tiếc, không đến được': { response: 'Không tham dự', presence: 'Không' },
      'Không đến được': { response: 'Không tham dự', presence: 'Không' }
    };
    const mapped = attendanceMap[attendance] || {
      response: 'Đang cân nhắc',
      presence: 'Chưa rõ'
    };

    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(HEADER_ROW, 1, 1, lastCol)
      .getDisplayValues()[0].map(h => String(h).trim());

    const col = {};
    headers.forEach((h, i) => { col[h] = i + 1; });
    const nameCol = col['Họ và tên'] || 2;
    const responseCol = col['Phản hồi'] || 7;
    const presenceCol = col['Có tham dự?'] || 8;
    const noteCol = col['Ghi chú'] || 11;
    const sttCol = col['STT'] || 1;

    const lastRow = sheet.getLastRow();
    const dataStart = HEADER_ROW + 1;
    const rowCount = Math.max(0, lastRow - HEADER_ROW);
    const names = rowCount
      ? sheet.getRange(dataStart, nameCol, rowCount, 1).getDisplayValues().flat()
      : [];

    const targetKey = normalize_(guestName);
    let targetRow = -1;
    for (let i = 0; i < names.length; i++) {
      if (normalize_(names[i]) === targetKey) {
        targetRow = dataStart + i;
        break;
      }
    }

    if (targetRow === -1) {
      targetRow = Math.max(lastRow + 1, dataStart);
      const nextNo = Math.max(0, targetRow - dataStart + 1);
      sheet.getRange(targetRow, sttCol).setValue(nextNo);
      sheet.getRange(targetRow, nameCol).setValue(guestName);
    }

    const timestamp = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone() || 'Asia/Ho_Chi_Minh',
      'dd/MM/yyyy HH:mm:ss'
    );
    const noteText = message
      ? 'Lời chúc: ' + message + ' | Gửi lúc: ' + timestamp
      : 'Đã gửi phản hồi qua website lúc: ' + timestamp;

    sheet.getRange(targetRow, responseCol).setValue(mapped.response);
    sheet.getRange(targetRow, presenceCol).setValue(mapped.presence);

    const oldNote = String(sheet.getRange(targetRow, noteCol).getDisplayValue() || '').trim();
    const combinedNote = oldNote
      ? oldNote + '\n' + noteText
      : noteText;
    sheet.getRange(targetRow, noteCol).setValue(safeForSheet_(combinedNote));

    return ContentService
      .createTextOutput(JSON.stringify({
        ok: true,
        row: targetRow,
        guest: guestName,
        response: mapped.response,
        attendance: mapped.presence
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function clean_(value) {
  return String(value || '').trim().slice(0, 2000);
}

function normalize_(value) {
  return String(value || '')
    .trim()
    .toLocaleLowerCase('vi')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/\s+/g, ' ');
}

function safeForSheet_(value) {
  const s = String(value || '');
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
