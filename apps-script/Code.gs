/**
 * LTS - Quản lý Đi trễ / Xin nghỉ buổi tập
 * Backend API bằng Google Apps Script
 *
 * CÁCH DÙNG:
 * 1. Chạy hàm setup() một lần duy nhất để tạo & format các sheet.
 * 2. (Tuỳ chọn) Chạy configNotify() để nhập email/Telegram nhận thông báo.
 * 3. Deploy > New deployment > Web app (Execute as: Me, Access: Anyone).
 *    Mỗi lần sửa code -> Manage deployments > Edit > Version: New version.
 */

const CONFIG = {
  DATA_SHEET: 'Data',
  MEMBERS_SHEET: 'Members',
  SUMMARY_SHEET: 'Tổng kết',
  PRACTICE_START_HOUR: 18,
  PRACTICE_START_MIN: 30,
  TIMEZONE: 'Asia/Ho_Chi_Minh',
  STATUSES: ['Chờ duyệt', 'Đã duyệt', 'Từ chối'],
  TYPES: ['Đi trễ', 'Nghỉ'],
  COLOR_PRIMARY: '#1d4ed8',
  COLOR_PRIMARY_LIGHT: '#dbeafe',
  COLOR_HEADER_TEXT: '#ffffff',
};

// Vị trí cột trong sheet Data (1-based)
const COL = { TS: 1, NAME: 2, TYPE: 3, DATE: 4, ARRIVAL: 5, LATE: 6, REASON: 7, STATUS: 8 };

/* ============================================================
 * API ENDPOINTS
 * ============================================================ */

function doGet(e) {
  const p = (e && e.parameter) || {};
  const action = p.action || 'members';

  try {
    if (action === 'members') {
      return jsonResponse({ status: 'success', members: getMembers_() });
    }
    // Cấu hình công khai cho mọi user (hình nền...)
    if (action === 'config') {
      return jsonResponse({ status: 'success', bgUrl: getBgUrl_() });
    }
    // Đăng nhập quản lý: kiểm tra PIN, trả về link Sheet nếu đúng
    if (action === 'login') {
      if (!checkPin_(p.pin)) return jsonResponse({ status: 'error', message: 'Sai mã PIN.' });
      return jsonResponse({ status: 'success', sheetUrl: getSheetUrl_() });
    }
    // Các endpoint dưới đây chỉ dành cho quản lý -> bắt buộc PIN đúng
    if (action === 'stats') {
      if (!checkPin_(p.pin)) return jsonResponse({ status: 'error', message: 'Sai mã PIN.' });
      const now = new Date();
      const month = Number(p.month) || (now.getMonth() + 1);
      const year = Number(p.year) || now.getFullYear();
      return jsonResponse({ status: 'success', month: month, year: year, stats: getStats_(month, year) });
    }
    if (action === 'history') {
      if (!checkPin_(p.pin)) return jsonResponse({ status: 'error', message: 'Sai mã PIN.' });
      const name = String(p.name || '').trim();
      if (!name) return jsonResponse({ status: 'error', message: 'Thiếu tên.' });
      return jsonResponse({ status: 'success', history: getHistory_(name, Number(p.limit) || 20) });
    }
    // Danh sách yêu cầu để duyệt (mặc định: đang chờ duyệt)
    if (action === 'pending') {
      if (!checkPin_(p.pin)) return jsonResponse({ status: 'error', message: 'Sai mã PIN.' });
      return jsonResponse({ status: 'success', items: getRequests_(p.filter || 'Chờ duyệt') });
    }
    return jsonResponse({ status: 'error', message: 'Unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ status: 'error', message: 'Lỗi server: ' + err.message });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    const body = JSON.parse(e.postData.contents);

    // --- Admin: đổi hình nền (cần PIN) ---
    if (body.action === 'setBackground') {
      if (!checkPin_(body.pin)) return jsonResponse({ status: 'error', message: 'Sai mã PIN.' });
      const url = String(body.bgUrl || '').trim();
      const props = PropertiesService.getScriptProperties();
      if (url) {
        if (!/^https?:\/\//i.test(url))
          return jsonResponse({ status: 'error', message: 'URL hình không hợp lệ (phải bắt đầu http/https).' });
        props.setProperty('BG_URL', url);
      } else {
        props.deleteProperty('BG_URL');
      }
      return jsonResponse({ status: 'success', message: url ? 'Đã đổi hình nền.' : 'Đã xoá hình nền.', bgUrl: url });
    }

    // --- Admin: duyệt / từ chối yêu cầu (cần PIN) ---
    if (body.action === 'updateStatus') {
      if (!checkPin_(body.pin)) return jsonResponse({ status: 'error', message: 'Sai mã PIN.' });
      const row = Number(body.row);
      const status = String(body.status || '').trim();
      if (CONFIG.STATUSES.indexOf(status) === -1)
        return jsonResponse({ status: 'error', message: 'Trạng thái không hợp lệ.' });
      const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.DATA_SHEET);
      if (!(row >= 2 && row <= sheet.getLastRow()))
        return jsonResponse({ status: 'error', message: 'Dòng không hợp lệ.' });
      sheet.getRange(row, COL.STATUS).setValue(status);
      try { refreshSummary_(); } catch (ignore) {}
      return jsonResponse({ status: 'success', message: 'Đã cập nhật: ' + status });
    }

    const name = String(body.name || '').trim();
    const type = String(body.type || '').trim();
    const dateStr = String(body.date || '').trim();
    const reason = String(body.reason || '').trim();
    const arrivalTime = String(body.arrivalTime || '').trim();

    if (!name) return jsonResponse({ status: 'error', message: 'Thiếu tên thành viên.' });
    if (CONFIG.TYPES.indexOf(type) === -1)
      return jsonResponse({ status: 'error', message: 'Loại yêu cầu không hợp lệ.' });
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr))
      return jsonResponse({ status: 'error', message: 'Ngày không hợp lệ (YYYY-MM-DD).' });
    if (!reason) return jsonResponse({ status: 'error', message: 'Thiếu lý do.' });

    // Tính phút trễ
    let lateMinutes = '';
    let arrivalDisplay = '';
    if (type === 'Đi trễ') {
      if (!/^\d{1,2}:\d{2}$/.test(arrivalTime))
        return jsonResponse({ status: 'error', message: 'Thiếu giờ đến dự kiến (HH:mm).' });
      const parts = arrivalTime.split(':');
      const arrivalMinOfDay = Number(parts[0]) * 60 + Number(parts[1]);
      const startMinOfDay = CONFIG.PRACTICE_START_HOUR * 60 + CONFIG.PRACTICE_START_MIN;
      lateMinutes = Math.max(0, arrivalMinOfDay - startMinOfDay);
      arrivalDisplay = arrivalTime;
    }

    const d = dateStr.split('-');
    const appliedDate = new Date(Number(d[0]), Number(d[1]) - 1, Number(d[2]));

    const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.DATA_SHEET);

    // Chặn gửi trùng: cùng tên + cùng ngày + cùng loại, chưa bị từ chối
    if (isDuplicate_(sheet, name, type, appliedDate)) {
      return jsonResponse({
        status: 'error',
        message: 'Bạn đã gửi yêu cầu "' + type + '" cho ngày này rồi.',
      });
    }

    sheet.appendRow([
      new Date(), name, type, appliedDate,
      arrivalDisplay, lateMinutes, reason, CONFIG.STATUSES[0],
    ]);

    try { refreshSummary_(); } catch (ignore) {}

    // Thông báo (không chặn phản hồi nếu lỗi)
    try { notify_({ name: name, type: type, dateStr: dateStr, arrivalTime: arrivalDisplay, lateMinutes: lateMinutes, reason: reason }); }
    catch (ignore) {}

    return jsonResponse({ status: 'success', message: 'Đã ghi nhận yêu cầu.' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: 'Lỗi server: ' + err.message });
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================
 * DATA HELPERS
 * ============================================================ */

function getMembers_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.MEMBERS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues()
    .map(function (r) { return String(r[0]).trim(); })
    .filter(function (v) { return v !== ''; });
}

function isDuplicate_(sheet, name, type, appliedDate) {
  if (sheet.getLastRow() < 2) return false;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, COL.STATUS).getValues();
  const target = ymd_(appliedDate);
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (String(r[COL.NAME - 1]).trim() === name &&
        String(r[COL.TYPE - 1]).trim() === type &&
        r[COL.DATE - 1] instanceof Date && ymd_(r[COL.DATE - 1]) === target &&
        String(r[COL.STATUS - 1]).trim() !== 'Từ chối') {
      return true;
    }
  }
  return false;
}

function getStats_(month, year) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.DATA_SHEET);
  const members = getMembers_();
  const map = {};
  members.forEach(function (m) { map[m] = { name: m, late: 0, lateMinutes: 0, off: 0 }; });

  if (sheet.getLastRow() >= 2) {
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, COL.STATUS).getValues();
    rows.forEach(function (r) {
      const name = String(r[COL.NAME - 1]).trim();
      const type = String(r[COL.TYPE - 1]).trim();
      const date = r[COL.DATE - 1];
      const status = String(r[COL.STATUS - 1]).trim();
      if (!name || status === 'Từ chối') return;
      if (!(date instanceof Date)) return;
      if (date.getMonth() + 1 !== month || date.getFullYear() !== year) return;
      if (!map[name]) map[name] = { name: name, late: 0, lateMinutes: 0, off: 0 };
      if (type === 'Đi trễ') {
        map[name].late += 1;
        map[name].lateMinutes += Number(r[COL.LATE - 1]) || 0;
      } else if (type === 'Nghỉ') {
        map[name].off += 1;
      }
    });
  }

  return Object.keys(map).map(function (k) { return map[k]; })
    .sort(function (a, b) { return (b.late + b.off) - (a.late + a.off); });
}

function getRequests_(statusFilter) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.DATA_SHEET);
  if (sheet.getLastRow() < 2) return [];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, COL.STATUS).getValues();
  const wantAll = !statusFilter || statusFilter === 'all';
  const out = [];
  for (let i = rows.length - 1; i >= 0; i--) {
    const r = rows[i];
    const status = String(r[COL.STATUS - 1]).trim();
    if (!wantAll && status !== statusFilter) continue;
    out.push({
      row: i + 2, // số dòng thực trong sheet
      timestamp: r[COL.TS - 1] instanceof Date ? formatDT_(r[COL.TS - 1]) : '',
      name: String(r[COL.NAME - 1]).trim(),
      type: String(r[COL.TYPE - 1]).trim(),
      date: r[COL.DATE - 1] instanceof Date ? formatD_(r[COL.DATE - 1]) : '',
      arrival: formatArrival_(r[COL.ARRIVAL - 1]),
      lateMinutes: r[COL.LATE - 1] === '' ? '' : Number(r[COL.LATE - 1]),
      reason: String(r[COL.REASON - 1] || ''),
      status: status,
    });
  }
  return out;
}

function getHistory_(name, limit) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.DATA_SHEET);
  if (sheet.getLastRow() < 2) return [];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, COL.STATUS).getValues();
  const out = [];
  for (let i = rows.length - 1; i >= 0; i--) {
    const r = rows[i];
    if (String(r[COL.NAME - 1]).trim() !== name) continue;
    out.push({
      timestamp: r[COL.TS - 1] instanceof Date ? formatDT_(r[COL.TS - 1]) : '',
      type: String(r[COL.TYPE - 1]).trim(),
      date: r[COL.DATE - 1] instanceof Date ? formatD_(r[COL.DATE - 1]) : '',
      arrival: formatArrival_(r[COL.ARRIVAL - 1]),
      lateMinutes: r[COL.LATE - 1] === '' ? '' : Number(r[COL.LATE - 1]),
      reason: String(r[COL.REASON - 1] || ''),
      status: String(r[COL.STATUS - 1]).trim(),
    });
    if (out.length >= limit) break;
  }
  return out;
}

/* ============================================================
 * NOTIFY (email + Telegram) — tuỳ chọn
 * ============================================================ */

function notify_(req) {
  const props = PropertiesService.getScriptProperties();
  const email = props.getProperty('MANAGER_EMAIL');
  const tgToken = props.getProperty('TELEGRAM_BOT_TOKEN');
  const tgChat = props.getProperty('TELEGRAM_CHAT_ID');

  const lateStr = req.type === 'Đi trễ'
    ? ('\n• Giờ đến: ' + req.arrivalTime + ' (trễ ' + req.lateMinutes + ' phút)')
    : '';
  const text =
    '🔔 Yêu cầu mới — LTS\n' +
    '• Tên: ' + req.name + '\n' +
    '• Loại: ' + req.type + '\n' +
    '• Ngày: ' + req.dateStr + lateStr + '\n' +
    '• Lý do: ' + req.reason;

  if (email) {
    MailApp.sendEmail(email, '[LTS] ' + req.type + ' — ' + req.name, text);
  }
  if (tgToken && tgChat) {
    UrlFetchApp.fetch('https://api.telegram.org/bot' + tgToken + '/sendMessage', {
      method: 'post',
      payload: { chat_id: tgChat, text: text },
      muteHttpExceptions: true,
    });
  }
}

/* ============================================================
 * PHÂN QUYỀN QUẢN LÝ (PIN)
 * ============================================================ */

function checkPin_(pin) {
  const saved = PropertiesService.getScriptProperties().getProperty('MANAGER_PIN');
  if (!saved) return false; // chưa đặt PIN -> khoá toàn bộ khu quản lý
  return String(pin || '').trim() === saved;
}

function getSheetUrl_() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty('SHEET_URL') || SpreadsheetApp.getActive().getUrl();
}

function getBgUrl_() {
  return PropertiesService.getScriptProperties().getProperty('BG_URL') || '';
}

/** Chạy tay để đặt PIN quản lý + link Sheet hiển thị cho manager. */
function configManager() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  const r1 = ui.prompt('Đặt PIN quản lý', 'Nhập mã PIN (VD 6 số). Để trống = khoá khu quản lý:', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() === ui.Button.OK) {
    const v = r1.getResponseText().trim();
    if (v) props.setProperty('MANAGER_PIN', v); else props.deleteProperty('MANAGER_PIN');
  }

  const r2 = ui.prompt('Link Google Sheet', 'Dán link Sheet cho nút "Mở Sheet" (để trống = tự lấy link hiện tại):', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() === ui.Button.OK) {
    const v = r2.getResponseText().trim();
    if (v) props.setProperty('SHEET_URL', v); else props.deleteProperty('SHEET_URL');
  }

  ui.alert('✅ Đã lưu cấu hình quản lý.');
}

/** Chạy tay để nhập cấu hình thông báo (dùng prompt). */
function configNotify() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  const r1 = ui.prompt('Email nhận thông báo', 'Nhập email Manager (để trống = bỏ qua):', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() === ui.Button.OK) {
    const v = r1.getResponseText().trim();
    if (v) props.setProperty('MANAGER_EMAIL', v); else props.deleteProperty('MANAGER_EMAIL');
  }

  const r2 = ui.prompt('Telegram Bot Token', 'Nhập bot token (để trống = bỏ qua):', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() === ui.Button.OK) {
    const v = r2.getResponseText().trim();
    if (v) props.setProperty('TELEGRAM_BOT_TOKEN', v); else props.deleteProperty('TELEGRAM_BOT_TOKEN');
  }

  const r3 = ui.prompt('Telegram Chat ID', 'Nhập chat_id (để trống = bỏ qua):', ui.ButtonSet.OK_CANCEL);
  if (r3.getSelectedButton() === ui.Button.OK) {
    const v = r3.getResponseText().trim();
    if (v) props.setProperty('TELEGRAM_CHAT_ID', v); else props.deleteProperty('TELEGRAM_CHAT_ID');
  }

  ui.alert('✅ Đã lưu cấu hình thông báo.');
}

/* ============================================================
 * UTILS
 * ============================================================ */

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
function ymd_(d) {
  return d.getFullYear() + '-' + (d.getMonth() + 1) + '-' + d.getDate();
}
function formatD_(d) {
  return Utilities.formatDate(d, CONFIG.TIMEZONE, 'dd/MM/yyyy');
}
function formatDT_(d) {
  return Utilities.formatDate(d, CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm');
}
/** Giờ đến: ô có thể lưu dạng Date (do Sheets tự chuyển "19:00") -> trả HH:mm. */
function formatArrival_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, CONFIG.TIMEZONE, 'HH:mm');
  return String(v || '');
}

/* ============================================================
 * SETUP - chạy 1 lần để tạo & format sheet
 * ============================================================ */

function setup() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(CONFIG.TIMEZONE);

  setupDataSheet_(ss);
  setupMembersSheet_(ss);
  setupSummarySheet_(ss);

  const defaultSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('Trang tính1');
  if (defaultSheet && ss.getSheets().length > 3) ss.deleteSheet(defaultSheet);

  SpreadsheetApp.getUi().alert('✅ Setup xong! Đã tạo 3 sheet: Data, Members, Tổng kết.');
}

function setupDataSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.DATA_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONFIG.DATA_SHEET, 0);

  const headers = [
    'Thời gian gửi', 'Tên', 'Loại', 'Ngày áp dụng',
    'Giờ đến dự kiến', 'Số phút trễ', 'Lý do', 'Trạng thái'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  sheet.getRange(1, 1, 1, headers.length)
    .setBackground(CONFIG.COLOR_PRIMARY).setFontColor(CONFIG.COLOR_HEADER_TEXT)
    .setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 36);
  sheet.setFrozenRows(1);

  const widths = [160, 160, 90, 120, 130, 110, 280, 110];
  widths.forEach(function (w, i) { sheet.setColumnWidth(i + 1, w); });

  const maxRows = sheet.getMaxRows();
  sheet.getRange(2, 1, maxRows - 1, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  sheet.getRange(2, 4, maxRows - 1, 1).setNumberFormat('dd/MM/yyyy');
  sheet.getRange(2, 1, maxRows - 1, headers.length).setVerticalAlignment('middle');
  sheet.getRange(2, 3, maxRows - 1, 1).setHorizontalAlignment('center');
  sheet.getRange(2, 5, maxRows - 1, 2).setHorizontalAlignment('center');
  sheet.getRange(2, 8, maxRows - 1, 1).setHorizontalAlignment('center');

  const statusRange = sheet.getRange(2, 8, maxRows - 1, 1);
  statusRange.setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(CONFIG.STATUSES, true).setAllowInvalid(false).build()
  );

  const rules = [
    ruleTextEq_(statusRange, 'Chờ duyệt', '#fef9c3'),
    ruleTextEq_(statusRange, 'Đã duyệt', '#dcfce7'),
    ruleTextEq_(statusRange, 'Từ chối', '#fee2e2'),
    ruleTextEq_(sheet.getRange(2, 3, maxRows - 1, 1), 'Đi trễ', '#ffedd5'),
    ruleTextEq_(sheet.getRange(2, 3, maxRows - 1, 1), 'Nghỉ', '#e0e7ff'),
  ];
  sheet.setConditionalFormatRules(rules);

  // Banding chỉ tô vùng dữ liệu (row 2 trở đi) để không đè lên header xanh
  sheet.getBandings().forEach(function (b) { b.remove(); });
  sheet.getRange(2, 1, maxRows - 1, headers.length)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);

  // Viền mỏng cho toàn bảng (gồm header)
  sheet.getRange(1, 1, maxRows, headers.length)
    .setBorder(true, true, true, true, true, true, '#e2e8f0', SpreadsheetApp.BorderStyle.SOLID);

  // Bộ lọc trên toàn bảng -> manager lọc/sắp xếp theo Tên, Loại, Ngày, Trạng thái
  const existing = sheet.getFilter();
  if (existing) existing.remove();
  sheet.getRange(1, 1, maxRows, headers.length).createFilter();
}

function setupMembersSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.MEMBERS_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONFIG.MEMBERS_SHEET, 1);

  sheet.getRange('A1').setValue('Tên thành viên')
    .setBackground(CONFIG.COLOR_PRIMARY).setFontColor(CONFIG.COLOR_HEADER_TEXT)
    .setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setColumnWidth(1, 220);
  sheet.setFrozenRows(1);

  if (sheet.getLastRow() < 2) {
    sheet.getRange(2, 1, 3, 1).setValues([['Nguyễn Văn A'], ['Trần Thị B'], ['Lê Văn C']]);
  }
}

function setupSummarySheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.SUMMARY_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SUMMARY_SHEET, 2);

  const now = new Date();

  sheet.getRange('A1').setValue('📊 THỐNG KÊ THEO THÁNG');
  sheet.getRange('A1:E1').merge()
    .setBackground(CONFIG.COLOR_PRIMARY).setFontColor(CONFIG.COLOR_HEADER_TEXT)
    .setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
  sheet.setRowHeight(1, 40);

  sheet.getRange('A2').setValue('Tháng:').setFontWeight('bold');
  sheet.getRange('B2').setValue(now.getMonth() + 1);
  sheet.getRange('C2').setValue('Năm:').setFontWeight('bold');
  sheet.getRange('D2').setValue(now.getFullYear());
  sheet.getRange('B2').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['1','2','3','4','5','6','7','8','9','10','11','12'], true).build()
  );
  sheet.getRange('B2:D2').setHorizontalAlignment('center');
  sheet.getRange('A2:D2').setBackground(CONFIG.COLOR_PRIMARY_LIGHT);

  const headers = ['Tên', 'Số lần trễ', 'Tổng phút trễ', 'Số lần nghỉ', 'Tổng vắng/trễ'];
  sheet.getRange(4, 1, 1, headers.length).setValues([headers])
    .setBackground(CONFIG.COLOR_PRIMARY).setFontColor(CONFIG.COLOR_HEADER_TEXT)
    .setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setFrozenRows(4);

  const widths = [220, 110, 130, 110, 130];
  widths.forEach(function (w, i) { sheet.setColumnWidth(i + 1, w); });

  // Banding tô sẵn 50 dòng (row 5 trở đi) để không đè header row 4
  const ROWS = 50;
  sheet.getBandings().forEach(function (b) { b.remove(); });
  sheet.getRange(5, 1, ROWS, headers.length)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
  sheet.getRange(4, 1, ROWS + 1, headers.length)
    .setBorder(true, true, true, true, true, true, '#e2e8f0', SpreadsheetApp.BorderStyle.SOLID);

  // Xoá mọi biểu đồ cũ (gồm chart rỗng "Thêm một chuỗi dữ liệu...")
  sheet.getCharts().forEach(function (c) { sheet.removeChart(c); });

  // Tính & ghi số liệu (không dùng công thức -> không lỗi locale)
  refreshSummary_(sheet);
}

/**
 * Tính thống kê theo Tháng/Năm ở B2/D2 rồi ghi số trực tiếp.
 * Dùng lại getStats_ (đã bỏ "Từ chối"). Không có công thức -> không lỗi #ERROR.
 */
function refreshSummary_(sheet) {
  if (!sheet) sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SUMMARY_SHEET);
  if (!sheet) return;

  const month = Number(sheet.getRange('B2').getValue()) || (new Date().getMonth() + 1);
  const year = Number(sheet.getRange('D2').getValue()) || new Date().getFullYear();
  const stats = getStats_(month, year); // [{name, late, lateMinutes, off}]

  const FIRST = 5, ROWS = 50;
  // Xoá vùng cũ
  sheet.getRange(FIRST, 1, ROWS, 5).clearContent();

  const values = stats.slice(0, ROWS).map(function (s) {
    return [s.name, s.late, s.lateMinutes, s.off, s.late + s.off];
  });
  if (values.length) {
    sheet.getRange(FIRST, 1, values.length, 5).setValues(values);
  }
  sheet.getRange(FIRST, 2, ROWS, 4).setHorizontalAlignment('center');
}

/** Nút menu: làm mới thống kê thủ công. */
function refreshSummary() {
  refreshSummary_();
  SpreadsheetApp.getActive().toast('Đã cập nhật thống kê.', 'LTS', 3);
}

/** Tự làm mới khi đổi ô Tháng (B2) hoặc Năm (D2) trong sheet Tổng kết. */
function onEdit(e) {
  try {
    const sh = e.range.getSheet();
    if (sh.getName() !== CONFIG.SUMMARY_SHEET) return;
    const a1 = e.range.getA1Notation();
    if (a1 === 'B2' || a1 === 'D2') refreshSummary_(sh);
  } catch (err) { /* bỏ qua */ }
}

/* Menu tiện lợi trong Sheet */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ LTS')
    .addItem('Setup / Format lại', 'setup')
    .addItem('Làm mới thống kê', 'refreshSummary')
    .addSeparator()
    .addItem('Đặt PIN quản lý + link Sheet', 'configManager')
    .addItem('Cấu hình thông báo (email/Telegram)', 'configNotify')
    .addToUi();
}

function ruleTextEq_(range, text, bg) {
  return SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(text).setBackground(bg).setRanges([range]).build();
}
