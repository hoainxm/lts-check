/**
 * LTS - Quản lý Đi trễ / Xin nghỉ buổi tập
 * Backend API bằng Google Apps Script
 *
 * MÔ HÌNH DỮ LIỆU:
 *   - Mỗi tháng 1 sheet tên "MM/YYYY" (VD 07/2026, 08/2026). Tự tạo khi có yêu cầu tháng đó.
 *   - User gửi -> ghi thẳng vào sheet tháng, Trạng thái = "Chờ duyệt".
 *   - Admin Duyệt -> Trạng thái = "Đã duyệt"; Từ chối -> "Từ chối".
 *   - Thống kê / Tổng kết: chỉ tính bản "Đã duyệt".
 *   - Lịch sử: mọi bản của 1 người (mọi trạng thái).
 *   - Lịch tuần: sheet "Lịch tuần", 1 dòng / (tuần, người). Chỉ đăng ký cho TUẦN SAU,
 *     sang Thứ 2 tuần đó thì khóa. Mọi lần lưu đều ghi thêm vào "Log lịch" (append-only)
 *     để đối chiếu khi có khiếu nại "đã đăng ký mà không thấy".
 *   - Điểm danh: sheet "Điểm danh", 1 dòng / (ngày, người).
 *   - Mọi ngày/tuần tính theo giờ VN (CONFIG.TIMEZONE), lưu dạng text yyyy-MM-dd
 *     -> không phụ thuộc múi giờ của project Apps Script.
 *
 * CÁCH DÙNG:
 * 1. setup() (menu ⚙️ LTS -> Setup) tạo sheet tháng hiện tại, Members, Tổng kết.
 * 2. configManager() đặt PIN admin (chạy từ menu, KHÔNG chạy từ editor).
 * 3. (Tuỳ chọn) configNotify().
 * 4. Deploy > Web app (Execute as: Me, Access: Anyone).
 *    Mỗi lần sửa code -> Manage deployments > Edit > New version > Deploy.
 */

const CONFIG = {
  MEMBERS_SHEET: 'Members',
  SUMMARY_SHEET: 'Tổng kết',
  SCHEDULE_SHEET: 'Lịch tuần',
  SCHEDULE_LOG_SHEET: 'Log lịch',
  CHECKIN_SHEET: 'Điểm danh',
  PRACTICE_START_HOUR: 18,
  PRACTICE_START_MIN: 30,
  TIMEZONE: 'Asia/Ho_Chi_Minh',
  TZ_OFFSET: '+07:00', // VN không có giờ mùa hè
  STATUSES: ['Chờ duyệt', 'Đã duyệt', 'Từ chối'],
  TYPES: ['Đi trễ', 'Nghỉ'],
  COLOR_PRIMARY: '#1d4ed8',
  COLOR_PRIMARY_LIGHT: '#dbeafe',
  COLOR_HEADER_TEXT: '#ffffff',
};

// Cột trong sheet tháng (1-based). Thời gian gửi tách 2 cột: Ngày gửi + Giờ gửi.
const COL = { TS_DATE: 1, TS_TIME: 2, NAME: 3, TYPE: 4, DATE: 5, ARRIVAL: 6, LATE: 7, REASON: 8, STATUS: 9, NOTE: 10 };
const MONTH_RE = /^(0[1-9]|1[0-2])\/\d{4}$/;
const ISO_RE = /^\d{4}-\d{2}-\d{2}$/;
const DAY_LABELS = ['T2', 'T3', 'T4', 'T5', 'T6', 'T7', 'CN'];
// Sheet "Lịch tuần": Tuần (Thứ 2) | Tên | T2..CN | Cập nhật lúc
const SCOL = { WEEK: 1, NAME: 2, DAY0: 3, UPDATED: 10 };
// Sheet "Điểm danh": Ngày | Tên | Giờ đến | Cập nhật lúc
const KCOL = { DATE: 1, NAME: 2, TIME: 3, UPDATED: 4 };

/* ============================================================
 * API ENDPOINTS
 * ============================================================ */

function doGet(e) {
  const p = (e && e.parameter) || {};
  const action = p.action || 'members';
  try {
    if (action === 'members') return jsonResponse({ status: 'success', members: getMembers_() });
    if (action === 'config')  return jsonResponse({ status: 'success', bgUrl: getBgUrl_() });

    if (action === 'login') {
      if (!checkPin_(p.pin)) return jsonResponse({ status: 'error', message: 'Sai mã PIN.' });
      return jsonResponse({ status: 'success', sheetUrl: getSheetUrl_() });
    }
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
      return jsonResponse({ status: 'success', history: getHistory_(name, Number(p.limit) || 30) });
    }
    if (action === 'pending') {
      if (!checkPin_(p.pin)) return jsonResponse({ status: 'error', message: 'Sai mã PIN.' });
      return jsonResponse({ status: 'success', items: getPendingList_() });
    }
    if (action === 'schedule') {
      const name = String(p.name || '').trim();
      if (!name) return jsonResponse({ status: 'error', message: 'Thiếu tên.' });
      const week = nextWeekMonday_();
      const found = findScheduleRow_(week, name);
      return jsonResponse({
        status: 'success',
        week: week,
        days: found ? found.days : [false, false, false, false, false, false, false],
        registered: !!found,
        updatedAt: found ? found.updatedAt : '',
      });
    }
    if (action === 'rollcall') {
      if (!checkPin_(p.pin)) return jsonResponse({ status: 'error', message: 'Sai mã PIN.' });
      const date = String(p.date || '').trim();
      if (!ISO_RE.test(date)) return jsonResponse({ status: 'error', message: 'Ngày không hợp lệ.' });
      return jsonResponse({ status: 'success', date: date, list: getRollcall_(date) });
    }
    return jsonResponse({ status: 'error', message: 'Unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ status: 'error', message: 'Lỗi server: ' + err.message });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  // Không lấy được lock mà vẫn ghi -> 2 người lưu cùng lúc có thể ghi đè/lệch dòng nhau
  if (!lock.tryLock(15000))
    return jsonResponse({ status: 'error', message: 'Hệ thống đang bận, thử lại sau vài giây.' });
  try {
    const body = JSON.parse(e.postData.contents);

    // --- User: đăng ký lịch tập tuần sau ---
    if (body.action === 'registerSchedule') return jsonResponse(registerSchedule_(body));

    // --- Admin: điểm danh ---
    if (body.action === 'checkin') {
      if (!checkPin_(body.pin)) return jsonResponse({ status: 'error', message: 'Sai mã PIN.' });
      return jsonResponse(saveCheckin_(body));
    }

    // --- Admin: đổi hình nền ---
    if (body.action === 'setBackground') {
      if (!checkPin_(body.pin)) return jsonResponse({ status: 'error', message: 'Sai mã PIN.' });
      const url = String(body.bgUrl || '').trim();
      const props = PropertiesService.getScriptProperties();
      if (url) {
        if (!/^https?:\/\//i.test(url))
          return jsonResponse({ status: 'error', message: 'URL hình không hợp lệ (http/https).' });
        props.setProperty('BG_URL', url);
      } else props.deleteProperty('BG_URL');
      return jsonResponse({ status: 'success', message: url ? 'Đã đổi hình nền.' : 'Đã xoá hình nền.', bgUrl: url });
    }

    // --- Admin: duyệt / từ chối (đổi trạng thái, không xoá) ---
    if (body.action === 'approve' || body.action === 'reject') {
      if (!checkPin_(body.pin)) return jsonResponse({ status: 'error', message: 'Sai mã PIN.' });
      const sheetName = String(body.sheet || '');
      const row = Number(body.row);
      const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
      if (!sheet || !MONTH_RE.test(sheetName) || !(row >= 2 && row <= sheet.getLastRow()))
        return jsonResponse({ status: 'error', message: 'Yêu cầu không còn tồn tại. Tải lại.' });
      const note = String(body.note || '').trim();
      if (body.action === 'reject' && !note)
        return jsonResponse({ status: 'error', message: 'Cần nhập lý do từ chối.' });
      const newStatus = body.action === 'approve' ? 'Đã duyệt' : 'Từ chối';
      sheet.getRange(row, COL.STATUS).setValue(newStatus);
      sheet.getRange(row, COL.NOTE).setValue(note);
      try { refreshSummary_(); } catch (ignore) {}
      return jsonResponse({ status: 'success', message: body.action === 'approve' ? 'Đã duyệt.' : 'Đã từ chối.' });
    }

    // --- User: gửi yêu cầu -> ghi vào sheet tháng ---
    const name = String(body.name || '').trim();
    const type = String(body.type || '').trim();
    const dateStr = String(body.date || '').trim();
    const reason = String(body.reason || '').trim();
    const arrivalTime = String(body.arrivalTime || '').trim();

    if (!name) return jsonResponse({ status: 'error', message: 'Thiếu tên thành viên.' });
    if (CONFIG.TYPES.indexOf(type) === -1) return jsonResponse({ status: 'error', message: 'Loại yêu cầu không hợp lệ.' });
    if (!ISO_RE.test(dateStr)) return jsonResponse({ status: 'error', message: 'Ngày không hợp lệ.' });
    if (!reason) return jsonResponse({ status: 'error', message: 'Thiếu lý do.' });

    // Hạn nộp tính theo giờ VN phía server (không tin đồng hồ máy user)
    const today = todayVN_();
    if (type === 'Nghỉ' && dateStr <= today)
      return jsonResponse({ status: 'error', message:
        'Đơn nghỉ phải gửi trước 0h của ngày nghỉ — hôm nay ' + isoToDM_(today) +
        ', sớm nhất là ngày ' + isoToDM_(isoAddDays_(today, 1)) + '.' });
    if (type === 'Đi trễ' && dateStr < today)
      return jsonResponse({ status: 'error', message: 'Không gửi đơn đi trễ cho ngày đã qua.' });

    let lateMinutes = '', arrivalDisplay = '';
    if (type === 'Đi trễ') {
      if (!/^\d{1,2}:\d{2}$/.test(arrivalTime)) return jsonResponse({ status: 'error', message: 'Thiếu giờ đến dự kiến.' });
      const parts = arrivalTime.split(':');
      const arrMin = Number(parts[0]) * 60 + Number(parts[1]);
      const startMin = CONFIG.PRACTICE_START_HOUR * 60 + CONFIG.PRACTICE_START_MIN;
      lateMinutes = Math.max(0, arrMin - startMin);
      arrivalDisplay = arrivalTime;
    }

    // Nửa đêm giờ VN — new Date(y, m, d) lấy múi giờ của project Apps Script,
    // project khác múi VN thì ngày áp dụng / sheet tháng có thể lệch 1 ngày.
    const appliedDate = new Date(dateStr + 'T00:00:00' + CONFIG.TZ_OFFSET);
    const sheet = getOrCreateMonthSheet_(monthSheetName_(appliedDate));

    if (isDuplicate_(sheet, name, type, appliedDate))
      return jsonResponse({ status: 'error', message: 'Bạn đã gửi yêu cầu "' + type + '" cho ngày này rồi.' });

    const now = new Date();
    // Cột 1 = ngày gửi, cột 2 = giờ gửi (cùng thời điểm, format khác nhau khi hiển thị)
    sheet.appendRow([now, now, name, type, appliedDate, arrivalDisplay, lateMinutes, reason, 'Chờ duyệt']);

    try { notify_({ name: name, type: type, dateStr: dateStr, arrivalTime: arrivalDisplay, lateMinutes: lateMinutes, reason: reason }); }
    catch (ignore) {}

    return jsonResponse({ status: 'success', message: 'Đã gửi, chờ admin duyệt.' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: 'Lỗi server: ' + err.message });
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================
 * SHEET THÁNG
 * ============================================================ */

function monthSheetName_(dateObj) {
  return Utilities.formatDate(dateObj, CONFIG.TIMEZONE, 'MM/yyyy');
}

function listMonthSheets_() {
  return SpreadsheetApp.getActive().getSheets().filter(function (sh) {
    return MONTH_RE.test(sh.getName());
  });
}

function getOrCreateMonthSheet_(name) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); formatMonthSheet_(sh); }
  return sh;
}

/* ============================================================
 * NGÀY THEO GIỜ VN (chuỗi yyyy-MM-dd, tính toán thuần chuỗi/UTC
 * -> không lệch khi project Apps Script để múi giờ khác)
 * ============================================================ */

function todayVN_() { return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd'); }
function nowHHmmVN_() { return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'HH:mm'); }
function isoToUTC_(iso) { const p = iso.split('-'); return new Date(Date.UTC(+p[0], +p[1] - 1, +p[2])); }
function isoAddDays_(iso, n) {
  const d = isoToUTC_(iso);
  d.setUTCDate(d.getUTCDate() + n);
  return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
}
function weekdayIdx_(iso) { return (isoToUTC_(iso).getUTCDay() + 6) % 7; } // T2 = 0 … CN = 6
function mondayOf_(iso) { return isoAddDays_(iso, -weekdayIdx_(iso)); }
// Chủ nhật 27/09 -> 28/09; Thứ 2 28/09 -> 05/10
function nextWeekMonday_() { return isoAddDays_(mondayOf_(todayVN_()), 7); }
function isoToDM_(iso) { return iso.slice(8, 10) + '/' + iso.slice(5, 7); }
// Ô ngày có thể là text "2026-09-28" hoặc Date (nếu ai đó sửa tay trong Sheet)
function cellToIso_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  return String(v || '').trim();
}
function hhmmToMin_(s) {
  const m = /^(\d{1,2}):(\d{2})$/.exec(String(s || '').trim());
  return m ? Number(m[1]) * 60 + Number(m[2]) : null;
}

/* ============================================================
 * LỊCH TUẦN
 * ============================================================ */

function getScheduleSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CONFIG.SCHEDULE_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.SCHEDULE_SHEET);
    const headers = ['Tuần (Thứ 2)', 'Tên'].concat(DAY_LABELS).concat(['Cập nhật lúc']);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    styleHeader_(sh.getRange(1, 1, 1, headers.length));
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@'); // tuần = text, không để Sheet tự đổi sang Date
    sh.setColumnWidth(SCOL.NAME, 150);
    sh.getRange(1, SCOL.DAY0, sh.getMaxRows(), 7).setHorizontalAlignment('center');
  }
  return sh;
}

function getScheduleLogSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CONFIG.SCHEDULE_LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.SCHEDULE_LOG_SHEET);
    const headers = ['Thời gian (VN)', 'Tên', 'Tuần (Thứ 2)', 'Ngày đăng ký', 'Kết quả'];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    styleHeader_(sh.getRange(1, 1, 1, headers.length));
    sh.setFrozenRows(1);
    sh.getRange('A:C').setNumberFormat('@');
  }
  return sh;
}

function logSchedule_(name, week, days, result) {
  try {
    const picked = DAY_LABELS.filter(function (_, i) { return days && days[i]; }).join(', ');
    getScheduleLogSheet_().appendRow([
      Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss'),
      name, week, picked || '(không ngày nào)', result,
    ]);
  } catch (ignore) {}
}

// Trả { row, days[7], updatedAt } hoặc null
function findScheduleRow_(week, name) {
  const sh = getScheduleSheet_();
  if (sh.getLastRow() < 2) return null;
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, SCOL.UPDATED).getValues();
  for (let i = rows.length - 1; i >= 0; i--) { // dòng mới nhất thắng nếu lỡ trùng
    const r = rows[i];
    if (cellToIso_(r[SCOL.WEEK - 1]) !== week || String(r[SCOL.NAME - 1]).trim() !== name) continue;
    const days = [];
    for (let k = 0; k < 7; k++) days.push(String(r[SCOL.DAY0 - 1 + k]).trim() !== '');
    const upd = r[SCOL.UPDATED - 1];
    return {
      row: i + 2, days: days,
      updatedAt: upd instanceof Date ? formatDT_(upd) : String(upd || ''),
    };
  }
  return null;
}

function registerSchedule_(body) {
  const name = String(body.name || '').trim();
  const week = String(body.week || '').trim();
  const days = body.days;
  if (!name || getMembers_().indexOf(name) === -1)
    return { status: 'error', message: 'Tên không có trong danh sách thành viên.' };
  if (!Array.isArray(days) || days.length !== 7)
    return { status: 'error', message: 'Dữ liệu ngày không hợp lệ, tải lại trang.' };

  const target = nextWeekMonday_();
  if (week !== target) {
    logSchedule_(name, week, days, 'TỪ CHỐI: tuần ' + week + ' ≠ ' + target);
    if (ISO_RE.test(week) && week < target)
      return { status: 'error', message: 'Tuần ' + isoToDM_(week) + ' đã khóa (đã sang Thứ 2). ' +
        'Tải lại trang để đăng ký tuần ' + isoToDM_(target) + ', hoặc nhờ quản lý điểm danh.' };
    return { status: 'error', message: 'Tuần không hợp lệ — tải lại trang rồi đăng ký lại.' };
  }

  const bools = days.map(function (v) { return v === true; });
  const values = [target, name].concat(bools.map(function (v) { return v ? '✓' : ''; })).concat([new Date()]);
  const sh = getScheduleSheet_();
  const found = findScheduleRow_(target, name);
  if (found) sh.getRange(found.row, 1, 1, values.length).setValues([values]);
  else sh.appendRow(values);
  SpreadsheetApp.flush();

  // Đọc lại để chắc chắn đã ghi — trả về đúng cái đang nằm trong Sheet
  const saved = findScheduleRow_(target, name);
  const ok = !!saved && saved.days.join() === bools.join();
  logSchedule_(name, target, bools, ok ? 'OK' : 'LỖI: đọc lại không khớp');
  if (!ok) return { status: 'error', message: 'Lưu chưa thành công, thử lại.' };

  const picked = DAY_LABELS.filter(function (_, i) { return bools[i]; });
  return {
    status: 'success',
    week: target, days: saved.days, updatedAt: saved.updatedAt,
    message: 'Đã lưu lịch tuần ' + isoToDM_(target) + ': ' +
      (picked.length ? picked.join(', ') : 'không đi buổi nào') + '.',
  };
}

/* ============================================================
 * ĐIỂM DANH
 * ============================================================ */

function getCheckinSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CONFIG.CHECKIN_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.CHECKIN_SHEET);
    const headers = ['Ngày', 'Tên', 'Giờ đến', 'Cập nhật lúc'];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    styleHeader_(sh.getRange(1, 1, 1, headers.length));
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@');
    sh.getRange('C:C').setNumberFormat('@');
    sh.setColumnWidth(KCOL.NAME, 150);
  }
  return sh;
}

function saveCheckin_(body) {
  const name = String(body.name || '').trim();
  const date = String(body.date || '').trim();
  let time = String(body.time || '').trim();
  if (!name) return { status: 'error', message: 'Thiếu tên.' };
  if (!ISO_RE.test(date)) return { status: 'error', message: 'Ngày không hợp lệ.' };
  if (time === 'now') {
    if (date !== todayVN_()) return { status: 'error', message: '"Bây giờ" chỉ dùng cho hôm nay — nhập giờ tay.' };
    time = nowHHmmVN_();
  }
  if (time && hhmmToMin_(time) === null) return { status: 'error', message: 'Giờ không hợp lệ.' };

  const sh = getCheckinSheet_();
  let row = 0;
  if (sh.getLastRow() >= 2) {
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
    for (let i = 0; i < rows.length; i++)
      if (cellToIso_(rows[i][0]) === date && String(rows[i][1]).trim() === name) { row = i + 2; break; }
  }
  if (!time) { // bỏ trống giờ = xoá điểm danh
    if (row) sh.deleteRow(row);
    return { status: 'success', message: 'Đã xoá điểm danh của ' + name + '.' };
  }
  const values = [[date, name, time, new Date()]];
  if (row) sh.getRange(row, 1, 1, 4).setValues(values);
  else sh.appendRow(values[0]);
  return { status: 'success', message: 'Đã điểm danh ' + name + ' lúc ' + time + '.' };
}

function getRollcall_(date) {
  const map = {};
  const get = function (nm) {
    if (!map[nm]) map[nm] = { name: nm, registered: false, off: false, expected: '', checkin: '' };
    return map[nm];
  };

  // 1. Lịch tuần chứa ngày đó
  const week = mondayOf_(date), idx = weekdayIdx_(date);
  const ssh = getScheduleSheet_();
  if (ssh.getLastRow() >= 2) {
    ssh.getRange(2, 1, ssh.getLastRow() - 1, SCOL.UPDATED).getValues().forEach(function (r) {
      if (cellToIso_(r[SCOL.WEEK - 1]) !== week) return;
      const nm = String(r[SCOL.NAME - 1]).trim();
      if (!nm) return;
      // dòng sau ghi đè dòng trước (lỡ trùng)
      const on = String(r[SCOL.DAY0 - 1 + idx]).trim() !== '';
      if (on) get(nm).registered = true; else if (map[nm]) map[nm].registered = false;
    });
  }

  // 2. Đơn trễ/nghỉ cho ngày đó (bỏ đơn bị từ chối)
  const p = date.split('-');
  const msh = SpreadsheetApp.getActive().getSheetByName(p[1] + '/' + p[0]);
  if (msh && msh.getLastRow() >= 2) {
    msh.getRange(2, 1, msh.getLastRow() - 1, COL.STATUS).getValues().forEach(function (r) {
      if (!(r[COL.DATE - 1] instanceof Date) || cellToIso_(r[COL.DATE - 1]) !== date) return;
      if (String(r[COL.STATUS - 1]).trim() === 'Từ chối') return;
      const nm = String(r[COL.NAME - 1]).trim();
      const type = String(r[COL.TYPE - 1]).trim();
      if (type === 'Nghỉ') get(nm).off = true;
      else if (type === 'Đi trễ') get(nm).expected = formatArrival_(r[COL.ARRIVAL - 1]);
    });
  }

  // 3. Điểm danh
  const ksh = getCheckinSheet_();
  if (ksh.getLastRow() >= 2) {
    ksh.getRange(2, 1, ksh.getLastRow() - 1, 3).getValues().forEach(function (r) {
      if (cellToIso_(r[KCOL.DATE - 1]) !== date) return;
      get(String(r[KCOL.NAME - 1]).trim()).checkin = formatArrival_(r[KCOL.TIME - 1]);
    });
  }

  const startMin = CONFIG.PRACTICE_START_HOUR * 60 + CONFIG.PRACTICE_START_MIN;
  const today = todayVN_();
  const list = Object.keys(map).map(function (k) {
    const it = map[k];
    const arr = hhmmToMin_(it.checkin);
    if (arr !== null) {
      if (!it.registered && !it.expected && !it.off) it.result = 'walkin';
      else if (arr <= startMin) it.result = 'ontime';
      else if (it.expected) it.result = arr <= (hhmmToMin_(it.expected) || startMin) ? 'late_ok' : 'late_over';
      else it.result = 'late';
    } else if (it.off) it.result = 'off_ok';
    else if (date < today || (date === today && hhmmToMin_(nowHHmmVN_()) > startMin + 60)) it.result = 'absent';
    else it.result = 'pending';
    return it;
  });
  const order = getMembers_();
  list.sort(function (a, b) {
    const ia = order.indexOf(a.name), ib = order.indexOf(b.name);
    return (ia === -1 ? 999 : ia) - (ib === -1 ? 999 : ib) || a.name.localeCompare(b.name);
  });
  return list;
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
  const target = cellToIso_(appliedDate);
  return rows.some(function (r) {
    return String(r[COL.NAME - 1]).trim() === name &&
      String(r[COL.TYPE - 1]).trim() === type &&
      r[COL.DATE - 1] instanceof Date && cellToIso_(r[COL.DATE - 1]) === target &&
      String(r[COL.STATUS - 1]).trim() !== 'Từ chối';
  });
}

function getPendingList_() {
  const out = [];
  listMonthSheets_().forEach(function (sh) {
    if (sh.getLastRow() < 2) return;
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, COL.STATUS).getValues();
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (String(r[COL.STATUS - 1]).trim() !== 'Chờ duyệt') continue;
      out.push({
        sheet: sh.getName(),
        row: i + 2,
        tsRaw: r[COL.TS_DATE - 1] instanceof Date ? r[COL.TS_DATE - 1].getTime() : 0,
        timestamp: r[COL.TS_DATE - 1] instanceof Date ? formatDT_(r[COL.TS_DATE - 1]) : '',
        name: String(r[COL.NAME - 1]).trim(),
        type: String(r[COL.TYPE - 1]).trim(),
        date: r[COL.DATE - 1] instanceof Date ? formatD_(r[COL.DATE - 1]) : '',
        arrival: formatArrival_(r[COL.ARRIVAL - 1]),
        lateMinutes: r[COL.LATE - 1] === '' ? '' : Number(r[COL.LATE - 1]),
        reason: String(r[COL.REASON - 1] || ''),
      });
    }
  });
  out.sort(function (a, b) { return a.tsRaw - b.tsRaw; }); // cũ nhất trước (FIFO)
  return out;
}

function getStats_(month, year) {
  const name = pad2_(month) + '/' + year;
  const sheet = SpreadsheetApp.getActive().getSheetByName(name);
  const members = getMembers_();
  const map = {};
  members.forEach(function (m) { map[m] = { name: m, late: 0, lateMinutes: 0, off: 0 }; });

  if (sheet && sheet.getLastRow() >= 2) {
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, COL.STATUS).getValues();
    rows.forEach(function (r) {
      const nm = String(r[COL.NAME - 1]).trim();
      const type = String(r[COL.TYPE - 1]).trim();
      const status = String(r[COL.STATUS - 1]).trim();
      if (!nm || status !== 'Đã duyệt') return; // chỉ tính đã duyệt
      if (!map[nm]) map[nm] = { name: nm, late: 0, lateMinutes: 0, off: 0 };
      if (type === 'Đi trễ') { map[nm].late += 1; map[nm].lateMinutes += Number(r[COL.LATE - 1]) || 0; }
      else if (type === 'Nghỉ') { map[nm].off += 1; }
    });
  }
  return Object.keys(map).map(function (k) { return map[k]; })
    .sort(function (a, b) { return (b.late + b.off) - (a.late + a.off); });
}

function getHistory_(name, limit) {
  const all = [];
  listMonthSheets_().forEach(function (sh) {
    if (sh.getLastRow() < 2) return;
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, COL.NOTE).getValues();
    rows.forEach(function (r) {
      if (String(r[COL.NAME - 1]).trim() !== name) return;
      all.push({
        tsRaw: r[COL.TS_DATE - 1] instanceof Date ? r[COL.TS_DATE - 1].getTime() : 0,
        timestamp: r[COL.TS_DATE - 1] instanceof Date ? formatDT_(r[COL.TS_DATE - 1]) : '',
        type: String(r[COL.TYPE - 1]).trim(),
        date: r[COL.DATE - 1] instanceof Date ? formatD_(r[COL.DATE - 1]) : '',
        arrival: formatArrival_(r[COL.ARRIVAL - 1]),
        lateMinutes: r[COL.LATE - 1] === '' ? '' : Number(r[COL.LATE - 1]),
        reason: String(r[COL.REASON - 1] || ''),
        status: String(r[COL.STATUS - 1]).trim(),
        note: String(r[COL.NOTE - 1] || ''),
      });
    });
  });
  all.sort(function (a, b) { return b.tsRaw - a.tsRaw; }); // mới nhất trước
  return all.slice(0, limit);
}

/* ============================================================
 * NOTIFY (tuỳ chọn)
 * ============================================================ */

function notify_(req) {
  const props = PropertiesService.getScriptProperties();
  const email = props.getProperty('MANAGER_EMAIL');
  const tgToken = props.getProperty('TELEGRAM_BOT_TOKEN');
  const tgChat = props.getProperty('TELEGRAM_CHAT_ID');

  const lateStr = req.type === 'Đi trễ'
    ? ('\n• Giờ đến: ' + req.arrivalTime + ' (trễ ' + req.lateMinutes + ' phút)') : '';
  const text =
    '🔔 Yêu cầu mới (chờ duyệt) — LTS\n' +
    '• Tên: ' + req.name + '\n• Loại: ' + req.type + '\n• Ngày: ' + req.dateStr + lateStr +
    '\n• Lý do: ' + req.reason;

  if (email) MailApp.sendEmail(email, '[LTS] ' + req.type + ' — ' + req.name, text);
  if (tgToken && tgChat) {
    UrlFetchApp.fetch('https://api.telegram.org/bot' + tgToken + '/sendMessage', {
      method: 'post', payload: { chat_id: tgChat, text: text }, muteHttpExceptions: true,
    });
  }
}

/* ============================================================
 * PIN + hình nền
 * ============================================================ */

function checkPin_(pin) {
  const saved = PropertiesService.getScriptProperties().getProperty('MANAGER_PIN');
  if (!saved) return false;
  return String(pin || '').trim() === saved;
}
function getSheetUrl_() {
  return PropertiesService.getScriptProperties().getProperty('SHEET_URL') || SpreadsheetApp.getActive().getUrl();
}
function getBgUrl_() {
  return PropertiesService.getScriptProperties().getProperty('BG_URL') || '';
}

function configManager() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const r1 = ui.prompt('Đặt PIN quản lý', 'Nhập mã PIN (để trống = khoá khu quản lý):', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() === ui.Button.OK) {
    const v = r1.getResponseText().trim();
    if (v) props.setProperty('MANAGER_PIN', v); else props.deleteProperty('MANAGER_PIN');
  }
  const r2 = ui.prompt('Link Google Sheet', 'Dán link Sheet (để trống = tự lấy):', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() === ui.Button.OK) {
    const v = r2.getResponseText().trim();
    if (v) props.setProperty('SHEET_URL', v); else props.deleteProperty('SHEET_URL');
  }
  ui.alert('✅ Đã lưu cấu hình quản lý.');
}

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
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function pad2_(n) { return ('0' + n).slice(-2); }
function formatD_(d) { return Utilities.formatDate(d, CONFIG.TIMEZONE, 'dd/MM/yyyy'); }
function formatDT_(d) { return Utilities.formatDate(d, CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm'); }
function formatArrival_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, CONFIG.TIMEZONE, 'HH:mm');
  return String(v || '');
}

/* ============================================================
 * SETUP
 * ============================================================ */

function setup() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(CONFIG.TIMEZONE);

  // Sheet tháng hiện tại
  getOrCreateMonthSheet_(monthSheetName_(new Date()));
  setupMembersSheet_(ss);
  setupSummarySheet_(ss);
  getScheduleSheet_();
  getScheduleLogSheet_();
  getCheckinSheet_();

  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('Trang tính1');
  if (def && ss.getSheets().length > 3) { try { ss.deleteSheet(def); } catch (e) {} }

  const msg = '✅ Setup xong! Sheet tháng hiện tại + Members + Tổng kết đã sẵn sàng.';
  try { SpreadsheetApp.getUi().alert(msg); }
  catch (e) { try { ss.toast(msg, 'LTS', 5); } catch (e2) { Logger.log(msg); } }
}

function styleHeader_(range, bg) {
  range.setBackground(bg || CONFIG.COLOR_PRIMARY).setFontColor(CONFIG.COLOR_HEADER_TEXT)
    .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center').setVerticalAlignment('middle');
}

function formatMonthSheet_(sheet) {
  const headers = ['Ngày gửi', 'Giờ gửi', 'Tên', 'Loại', 'Ngày áp dụng', 'Giờ đến dự kiến', 'Số phút trễ', 'Lý do', 'Trạng thái', 'Ghi chú (duyệt/từ chối)'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader_(sheet.getRange(1, 1, 1, headers.length));
  sheet.setRowHeight(1, 36);
  sheet.setFrozenRows(1);

  const widths = [110, 90, 150, 90, 120, 130, 100, 260, 110, 240];
  widths.forEach(function (w, i) { sheet.setColumnWidth(i + 1, w); });

  const maxRows = sheet.getMaxRows();
  sheet.getRange(2, COL.TS_DATE, maxRows - 1, 1).setNumberFormat('dd/MM/yyyy');
  sheet.getRange(2, COL.TS_TIME, maxRows - 1, 1).setNumberFormat('HH:mm:ss');
  sheet.getRange(2, COL.DATE, maxRows - 1, 1).setNumberFormat('dd/MM/yyyy');
  sheet.getRange(2, COL.ARRIVAL, maxRows - 1, 1).setNumberFormat('@'); // giờ đến = text
  sheet.getRange(2, 1, maxRows - 1, headers.length).setVerticalAlignment('middle');
  sheet.getRange(2, COL.TS_DATE, maxRows - 1, 2).setHorizontalAlignment('center');
  sheet.getRange(2, COL.TYPE, maxRows - 1, 1).setHorizontalAlignment('center');
  sheet.getRange(2, COL.DATE, maxRows - 1, 3).setHorizontalAlignment('center');
  sheet.getRange(2, COL.STATUS, maxRows - 1, 1).setHorizontalAlignment('center');

  const statusRange = sheet.getRange(2, COL.STATUS, maxRows - 1, 1);
  statusRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.STATUSES, true).setAllowInvalid(true).build());

  sheet.setConditionalFormatRules([
    ruleTextEq_(statusRange, 'Chờ duyệt', '#fef9c3'),
    ruleTextEq_(statusRange, 'Đã duyệt', '#dcfce7'),
    ruleTextEq_(statusRange, 'Từ chối', '#fee2e2'),
    ruleTextEq_(sheet.getRange(2, COL.TYPE, maxRows - 1, 1), 'Đi trễ', '#ffedd5'),
    ruleTextEq_(sheet.getRange(2, COL.TYPE, maxRows - 1, 1), 'Nghỉ', '#e0e7ff'),
  ]);

  sheet.getBandings().forEach(function (b) { b.remove(); });
  sheet.getRange(2, 1, maxRows - 1, headers.length).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
  sheet.getRange(1, 1, maxRows, headers.length).setBorder(true, true, true, true, true, true, '#e2e8f0', SpreadsheetApp.BorderStyle.SOLID);

  const ex = sheet.getFilter(); if (ex) ex.remove();
  sheet.getRange(1, 1, maxRows, headers.length).createFilter();
}

function setupMembersSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.MEMBERS_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONFIG.MEMBERS_SHEET);
  sheet.getRange('A1').setValue('Tên thành viên')
    .setBackground(CONFIG.COLOR_PRIMARY).setFontColor(CONFIG.COLOR_HEADER_TEXT)
    .setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setColumnWidth(1, 220);
  sheet.setFrozenRows(1);
  if (sheet.getLastRow() < 2) sheet.getRange(2, 1, 3, 1).setValues([['Nguyễn Văn A'], ['Trần Thị B'], ['Lê Văn C']]);
}

function setupSummarySheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.SUMMARY_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SUMMARY_SHEET);
  const now = new Date();

  sheet.getRange('A1').setValue('📊 THỐNG KÊ THEO THÁNG');
  sheet.getRange('A1:E1').merge().setBackground(CONFIG.COLOR_PRIMARY).setFontColor(CONFIG.COLOR_HEADER_TEXT)
    .setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
  sheet.setRowHeight(1, 40);

  sheet.getRange('A2').setValue('Tháng:').setFontWeight('bold');
  sheet.getRange('B2').setValue(now.getMonth() + 1);
  sheet.getRange('C2').setValue('Năm:').setFontWeight('bold');
  sheet.getRange('D2').setValue(now.getFullYear());
  sheet.getRange('B2').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['1','2','3','4','5','6','7','8','9','10','11','12'], true).build());
  sheet.getRange('B2:D2').setHorizontalAlignment('center');
  sheet.getRange('A2:D2').setBackground(CONFIG.COLOR_PRIMARY_LIGHT);

  const headers = ['Tên', 'Số lần trễ', 'Tổng phút trễ', 'Số lần nghỉ', 'Tổng vắng/trễ'];
  sheet.getRange(4, 1, 1, headers.length).setValues([headers]);
  styleHeader_(sheet.getRange(4, 1, 1, headers.length));
  sheet.setFrozenRows(4);

  const widths = [220, 110, 130, 110, 130];
  widths.forEach(function (w, i) { sheet.setColumnWidth(i + 1, w); });

  const ROWS = 50;
  sheet.getBandings().forEach(function (b) { b.remove(); });
  sheet.getRange(5, 1, ROWS, headers.length).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
  sheet.getRange(4, 1, ROWS + 1, headers.length).setBorder(true, true, true, true, true, true, '#e2e8f0', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getCharts().forEach(function (c) { sheet.removeChart(c); });

  refreshSummary_(sheet);
}

function refreshSummary_(sheet) {
  if (!sheet) sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SUMMARY_SHEET);
  if (!sheet) return;
  const month = Number(sheet.getRange('B2').getValue()) || (new Date().getMonth() + 1);
  const year = Number(sheet.getRange('D2').getValue()) || new Date().getFullYear();
  const stats = getStats_(month, year);

  const FIRST = 5, ROWS = 50;
  sheet.getRange(FIRST, 1, ROWS, 5).clearContent();
  const values = stats.slice(0, ROWS).map(function (s) { return [s.name, s.late, s.lateMinutes, s.off, s.late + s.off]; });
  if (values.length) sheet.getRange(FIRST, 1, values.length, 5).setValues(values);
  sheet.getRange(FIRST, 2, ROWS, 4).setHorizontalAlignment('center');
}

function refreshSummary() {
  refreshSummary_();
  SpreadsheetApp.getActive().toast('Đã cập nhật thống kê.', 'LTS', 3);
}

function onEdit(e) {
  try {
    const sh = e.range.getSheet();
    if (sh.getName() !== CONFIG.SUMMARY_SHEET) return;
    const a1 = e.range.getA1Notation();
    if (a1 === 'B2' || a1 === 'D2') refreshSummary_(sh);
  } catch (err) { /* bỏ qua */ }
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('⚙️ LTS')
    .addItem('Setup / Format lại', 'setup')
    .addItem('Làm mới thống kê', 'refreshSummary')
    .addSeparator()
    .addItem('Đặt PIN quản lý + link Sheet', 'configManager')
    .addItem('Cấu hình thông báo (email/Telegram)', 'configNotify')
    .addToUi();
}

function ruleTextEq_(range, text, bg) {
  return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(text).setBackground(bg).setRanges([range]).build();
}
