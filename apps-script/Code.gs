/**
 * LTS - Quản lý Đi trễ / Xin nghỉ / Lịch tập / Điểm danh
 * Backend API bằng Google Apps Script
 *
 * MÔ HÌNH DỮ LIỆU (ma trận người × ngày):
 *   - Mỗi tháng 1 sheet "Tháng MM/yyyy" (tự tạo khi có hoạt động), gom cả 3 luồng:
 *     đăng ký lịch tập + đơn Đi trễ/Nghỉ + điểm danh.
 *   - Trong sheet: các BLOCK TUẦN xếp dọc. Mỗi block:
 *       dòng 1: tiêu đề tuần (ô A chứa Date Thứ 2, format "Tuần dd/MM/yyyy") — code định vị block qua Date này
 *       dòng 2: header ngày (T2..CN, mỗi ngày merge 5 cột)
 *       dòng 3: header con: ĐKy tập | Đi muộn/Nghỉ | Giờ dự kiến đến | Giờ đến | Lý do
 *       tiếp theo: mỗi thành viên 1 dòng (pre-fill từ Members)
 *       kết thúc: 1 dòng trống ngăn cách
 *   - Tuần vắt 2 tháng: thuộc sheet của THÁNG CHIẾM ĐA SỐ ngày (tháng của ngày Thứ 5).
 *   - Gửi lại đơn cùng ngày = ghi đè (cập nhật), không chặn trùng.
 *   - Sheet log cũ "MM/yyyy" + sheet "Lịch tập" cũ: ĐÓNG BĂNG, chỉ đọc khi migrate.
 *   - Sheet "Nhật ký": append-only, ghi MỌI lần user/admin bấm gửi (kể cả bị từ chối) + giờ VN
 *     -> đối chiếu khi có khiếu nại "đã đăng ký/đã gửi mà không thấy". Không ghi PIN.
 *
 * CÁCH DÙNG:
 * 1. setup() (menu ⚙️ LTS -> Setup) tạo sheet Tháng hiện tại + block tuần này, Members, Tổng kết.
 * 2. Menu "Chuyển dữ liệu cũ -> sheet Tháng" để migrate log cũ (chạy 1 lần).
 * 3. configManager() đặt PIN admin; (tuỳ chọn) configNotify().
 * 4. Deploy > Web app. Mỗi lần sửa code -> Manage deployments > Edit > New version > Deploy.
 */

const CONFIG = {
  MEMBERS_SHEET: 'Members',
  SUMMARY_SHEET: 'Tổng kết',
  PRACTICE_START_HOUR: 18,
  PRACTICE_START_MIN: 30,
  TIMEZONE: 'Asia/Ho_Chi_Minh',
  TYPES: ['Đi trễ', 'Nghỉ'],
  COLOR_PRIMARY: '#1d4ed8',
  COLOR_PRIMARY_LIGHT: '#dbeafe',
  COLOR_HEADER_TEXT: '#ffffff',
};

// ===== Ma trận "Tháng MM/yyyy" =====
// Cột 1 = Tên. Ngày thứ d (0=T2..6=CN), field f (0=ĐK, 1=Loại, 2=Giờ dự kiến, 3=Giờ đến, 4=Lý do):
// cột = 2 + d*5 + f
const MX = {
  DAYS: 7,
  DAY_COLS: 5,
  TOTAL_COLS: 36, // 1 (Tên) + 7*5
  HEADER_ROWS: 3, // tiêu đề tuần + header ngày + header con
  SUB_HEADERS: ['ĐKy tập', 'Đi muộn/Nghỉ', 'Giờ dự kiến đến', 'Giờ đến', 'Lý do'],
};
const MX_MARK = '✓';
const ABSENT_MARK = '✗'; // vắng không phép (ghi vào ô Giờ đến của ngày đã qua)
const COLOR_ABSENT_BG = '#fecaca'; // nền ô ✗
const COLOR_IDLE_BG = '#f1f5f9';   // ngày đã qua: không đăng ký, không đơn, không đến
const SUMMARY_HEADERS = ['Tên', 'ĐK tập', 'Có mặt', 'Đúng giờ', 'Đi trễ', 'Nghỉ phép', 'Vắng KP', 'Vãng lai', 'Chuyên cần'];
const SUMMARY_ROWS = 60;
const MATRIX_RE = /^Tháng (0[1-9]|1[0-2])\/\d{4}$/;
const DAY_NAMES = ['T2', 'T3', 'T4', 'T5', 'T6', 'T7', 'CN'];
const TIME_RE = /^([01]?\d|2[0-3]):[0-5]\d$/;

// ===== Legacy (chỉ dùng khi migrate, KHÔNG ghi mới) =====
const COL = { TS_DATE: 1, TS_TIME: 2, NAME: 3, TYPE: 4, DATE: 5, ARRIVAL: 6, PRESENT: 7, REASON: 8, STATUS: 9, NOTE: 10 };
const MONTH_RE = /^(0[1-9]|1[0-2])\/\d{4}$/;
const SCHED = { WEEK: 1, NAME: 2, FIRST: 3, DAYS: 7 };
const SCHED_MARK = '✓';
const LEGACY_SCHEDULE_SHEET = 'Lịch tập';
const LOG_SHEET = 'Nhật ký';

function mxCol_(dayIdx, field) { return 2 + dayIdx * MX.DAY_COLS + field; }

/* ============================================================
 * API ENDPOINTS
 * ============================================================ */

function doGet(e) {
  const p = (e && e.parameter) || {};
  const action = p.action || 'members';
  try {
    if (action === 'members') return jsonResponse({ status: 'success', members: getMembers_() });
    if (action === 'config')  return jsonResponse({ status: 'success', bgUrl: getBgUrl_() });

    // Public: lịch đã đăng ký của 1 người cho tuần sau (hiện lại trạng thái tick trên form)
    if (action === 'schedule') {
      const name = String(p.name || '').trim();
      if (!name) return jsonResponse({ status: 'error', message: 'Thiếu tên.' });
      const week = nextWeekStart_();
      const days = [false, false, false, false, false, false, false];
      const sheet = SpreadsheetApp.getActive().getSheetByName(matrixSheetName_(week)); // KHÔNG tạo khi GET
      if (sheet) {
        const block = findWeekBlock_(sheet, week);
        if (block) {
          const row = findMemberRow_(sheet, block, name);
          if (row) {
            const vals = sheet.getRange(row, 2, 1, MX.DAYS * MX.DAY_COLS).getValues()[0];
            for (let i = 0; i < MX.DAYS; i++) days[i] = String(vals[i * MX.DAY_COLS]).trim() === MX_MARK;
          }
        }
      }
      const registered = days.some(function (v) { return v; });
      return jsonResponse({ status: 'success', week: isoDate_(week), days: days, registered: registered });
    }

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
    // Admin: danh sách điểm danh + kết quả đối chiếu của 1 ngày
    if (action === 'rollcall') {
      if (!checkPin_(p.pin)) return jsonResponse({ status: 'error', message: 'Sai mã PIN.' });
      const dateStr = String(p.date || '').trim();
      if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return jsonResponse({ status: 'error', message: 'Ngày không hợp lệ.' });
      return jsonResponse({ status: 'success', date: dateStr, list: getRollcall_(dateStr) });
    }
    return jsonResponse({ status: 'error', message: 'Unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ status: 'error', message: 'Lỗi server: ' + err.message });
  }
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (ignore) {}
  const out = doPostLocked_(e);
  if (body.action !== 'setBackground') logPost_(body, out);
  return out;
}

function doPostLocked_(e) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000))
    return jsonResponse({ status: 'error', message: 'Server đang bận, vui lòng thử lại.' });
  try {
    const body = JSON.parse(e.postData.contents);

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

    // --- User: đăng ký lịch tập tuần sau (ghi đè ô ĐK, giữ nguyên các ô khác) ---
    if (body.action === 'registerSchedule') {
      const name = String(body.name || '').trim();
      const days = body.days;
      if (!name) return jsonResponse({ status: 'error', message: 'Thiếu tên thành viên.' });
      if (!Array.isArray(days) || days.length !== MX.DAYS)
        return jsonResponse({ status: 'error', message: 'Dữ liệu ngày tập không hợp lệ.' });

      // Chỉ nhận tuần KẾ TIẾP — sang Thứ 2 là tuần đó tự khóa (đi tập thì nhờ admin điểm danh)
      const next = nextWeekStart_();
      if (String(body.week || '') !== isoDate_(next))
        return jsonResponse({
          status: 'error',
          message: 'Chỉ đăng ký được cho tuần sau (bắt đầu ' + dmy_(next) + '). Tuần hiện tại đã chốt.',
        });

      const loc = getOrCreateWeekBlock_(next);
      const row = ensureMemberRow_(loc.sheet, loc.block, name);
      const range = loc.sheet.getRange(row, 2, 1, MX.DAYS * MX.DAY_COLS);
      const vals = range.getValues()[0];
      for (let i = 0; i < MX.DAYS; i++) vals[i * MX.DAY_COLS] = days[i] ? MX_MARK : '';
      range.setValues([vals]);
      SpreadsheetApp.flush();

      // Đọc lại đúng ô vừa ghi -> chỉ báo thành công khi Sheet thực sự có dữ liệu
      const loc2 = getOrCreateWeekBlock_(next);
      const row2 = findMemberRow_(loc2.sheet, loc2.block, name);
      const saved = [];
      if (row2) {
        const v2 = loc2.sheet.getRange(row2, 2, 1, MX.DAYS * MX.DAY_COLS).getValues()[0];
        for (let i = 0; i < MX.DAYS; i++) saved.push(String(v2[i * MX.DAY_COLS]).trim() === MX_MARK);
      }
      if (saved.join() !== days.map(function (d) { return !!d; }).join())
        return jsonResponse({ status: 'error', message: 'Lưu chưa thành công, vui lòng thử lại.' });

      const picked = DAY_NAMES.filter(function (_, i) { return saved[i]; });
      return jsonResponse({
        status: 'success', week: isoDate_(next), days: saved,
        updatedAt: Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm'),
        message: 'Đã lưu lịch tuần ' + dmy_(next) + ': ' + (picked.length ? picked.join(', ') : 'không đi buổi nào') + '.',
      });
    }

    // --- Admin: điểm danh — ghi giờ đến vào ô (tên × ngày). time rỗng = xoá điểm danh ---
    if (body.action === 'checkin') {
      if (!checkPin_(body.pin)) return jsonResponse({ status: 'error', message: 'Sai mã PIN.' });
      const name = String(body.name || '').trim();
      const dateStr = String(body.date || '').trim();
      let time = String(body.time || '').trim();
      // 'now' = lấy giờ hiện tại theo múi giờ VN phía server (đồng hồ máy admin có thể lệch)
      if (time === 'now') time = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'HH:mm');
      if (!name) return jsonResponse({ status: 'error', message: 'Thiếu tên.' });
      if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return jsonResponse({ status: 'error', message: 'Ngày không hợp lệ.' });
      if (time && !TIME_RE.test(time)) return jsonResponse({ status: 'error', message: 'Giờ đến phải dạng HH:mm.' });

      const d = dateFromYmd_(dateStr);
      const dayIdx = (d.getDay() + 6) % 7;
      const loc = getOrCreateWeekBlock_(weekStart_(d));
      const row = ensureMemberRow_(loc.sheet, loc.block, name);
      // setBackground(null): xoá nền đỏ nếu ô này từng bị đánh dấu vắng (✗) rồi admin điểm danh bù
      loc.sheet.getRange(row, mxCol_(dayIdx, 3)).setNumberFormat('@').setBackground(null).setValue(time);
      return jsonResponse({
        status: 'success',
        message: time ? ('Đã điểm danh ' + name + ' lúc ' + time + '.') : ('Đã xoá điểm danh của ' + name + '.'),
      });
    }

    // --- User: gửi đơn Đi trễ / Nghỉ -> điền vào ô (tên × ngày) trong ma trận ---
    const name = String(body.name || '').trim();
    const type = String(body.type || '').trim();
    const dateStr = String(body.date || '').trim();
    const reason = String(body.reason || '').trim();
    const arrivalTime = String(body.arrivalTime || '').trim();

    if (!name) return jsonResponse({ status: 'error', message: 'Thiếu tên thành viên.' });
    if (CONFIG.TYPES.indexOf(type) === -1) return jsonResponse({ status: 'error', message: 'Loại yêu cầu không hợp lệ.' });
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return jsonResponse({ status: 'error', message: 'Ngày không hợp lệ.' });
    if (!reason) return jsonResponse({ status: 'error', message: 'Thiếu lý do.' });

    let arrivalDisplay = '';
    if (type === 'Đi trễ') {
      if (!TIME_RE.test(arrivalTime)) return jsonResponse({ status: 'error', message: 'Thiếu giờ đến dự kiến.' });
      arrivalDisplay = arrivalTime;
    }

    // Nghỉ: đơn cho ngày D phải gửi trước 0h ngày D (giờ VN) -> chỉ nhận từ ngày mai trở đi
    if (type === 'Nghỉ' && dateStr <= isoDate_(todayVN_()))
      return jsonResponse({ status: 'error', message: 'Đơn nghỉ phải gửi trước 0h của ngày nghỉ — hôm nay ' +
        dmy_(todayVN_()).slice(0, 5) + ', chỉ nhận từ ngày ' +
        dmy_(new Date(todayVN_().getTime() + 86400000)).slice(0, 5) + ' trở đi.' });

    const appliedDate = dateFromYmd_(dateStr);
    const dayIdx = (appliedDate.getDay() + 6) % 7;
    const loc = getOrCreateWeekBlock_(weekStart_(appliedDate));
    const row = ensureMemberRow_(loc.sheet, loc.block, name);

    // Ghi đè Loại + Giờ dự kiến + Lý do (KHÔNG đụng ô ĐK và Giờ đến)
    const typeCell = loc.sheet.getRange(row, mxCol_(dayIdx, 1));
    const existed = String(typeCell.getValue()).trim() !== '';
    typeCell.setValue(type);
    loc.sheet.getRange(row, mxCol_(dayIdx, 2)).setNumberFormat('@').setValue(arrivalDisplay);
    loc.sheet.getRange(row, mxCol_(dayIdx, 4)).setValue(reason);

    try { refreshSummary_(); } catch (ignore) {}
    try { notify_({ name: name, type: type, dateStr: dateStr, arrivalTime: arrivalDisplay, reason: reason }); }
    catch (ignore) {}

    return jsonResponse({ status: 'success', message: existed ? 'Đã cập nhật yêu cầu.' : 'Đã gửi thành công.' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: 'Lỗi server: ' + err.message });
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================
 * MA TRẬN: SHEET THÁNG + BLOCK TUẦN
 * ============================================================ */

function matrixSheetName_(monday) {
  // Tuần thuộc THÁNG CHIẾM ĐA SỐ ngày trong tuần (= tháng của ngày Thứ 5).
  // VD tuần 29/06–05/07 có 5 ngày tháng 7 -> nằm ở "Tháng 07/yyyy", không rơi về tháng 6.
  const anchor = new Date(monday.getFullYear(), monday.getMonth(), monday.getDate() + 3);
  return 'Tháng ' + pad2_(anchor.getMonth() + 1) + '/' + anchor.getFullYear();
}
function matrixNameMY_(month, year) {
  return 'Tháng ' + pad2_(month) + '/' + year;
}

function getOrCreateMatrixSheet_(monday) {
  const ss = SpreadsheetApp.getActive();
  const name = matrixSheetName_(monday);
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); formatMatrixSheet_(sh); }
  return sh;
}

function formatMatrixSheet_(sheet) {
  // Sheet mới mặc định 26 cột — ma trận cần 36
  if (sheet.getMaxColumns() < MX.TOTAL_COLS)
    sheet.insertColumnsAfter(sheet.getMaxColumns(), MX.TOTAL_COLS - sheet.getMaxColumns());
  sheet.setColumnWidth(1, 150);
  for (let d = 0; d < MX.DAYS; d++) {
    sheet.setColumnWidth(mxCol_(d, 0), 50);
    sheet.setColumnWidth(mxCol_(d, 1), 95);
    sheet.setColumnWidth(mxCol_(d, 2), 90);
    sheet.setColumnWidth(mxCol_(d, 3), 70);
    sheet.setColumnWidth(mxCol_(d, 4), 180);
  }
  sheet.setFrozenColumns(1);
  const maxRows = sheet.getMaxRows();
  // Format cấp cột 1 lần: mọi ô dữ liệu là text ('@', giờ nhập tay không bị Sheets đổi kiểu), canh giữa,
  // wrap để nội dung dài tự xuống dòng -> chiều cao dòng tự giãn. Riêng cột Lý do canh trái.
  // Header ghi sau sẽ tự đè alignment riêng.
  sheet.getRange(1, 2, maxRows, MX.DAYS * MX.DAY_COLS)
    .setNumberFormat('@').setHorizontalAlignment('center').setWrap(true).setVerticalAlignment('middle');
  for (let d = 0; d < MX.DAYS; d++)
    sheet.getRange(1, mxCol_(d, 4), maxRows, 1).setHorizontalAlignment('left');
  const all = sheet.getRange(1, 1, maxRows, MX.TOTAL_COLS);
  // Màu đánh dấu: Đi trễ cam, Nghỉ chàm, ĐK ✓ xanh lá nhạt
  sheet.setConditionalFormatRules([
    ruleTextEq_(all, 'Đi trễ', '#ffedd5'),
    ruleTextEq_(all, 'Nghỉ', '#e0e7ff'),
    ruleTextEq_(all, MX_MARK, '#dcfce7'),
  ]);
}

// Vạch dọc đậm ngăn cách giữa các cụm ngày (mép trái cột ĐK của mỗi ngày, gồm cả ranh Tên|T2)
function applyDayDividers_(sheet, startRow, numRows) {
  if (numRows < 1) return;
  for (let d = 0; d < MX.DAYS; d++) {
    sheet.getRange(startRow, mxCol_(d, 0), numRows, 1)
      .setBorder(null, true, null, null, null, null, '#475569', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }
}

// Tìm block tuần: ô cột A chứa Date đúng Thứ 2. Trả { titleRow, dataStart, dataEnd } hoặc null.
function findWeekBlock_(sheet, monday) {
  const last = sheet.getLastRow();
  if (last < 1) return null;
  const key = ymd_(monday);
  const colA = sheet.getRange(1, 1, last, 1).getValues();
  for (let i = 0; i < last; i++) {
    const v = colA[i][0];
    if (!(v instanceof Date) || ymd_(v) !== key) continue;
    const titleRow = i + 1;
    const dataStart = titleRow + MX.HEADER_ROWS;
    let r = dataStart;
    while (r <= last) {
      const nv = colA[r - 1][0];
      if (nv instanceof Date || String(nv).trim() === '') break;
      r++;
    }
    return { titleRow: titleRow, dataStart: dataStart, dataEnd: r - 1 };
  }
  return null;
}

// Ghi 3 dòng header của block tuần (dùng cả khi tạo mới lẫn sửa block hỏng)
function writeBlockHeaders_(sheet, monday, titleRow) {
  // Dòng 1: tiêu đề tuần — ô A giữ Date Thứ 2 để code định vị.
  // KHÔNG merge cả dòng: cột 1 đang freeze, merge vắt qua ranh giới freeze sẽ bị Sheets chặn.
  const sunday = new Date(monday.getFullYear(), monday.getMonth(), monday.getDate() + 6);
  sheet.getRange(titleRow, 1).setValue(monday)
    .setNumberFormat('"Tuần "dd/MM' + '" – ' + dmy_(sunday).slice(0, 5) + '"');
  sheet.getRange(titleRow, 1, 1, MX.TOTAL_COLS)
    .setBackground(CONFIG.COLOR_PRIMARY).setFontColor(CONFIG.COLOR_HEADER_TEXT)
    .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('left').setVerticalAlignment('middle');
  sheet.setRowHeight(titleRow, 30);

  // Dòng 2: header ngày (merge 5 cột theo chiều ngang) + dòng 3: header con.
  // Ô "Tên" KHÔNG merge dọc — merge dọc bị insertRowAfter kéo giãn, nuốt dòng thành viên chèn sau đó.
  sheet.getRange(titleRow + 1, 1).setValue('Tên');
  const subRow = [];
  for (let d = 0; d < MX.DAYS; d++) {
    const dd = new Date(monday.getFullYear(), monday.getMonth(), monday.getDate() + d);
    sheet.getRange(titleRow + 1, mxCol_(d, 0), 1, MX.DAY_COLS).merge()
      .setValue(DAY_NAMES[d] + ' ' + pad2_(dd.getDate()) + '/' + pad2_(dd.getMonth() + 1));
    MX.SUB_HEADERS.forEach(function (h) { subRow.push(h); });
  }
  sheet.getRange(titleRow + 2, 2, 1, MX.DAYS * MX.DAY_COLS).setValues([subRow]);
  sheet.getRange(titleRow + 1, 1, 2, MX.TOTAL_COLS)
    .setBackground(CONFIG.COLOR_PRIMARY_LIGHT).setFontWeight('bold').setFontSize(9)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
}

// Tạo block tuần mới ở cuối sheet, pre-fill toàn bộ thành viên
function createWeekBlock_(sheet, monday) {
  const members = getMembers_();
  const last = sheet.getLastRow();
  const titleRow = last === 0 ? 1 : last + 2; // chừa 1 dòng trống ngăn cách
  const dataStart = titleRow + MX.HEADER_ROWS;
  const needRows = dataStart + Math.max(members.length, 1);
  if (sheet.getMaxRows() < needRows) sheet.insertRowsAfter(sheet.getMaxRows(), needRows - sheet.getMaxRows());

  writeBlockHeaders_(sheet, monday, titleRow);

  // Dòng thành viên: pre-fill toàn bộ Members (format text/canh lề đã set cấp cột)
  if (members.length) {
    sheet.getRange(dataStart, 1, members.length, 1)
      .setValues(members.map(function (m) { return [m]; }));
  }
  const totalRows = MX.HEADER_ROWS + Math.max(members.length, 0);
  sheet.getRange(titleRow, 1, totalRows, MX.TOTAL_COLS)
    .setBorder(true, true, true, true, true, true, '#94a3b8', SpreadsheetApp.BorderStyle.SOLID);
  applyDayDividers_(sheet, titleRow, totalRows);

  return { titleRow: titleRow, dataStart: dataStart, dataEnd: dataStart + members.length - 1 };
}

function getOrCreateWeekBlock_(monday) {
  const sheet = getOrCreateMatrixSheet_(monday);
  let block = findWeekBlock_(sheet, monday);
  if (!block) block = createWeekBlock_(sheet, monday);
  else block = repairWeekBlock_(sheet, monday, block);
  return { sheet: sheet, block: block };
}

// Block "xác sống" (lần chạy lỗi trước ghi được ô Date rồi chết): thiếu header
// -> dựng lại header tại chỗ; block trống thì chèn dòng pre-fill thành viên (chèn để không đè block dưới).
// Block lành: chỉ tốn 1 lần đọc kiểm tra.
function repairWeekBlock_(sheet, monday, block) {
  const headerOk = String(sheet.getRange(block.titleRow + 2, 2).getValue()).trim() === MX.SUB_HEADERS[0];
  if (!headerOk) {
    sheet.getRange(block.titleRow, 1, MX.HEADER_ROWS, MX.TOTAL_COLS).breakApart();
    writeBlockHeaders_(sheet, monday, block.titleRow);
    applyDayDividers_(sheet, block.titleRow, MX.HEADER_ROWS);
    if (block.dataEnd < block.dataStart) {
      const members = getMembers_();
      if (members.length) {
        sheet.insertRowsAfter(block.dataStart - 1, members.length);
        sheet.getRange(block.dataStart, 1, members.length, 1)
          .setValues(members.map(function (m) { return [m]; }));
        sheet.getRange(block.dataStart, 1, members.length, MX.TOTAL_COLS)
          .setBorder(true, true, true, true, true, true, '#94a3b8', SpreadsheetApp.BorderStyle.SOLID);
        applyDayDividers_(sheet, block.dataStart, members.length);
        block.dataEnd = block.dataStart + members.length - 1;
      }
    }
  }
  return block;
}

function findMemberRow_(sheet, block, name) {
  if (block.dataEnd < block.dataStart) return 0;
  const vals = sheet.getRange(block.dataStart, 1, block.dataEnd - block.dataStart + 1, 1).getValues();
  for (let i = 0; i < vals.length; i++)
    if (String(vals[i][0]).trim() === name) return block.dataStart + i;
  return 0;
}

// Tìm dòng thành viên trong block; chưa có (người mới thêm vào Members sau khi block đã tạo) -> chèn cuối block
function ensureMemberRow_(sheet, block, name) {
  const found = findMemberRow_(sheet, block, name);
  if (found) return found;
  const anchor = Math.max(block.dataEnd, block.dataStart - 1); // block rỗng -> chèn ngay sau header
  sheet.insertRowAfter(anchor);
  const row = anchor + 1;
  sheet.getRange(row, 1).setValue(name);
  sheet.getRange(row, 1, 1, MX.TOTAL_COLS)
    .setBorder(true, true, true, true, true, true, '#94a3b8', SpreadsheetApp.BorderStyle.SOLID);
  applyDayDividers_(sheet, row, 1);
  block.dataEnd = row;
  return row;
}

// Quét toàn bộ block trong 1 sheet ma trận, gọi cb(name, date, cell) cho từng người × ngày
function scanMatrix_(sheet, cb) {
  const last = sheet.getLastRow();
  if (last < 1) return;
  // Phòng sheet bị thiếu cột (tạo tay / hỏng format) — nới đủ 36 cột trước khi đọc
  if (sheet.getMaxColumns() < MX.TOTAL_COLS)
    sheet.insertColumnsAfter(sheet.getMaxColumns(), MX.TOTAL_COLS - sheet.getMaxColumns());
  const vals = sheet.getRange(1, 1, last, MX.TOTAL_COLS).getValues();
  let r = 0;
  while (r < last) {
    const v = vals[r][0];
    if (!(v instanceof Date)) { r++; continue; }
    const monday = v;
    let dr = r + MX.HEADER_ROWS;
    while (dr < last) {
      const nm = vals[dr][0];
      if (nm instanceof Date || String(nm).trim() === '') break;
      const name = String(nm).trim();
      for (let d = 0; d < MX.DAYS; d++) {
        const base = 1 + d * MX.DAY_COLS; // index 0-based của ô ĐK trong dòng
        cb(name, new Date(monday.getFullYear(), monday.getMonth(), monday.getDate() + d), {
          reg: String(vals[dr][base]).trim() === MX_MARK,
          type: String(vals[dr][base + 1]).trim(),
          expected: normTime_(vals[dr][base + 2]),
          checkin: normTime_(vals[dr][base + 3]),
          reason: String(vals[dr][base + 4] || '').trim(),
        });
      }
      dr++;
    }
    r = dr;
  }
}

/* ============================================================
 * ĐỌC DỮ LIỆU: STATS / HISTORY / ROLLCALL
 * ============================================================ */

function getStats_(month, year) {
  const members = getMembers_();
  const map = {};
  members.forEach(function (m) { map[m] = { name: m, late: 0, off: 0 }; });

  // Ngày trong tháng M có thể nằm ở sheet tháng M-1 / M / M+1 (tuần vắt tháng, gán theo tháng đa số)
  const prevM = month === 1 ? 12 : month - 1;
  const prevY = month === 1 ? year - 1 : year;
  const nextM = month === 12 ? 1 : month + 1;
  const nextY = month === 12 ? year + 1 : year;
  [matrixNameMY_(prevM, prevY), matrixNameMY_(month, year), matrixNameMY_(nextM, nextY)].forEach(function (nm) {
    const sheet = SpreadsheetApp.getActive().getSheetByName(nm);
    if (!sheet) return;
    scanMatrix_(sheet, function (name, date, cell) {
      if (date.getMonth() + 1 !== month || date.getFullYear() !== year) return;
      if (!cell.type) return;
      if (!map[name]) map[name] = { name: name, late: 0, off: 0 };
      if (cell.type === 'Đi trễ') map[name].late += 1;
      else if (cell.type === 'Nghỉ') map[name].off += 1;
    });
  });
  return Object.keys(map).map(function (k) { return map[k]; })
    .sort(function (a, b) { return (b.late + b.off) - (a.late + a.off); });
}

// Thống kê chuyên cần theo tháng — dùng cho sheet Tổng kết.
// Đếm theo KẾT QUẢ THỰC TẾ từng ngày (evalRollcall_), chỉ tính ngày đã qua cho mục Vắng.
function getAttendance_(month, year) {
  const map = {};
  function ent_(n) {
    if (!map[n]) map[n] = { name: n, reg: 0, present: 0, ontime: 0, late: 0, off: 0, absent: 0, walkin: 0 };
    return map[n];
  }
  getMembers_().forEach(function (m) { ent_(m); });
  const today = todayVN_();

  const prevM = month === 1 ? 12 : month - 1;
  const prevY = month === 1 ? year - 1 : year;
  const nextM = month === 12 ? 1 : month + 1;
  const nextY = month === 12 ? year + 1 : year;
  [matrixNameMY_(prevM, prevY), matrixNameMY_(month, year), matrixNameMY_(nextM, nextY)].forEach(function (nm) {
    const sheet = SpreadsheetApp.getActive().getSheetByName(nm);
    if (!sheet) return;
    scanMatrix_(sheet, function (name, date, cell) {
      if (date.getMonth() + 1 !== month || date.getFullYear() !== year) return;
      if (!cell.reg && !cell.type && !cell.checkin) return;
      const e = ent_(name);
      if (cell.reg) e.reg += 1;
      const res = evalRollcall_(cell.reg, cell.checkin,
        cell.type === 'Đi trễ' ? cell.expected : '', cell.type === 'Nghỉ', date.getTime() < today.getTime());
      if (res === 'ontime') { e.present += 1; e.ontime += 1; }
      else if (res === 'late' || res === 'late_ok' || res === 'late_over') { e.present += 1; e.late += 1; }
      else if (res === 'off_ok') e.off += 1;
      else if (res === 'absent') e.absent += 1;
      else if (res === 'walkin') e.walkin += 1;
    });
  });
  return Object.keys(map).map(function (k) { return map[k]; })
    .sort(function (a, b) {
      return (b.absent - a.absent) || (b.late - a.late) || (b.off - a.off) || a.name.localeCompare(b.name);
    });
}

function getHistory_(name, limit) {
  const all = [];
  SpreadsheetApp.getActive().getSheets().forEach(function (sh) {
    if (!MATRIX_RE.test(sh.getName())) return;
    scanMatrix_(sh, function (nm, date, cell) {
      if (nm !== name || !cell.type) return;
      all.push({
        tsRaw: date.getTime(),
        timestamp: '', // ma trận không lưu giờ gửi đơn
        type: cell.type,
        date: dmy_(date),
        arrival: cell.expected,
        present: cell.checkin,
        reason: cell.reason,
        status: '',
        note: '',
      });
    });
  });
  all.sort(function (a, b) { return b.tsRaw - a.tsRaw; });
  return all.slice(0, limit);
}

function getRollcall_(dateStr) {
  const d = dateFromYmd_(dateStr);
  const target = ymd_(d);
  const isPast = dateStr < isoDate_(todayVN_());
  const list = [];
  const sheet = SpreadsheetApp.getActive().getSheetByName(matrixSheetName_(weekStart_(d)));
  if (sheet) {
    scanMatrix_(sheet, function (name, date, cell) {
      if (ymd_(date) !== target) return;
      if (!cell.reg && !cell.type && !cell.checkin) return; // không liên quan ngày này
      const off = cell.type === 'Nghỉ';
      const expected = cell.type === 'Đi trễ' ? cell.expected : '';
      list.push({
        name: name,
        registered: cell.reg,
        checkin: cell.checkin,
        expected: expected,
        off: off,
        result: evalRollcall_(cell.reg, cell.checkin, expected, off, isPast),
      });
    });
  }
  list.sort(function (a, b) {
    if (a.registered !== b.registered) return a.registered ? -1 : 1;
    return a.name.localeCompare(b.name);
  });
  return list;
}

// Đối chiếu: liên kết với đơn Đi trễ / Nghỉ cùng ngày
function evalRollcall_(registered, checkin, expected, off, isPast) {
  if (!registered) return checkin ? 'walkin' : '';
  if (checkin) {
    const c = toMin_(checkin);
    if (c <= CONFIG.PRACTICE_START_HOUR * 60 + CONFIG.PRACTICE_START_MIN) return 'ontime';
    if (expected && TIME_RE.test(expected)) return c <= toMin_(expected) ? 'late_ok' : 'late_over';
    return 'late';
  }
  if (off) return 'off_ok';
  return isPast ? 'absent' : 'pending';
}

/* ============================================================
 * MIGRATE DỮ LIỆU CŨ (chạy 1 lần từ menu)
 * ============================================================ */

function migrateToMatrix() {
  const ui = SpreadsheetApp.getUi();
  const ok = ui.alert(
    'Tạo lại sheet Tháng từ dữ liệu cũ',
    'Toàn bộ sheet "Tháng MM/yyyy" hiện có sẽ bị XÓA và tạo lại từ dữ liệu cũ (sheet log "MM/yyyy" + "Lịch tập").\n' +
    'Dữ liệu nhập tay TRỰC TIẾP trên sheet Tháng (nếu có) sẽ mất. Sheet cũ giữ nguyên.\n\nTiếp tục?',
    ui.ButtonSet.YES_NO
  );
  if (ok !== ui.Button.YES) return;

  // Đã có sheet Tháng = đã chạy thật -> chạy lại sẽ XOÁ mọi đăng ký/đơn/điểm danh ghi sau migrate
  const hasMatrix = SpreadsheetApp.getActive().getSheets().some(function (sh) { return MATRIX_RE.test(sh.getName()); });
  if (hasMatrix) {
    ui.alert('Đã có sheet "Tháng MM/yyyy" — KHÔNG chạy lại migrate (sẽ xoá dữ liệu mới).\n' +
      'Nếu thật sự cần: tự đổi tên/xoá các sheet Tháng trước, rồi chạy lại.');
    return;
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    ui.alert('Server đang bận (có người đang gửi yêu cầu), thử lại sau.');
    return;
  }
  try {
    rebuildMatrixLocked_();
  } finally {
    lock.releaseLock();
  }
}

// Rebuild toàn phần: gom dữ liệu cũ vào bộ nhớ -> xóa sheet Tháng -> ghi lại theo LÔ
// (mỗi block 1 lệnh setValues) — nhanh, idempotent, không thể tạo block trùng/hỏng.
function rebuildMatrixLocked_() {
  const ss = SpreadsheetApp.getActive();
  const members = getMembers_();
  let nReq = 0, nSched = 0;

  // weeks: ymd(Thứ 2) -> { monday, rows: { tên -> mảng 35 ô } }
  const weeks = {};
  function emptyRow_() { const a = []; for (let i = 0; i < MX.DAYS * MX.DAY_COLS; i++) a.push(''); return a; }
  function rowFor_(monday, name) {
    const k = ymd_(monday);
    if (!weeks[k]) {
      const rows = {};
      members.forEach(function (m) { rows[m] = emptyRow_(); });
      weeks[k] = { monday: monday, rows: rows };
    }
    if (!weeks[k].rows[name]) weeks[k].rows[name] = emptyRow_();
    return weeks[k].rows[name];
  }

  // Luôn có block tuần hiện tại + tuần sau (dùng ngay cho đăng ký/điểm danh)
  const thisMon = weekStart_(todayVN_());
  const nextMon = new Date(thisMon.getFullYear(), thisMon.getMonth(), thisMon.getDate() + 7);
  [thisMon, nextMon].forEach(function (m) {
    const k = ymd_(m);
    if (!weeks[k]) {
      const rows = {};
      members.forEach(function (mm) { rows[mm] = emptyRow_(); });
      weeks[k] = { monday: m, rows: rows };
    }
  });

  // 1) Log đơn cũ (sheet "MM/yyyy")
  listMonthSheets_().forEach(function (sh) {
    if (sh.getLastRow() < 2) return;
    const hasStatus = sheetHasStatus_(sh);
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, hasStatus ? COL.STATUS : COL.REASON).getValues();
    rows.forEach(function (r) {
      const name = String(r[COL.NAME - 1]).trim();
      const type = String(r[COL.TYPE - 1]).trim();
      if (!name || CONFIG.TYPES.indexOf(type) === -1) return;
      if (!(r[COL.DATE - 1] instanceof Date)) return;
      if (hasStatus && String(r[COL.STATUS - 1]).trim() === 'Từ chối') return;

      const date = r[COL.DATE - 1];
      const base = ((date.getDay() + 6) % 7) * MX.DAY_COLS;
      const arr = rowFor_(weekStart_(date), name);
      arr[base + 1] = type;
      arr[base + 2] = type === 'Đi trễ' ? formatArrival_(r[COL.ARRIVAL - 1]) : '';
      arr[base + 4] = String(r[COL.REASON - 1] || '');
      const present = normTime_(r[COL.PRESENT - 1]); // cột 7 cũ: chỉ nhận HH:mm (số phút trễ bị bỏ qua)
      if (present) arr[base + 3] = present;
      nReq++;
    });
  });

  // 2) Sheet "Lịch tập" cũ — chạy SAU để giờ điểm danh thắng log
  const ws = ss.getSheetByName(LEGACY_SCHEDULE_SHEET);
  if (ws && ws.getLastRow() >= 2) {
    const rows = ws.getRange(2, 1, ws.getLastRow() - 1, SCHED.FIRST - 1 + SCHED.DAYS * 2).getValues();
    rows.forEach(function (r) {
      const week = r[SCHED.WEEK - 1];
      const name = String(r[SCHED.NAME - 1]).trim();
      if (!(week instanceof Date) || !name) return;
      const monday = weekStart_(week);
      let touched = false;
      for (let d = 0; d < SCHED.DAYS; d++) {
        const reg = String(r[SCHED.FIRST - 1 + d * 2]).trim() === SCHED_MARK;
        const time = normTime_(r[SCHED.FIRST + d * 2]);
        if (reg || time) {
          const arr = rowFor_(monday, name);
          if (reg) arr[d * MX.DAY_COLS] = MX_MARK;
          if (time) arr[d * MX.DAY_COLS + 3] = time;
          touched = true;
        }
      }
      if (touched) nSched++;
    });
  }

  // 3) Xóa toàn bộ sheet Tháng cũ (kể cả sheet hỏng/trùng)
  ss.getSheets().forEach(function (sh) {
    if (MATRIX_RE.test(sh.getName())) ss.deleteSheet(sh);
  });

  // 4) Gom tuần theo sheet tháng, ghi lại — mỗi block 1 lệnh setValues
  const sorted = Object.keys(weeks).map(function (k) { return weeks[k]; })
    .sort(function (a, b) { return a.monday - b.monday; });
  const byName = {};
  sorted.forEach(function (w) {
    const nm = matrixSheetName_(w.monday);
    (byName[nm] = byName[nm] || []).push(w);
  });

  Object.keys(byName).forEach(function (nm) {
    const sh = ss.insertSheet(nm);
    formatMatrixSheet_(sh);
    let titleRow = 1;
    byName[nm].forEach(function (w) {
      const order = members.slice();
      Object.keys(w.rows).forEach(function (n) { if (order.indexOf(n) === -1) order.push(n); });
      const dataStart = titleRow + MX.HEADER_ROWS;
      const need = dataStart + order.length;
      if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());

      writeBlockHeaders_(sh, w.monday, titleRow);
      const data = order.map(function (n) { return [n].concat(w.rows[n] || emptyRow_()); });
      sh.getRange(dataStart, 1, data.length, MX.TOTAL_COLS).setValues(data);
      sh.getRange(titleRow, 1, MX.HEADER_ROWS + data.length, MX.TOTAL_COLS)
        .setBorder(true, true, true, true, true, true, '#94a3b8', SpreadsheetApp.BorderStyle.SOLID);
      applyDayDividers_(sh, titleRow, MX.HEADER_ROWS + data.length);
      titleRow = dataStart + data.length + 1; // 1 dòng trống ngăn cách
    });
    sh.autoResizeColumn(1); // cột Tên tự vừa tên dài nhất
    if (sh.getColumnWidth(1) < 110) sh.setColumnWidth(1, 110);
  });

  try { refreshSummary_(); } catch (ignore) {}
  const msg = '✅ Đã tạo lại sheet Tháng: ' + nReq + ' đơn + ' + nSched + ' dòng lịch tuần. Sheet cũ giữ nguyên làm backup.';
  try { SpreadsheetApp.getUi().alert(msg); }
  catch (e) { try { ss.toast(msg, 'LTS', 8); } catch (e2) { Logger.log(msg); } }
}

// Dọn sheet cũ sau khi đã migrate xong và kiểm tra dữ liệu ổn:
// xóa các sheet log "MM/yyyy" + sheet "Lịch tập". Sheet Tháng (ma trận) giữ nguyên.
function cleanupLegacySheets() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();

  const hasMatrix = ss.getSheets().some(function (sh) { return MATRIX_RE.test(sh.getName()); });
  if (!hasMatrix) {
    ui.alert('Chưa có sheet "Tháng MM/yyyy" nào — hãy chạy "Chuyển dữ liệu cũ → sheet Tháng" trước khi dọn.');
    return;
  }
  const targets = ss.getSheets().filter(function (sh) {
    return MONTH_RE.test(sh.getName()) || sh.getName() === LEGACY_SCHEDULE_SHEET;
  });
  if (!targets.length) {
    ui.alert('Không còn sheet cũ nào để dọn.');
    return;
  }
  const names = targets.map(function (sh) { return sh.getName(); }).join(', ');
  const ok = ui.alert(
    'Xóa sheet dữ liệu cũ',
    'Sẽ XÓA VĨNH VIỄN các sheet: ' + names + '.\n\n' +
    'Lưu ý: sau khi xóa, chức năng "Chuyển dữ liệu cũ → sheet Tháng" không còn nguồn để tạo lại — ' +
    'chỉ xóa khi đã kiểm tra sheet Tháng đầy đủ dữ liệu.\n\nTiếp tục?',
    ui.ButtonSet.YES_NO
  );
  if (ok !== ui.Button.YES) return;

  targets.forEach(function (sh) { ss.deleteSheet(sh); });
  ui.alert('✅ Đã xóa ' + targets.length + ' sheet cũ (' + names + ').');
}

// Sheet log cũ "MM/yyyy" — chỉ dùng để migrate
function listMonthSheets_() {
  return SpreadsheetApp.getActive().getSheets().filter(function (sh) {
    return MONTH_RE.test(sh.getName());
  });
}
function sheetHasStatus_(sheet) {
  return sheet.getLastColumn() >= COL.STATUS &&
    String(sheet.getRange(1, COL.STATUS).getValue()).trim() === 'Trạng thái';
}

/* ============================================================
 * ĐỒNG BỘ MEMBERS + ĐÁNH DẤU NGÀY ĐÃ QUA
 * ============================================================ */

// Menu: đồng bộ Members với các block tuần HIỆN TẠI + TƯƠNG LAI.
// - Người bị xóa khỏi Members: xóa dòng nếu dòng chưa có dữ liệu; dòng có dữ liệu giữ lại (lịch sử).
// - Người mới thêm vào Members: chèn dòng vào cuối block.
// Block các tuần ĐÃ QUA giữ nguyên tuyệt đối.
function syncMembers() {
  const ui = SpreadsheetApp.getUi();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { ui.alert('Server đang bận, thử lại sau.'); return; }
  let res;
  try { res = syncMembersLocked_(); } finally { lock.releaseLock(); }
  let msg = '✅ Đồng bộ xong (tuần hiện tại + tương lai): thêm ' + res.added + ' dòng, xóa ' + res.removed + ' dòng.';
  if (res.kept) msg += '\nGiữ lại ' + res.kept + ' dòng của người đã xóa vì dòng còn dữ liệu.';
  ui.alert(msg);
}

function syncMembersLocked_() {
  const ss = SpreadsheetApp.getActive();
  const members = getMembers_();
  const thisMon = weekStart_(todayVN_());
  let added = 0, removed = 0, kept = 0;

  ss.getSheets().forEach(function (sh) {
    if (!MATRIX_RE.test(sh.getName())) return;
    const last = sh.getLastRow();
    if (last < 1) return;
    // Gom Monday của các block thuộc tuần hiện tại/tương lai
    const mondays = [];
    sh.getRange(1, 1, last, 1).getValues().forEach(function (r) {
      const v = r[0];
      if (v instanceof Date && weekStart_(v).getTime() >= thisMon.getTime()) mondays.push(weekStart_(v));
    });

    mondays.forEach(function (monday) {
      let block = findWeekBlock_(sh, monday);
      if (!block) return;

      // 1) Xóa dòng người không còn trong Members (từ dưới lên để index không lệch)
      if (block.dataEnd >= block.dataStart) {
        const n = block.dataEnd - block.dataStart + 1;
        const rows = sh.getRange(block.dataStart, 1, n, MX.TOTAL_COLS).getValues();
        for (let i = n - 1; i >= 0; i--) {
          const name = String(rows[i][0]).trim();
          if (!name || members.indexOf(name) !== -1) continue;
          const hasData = rows[i].slice(1).some(function (c) { return String(c).trim() !== ''; });
          if (hasData) { kept++; continue; }
          sh.deleteRow(block.dataStart + i);
          removed++;
        }
        block = findWeekBlock_(sh, monday);
        if (!block) return;
      }

      // 2) Thêm người mới chưa có dòng
      members.forEach(function (m) {
        if (!findMemberRow_(sh, block, m)) { ensureMemberRow_(sh, block, m); added++; }
      });
    });
  });
  return { added: added, removed: removed, kept: kept };
}

// Menu / trigger: rà mọi ngày ĐÃ QUA trong các sheet Tháng:
// - Đăng ký (✓) hoặc báo Đi trễ mà KHÔNG có giờ điểm danh, không đơn Nghỉ -> ghi ✗ nền đỏ vào ô Giờ đến.
// - Không đăng ký + không đơn + không đến -> tô xám cả cụm 5 ô của ngày đó.
// Idempotent: chạy lại bao nhiêu lần cũng được; điểm danh bù sẽ đè ✗ và lần chạy sau tự xoá nền.
function markPastDays() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  let n;
  try { n = markPastDaysLocked_(); } finally { lock.releaseLock(); }
  const msg = '✅ Đã rà ngày đã qua: thêm ' + n + ' dấu vắng (' + ABSENT_MARK + ').';
  try { SpreadsheetApp.getUi().alert(msg); }
  catch (e) { try { SpreadsheetApp.getActive().toast(msg, 'LTS', 5); } catch (e2) { Logger.log(msg); } }
}

function markPastDaysLocked_() {
  const ss = SpreadsheetApp.getActive();
  const today = todayVN_();
  let marked = 0;

  ss.getSheets().forEach(function (sh) {
    if (!MATRIX_RE.test(sh.getName())) return;
    const last = sh.getLastRow();
    if (last < 1) return;
    if (sh.getMaxColumns() < MX.TOTAL_COLS)
      sh.insertColumnsAfter(sh.getMaxColumns(), MX.TOTAL_COLS - sh.getMaxColumns());
    const vals = sh.getRange(1, 1, last, MX.TOTAL_COLS).getValues();

    let r = 0;
    while (r < last) {
      const v = vals[r][0];
      if (!(v instanceof Date)) { r++; continue; }
      const monday = v;
      const dataStart = r + MX.HEADER_ROWS;
      let dr = dataStart;
      while (dr < last) {
        const nm = vals[dr][0];
        if (nm instanceof Date || String(nm).trim() === '') break;
        dr++;
      }
      const nRows = dr - dataStart;

      // Chỉ block có ít nhất 1 ngày đã qua
      if (nRows > 0 && weekStart_(monday).getTime() < today.getTime()) {
        const bg = [];
        const newMarks = [];
        for (let i = 0; i < nRows; i++) {
          const rowVals = vals[dataStart + i];
          const rowBg = [];
          for (let c = 0; c < MX.DAYS * MX.DAY_COLS; c++) rowBg.push(null);
          for (let d = 0; d < MX.DAYS; d++) {
            const date = new Date(monday.getFullYear(), monday.getMonth(), monday.getDate() + d);
            if (date.getTime() >= today.getTime()) continue; // hôm nay chưa kết thúc -> chưa đánh
            const base = d * MX.DAY_COLS;
            const reg = String(rowVals[1 + base]).trim() === MX_MARK;
            const type = String(rowVals[1 + base + 1]).trim();
            const rawCheckin = String(rowVals[1 + base + 3]).trim();
            const checkin = normTime_(rowVals[1 + base + 3]);
            const off = type === 'Nghỉ';
            if ((reg || type === 'Đi trễ') && !checkin && !off) {
              rowBg[base + 3] = COLOR_ABSENT_BG;
              if (rawCheckin !== ABSENT_MARK) newMarks.push({ row: dataStart + 1 + i, col: 2 + base + 3 });
            } else if (!reg && !type && !checkin && rawCheckin !== ABSENT_MARK) {
              for (let f = 0; f < MX.DAY_COLS; f++) rowBg[base + f] = COLOR_IDLE_BG;
            }
          }
          bg.push(rowBg);
        }
        sh.getRange(dataStart + 1, 2, nRows, MX.DAYS * MX.DAY_COLS).setBackgrounds(bg);
        newMarks.forEach(function (m) {
          sh.getRange(m.row, m.col).setNumberFormat('@').setValue(ABSENT_MARK);
          marked++;
        });
      }
      r = dr;
    }
  });
  return marked;
}

// Menu: cài trigger chạy markPastDays mỗi ngày ~1h sáng (đánh dấu ngày hôm trước)
function enableDailyMark() {
  const ui = SpreadsheetApp.getUi();
  const has = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === 'markPastDays';
  });
  if (has) { ui.alert('Trigger tự đánh dấu đã bật từ trước.'); return; }
  ScriptApp.newTrigger('markPastDays').timeBased().everyDays(1).atHour(1).create();
  ui.alert('✅ Đã bật: tự động đánh dấu ngày đã qua mỗi ngày (~1h sáng).');
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

/* ============================================================
 * NOTIFY (tuỳ chọn)
 * ============================================================ */

function notify_(req) {
  const props = PropertiesService.getScriptProperties();
  const email = props.getProperty('MANAGER_EMAIL');
  const tgToken = props.getProperty('TELEGRAM_BOT_TOKEN');
  const tgChat = props.getProperty('TELEGRAM_CHAT_ID');

  const lateStr = req.type === 'Đi trễ'
    ? ('\n• Giờ đến dự kiến: ' + req.arrivalTime) : '';
  const text =
    '🔔 Yêu cầu mới — LTS\n' +
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
 * NHẬT KÝ (append-only) + BLOCK TUẦN NẰM SAI SHEET
 * ============================================================ */

function getLogSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(LOG_SHEET);
    const headers = ['Thời gian (VN)', 'Thao tác', 'Tên', 'Chi tiết', 'Kết quả'];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    styleHeader_(sh.getRange(1, 1, 1, headers.length));
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@');
  }
  return sh;
}

function logPost_(body, out) {
  try {
    let detail = '';
    if (body.action === 'registerSchedule') {
      const days = Array.isArray(body.days) ? body.days : [];
      const picked = DAY_NAMES.filter(function (_, i) { return days[i]; });
      detail = 'tuần ' + String(body.week || '?') + ': ' + (picked.join(', ') || '(không ngày nào)');
    } else if (body.action === 'checkin') {
      detail = String(body.date || '') + ' ' + String(body.time || '(xoá)');
    } else {
      detail = [body.type, body.date, body.arrivalTime, body.reason].filter(function (x) { return x; }).join(' | ');
    }
    let result = '';
    try { const r = JSON.parse(out.getContent()); result = r.status + ': ' + (r.message || ''); }
    catch (ignore) { result = '?'; }
    getLogSheet_().appendRow([
      Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss'),
      String(body.action || 'request'), String(body.name || ''), detail, result,
    ]);
  } catch (ignore) {}
}

// Mọi block tuần trong mọi sheet Tháng: [{ sheet, monday, block, misplaced }]
// misplaced = block nằm ở sheet khác sheet chuẩn (VD tạo bởi bản code cũ tính tháng theo Thứ 2)
function listWeekBlocks_() {
  const out = [];
  SpreadsheetApp.getActive().getSheets().forEach(function (sh) {
    if (!MATRIX_RE.test(sh.getName())) return;
    const last = sh.getLastRow();
    if (last < 1) return;
    sh.getRange(1, 1, last, 1).getValues().forEach(function (r) {
      const v = r[0];
      if (!(v instanceof Date)) return;
      const monday = weekStart_(v);
      out.push({
        sheet: sh, monday: monday, block: findWeekBlock_(sh, monday),
        misplaced: sh.getName() !== matrixSheetName_(monday),
      });
    });
  });
  return out;
}

function countBlockData_(sheet, block) {
  let n = 0;
  if (!block || block.dataEnd < block.dataStart) return 0;
  sheet.getRange(block.dataStart, 2, block.dataEnd - block.dataStart + 1, MX.DAYS * MX.DAY_COLS).getValues()
    .forEach(function (row) { row.forEach(function (c) { if (String(c).trim() !== '' && c !== ABSENT_MARK) n++; }); });
  return n;
}

// Menu: báo các block nằm sai sheet / trùng tuần (chỉ đọc)
function auditWeekBlocks() {
  const seen = {};
  const lines = [];
  listWeekBlocks_().forEach(function (b) {
    const k = ymd_(b.monday);
    const tag = 'Tuần ' + dmy_(b.monday).slice(0, 5) + ' ở "' + b.sheet.getName() + '"';
    if (b.misplaced) lines.push('⚠️ ' + tag + ' — đúng ra ở "' + matrixSheetName_(b.monday) + '" (' + countBlockData_(b.sheet, b.block) + ' ô có dữ liệu)');
    if (seen[k]) lines.push('⚠️ ' + tag + ' — TRÙNG với block ở "' + seen[k] + '"');
    seen[k] = b.sheet.getName();
  });
  SpreadsheetApp.getUi().alert(lines.length ? lines.join('\n') : '✅ Không có block tuần nào nằm sai sheet hay bị trùng.');
}

// Menu: chép dữ liệu từ block nằm sai sheet sang block chuẩn — CHỈ điền ô đang trống, không ghi đè,
// không xoá block cũ (xoá tay sau khi kiểm tra). Chạy lại nhiều lần vẫn an toàn.
function mergeMisplacedBlocks() {
  const ui = SpreadsheetApp.getUi();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { ui.alert('Server đang bận, thử lại sau.'); return; }
  let filled = 0, blocks = 0;
  try {
    listWeekBlocks_().forEach(function (b) {
      if (!b.misplaced || !b.block || b.block.dataEnd < b.block.dataStart) return;
      const src = b.sheet.getRange(b.block.dataStart, 1, b.block.dataEnd - b.block.dataStart + 1, MX.TOTAL_COLS).getValues();
      const loc = getOrCreateWeekBlock_(b.monday);
      blocks++;
      src.forEach(function (row) {
        const name = String(row[0]).trim();
        if (!name) return;
        const r = ensureMemberRow_(loc.sheet, loc.block, name);
        const range = loc.sheet.getRange(r, 2, 1, MX.DAYS * MX.DAY_COLS);
        const cur = range.getValues()[0];
        let changed = false;
        for (let c = 0; c < cur.length; c++) {
          const v = row[c + 1];
          if (String(cur[c]).trim() === '' && String(v).trim() !== '' && v !== ABSENT_MARK) { cur[c] = v; changed = true; filled++; }
        }
        if (changed) {
          range.setValues([cur]);
          getLogSheet_().appendRow([
            Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss'), 'backfill', name,
            'tuần ' + isoDate_(b.monday) + ' từ "' + b.sheet.getName() + '" -> "' + loc.sheet.getName() + '"', 'OK',
          ]);
        }
      });
    });
  } finally { lock.releaseLock(); }
  try { refreshSummary_(); } catch (ignore) {}
  ui.alert('✅ Đã gộp ' + blocks + ' block sai sheet, điền ' + filled + ' ô trống.\n' +
    'Block cũ vẫn giữ nguyên — kiểm tra xong có thể xoá tay. Chi tiết trong sheet "' + LOG_SHEET + '".');
}

/* ============================================================
 * UTILS
 * ============================================================ */

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function pad2_(n) { return ('0' + n).slice(-2); }
function ymd_(d) { return d.getFullYear() + '-' + (d.getMonth() + 1) + '-' + d.getDate(); }
function isoDate_(d) { return d.getFullYear() + '-' + pad2_(d.getMonth() + 1) + '-' + pad2_(d.getDate()); }
function dmy_(d) { return pad2_(d.getDate()) + '/' + pad2_(d.getMonth() + 1) + '/' + d.getFullYear(); }
function dateFromYmd_(s) { const p = s.split('-'); return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2])); }
// Hôm nay theo giờ VN (không phụ thuộc timezone server của Apps Script)
function todayVN_() { return dateFromYmd_(Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd')); }
// Thứ 2 đầu tuần chứa d
function weekStart_(d) { const x = new Date(d.getTime()); x.setDate(x.getDate() - ((x.getDay() + 6) % 7)); x.setHours(0, 0, 0, 0); return x; }
// Thứ 2 tuần KẾ TIẾP — tuần duy nhất được phép đăng ký lịch
function nextWeekStart_() { const w = weekStart_(todayVN_()); w.setDate(w.getDate() + 7); return w; }
function toMin_(hhmm) { const p = hhmm.split(':'); return Number(p[0]) * 60 + Number(p[1]); }
function formatArrival_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, CONFIG.TIMEZONE, 'HH:mm');
  return String(v || '').trim();
}
// Chuẩn hoá giờ: chỉ nhận HH:mm (hoặc Date) — giá trị khác (VD số phút trễ cũ) trả ''
function normTime_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, CONFIG.TIMEZONE, 'HH:mm');
  const s = String(v || '').trim();
  return TIME_RE.test(s) ? s : '';
}

/* ============================================================
 * SETUP
 * ============================================================ */

function setup() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(CONFIG.TIMEZONE);

  setupMembersSheet_(ss);
  // Sheet Tháng hiện tại + block tuần này (tạo sau Members để pre-fill đủ người)
  getOrCreateWeekBlock_(weekStart_(todayVN_()));
  setupSummarySheet_(ss);
  getLogSheet_();

  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('Trang tính1');
  if (def && ss.getSheets().length > 3) { try { ss.deleteSheet(def); } catch (e) {} }

  const msg = '✅ Setup xong! Sheet Tháng (ma trận tuần) + Members + Tổng kết đã sẵn sàng.';
  try { SpreadsheetApp.getUi().alert(msg); }
  catch (e) { try { ss.toast(msg, 'LTS', 5); } catch (e2) { Logger.log(msg); } }
}

function styleHeader_(range, bg) {
  range.setBackground(bg || CONFIG.COLOR_PRIMARY).setFontColor(CONFIG.COLOR_HEADER_TEXT)
    .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center').setVerticalAlignment('middle');
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
  const COLS = SUMMARY_HEADERS.length;
  if (sheet.getMaxColumns() < COLS) sheet.insertColumnsAfter(sheet.getMaxColumns(), COLS - sheet.getMaxColumns());

  sheet.setFrozenColumns(0); // không được merge vắt qua cột freeze — bỏ freeze cột (nếu có) trước
  sheet.getRange(1, 1, 3, sheet.getMaxColumns()).breakApart().setBackground(null); // gỡ merge/nền của layout cũ (4 hoặc 5 cột)
  sheet.getRange('A1').setValue('📊 THỐNG KÊ THEO THÁNG');
  sheet.getRange(1, 1, 1, COLS).merge().setBackground(CONFIG.COLOR_PRIMARY).setFontColor(CONFIG.COLOR_HEADER_TEXT)
    .setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
  sheet.setRowHeight(1, 40);

  sheet.getRange('A2').setValue('Tháng:').setFontWeight('bold');
  sheet.getRange('B2').setValue(now.getMonth() + 1);
  sheet.getRange('C2').setValue('Năm:').setFontWeight('bold');
  sheet.getRange('D2').setValue(now.getFullYear());
  sheet.getRange('B2').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['1','2','3','4','5','6','7','8','9','10','11','12'], true).build());
  sheet.getRange('B2:D2').setHorizontalAlignment('center');
  sheet.getRange(2, 1, 1, COLS).setBackground(CONFIG.COLOR_PRIMARY_LIGHT);

  // Dòng chú thích cách đọc bảng
  sheet.getRange(3, 1, 1, COLS).merge()
    .setValue('Có mặt = có giờ điểm danh trong số buổi ĐK · Vắng KP = ĐK (hoặc báo trễ) nhưng không đến, không đơn nghỉ (chỉ tính ngày đã qua) · Vãng lai = đến tập nhưng không ĐK · Chuyên cần = Có mặt / ĐK tập.')
    .setFontStyle('italic').setFontSize(9).setFontColor('#64748b').setWrap(true)
    .setVerticalAlignment('middle');
  sheet.setRowHeight(3, 34);

  sheet.getRange(4, 1, 1, COLS).setValues([SUMMARY_HEADERS]);
  styleHeader_(sheet.getRange(4, 1, 1, COLS));
  sheet.setFrozenRows(4);

  const widths = [200, 70, 75, 80, 70, 90, 80, 80, 95];
  widths.forEach(function (w, i) { sheet.setColumnWidth(i + 1, w); });

  sheet.getBandings().forEach(function (b) { b.remove(); });
  sheet.getRange(5, 1, SUMMARY_ROWS, COLS).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
  sheet.getRange(4, 1, SUMMARY_ROWS + 1, COLS).setBorder(true, true, true, true, true, true, '#e2e8f0', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getCharts().forEach(function (c) { sheet.removeChart(c); });

  refreshSummary_(sheet);
}

function refreshSummary_(sheet) {
  if (!sheet) sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SUMMARY_SHEET);
  if (!sheet) return;
  const month = Number(sheet.getRange('B2').getValue()) || (new Date().getMonth() + 1);
  const year = Number(sheet.getRange('D2').getValue()) || new Date().getFullYear();
  const stats = getAttendance_(month, year);

  const FIRST = 5, COLS = SUMMARY_HEADERS.length;
  sheet.getRange(FIRST, 1, SUMMARY_ROWS, COLS).clearContent().setFontWeight('normal');
  const rows = stats.slice(0, SUMMARY_ROWS - 1).map(function (s) {
    return [s.name, s.reg, s.present, s.ontime, s.late, s.off, s.absent, s.walkin,
      s.reg ? Math.round(100 * s.present / s.reg) + '%' : '—'];
  });
  if (rows.length) {
    sheet.getRange(FIRST, 1, rows.length, COLS).setValues(rows);
    const t = stats.reduce(function (a, s) {
      a.reg += s.reg; a.present += s.present; a.ontime += s.ontime; a.late += s.late;
      a.off += s.off; a.absent += s.absent; a.walkin += s.walkin; return a;
    }, { reg: 0, present: 0, ontime: 0, late: 0, off: 0, absent: 0, walkin: 0 });
    sheet.getRange(FIRST + rows.length, 1, 1, COLS).setValues([[
      'TỔNG', t.reg, t.present, t.ontime, t.late, t.off, t.absent, t.walkin,
      t.reg ? Math.round(100 * t.present / t.reg) + '%' : '—',
    ]]).setFontWeight('bold');
  }
  sheet.getRange(FIRST, 2, SUMMARY_ROWS, COLS - 1).setHorizontalAlignment('center');
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
    .addItem('Đồng bộ Members → sheet Tháng', 'syncMembers')
    .addItem('Đánh dấu vắng các ngày đã qua', 'markPastDays')
    .addItem('Bật tự đánh dấu hằng ngày (1h sáng)', 'enableDailyMark')
    .addItem('Kiểm tra block tuần nằm sai sheet', 'auditWeekBlocks')
    .addItem('Gộp block tuần nằm sai sheet (backfill)', 'mergeMisplacedBlocks')
    .addItem('Chuyển dữ liệu cũ → sheet Tháng', 'migrateToMatrix')
    .addItem('Xóa sheet dữ liệu cũ (sau khi đã kiểm tra)', 'cleanupLegacySheets')
    .addSeparator()
    .addItem('Đặt PIN quản lý + link Sheet', 'configManager')
    .addItem('Cấu hình thông báo (email/Telegram)', 'configNotify')
    .addToUi();
}

function ruleTextEq_(range, text, bg) {
  return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(text).setBackground(bg).setRanges([range]).build();
}