// Test logic Code.gs bằng Sheet giả lập (không gọi Google).
// Chạy: node --test tests/   (nên chạy thêm với TZ=Pacific/Kiritimati để bắt lỗi lệch múi giờ)
const test = require("node:test");
const assert = require("node:assert");
const fs = require("node:fs");
const path = require("node:path");

const SRC = fs.readFileSync(path.join(__dirname, "../apps-script/Code.gs"), "utf8");

function fmt(d, tz, pat) {
  const parts = new Intl.DateTimeFormat("en-GB", {
    timeZone: tz, year: "numeric", month: "2-digit", day: "2-digit",
    hour: "2-digit", minute: "2-digit", second: "2-digit", hourCycle: "h23",
  }).formatToParts(d);
  const g = (t) => parts.find((x) => x.type === t).value;
  return pat.replace("yyyy", g("year")).replace("MM", g("month")).replace("dd", g("day"))
    .replace("HH", g("hour")).replace("mm", g("minute")).replace("ss", g("second"));
}

// Mọi method định dạng không quan trọng -> trả về chính stub
const stub = new Proxy({}, { get: (t, k) => (k === "getValues" ? () => [[]] : () => stub) });

class Sheet {
  constructor(n) { this.n = n; this.d = []; }
  getName() { return this.n; }
  getLastRow() { return this.d.length; }
  getMaxRows() { return 1000; }
  getRange(r, c, nr = 1, nc = 1) {
    if (typeof r === "string") return stub;
    const s = this;
    const o = {
      getValues() {
        const out = [];
        for (let i = 0; i < nr; i++) {
          const row = s.d[r - 1 + i] || [];
          out.push(Array.from({ length: nc }, (_, j) => (row[c - 1 + j] === undefined ? "" : row[c - 1 + j])));
        }
        return out;
      },
      setValues(v) {
        v.forEach((row, i) => {
          s.d[r - 1 + i] = s.d[r - 1 + i] || [];
          row.forEach((x, j) => (s.d[r - 1 + i][c - 1 + j] = x));
        });
        return this;
      },
      setValue(x) { return this.setValues([[x]]); },
    };
    return new Proxy(o, { get: (t, k) => (k in t ? t[k] : () => stub) });
  }
  appendRow(v) { this.d.push(v.slice()); }
  deleteRow(r) { this.d.splice(r - 1, 1); }
}

function load(members) {
  let NOW = new Date();
  const RealDate = Date;
  class FakeDate extends RealDate {
    constructor(...a) { if (a.length === 0) super(NOW.getTime()); else super(...a); }
    static now() { return NOW.getTime(); }
  }
  const sheets = {};
  const wrap = (sh) => new Proxy(sh, {
    get: (t, k) => (k in t ? (typeof t[k] === "function" ? t[k].bind(t) : t[k])
      : k === "getBandings" ? () => [] : k === "getFilter" ? () => null : () => stub),
  });
  const ss = {
    getSheetByName: (n) => sheets[n] || null,
    insertSheet: (n) => (sheets[n] = wrap(new Sheet(n))),
    getSheets: () => Object.values(sheets),
  };
  const env = {
    Date: FakeDate,
    SpreadsheetApp: { getActive: () => ss, flush() {}, newDataValidation: () => stub, newConditionalFormatRule: () => stub },
    Utilities: { formatDate: fmt },
    PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => ({ MANAGER_PIN: "1" })[k] || null }) },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) },
    ContentService: { createTextOutput: (s) => ({ setMimeType: () => JSON.parse(s) }), MimeType: { JSON: 1 } },
    MailApp: { sendEmail() {} },
    UrlFetchApp: { fetch() {} },
  };
  const api = new Function(...Object.keys(env), SRC + "\nreturn { doGet, doPost };")(...Object.values(env));
  const m = ss.insertSheet("Members");
  m.d = [["Tên"]].concat(members.map((x) => [x]));
  // sheet tháng thật luôn có dòng tiêu đề (formatMonthSheet_)
  ss.insertSheet("09/2026").d.push(["h"]);
  ss.insertSheet("10/2026").d.push(["h"]);
  return {
    sheets,
    at: (iso) => (NOW = new RealDate(iso)),
    get: (p) => api.doGet({ parameter: p }),
    post: (b) => api.doPost({ postData: { contents: JSON.stringify(b) } }),
  };
}

const MEMBERS = ["Ebi", "Yến Thư", "Việt Anh"];

test("Chủ nhật tối đăng ký cho tuần bắt đầu Thứ 2 ngày mai", () => {
  const s = load(MEMBERS);
  s.at("2026-09-27T21:30:00+07:00");
  assert.strictEqual(s.get({ action: "schedule", name: "Ebi" }).week, "2026-09-28");
  s.at("2026-09-28T12:00:00+07:00");
  assert.strictEqual(s.get({ action: "schedule", name: "Ebi" }).week, "2026-10-05");
});

test("Lưu 2 lần = ghi đè 1 dòng, trả đúng ngày đã lưu, có log", () => {
  const s = load(MEMBERS);
  s.at("2026-09-27T21:30:00+07:00");
  const days = [true, false, true, false, false, false, false];
  let r = s.post({ action: "registerSchedule", name: "Ebi", week: "2026-09-28", days });
  assert.strictEqual(r.status, "success");
  r = s.post({ action: "registerSchedule", name: "Ebi", week: "2026-09-28", days: [true, false, true, false, true, false, false] });
  assert.deepStrictEqual(r.days, [true, false, true, false, true, false, false]);
  assert.strictEqual(s.sheets["Lịch tuần"].d.length - 1, 1);
  assert.strictEqual(s.sheets["Log lịch"].d.length - 1, 2);
  const g = s.get({ action: "schedule", name: "Ebi" });
  assert.ok(g.registered);
});

test("Sang Thứ 2 thì tuần đó khóa, vẫn ghi log", () => {
  const s = load(MEMBERS);
  s.at("2026-09-28T00:05:00+07:00");
  const r = s.post({ action: "registerSchedule", name: "Yến Thư", week: "2026-09-28", days: [true, true, false, false, false, false, false] });
  assert.strictEqual(r.status, "error");
  assert.match(r.message, /đã khóa/);
  assert.match(s.sheets["Log lịch"].d[1][4], /TỪ CHỐI/);
});

test("Tên ngoài danh sách bị từ chối", () => {
  const s = load(MEMBERS);
  s.at("2026-09-27T21:30:00+07:00");
  const r = s.post({ action: "registerSchedule", name: "Người lạ", week: "2026-09-28", days: Array(7).fill(true) });
  assert.strictEqual(r.status, "error");
});

test("Đơn nghỉ: trưa hôm trước OK, cùng ngày bị chặn theo giờ VN", () => {
  const s = load(MEMBERS);
  s.at("2026-09-27T12:00:00+07:00");
  assert.strictEqual(s.post({ name: "Việt Anh", type: "Nghỉ", date: "2026-09-28", reason: "x" }).status, "success");
  s.at("2026-09-28T00:05:00+07:00");
  const r = s.post({ name: "Yến Thư", type: "Nghỉ", date: "2026-09-28", reason: "x" });
  assert.strictEqual(r.status, "error");
  assert.match(r.message, /hôm nay 28\/09/);
});

test("Điểm danh: đúng/ trễ/ ngoài lịch/ nghỉ có phép", () => {
  const s = load(MEMBERS);
  s.at("2026-09-27T12:00:00+07:00");
  s.post({ action: "registerSchedule", name: "Ebi", week: "2026-09-28", days: [true, false, false, false, false, false, false] });
  s.post({ name: "Việt Anh", type: "Nghỉ", date: "2026-09-28", reason: "x" });
  s.at("2026-09-28T18:40:00+07:00");
  assert.strictEqual(s.post({ action: "checkin", pin: "1", name: "Ebi", date: "2026-09-28", time: "now" }).status, "success");
  s.post({ action: "checkin", pin: "1", name: "Yến Thư", date: "2026-09-28", time: "18:20" });
  const list = s.get({ action: "rollcall", pin: "1", date: "2026-09-28" }).list;
  const by = Object.fromEntries(list.map((x) => [x.name, x]));
  assert.strictEqual(by["Ebi"].result, "late");
  assert.strictEqual(by["Ebi"].checkin, "18:40");
  assert.strictEqual(by["Yến Thư"].result, "walkin");
  assert.strictEqual(by["Việt Anh"].result, "off_ok");
});

test("Rollcall/checkin cần PIN", () => {
  const s = load(MEMBERS);
  assert.strictEqual(s.get({ action: "rollcall", pin: "sai", date: "2026-09-28" }).status, "error");
  assert.strictEqual(s.post({ action: "checkin", pin: "sai", name: "Ebi", date: "2026-09-28", time: "18:00" }).status, "error");
});
