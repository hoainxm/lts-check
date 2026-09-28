// Test logic Code.gs bằng Sheet giả lập (không gọi Google).
// Chạy: TZ=Asia/Ho_Chi_Minh node --test tests/backend.test.cjs
// (project Apps Script đặt múi giờ VN trong appsscript.json -> test cùng múi giờ đó)
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
  getLastRow() {
    for (let i = this.d.length - 1; i >= 0; i--)
      if ((this.d[i] || []).some((c) => c !== undefined && c !== "")) return i + 1;
    return 0;
  }
  getMaxRows() { return 1000; }
  getMaxColumns() { return 40; }
  getLastColumn() { return 40; }
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
      getValue() { return this.getValues()[0][0]; },
      setValues(v) {
        v.forEach((row, i) => {
          s.d[r - 1 + i] = s.d[r - 1 + i] || [];
          row.forEach((x, j) => (s.d[r - 1 + i][c - 1 + j] = x));
        });
        return p;
      },
      setValue(x) { return this.setValues([[x]]); },
    };
    const p = new Proxy(o, { get: (t, k) => (k in t ? t[k] : () => p) });
    return p;
  }
  appendRow(v) { this.d[this.getLastRow()] = v.slice(); }
  deleteRow(r) { this.d.splice(r - 1, 1); }
  insertRowAfter(r) { this.d.splice(r, 0, []); }
  insertRowsAfter(r, n) { this.d.splice(r, 0, ...Array.from({ length: n }, () => [])); }
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
      : k === "getBandings" ? () => [] : k === "getFilter" ? () => null : k === "getCharts" ? () => [] : () => stub),
  });
  const alerts = [];
  const props = { MANAGER_PIN: "1" };
  const prompts = []; // câu trả lời lần lượt cho ui.prompt
  const ss = {
    getSheetByName: (n) => sheets[n] || null,
    insertSheet: (n) => (sheets[n] = wrap(new Sheet(n))),
    getSheets: () => Object.values(sheets),
    deleteSheet: (sh) => delete sheets[sh.getName()],
    setSpreadsheetTimeZone() {}, toast() {}, getUrl: () => "",
  };
  const env = {
    Date: FakeDate,
    SpreadsheetApp: {
      getActive: () => ss, flush() {}, newDataValidation: () => stub, newConditionalFormatRule: () => stub,
      getUi: () => ({
        alert: (m) => alerts.push(m), ButtonSet: {}, Button: { OK: "OK" },
        prompt: () => ({ getSelectedButton: () => "OK", getResponseText: () => prompts.shift() || "" }),
      }),
      BorderStyle: {}, BandingTheme: {},
    },
    Utilities: { formatDate: fmt },
    PropertiesService: { getScriptProperties: () => ({
      getProperty: (k) => props[k] || null, setProperty: (k, v) => (props[k] = v), deleteProperty: (k) => delete props[k],
    }) },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) },
    ContentService: { createTextOutput: (s) => ({ setMimeType() { return this; }, getContent: () => s }), MimeType: { JSON: 1 } },
    MailApp: { sendEmail() {} },
    UrlFetchApp: { fetch() {} },
    ScriptApp: { getProjectTriggers: () => [] },
    Logger: { log() {} },
  };
  const api = new Function(...Object.keys(env), SRC +
    "\nreturn { doGet, doPost, mergeMisplacedBlocks, auditWeekBlocks, createWeekBlock_, formatMatrixSheet_, reopenThisWeek };")(...Object.values(env));
  const m = ss.insertSheet("Members");
  m.d = [["Tên"]].concat(members.map((x) => [x]));
  const parse = (o) => JSON.parse(o.getContent());
  return {
    sheets, alerts, api, ss, FakeDate, props, prompts, members: m,
    at: (iso) => (NOW = new RealDate(iso)),
    get: (p) => parse(api.doGet({ parameter: p })),
    post: (b) => parse(api.doPost({ postData: { contents: JSON.stringify(b) } })),
  };
}

const MEMBERS = ["Ebi", "Yến Thư", "Việt Anh", "Quỳnh Chi"];
const logs = (s) => (s.sheets["Nhật ký"] ? s.sheets["Nhật ký"].d.slice(1) : []);

test("Chủ nhật tối: tuần đăng ký là 28/09, nằm ở sheet Tháng 10/2026 (tháng của Thứ 5)", () => {
  const s = load(MEMBERS);
  s.at("2026-09-27T21:30:00+07:00");
  assert.strictEqual(s.get({ action: "schedule", name: "Ebi" }).week, "2026-09-28");
  const r = s.post({ action: "registerSchedule", name: "Ebi", week: "2026-09-28", days: [true, false, true, false, false, false, false] });
  assert.strictEqual(r.status, "success", r.message);
  assert.deepStrictEqual(r.days, [true, false, true, false, false, false, false]);
  assert.ok(s.sheets["Tháng 10/2026"], "phải tạo sheet Tháng 10/2026");
  assert.ok(!s.sheets["Tháng 09/2026"], "không được ghi vào Tháng 09/2026");
  const g = s.get({ action: "schedule", name: "Ebi" });
  assert.deepStrictEqual(g.days, [true, false, true, false, false, false, false]);
  assert.ok(g.registered);
});

test("Lưu 2 lần = ghi đè, Nhật ký ghi cả 2 lần", () => {
  const s = load(MEMBERS);
  s.at("2026-09-27T21:30:00+07:00");
  s.post({ action: "registerSchedule", name: "Yến Thư", week: "2026-09-28", days: [true, true, false, false, false, false, false] });
  const r = s.post({ action: "registerSchedule", name: "Yến Thư", week: "2026-09-28", days: [false, true, false, true, false, false, false] });
  assert.deepStrictEqual(r.days, [false, true, false, true, false, false, false]);
  assert.deepStrictEqual(s.get({ action: "schedule", name: "Yến Thư" }).days, [false, true, false, true, false, false, false]);
  const l = logs(s);
  assert.strictEqual(l.length, 2);
  assert.match(l[1][3], /tuần 2026-09-28: T3, T5/);
  assert.match(l[1][4], /^success/);
});

test("Sang Thứ 2: tuần đó khóa, lần bị từ chối vẫn có trong Nhật ký", () => {
  const s = load(MEMBERS);
  s.at("2026-09-28T00:05:00+07:00");
  const r = s.post({ action: "registerSchedule", name: "Quỳnh Chi", week: "2026-09-28", days: [true, false, false, false, false, false, false] });
  assert.strictEqual(r.status, "error");
  assert.match(logs(s)[0][4], /^error/);
});

test("Đơn nghỉ: trưa hôm trước OK; cùng ngày bị chặn, thông báo ghi rõ ngày", () => {
  const s = load(MEMBERS);
  s.at("2026-09-27T12:00:00+07:00");
  assert.strictEqual(s.post({ name: "Việt Anh", type: "Nghỉ", date: "2026-09-28", reason: "Đi tiệc" }).status, "success");
  s.at("2026-09-28T10:46:00+07:00");
  const r = s.post({ name: "Việt Anh", type: "Nghỉ", date: "2026-09-28", reason: "Đi tiệc" });
  assert.strictEqual(r.status, "error");
  assert.match(r.message, /hôm nay 28\/09, chỉ nhận từ ngày 29\/09/);
  assert.strictEqual(logs(s).length, 2);
});

test("Điểm danh: đăng ký + trễ, nghỉ có phép, ngoài lịch", () => {
  const s = load(MEMBERS);
  s.at("2026-09-27T12:00:00+07:00");
  s.post({ action: "registerSchedule", name: "Ebi", week: "2026-09-28", days: [true, false, false, false, false, false, false] });
  s.post({ name: "Việt Anh", type: "Nghỉ", date: "2026-09-28", reason: "x" });
  s.at("2026-09-28T18:40:00+07:00");
  assert.strictEqual(s.post({ action: "checkin", pin: "1", name: "Ebi", date: "2026-09-28", time: "now" }).status, "success");
  s.post({ action: "checkin", pin: "1", name: "Yến Thư", date: "2026-09-28", time: "18:20" });
  const by = Object.fromEntries(s.get({ action: "rollcall", pin: "1", date: "2026-09-28" }).list.map((x) => [x.name, x]));
  assert.strictEqual(by["Ebi"].result, "late");
  assert.strictEqual(by["Yến Thư"].result, "walkin");
  assert.strictEqual(by["Việt Anh"].off, true);
});

test("Backfill: block tuần 28/09 nằm sai ở Tháng 09 được gộp sang Tháng 10, không ghi đè", () => {
  const s = load(MEMBERS);
  s.at("2026-09-27T21:00:00+07:00");
  // Mô phỏng bản code cũ (v15) tạo block tuần 28/09 trong "Tháng 09/2026"
  const old = s.ss.insertSheet("Tháng 09/2026");
  s.api.createWeekBlock_(old, new s.FakeDate(2026, 8, 28));
  const oldRow = old.d.findIndex((r) => r && r[0] === "Yến Thư") + 1;
  old.getRange(oldRow, 2).setValue("✓"); // T2
  old.getRange(oldRow, 7).setValue("✓"); // T3
  // Ebi đăng ký bằng bản mới -> vào Tháng 10
  s.post({ action: "registerSchedule", name: "Ebi", week: "2026-09-28", days: [false, false, false, false, true, false, false] });
  assert.strictEqual(s.get({ action: "schedule", name: "Yến Thư" }).registered, false, "trước khi gộp: không thấy");
  s.api.auditWeekBlocks();
  assert.match(s.alerts[0], /Tuần 28\/09 ở "Tháng 09\/2026".*"Tháng 10\/2026"/);
  s.api.mergeMisplacedBlocks();
  assert.deepStrictEqual(s.get({ action: "schedule", name: "Yến Thư" }).days, [true, true, false, false, false, false, false]);
  assert.deepStrictEqual(s.get({ action: "schedule", name: "Ebi" }).days, [false, false, false, false, true, false, false]);
  s.api.mergeMisplacedBlocks(); // chạy lại không đổi gì
  assert.deepStrictEqual(s.get({ action: "schedule", name: "Yến Thư" }).days, [true, true, false, false, false, false, false]);
});

test("Rollcall/checkin cần PIN; Nhật ký không lưu PIN", () => {
  const s = load(MEMBERS);
  assert.strictEqual(s.get({ action: "rollcall", pin: "sai", date: "2026-09-28" }).status, "error");
  assert.strictEqual(s.post({ action: "checkin", pin: "sai", name: "Ebi", date: "2026-09-28", time: "18:00" }).status, "error");
  assert.ok(!JSON.stringify(logs(s)).includes("sai"));
});

/* ================= AUDIT: đăng ký lịch tuần ================= */
const W = [true, false, true, false, true, false, false];
const reg = (s, name, week, days) => s.post({ action: "registerSchedule", name, week, days });
const rowsOf = (sh, name) => sh.d.filter((r) => r && r[0] === name).length;

test("Audit: CN 23:59 mở trang, 0:01 T2 mới lưu -> WEEK_CHANGED + trả tuần mới, không ghi nhầm", () => {
  const s = load(MEMBERS);
  s.at("2026-09-27T23:59:00+07:00");
  const wk = s.get({ action: "schedule", name: "Ebi" }).week;
  assert.strictEqual(wk, "2026-09-28");
  s.at("2026-09-28T00:01:00+07:00");
  const r = reg(s, "Ebi", wk, W);
  assert.strictEqual(r.code, "WEEK_CHANGED");
  assert.strictEqual(r.week, "2026-10-05");
  assert.strictEqual(s.get({ action: "schedule", name: "Ebi" }).registered, false);
});

test("Audit: tuần vắt tháng / vắt năm vào đúng sheet (tháng chứa Thứ 5)", () => {
  const cases = [
    ["2026-06-27T20:00:00+07:00", "2026-06-29", "Tháng 07/2026"],
    ["2026-09-27T20:00:00+07:00", "2026-09-28", "Tháng 10/2026"],
    ["2026-12-27T20:00:00+07:00", "2026-12-28", "Tháng 12/2026"],
    ["2027-01-03T20:00:00+07:00", "2027-01-04", "Tháng 01/2027"],
  ];
  for (const [now, week, sheet] of cases) {
    const s = load(MEMBERS);
    s.at(now);
    assert.strictEqual(s.get({ action: "schedule", name: "Ebi" }).week, week, now);
    assert.strictEqual(reg(s, "Ebi", week, W).status, "success", now);
    assert.ok(s.sheets[sheet], now + " -> " + sheet + " | có: " + Object.keys(s.sheets));
    assert.deepStrictEqual(s.get({ action: "schedule", name: "Ebi" }).days, W);
  }
});

test("Audit: nhiều người lưu lần lượt cùng block — không ai đè ai", () => {
  const s = load(MEMBERS);
  s.at("2026-09-27T21:00:00+07:00");
  const plan = { "Ebi": W, "Yến Thư": [false, true, false, true, false, true, false],
    "Việt Anh": [true, true, true, true, true, true, true], "Quỳnh Chi": [false, false, false, false, false, false, true] };
  for (const [n, d] of Object.entries(plan)) assert.strictEqual(reg(s, n, "2026-09-28", d).status, "success");
  for (const [n, d] of Object.entries(plan)) assert.deepStrictEqual(s.get({ action: "schedule", name: n }).days, d, n);
});

test('Audit: "false"/1/"true" không phải boolean true -> không tính là tick', () => {
  const s = load(MEMBERS);
  s.at("2026-09-27T21:00:00+07:00");
  const r = reg(s, "Ebi", "2026-09-28", ["false", 1, "true", true, null, 0, false]);
  assert.deepStrictEqual(r.days, [false, false, false, true, false, false, false]);
});

test("Audit: tên không có trong Members bị từ chối, không tạo dòng rác", () => {
  const s = load(MEMBERS);
  s.at("2026-09-27T21:00:00+07:00");
  assert.strictEqual(reg(s, "Hacker", "2026-09-28", W).status, "error");
  assert.ok(!s.sheets["Tháng 10/2026"] || rowsOf(s.sheets["Tháng 10/2026"], "Hacker") === 0);
});

test("Audit: tên gửi dạng Unicode NFD khớp đúng dòng NFC, không tạo dòng trùng", () => {
  const s = load(MEMBERS);
  s.at("2026-09-27T21:00:00+07:00");
  reg(s, "Yến Thư", "2026-09-28", W);
  const r = reg(s, "Yến Thư".normalize("NFD"), "2026-09-28", [true, false, false, false, false, false, false]);
  assert.strictEqual(r.status, "success");
  assert.strictEqual(rowsOf(s.sheets["Tháng 10/2026"], "Yến Thư"), 1);
  assert.deepStrictEqual(s.get({ action: "schedule", name: "Yến Thư" }).days, [true, false, false, false, false, false, false]);
});

test("Audit: thành viên mới thêm giữa chừng -> chèn dòng, block tuần bên dưới không bị lệch", () => {
  const s = load(MEMBERS);
  s.at("2026-09-27T21:00:00+07:00");
  reg(s, "Ebi", "2026-09-28", W); // block 28/09 (Tháng 10)
  // tạo sẵn block 05/10 bên dưới + dữ liệu điểm danh
  s.at("2026-10-05T18:00:00+07:00");
  s.post({ action: "checkin", pin: "1", name: "Yến Thư", date: "2026-10-05", time: "18:10" });
  // thêm người mới, đăng ký vào block PHÍA TRÊN (mở lại tuần 28/09 không áp dụng -> dùng block 12/10 để chắc chắn)
  s.members.d.push(["Người Mới"]);
  s.at("2026-10-04T20:00:00+07:00");
  assert.strictEqual(reg(s, "Người Mới", "2026-10-05", W).status, "success");
  const by = Object.fromEntries(s.get({ action: "rollcall", pin: "1", date: "2026-10-05" }).list.map((x) => [x.name, x]));
  assert.strictEqual(by["Yến Thư"].checkin, "18:10");
  assert.strictEqual(by["Người Mới"].registered, true);
  s.at("2026-09-28T12:00:00+07:00");
  assert.strictEqual(s.get({ action: "rollcall", pin: "1", date: "2026-09-28" }).list.find((x) => x.name === "Ebi").registered, true);
});

test("Audit: ô tiêu đề tuần bị gõ đè thành chữ -> vẫn tìm thấy block, không tạo block thứ 2, tự sửa lại", () => {
  const s = load(MEMBERS);
  s.at("2026-09-27T21:00:00+07:00");
  reg(s, "Ebi", "2026-09-28", W);
  const sh = s.sheets["Tháng 10/2026"];
  const titleIdx = sh.d.findIndex((r) => r && r[0] instanceof Date);
  sh.d[titleIdx][0] = "Tuần 28/09 – 04/10"; // admin gõ đè
  assert.deepStrictEqual(s.get({ action: "schedule", name: "Ebi" }).days, W, "đọc được qua tiêu đề chữ");
  assert.strictEqual(reg(s, "Yến Thư", "2026-09-28", W).status, "success");
  assert.strictEqual(sh.d.filter((r) => r && (r[0] instanceof Date || /^Tuần/.test(String(r[0])))).length, 1, "chỉ 1 block");
  assert.ok(sh.d[titleIdx][0] instanceof Date, "tiêu đề được ghi lại thành Date");
});

test("Mở lại tuần này: chỉ sửa được từ hôm nay, ngày đã qua giữ nguyên; hết hạn tự đóng", () => {
  const s = load(MEMBERS);
  s.at("2026-09-27T12:00:00+07:00");
  reg(s, "Ebi", "2026-09-28", [true, false, false, false, false, false, false]); // T2 đã ĐK từ trước
  s.at("2026-09-29T10:00:00+07:00"); // Thứ 3
  assert.strictEqual(s.get({ action: "config" }).reopen, null);
  s.prompts.push("20:00");
  s.api.reopenThisWeek();
  assert.match(s.alerts.pop(), /Đã mở lại đăng ký tuần 28\/09 đến 29\/09 20:00/);
  const cfg = s.get({ action: "config" }).reopen;
  assert.deepStrictEqual(cfg, { week: "2026-09-28", fromDay: 1, until: "29/09 20:00" });
  const g = s.get({ action: "schedule", name: "Ebi", week: "2026-09-28" });
  assert.strictEqual(g.week, "2026-09-28");
  assert.strictEqual(g.fromDay, 1);
  // cố bỏ tick T2 (đã qua) + tick T3, T5
  const r = reg(s, "Ebi", "2026-09-28", [false, true, false, true, false, false, false]);
  assert.strictEqual(r.status, "success", r.message);
  assert.match(r.message, /\(mở lại\)/);
  assert.deepStrictEqual(r.days, [true, true, false, true, false, false, false], "T2 giữ nguyên");
  // tuần sau vẫn đăng ký bình thường song song
  assert.strictEqual(reg(s, "Ebi", "2026-10-05", W).status, "success");
  // hết hạn
  s.at("2026-09-29T20:01:00+07:00");
  assert.strictEqual(reg(s, "Yến Thư", "2026-09-28", W).code, "WEEK_CHANGED");
  assert.strictEqual(s.get({ action: "schedule", name: "Ebi", week: "2026-09-28" }).week, "2026-10-05");
  assert.strictEqual(s.get({ action: "config" }).reopen, null);
});

test("Mở lại: chặn hạn ở quá khứ / quá Chủ nhật; để trống = đóng", () => {
  const s = load(MEMBERS);
  s.at("2026-09-29T10:00:00+07:00");
  s.prompts.push("09:00"); s.api.reopenThisWeek();
  assert.match(s.alerts.pop(), /tương lai/);
  s.prompts.push("05/10 10:00"); s.api.reopenThisWeek();
  assert.match(s.alerts.pop(), /Chủ nhật/);
  s.prompts.push("04/10 23:00"); s.api.reopenThisWeek();
  assert.ok(s.get({ action: "config" }).reopen);
  s.prompts.push(""); s.api.reopenThisWeek();
  assert.strictEqual(s.get({ action: "config" }).reopen, null);
  // REOPEN của tuần trước không có hiệu lực tuần này
  s.props.REOPEN = JSON.stringify({ week: "2026-09-21", until: Date.parse("2026-12-31T00:00:00+07:00") });
  assert.strictEqual(s.get({ action: "config" }).reopen, null);
});

test("Audit: tên trong Members/Sheet gõ dạng NFD, web gửi NFC -> vẫn đúng 1 dòng", () => {
  const nfd = "Quỳnh Chi".normalize("NFD");
  const s = load(["Ebi", nfd]);
  s.at("2026-09-27T21:00:00+07:00");
  const r = reg(s, "Quỳnh Chi", "2026-09-28", W);
  assert.strictEqual(r.status, "success", r.message);
  const sh = s.sheets["Tháng 10/2026"];
  assert.strictEqual(sh.d.filter((x) => x && String(x[0]).normalize("NFC") === "Quỳnh Chi").length, 1);
  assert.deepStrictEqual(s.get({ action: "schedule", name: "Quỳnh Chi" }).days, W);
});
