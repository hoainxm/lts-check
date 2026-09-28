// Chặn lệch frontend ↔ backend: mọi action index.html gọi phải có nhánh xử lý trong Code.gs.
// (Sự cố 28/09/2026: deploy Code.gs thiếu schedule/rollcall -> Lịch tuần + Điểm danh chết.)
import { readFileSync } from "node:fs";

const html = readFileSync(new URL("../index.html", import.meta.url), "utf8");
const gs = readFileSync(new URL("../apps-script/Code.gs", import.meta.url), "utf8");

const called = new Set([
  ...[...html.matchAll(/[?&]action=([a-zA-Z]+)/g)].map((m) => m[1]),
  ...[...html.matchAll(/action:\s*"([a-zA-Z]+)"/g)].map((m) => m[1]),
]);
const handled = new Set(
  [...gs.matchAll(/action\s*===\s*'([a-zA-Z]+)'/g)].map((m) => m[1]),
);

const missing = [...called].filter((a) => !handled.has(a));
if (missing.length) {
  console.error("❌ Code.gs thiếu xử lý action: " + missing.join(", "));
  process.exit(1);
}
console.log("✅ " + called.size + " action frontend đều có trong Code.gs: " + [...called].join(", "));
