#!/usr/bin/env bash
# Kiểm tra nhanh API đang chạy (chỉ gọi GET đọc, không ghi dữ liệu).
# Dùng: bash scripts/check-api.sh          (lấy API_URL từ index.html)
#       bash scripts/check-api.sh <URL/exec>
set -u
cd "$(dirname "$0")/.."
URL="${1:-$(grep -oE 'https://script\.google\.com/macros/s/[^"]+/exec' index.html | head -1)}"
[ -z "$URL" ] && { echo "Không tìm thấy API_URL"; exit 2; }

fail=0
check() { # $1 = mô tả, $2 = query, $3 = regex mong đợi
  local out
  out=$(curl -sL --max-time 30 "$URL?$2")
  if echo "$out" | grep -qE "$3"; then echo "✅ $1"
  else echo "❌ $1 → ${out:0:150}"; fail=1; fi
}
check "members"            "action=members"                    '"status":"success"'
check "config"             "action=config"                     '"status":"success"'
check "schedule (Lịch tuần)" "action=schedule&name=__healthcheck__" '"status":"success".*"week"'
check "rollcall có route (đòi PIN)" "action=rollcall&date=2000-01-01&pin=x" 'PIN'
check "history có route (đòi PIN)"  "action=history&name=x&pin=x"      'PIN'
exit $fail
