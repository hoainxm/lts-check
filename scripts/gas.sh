#!/usr/bin/env bash
# Quản lý Apps Script bằng clasp — thay cho copy/dán tay trong trình soạn.
#
# Lần đầu (1 lần):
#   1. https://script.google.com/home/usersettings -> bật "Google Apps Script API"
#   2. bash scripts/gas.sh login          (đăng nhập ĐÚNG tài khoản chủ Sheet LTS)
#   3. Apps Script -> Project Settings -> copy "Script ID" -> ghi vào scripts/gas.env:
#        SCRIPT_ID=xxxxxxxx
#      (DEPLOYMENT_ID tự lấy từ API_URL trong index.html)
#
# Lệnh:
#   bash scripts/gas.sh backup            tải code ĐANG CÓ trên Apps Script về backups/<giờ>/
#   bash scripts/gas.sh versions          liệt kê version đã tạo
#   bash scripts/gas.sh rollback <số>     trỏ Web App về version <số> (URL giữ nguyên)
#   bash scripts/gas.sh deploy            test -> backup -> push -> version mới -> deploy -> check-api
set -euo pipefail
cd "$(dirname "$0")/.."
CLASP="npx -y @google/clasp@2.4.2"

[ -f scripts/gas.env ] && source scripts/gas.env
DEPLOYMENT_ID="${DEPLOYMENT_ID:-$(grep -oE 'macros/s/[^/"]+/exec' index.html | head -1 | sed -E 's#macros/s/([^/]+)/exec#\1#')}"

need_id() {
  [ -n "${SCRIPT_ID:-}" ] || { echo "Thiếu SCRIPT_ID — ghi vào scripts/gas.env (xem đầu file)."; exit 2; }
  printf '{"scriptId":"%s","rootDir":"apps-script"}\n' "$SCRIPT_ID" > .clasp.json
}

backup() {
  local dir="backups/$(date +%Y%m%d-%H%M%S)"
  mkdir -p "$dir"
  printf '{"scriptId":"%s"}\n' "$SCRIPT_ID" > "$dir/.clasp.json"
  (cd "$dir" && $CLASP pull >/dev/null)
  rm "$dir/.clasp.json"
  echo "📦 Đã lưu code đang chạy vào $dir"
}

case "${1:-}" in
  login)    $CLASP login ;;
  backup)   need_id; backup ;;
  versions) need_id; $CLASP versions ;;
  rollback)
    need_id
    [ -n "${2:-}" ] || { echo "Dùng: gas.sh rollback <số version>  (xem: gas.sh versions)"; exit 2; }
    $CLASP deploy -i "$DEPLOYMENT_ID" -V "$2" -d "rollback về v$2"
    bash scripts/check-api.sh || true ;;
  deploy)
    need_id
    node scripts/check-actions.mjs
    TZ=Asia/Ho_Chi_Minh node --test tests/backend.test.cjs >/dev/null && echo "✅ test backend"
    backup
    $CLASP push -f
    ver=$($CLASP version "$(git log -1 --format='%h %s')" | grep -oE '[0-9]+' | tail -1)
    echo "🏷  Version mới: $ver"
    $CLASP deploy -i "$DEPLOYMENT_ID" -V "$ver" -d "$(git log -1 --format='%h')"
    echo "⏳ Chờ Google cập nhật…"; sleep 5
    bash scripts/check-api.sh ;;
  *) sed -n '2,17p' "$0"; exit 1 ;;
esac
