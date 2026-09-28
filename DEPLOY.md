# Hướng dẫn Deploy (bản hoàn thiện)

Làm lần lượt **A → B → C**. Backend (Apps Script) và Frontend (GitHub Pages) là 2 chỗ riêng.

---

## ⚡ Cách nhanh: deploy bằng script (khuyên dùng)

Setup 1 lần: bật Apps Script API tại https://script.google.com/home/usersettings →
`bash scripts/gas.sh login` (tài khoản chủ Sheet) → ghi `SCRIPT_ID=...` (Project Settings) vào `scripts/gas.env`.

```bash
bash scripts/gas.sh deploy        # test → backup bản đang chạy → push → version mới → deploy → health check
bash scripts/gas.sh versions      # xem các version
bash scripts/gas.sh rollback 12   # quay về version 12, URL giữ nguyên
bash scripts/check-api.sh         # kiểm tra API đang chạy (chỉ đọc)
```

CI (GitHub Actions) chặn merge nếu `index.html` gọi action mà `Code.gs` không có, hoặc test backend fail.

---

## A. Backend (Google Apps Script) — cách thủ công

1. Mở Google Sheet → **Extensions → Apps Script**
2. Xoá hết code cũ → dán **toàn bộ** `apps-script/Code.gs` → **Ctrl+S**
3. Đóng/mở lại Sheet → có menu **⚙️ LTS** trên thanh công cụ
4. Menu **⚙️ LTS → Setup / Format lại** (chạy từ menu, KHÔNG chạy từ editor)
   - Tạo: sheet **tháng hiện tại** (VD `07/2026`), **Members**, **Tổng kết**
5. Menu **⚙️ LTS → Đặt PIN quản lý** → nhập PIN (VD `270203`)
   - Hoặc: Project Settings ⚙️ → Script Properties → `MANAGER_PIN` = `270203`
6. **Deploy lại (BẮT BUỘC):** Deploy → Manage deployments → ✏️ Edit → **Version: New version** → Deploy
   - URL `/exec` giữ nguyên

---

## B. Frontend (GitHub Pages)

1. Repo GitHub → mở `index.html` → **Delete** (hoặc Upload đè)
2. **Upload** `index.html` mới → **Commit**
3. Chờ ~1 phút Pages build

---

## C. Kiểm tra

1. Mở web → **Ctrl + F5** (xoá cache)
2. **User thường:** chỉ thấy 1 tab **📝 Gửi**
3. Gửi thử → báo *"Đã gửi, chờ admin duyệt"* → dòng vào sheet tháng (VD `07/2026`), Trạng thái **Chờ duyệt**
4. Bấm **🔒** (góc trên phải) → nhập PIN → hiện **✅ Duyệt / 📊 Thống kê / 🕘 Lịch sử**
5. Tab **Duyệt** → **✅ Duyệt** → Trạng thái đổi *Đã duyệt* → hiện trong **Thống kê**
6. Bấm **🔓** (góc trên phải) hoặc **🚪 Thoát quản lý** (đáy trang) → về giao diện user

---

## Cách hoạt động

```
User gửi ──► Sheet tháng "MM/YYYY" (Chờ duyệt)
                 ├─ Admin ✅ Duyệt  ─► Đã duyệt ─► tính vào Thống kê / Tổng kết
                 └─ Admin ❌ Từ chối ─► Từ chối (không tính)
```
- **Mỗi tháng 1 sheet** tên `07/2026`, `08/2026`… tự tạo khi có yêu cầu tháng đó.
- **Lịch tuần / đơn / điểm danh** (bản ma trận, v20+): mỗi tháng 1 sheet `Tháng MM/yyyy`, trong đó mỗi tuần 1 block
  (người × ngày × 5 ô: ĐKy tập | Đi muộn/Nghỉ | Giờ dự kiến | Giờ đến | Lý do).
  Tuần vắt 2 tháng nằm ở sheet của **tháng chứa Thứ 5** — VD tuần 28/09–04/10 nằm ở `Tháng 10/2026`, KHÔNG ở `Tháng 09/2026`.
- **Nhật ký**: sheet `Nhật ký` ghi mọi lần gửi (đăng ký lịch, đơn, điểm danh — kể cả bị từ chối) kèm giờ VN → tra khi có khiếu nại.
- Menu ⚙️ LTS → **Kiểm tra block tuần nằm sai sheet** / **Gộp block tuần nằm sai sheet (backfill)**: gộp dữ liệu block do bản code cũ
  tạo nhầm sheet về đúng chỗ, chỉ điền ô trống, không xoá block cũ.
- **Mở lại đăng ký tuần này**: menu ⚙️ LTS → **Mở lại / đóng đăng ký tuần này** → nhập hạn (`20:00` hoặc `30/09 21:00`,
  tối đa hết Chủ nhật; để trống = đóng). Web hiện thêm nút "Tuần này (mở lại)"; chỉ tick được từ hôm nay, ngày đã qua giữ nguyên;
  hết hạn tự đóng. Lượt lưu ghi "(mở lại)" trong `Nhật ký`.
- Menu **Chuyển dữ liệu cũ → sheet Tháng** tự chặn khi đã có sheet Tháng (chạy lại sẽ xoá sạch dữ liệu mới).
- ⚠️ `apps-script/Code.gs` trong repo là **nguồn duy nhất**. Sửa trực tiếp trong trình soạn Apps Script thì phải copy ngược về repo ngay, nếu không lần deploy sau sẽ làm mất tính năng.
- **Thống kê / Tổng kết** chỉ tính bản **Đã duyệt**.
- **Lịch sử** hiện mọi yêu cầu của 1 người (mọi trạng thái, mới nhất trước).

## Ghi chú
- Đổi PIN: menu ⚙️ LTS → Đặt PIN (không cần deploy lại).
- Đổi thành viên: sửa sheet **Members** cột A.
- Đổi hình nền (admin): tab Thống kê → ô 🎨 (ảnh hiển thị theo cột mobile).
- Đổi API URL: sửa `API_URL` đầu `index.html` rồi up lại.
- Sheet `Data` cũ (nếu có từ bản trước) không dùng nữa — có thể xoá tay.
