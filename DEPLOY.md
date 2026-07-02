# Hướng dẫn Deploy (bản hoàn thiện)

Làm lần lượt **A → B → C**. Backend (Apps Script) và Frontend (GitHub Pages) là 2 chỗ riêng.

---

## A. Backend (Google Apps Script)

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
- **Thống kê / Tổng kết** chỉ tính bản **Đã duyệt**.
- **Lịch sử** hiện mọi yêu cầu của 1 người (mọi trạng thái, mới nhất trước).

## Ghi chú
- Đổi PIN: menu ⚙️ LTS → Đặt PIN (không cần deploy lại).
- Đổi thành viên: sửa sheet **Members** cột A.
- Đổi hình nền (admin): tab Thống kê → ô 🎨 (ảnh hiển thị theo cột mobile).
- Đổi API URL: sửa `API_URL` đầu `index.html` rồi up lại.
- Sheet `Data` cũ (nếu có từ bản trước) không dùng nữa — có thể xoá tay.
