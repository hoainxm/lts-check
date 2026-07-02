# LTS Check — Hướng dẫn Setup Backend (Google Sheet + Apps Script)

## Bước A: Tạo Google Sheet

1. Mở https://sheets.new → tạo sheet mới, đặt tên: **LTS - Quản lý Trễ/Nghỉ**

## Bước B: Cài Apps Script

1. Trong Sheet: menu **Extensions (Tiện ích mở rộng) → Apps Script**
2. Xoá code mặc định, dán toàn bộ nội dung file [`apps-script/Code.gs`](apps-script/Code.gs)
3. Bấm **Save** (Ctrl+S)

## Bước C: Chạy setup (1 lần duy nhất)

1. Trên thanh công cụ, chọn hàm **`setup`** trong dropdown → bấm **Run (▶)**
2. Google hỏi quyền → **Review permissions** → chọn tài khoản → **Advanced → Go to ... (unsafe) → Allow**
   (cảnh báo "unsafe" là bình thường vì script tự viết, chưa qua Google review)
3. Chạy xong sẽ có 3 sheet:
   - **Data**: dữ liệu thô, có dropdown Trạng thái + màu tự động
   - **Members**: danh sách thành viên → **thay 3 tên mẫu bằng tên thật của team**
   - **Tổng kết**: thống kê theo tháng, đổi ô Tháng/Năm để xem tháng khác

## Bước D: Deploy thành Web App (API)

1. Trong Apps Script: **Deploy → New deployment**
2. Bấm ⚙️ chọn type: **Web app**
3. Cấu hình:
   - Description: `LTS API v1`
   - Execute as: **Me** (tài khoản của bạn)
   - Who has access: **Anyone** ⚠️ bắt buộc, để form gọi được không cần đăng nhập
4. Bấm **Deploy** → copy **Web app URL** (dạng `https://script.google.com/macros/s/AKfy.../exec`)

## Bước E: Gửi URL cho tôi

Dán URL đó vào chat → tôi viết frontend (Bước 3) kết nối vào.

---

## Ghi chú

- **Sửa danh sách thành viên**: chỉ cần sửa sheet `Members`, dropdown trên web tự cập nhật.
- **Duyệt yêu cầu**: Manager đổi cột Trạng thái trong sheet `Data` (dropdown: Chờ duyệt / Đã duyệt / Từ chối). Yêu cầu "Từ chối" không tính vào thống kê.
- **Đổi giờ tập chuẩn**: sửa `PRACTICE_START_HOUR` / `PRACTICE_START_MIN` trong `Code.gs`, sau đó deploy lại (Deploy → Manage deployments → Edit → Version: New version).
- **Sau này sửa code GAS**: mỗi lần sửa phải tạo **New version** trong Manage deployments thì URL cũ mới nhận code mới.

## Test nhanh API (tuỳ chọn)

Mở URL sau trong trình duyệt:
```
<WEB_APP_URL>?action=members
```
Thấy JSON danh sách tên → API sống.
