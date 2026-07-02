# Deploy Frontend lên GitHub Pages

File [`index.html`](index.html) đã nhúng sẵn API URL — chỉ cần đưa lên GitHub Pages.

## Cách 1: Qua web GitHub (không cần cài git)

1. Vào https://github.com/new → tạo repo mới, VD tên `lts-check`, để **Public**
2. Trong repo: **Add file → Upload files** → kéo thả `index.html` → **Commit**
3. **Settings → Pages** → mục *Build and deployment*:
   - Source: **Deploy from a branch**
   - Branch: **main** / **/ (root)** → **Save**
4. Chờ ~1 phút → link hiện ra: `https://<user>.github.io/lts-check/`
5. Gửi link đó cho team.

## Cách 2: Qua git (nếu có cài)

```bash
cd D:/lts-check
git init
git add index.html
git commit -m "LTS check web app"
git branch -M main
git remote add origin https://github.com/<user>/lts-check.git
git push -u origin main
```
Rồi bật Pages như bước 3 ở Cách 1.

## Lưu ý

- Đổi API URL: sửa biến `API_URL` đầu file `index.html`.
- Mỗi lần sửa `index.html` → upload/push lại, Pages tự cập nhật sau ~1 phút.
- Muốn test ngay trên máy: mở thẳng file `index.html` bằng trình duyệt cũng chạy được (GET/POST tới GAS không dính CORS vì gửi `text/plain`).
