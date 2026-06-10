# QuoteBot — Quà Tặng VIVA

Hệ thống báo giá tự động với 604 sản phẩm từ 19 nhóm hàng.

---

## 🚀 Deploy lên Vercel (miễn phí, ~10 phút)

### Bước 1 — Lấy Anthropic API Key
1. Vào https://console.anthropic.com
2. Đăng ký / đăng nhập
3. Vào **API Keys** → **Create Key**
4. Copy key (dạng `sk-ant-...`)

### Bước 2 — Đưa code lên GitHub
1. Vào https://github.com → **New repository** → đặt tên (vd: `quotebot-viva`)
2. Tải [GitHub Desktop](https://desktop.github.com/) nếu chưa có
3. Clone repo vừa tạo về máy
4. Copy toàn bộ thư mục `quotebot-deploy` vào trong repo
5. Commit & Push

### Bước 3 — Deploy lên Vercel
1. Vào https://vercel.com → đăng nhập bằng GitHub
2. Click **Add New Project** → chọn repo `quotebot-viva`
3. Vercel tự detect Vite → click **Deploy**
4. Sau khi deploy xong, vào **Settings → Environment Variables**
5. Thêm:
   - **Name:** `ANTHROPIC_API_KEY`
   - **Value:** key vừa copy ở Bước 1
6. Vào **Deployments → Redeploy** để áp dụng key

✅ Xong! Bạn có link dạng `https://quotebot-viva.vercel.app`

---

## 🔒 Bảo mật

API key được lưu trong Vercel Environment Variables, **không bao giờ** lộ ra ngoài.
Người dùng chỉ thấy giao diện web, không thấy key.

---

## 💻 Chạy local (để test)

```bash
npm install
# Tạo file .env từ .env.example và điền key
cp .env.example .env
npm run dev
```

Mở http://localhost:5173

---

## 📋 Cấu trúc thư mục

```
quotebot-deploy/
├── api/
│   └── chat.js          ← Proxy bảo mật (Vercel serverless)
├── src/
│   ├── App.jsx          ← Toàn bộ UI + catalog 604 sản phẩm
│   └── main.jsx         ← Entry point React
├── index.html
├── package.json
├── vite.config.js
├── vercel.json
└── .env.example
```
