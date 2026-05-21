# 🔗 Inlinks Filter — Lịch Sử Dự Án

Tool lọc contextual internal links từ file export Screaming Frog.  
Deploy: **Streamlit Cloud** — [seosona.streamlit.app](https://seosona.streamlit.app)

---

## 📋 Tóm Tắt Dự Án

| Thông tin | Chi tiết |
|---|---|
| Bắt đầu | 2026-05-21 |
| Stack | Python · Streamlit · pandas · openpyxl |
| Deploy | Streamlit Cloud (free tier) |
| Repo | github.com/tranchiquyet285-svg/inlinks-filter |
| Dùng cho | Team nội bộ SEOSONA |

---

## 🎯 Mục Tiêu

Từ file `inlinks.csv` export của Screaming Frog → lọc ra chỉ những **internal link contextual** (nằm trong nội dung biên tập thật sự) → loại bỏ link điều hướng, card sản phẩm, sidebar, pagination...

Output: file Excel 4 sheet gồm link đã lọc, thống kê theo From, thống kê theo To (SEO), và quy tắc lọc.

---

## 🗂️ Cấu Trúc File

```
inlinks-filter/
  app.py            ← Streamlit UI + login
  filter_logic.py   ← Core filtering logic (patterns, rules)
  create_user.py    ← Helper tạo password hash cho user mới
  requirements.txt
  README.md
```

---

## 👤 Quản Lý Tài Khoản

### Thêm user mới

**Bước 1** — Tính hash password (chạy trên máy có Python):
```bash
python create_user.py
```
Nhập username, tên hiển thị, mật khẩu → copy output.

**Hoặc** — Dùng PowerShell (không cần Python):
```powershell
$pass = "matkhau_can_hash"
$bytes = [System.Text.Encoding]::UTF8.GetBytes($pass)
$sha = [System.Security.Cryptography.SHA256]::Create()
$hash = [System.BitConverter]::ToString($sha.ComputeHash($bytes)).Replace("-","").ToLower()
Write-Host $hash
```

**Bước 2** — Vào Streamlit Cloud → app → Settings → Secrets → thêm:
```toml
[users.tenusermoi]
name     = "Tên Hiển Thị"
password = "hash_tu_buoc_1"
```
Bấm **Save changes**.

### Xoá user
Vào Secrets → xoá block `[users.tenuser]` tương ứng → Save.

### Đổi mật khẩu
Tính hash mới → cập nhật `password` trong Secrets → Save.

### Danh sách user hiện tại
| Username | Tên | Ngày tạo |
|---|---|---|
| tranchiquyet | Quyết | 2026-05-21 |

---

## ⚙️ Logic Lọc

### Điều kiện GIỮ (tất cả phải thỏa)
1. `Type == Hyperlink`
2. `Link Position == Content`
3. `Anchor Text` không rỗng, không phải button/số
4. `Link Path` khớp ít nhất 1 **contextual pattern**
5. `Link Path` không khớp **exclude pattern** — trừ khi nằm trong strong content container

### Structural Patterns (universal, mọi CMS)
- `/p/a` — link trong paragraph ← pattern quan trọng nhất
- `/li/a` — link trong list item
- `/td/a` — link trong table cell (bảng giá trong mô tả)

### CMS Patterns
- `entry-content`, `post-content`, `the-content` — WordPress
- `tab-description`, `woocommerce-Tabs-panel--description` — WooCommerce
- `term-description`, `category-description` — mô tả category
- `elementor-widget-text-editor`, `wp-block-paragraph` — Page builders

### Strong Content Container (override exclude)
Nếu link nằm trong container nội dung đã biết (entry-content, term-description, tab-description...) → bỏ qua toàn bộ exclude patterns. Ví dụ: bảng giá trong mô tả category vẫn được giữ dù có `/table/tbody/`.

### Exclude Patterns
Tab chuyển tab, product grid, card sản phẩm, comment, login, TOC, carousel, breadcrumb, pagination, Divi blocks (section_XXX, col-XXX), box-subcategoryseo...

---

## 🌐 Áp Dụng Cho Website Mới

Nếu tool lọc được 0 link hoặc quá ít:

1. Upload file → mở expander **"⚠️ X rows không khớp pattern"** → xem top id/class
2. Mở expander **"🔬 Debug — Soi link của 1 URL cụ thể"** → nhập URL → xem dòng ❌ và Link Path
3. Tìm id/class tương ứng với body bài viết, mô tả sản phẩm, mô tả category
4. Báo → cập nhật `CONTEXTUAL_PATTERNS` hoặc `EXCLUDE_PATTERNS` trong `filter_logic.py`

---

## 📝 Lịch Sử Thay Đổi

| Ngày | Thay đổi |
|---|---|
| 2026-05-21 | Khởi tạo dự án, build Streamlit app từ quy trình xenangplus.com |
| 2026-05-21 | Deploy lên Streamlit Cloud |
| 2026-05-21 | Mở rộng patterns universal (mọi CMS) |
| 2026-05-21 | Thêm exclude: box-subcategoryseo, Divi blocks |
| 2026-05-21 | Xoá `/article/` pattern (quá rộng, bắt nhầm sidebar) |
| 2026-05-21 | Thêm `/li/a`, `/td/a` patterns |
| 2026-05-21 | Strong container override exclude (bảng giá trong mô tả category) |
| 2026-05-21 | Thêm Debug tool theo URL |
| 2026-05-21 | Thêm login system (SHA256 + Streamlit secrets) |

---

## 🐛 Quy Trình Fix Bug

Khi gặp link bị lọc sai:
1. Mở Debug expander → nhập URL → xem cột Kết quả + Link Path
2. Nếu **❌ Loại: xxx** → thêm exclude pattern đó vào whitelist hoặc tạo strong container pattern
3. Nếu **❌ Không khớp contextual** → thêm pattern mới vào `CONTEXTUAL_PATTERNS`
4. Sửa `filter_logic.py` → push → Streamlit tự reload
