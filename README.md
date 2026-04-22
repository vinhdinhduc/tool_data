# 🏛️ Công cụ Tổng hợp Dữ liệu Cấp Xã
**Google Apps Script – Xã Mường La**

---

## 📁 Cấu trúc file dự án

```
GAS_TongHopDuLieuXa/
├── Config.gs        → Cấu hình toàn bộ hệ thống (tên sheet, số cột, số dòng)
├── Code.gs          → Menu, hàm tiện ích, hàm gọi từ menu
├── TongHop.gs       → Chức năng tổng hợp (cơ bản)
├── QuanLyDong.gs    → Quản lý dòng đồng bộ (nâng cao)
└── README.md        → Hướng dẫn này
```

---

## ⚙️ HƯỚNG DẪN TRIỂN KHAI

### Bước 1 – Upload file Excel lên Google Drive

1. Mở [drive.google.com](https://drive.google.com)
2. Kéo file `Biểu_Báo_cáo_tháng_các_bản__tiểu_khu__Xog_.xlsx` vào Drive
3. Chuột phải → **"Mở bằng" → "Google Trang tính"**
4. File sẽ chuyển thành Google Sheets (giữ nguyên tên)

### Bước 2 – Mở Apps Script Editor

1. Trong Google Sheets, vào **Extensions → Apps Script**
2. Một tab mới mở ra với file `Code.gs` mặc định

### Bước 3 – Thêm các file script

Trong Apps Script Editor:

1. **Xóa** nội dung mặc định trong `Code.gs`
2. **Paste** nội dung từ file `Code.gs` trong thư mục này vào
3. Tạo file mới: click **"+"** bên cạnh "Files"
   - Đặt tên `Config` → paste nội dung `Config.gs`
   - Đặt tên `TongHop` → paste nội dung `TongHop.gs`
   - Đặt tên `QuanLyDong` → paste nội dung `QuanLyDong.gs`
4. Nhấn **Save (Ctrl+S)**

> **Lưu ý:** Apps Script tự thêm đuôi `.gs`, bạn chỉ cần nhập tên không cần đuôi.

### Bước 4 – Cấp quyền

1. Nhấn nút **Run** (▷) → chọn hàm `onOpen`
2. Google yêu cầu cấp quyền → **Review permissions**
3. Chọn tài khoản Google → **Allow**
4. Quay lại Google Sheets → **Reload trang**
5. Menu **"🏛️ Công cụ Xã"** xuất hiện trên thanh menu ✅

---

## 📖 HƯỚNG DẪN SỬ DỤNG

### Chức năng 1: Tổng hợp BIỂU TỔNG

> Menu: **🏛️ Công cụ Xã → 📊 Tổng hợp BIỂU TỔNG từ các bản**

**Khi nào dùng:**
- Sau khi các bản đã nhập xong số liệu vào cột D-G (ô màu vàng)
- Khi muốn rebuild lại toàn bộ công thức (nếu bị hỏng)
- Sau khi thêm sheet bản mới vào file

**Cách hoạt động:**
- Script scan tất cả dòng trong BIỂU TỔNG (dòng 7 → cuối)
- Với mỗi ô có công thức tổng hợp (`='BẢN NÀ LỐC'!D9+...`)
- Rebuild lại với **danh sách sheet hiện tại** + **số dòng thực tế**

---

### Chức năng 2: Cập nhật BIỂU TỔNG TOÀN XÃ

> Menu: **🏛️ Công cụ Xã → 📋 Cập nhật BIỂU TỔNG TOÀN XÃ**

**Kết quả:**
- Cột C (Số hộ): lấy từ Luỹ kế "Tổng số hộ" của mỗi bản
- Cột D (Số nhân khẩu): lấy từ Luỹ kế "Tổng số nhân khẩu"
- Dữ liệu được đặt dạng **công thức tham chiếu** → tự động cập nhật

---

### Chức năng 3: Thêm dòng mới (Nâng cao ⭐)

> Menu: **🏛️ Công cụ Xã → 📝 Quản lý dòng → ➕ Thêm dòng mới**

**Ví dụ thực tế:**
Cần thêm chỉ tiêu "Số hộ vay vốn ngân hàng" sau dòng 20:
1. Chọn menu → Nhập **20** → OK
2. Script tự động:
   - Chèn dòng mới (dòng 21) vào **BIỂU TỔNG** và **54 sheet bản**
   - Đặt công thức tổng hợp cho dòng 21 trong BIỂU TỔNG
   - Đặt công thức Luỹ kế (H = D+E+F+G) cho dòng 21 trong tất cả bản
3. Người dùng chỉ cần điền:
   - Cột B, dòng 21: "Số hộ vay vốn ngân hàng"
   - Cột C, dòng 21: "Hộ"

**Nguyên lý "Tính vị trí theo tổng số dòng":**
```
// KHÔNG làm thế này (hardcode số dòng - sai):
formula = "='BẢN NÀ LỐC'!D21"

// Làm thế này (tính động - đúng):
var newRow = insertAfterRow + 1;  // Tính từ vị trí thực tế
formula = "='BẢN NÀ LỐC'!D" + newRow  // Dùng biến newRow
```

---

### Chức năng 4: Xóa dòng (Nâng cao ⭐)

> Menu: **🏛️ Công cụ Xã → 📝 Quản lý dòng → ❌ Xóa dòng**

⚠️ **Cảnh báo:** Không thể hoàn tác. Script sẽ xác nhận 2 lần trước khi xóa.

---

## 🔧 CẤU TRÚC KỸ THUẬT

### Cấu trúc Sheet

| Sheet | Vai trò |
|-------|---------|
| `BIỂU TỔNG` | Tổng hợp chỉ tiêu toàn xã (formula-based) |
| `BIỂU TỔNG TOÀN XÃ` | Tổng hợp hộ/nhân khẩu theo bản |
| `DS BẢN, TK ` | Danh sách 54 bản/tiểu khu (tĩnh) |
| `BẢN NÀ LỐC` ... `BẢN CÁT LÌNH` | 54 sheet bản (nhập liệu) |

### Cấu trúc Cột (BIỂU TỔNG & Sheet bản)

| Cột | Ký hiệu | Nội dung |
|-----|---------|----------|
| A | Stt | Số thứ tự |
| B | Nội dung | Tên chỉ tiêu |
| C | ĐVT | Đơn vị tính |
| D | Đầu kỳ | Số liệu đầu kỳ |
| E | Tháng 10 | Số phát sinh T10 |
| F | Tháng 11 | Số phát sinh T11 |
| G | Tháng 12 | Số phát sinh T12 |
| H | Luỹ kế | = D+E+F+G |
| I | Ghi chú | |

### Công thức tổng hợp (BIỂU TỔNG)

```
Ô D9 = ='BẢN NÀ LỐC'!D9+'BẢN NONG HEO'!D9+...+'BẢN CÁT LÌNH'!D9
Ô H9 = =D9+E9+F9+G9
```

---

## ❓ XỬ LÝ LỖI THƯỜNG GẶP

| Lỗi | Nguyên nhân | Giải pháp |
|-----|-------------|-----------|
| Menu không hiện | Chưa chạy `onOpen` | Extensions → Apps Script → Run `onOpen` |
| "Không tìm thấy sheet" | Tên sheet khác | Sửa `CONFIG.SHEETS_HE_THONG` |
| Công thức hiện `#REF!` | Xóa/đổi tên sheet bản | Chạy lại "Tổng hợp BIỂU TỔNG" |
| "Không tìm thấy dòng Tổng số hộ" | Tên khác | Sửa `CONFIG.NOI_DUNG_TONG_SO_HO` |

---

## 📝 GHI CHÚ CHO ĐỒ ÁN

### Yêu cầu đã đáp ứng

✅ **Cơ bản:** Tổng hợp số liệu từ các bản → BIỂU TỔNG  
✅ **Nâng cao:** Thêm dòng → đồng bộ tất cả sheet, giữ chức năng tổng hợp  
✅ **"Tính vị trí theo tổng số dòng":** Dùng biến `newRow` và hàm `timDong()`  

### Giải thích kỹ thuật cho báo cáo

**Vấn đề:** Khi thêm dòng, số dòng của các ô thay đổi → công thức tham chiếu sai.

**Giải pháp:** 
1. Chèn dòng vào sheet bản TRƯỚC → Google Sheets tự điều chỉnh tham chiếu
2. Chèn dòng vào BIỂU TỔNG SAU → dòng mới trống
3. Đặt công thức cho dòng mới dùng `newRow` thực tế (không hardcode)

**Hàm quan trọng:**
```javascript
// Xây dựng công thức với số dòng ĐỘNG
function xayDungCongThucTongHop(row, colNumber, banNames) { ... }

// Tìm dòng theo NỘI DUNG (không dùng số cố định)
function timDong(sheet, noiDung) { ... }
```
