/**
 * ============================================================
 * CODE.GS - MENU & HÀM TIỆN ÍCH
 * ============================================================
 * Đây là file điều phối chính.
 * Gồm:
 *   1. onOpen()  - Tạo menu tùy chỉnh khi mở file
 *   2. Các hàm tiện ích dùng chung (helper functions)
 *   3. Các hàm gọi từ menu (wrapper functions)
 * ============================================================
 */

// ============================================================
// PHẦN 1: MENU
// ============================================================

/**
 * Chạy tự động khi người dùng mở Google Sheets.
 * Tạo menu "🏛️ Công cụ Xã" trên thanh menu.
 */
function onOpen() {
  var ui;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (err) {
    // Cho phép chạy onOpen từ context không có UI (Apps Script editor/trigger nền).
    Logger.log("[onOpen] UI not available in this context: " + err.message);
    return;
  }

  ui.createMenu("🏛️ Công cụ Xã")

    // Nhóm tổng hợp dữ liệu (Cơ bản)
    .addItem("📊 Tổng hợp BIỂU TỔNG từ các bản", "tongHopBieuTong")
    .addItem("📋 Cập nhật BIỂU TỔNG TOÀN XÃ", "tongHopToanXa")

    .addSeparator()

    // Nhóm quản lý dòng (Nâng cao)
    .addSubMenu(
      ui
        .createMenu("📝 Quản lý dòng chỉ tiêu")
        .addItem("➕ Thêm dòng mới (đồng bộ tất cả sheet)", "themDongDongBo")
        .addItem("❌ Xóa dòng     (đồng bộ tất cả sheet)", "xoaDongDongBo"),
    )

    .addSeparator()

    // Tiện ích
    .addItem("🔄 Xây dựng lại toàn bộ công thức", "xayDungLaiCongThuc")
    .addItem("🏘️ Xem danh sách sheet bản", "xemDanhSachBan")

    .addSeparator()

    .addItem("❓ Hướng dẫn sử dụng", "hienThiHuongDan")

    .addToUi();
}

// ============================================================
// PHẦN 2: HÀM TIỆN ÍCH DÙNG CHUNG
// ============================================================

/**
 * Lấy danh sách đối tượng Sheet của tất cả bản/tiểu khu.
 * Loại trừ các sheet hệ thống trong CONFIG.SHEETS_HE_THONG.
 *
 * @param {Spreadsheet} [ss] - Spreadsheet cần xử lý (mặc định: file hiện tại)
 * @returns {Sheet[]} Mảng Sheet
 */
function laySheetBan(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets().filter(function (sheet) {
    return CONFIG.SHEETS_HE_THONG.indexOf(sheet.getName()) === -1;
  });
}

/**
 * Lấy danh sách TÊN của tất cả sheet bản/tiểu khu.
 *
 * @param {Spreadsheet} [ss]
 * @returns {string[]} Mảng tên sheet
 */
function layTenSheetBan(ss) {
  return laySheetBan(ss).map(function (s) {
    return s.getName();
  });
}

/**
 * Xây dựng công thức tổng hợp cho một ô cụ thể trong BIỂU TỔNG.
 *
 * Kết quả dạng: ='BẢN NÀ LỐC'!D9+'BẢN NONG HEO'!D9+...+'BẢN CÁT LÌNH'!D9
 *
 * Đây là cốt lõi của "tính vị trí theo tổng số dòng":
 * tham số row là số dòng THỰC TẾ (tính động), không phải hằng số cố định.
 *
 * @param {number} row        - Số dòng thực tế trong sheet
 * @param {number} colNumber  - Số cột (4=D, 5=E, 6=F, 7=G)
 * @param {string[]} banNames - Danh sách tên sheet bản
 * @returns {string} Công thức tổng hợp
 *
 * @example
 * xayDungCongThucTongHop(9, 4, ['BẢN NÀ LỐC','BẢN NONG HEO'])
 * // => "='BẢN NÀ LỐC'!D9+'BẢN NONG HEO'!D9"
 */
function xayDungCongThucTongHop(row, colNumber, banNames) {
  // Chuyển số cột → chữ cột: 4→'D', 5→'E', 6→'F', 7→'G'
  var colLetter = String.fromCharCode(64 + colNumber);

  var parts = banNames.map(function (name) {
    return "'" + name + "'!" + colLetter + row;
  });

  return "=" + parts.join("+");
}

/**
 * Xây dựng công thức Luỹ kế: H = D + E + F + G
 *
 * @param {number} row - Số dòng thực tế
 * @returns {string} Công thức luỹ kế
 *
 * @example
 * xayDungCongThucLuyKe(9) // => "=D9+E9+F9+G9"
 */
function xayDungCongThucLuyKe(row) {
  return "=D" + row + "+E" + row + "+F" + row + "+G" + row;
}

/**
 * Kiểm tra xem một chuỗi có phải là công thức tổng hợp không.
 * Công thức tổng hợp có dạng: ='TÊN_SHEET'!Cột Dòng+...
 *
 * @param {string} formula - Chuỗi cần kiểm tra
 * @returns {boolean}
 *
 * @example
 * laFormulaTongHop("='BẢN NÀ LỐC'!D9+...")  // => true
 * laFormulaTongHop("=D9+E9+F9+G9")           // => false
 * laFormulaTongHop("")                        // => false
 */
function laFormulaTongHop(formula) {
  return (
    formula !== null &&
    formula !== undefined &&
    typeof formula === "string" &&
    formula.startsWith("='") &&
    formula.indexOf("'!") > 0
  );
}

/**
 * Tìm số dòng của một chỉ tiêu dựa theo NỘI DUNG cột B.
 *
 * ★ ĐÂY LÀ TRÁI TIM CỦA "TÍNH VỊ TRÍ THEO TỔNG SỐ DÒNG" ★
 *
 * Thay vì dùng số dòng cố định (ví dụ: "Tổng số hộ luôn ở dòng 9"),
 * hàm này TÌM KIẾM dòng đó theo NỘI DUNG.
 * → Khi thêm/xóa dòng, hàm vẫn tìm đúng vị trí.
 *
 * @param {Sheet}  sheet    - Sheet cần tìm trong
 * @param {string} noiDung  - Nội dung cần tìm (cột B)
 * @param {number} [startRow] - Dòng bắt đầu tìm (mặc định: DONG_DU_LIEU_BAT_DAU)
 * @param {number} [col]      - Cột tìm kiếm (mặc định: COT_NOI_DUNG = cột B)
 * @returns {number} Số dòng tìm thấy, hoặc -1 nếu không có
 *
 * @example
 * var dong = timDong(sheetBan, 'Tổng số hộ');
 * // dong = 9 (hoặc bất kỳ dòng nào chứa "Tổng số hộ")
 */
function timDong(sheet, noiDung, startRow, col) {
  col = col || CONFIG.COT_NOI_DUNG;
  startRow = startRow || CONFIG.DONG_DU_LIEU_BAT_DAU;

  var lastRow = sheet.getLastRow();
  if (lastRow < startRow) return -1;

  // Đọc toàn bộ cột B một lần (batch read)
  var values = sheet
    .getRange(startRow, col, lastRow - startRow + 1, 1)
    .getValues();

  for (var i = 0; i < values.length; i++) {
    var cellVal = values[i][0];
    if (cellVal && cellVal.toString().trim() === noiDung.trim()) {
      return startRow + i; // Trả về số dòng THỰC TẾ
    }
  }

  return -1; // Không tìm thấy
}

// ============================================================
// PHẦN 3: HÀM GỌI TỪ MENU
// ============================================================

/**
 * Hiển thị danh sách tất cả sheet bản
 */
function xemDanhSachBan() {
  var banNames = layTenSheetBan();
  var msg = "🏘️ Danh sách " + banNames.length + " bản/tiểu khu:\n\n";
  banNames.forEach(function (name, i) {
    msg += i + 1 + ". " + name + "\n";
  });
  SpreadsheetApp.getUi().alert(msg);
}

/**
 * Xây dựng lại toàn bộ công thức (với xác nhận trước)
 */
function xayDungLaiCongThuc() {
  var ui = SpreadsheetApp.getUi();
  var confirm = ui.alert(
    "🔄 Xây dựng lại công thức",
    "Thao tác này sẽ REBUILD TẤT CẢ công thức trong BIỂU TỔNG.\n\n" +
      "• Hữu ích khi thêm sheet bản mới vào file\n" +
      "• Hữu ích khi công thức bị lỗi\n\n" +
      "Bạn có muốn tiếp tục?",
    ui.ButtonSet.YES_NO,
  );
  if (confirm === ui.Button.YES) {
    tongHopBieuTong();
  }
}

/**
 * Hiển thị hướng dẫn sử dụng đầy đủ
 */
function hienThiHuongDan() {
  var msg =
    "📖 HƯỚNG DẪN SỬ DỤNG CÔNG CỤ XÃ\n" +
    "═══════════════════════════════════\n\n" +
    "1. 📊 TỔNG HỢP BIỂU TỔNG\n" +
    "   Dùng sau khi các bản đã nhập xong số liệu.\n" +
    "   Script tự động tính tổng từ tất cả sheet bản\n" +
    "   và cập nhật công thức trong BIỂU TỔNG.\n\n" +
    "2. 📋 CẬP NHẬT BIỂU TỔNG TOÀN XÃ\n" +
    "   Cập nhật số hộ và nhân khẩu theo từng bản\n" +
    "   (kéo dữ liệu từ cột H của sheet bản tương ứng).\n\n" +
    "3. ➕ THÊM DÒNG MỚI\n" +
    "   Khi cần thêm chỉ tiêu báo cáo mới:\n" +
    "   → Chèn dòng vào BIỂU TỔNG VÀ tất cả sheet bản\n" +
    "   → Công thức tổng hợp tự động thiết lập\n" +
    "   → Bạn chỉ cần điền Nội dung và Đơn vị tính\n\n" +
    "4. ❌ XÓA DÒNG\n" +
    "   Xóa chỉ tiêu khỏi tất cả sheet đồng bộ.\n" +
    "   ⚠️ Không thể hoàn tác!\n\n" +
    "5. 🔄 XÂY DỰNG LẠI CÔNG THỨC\n" +
    "   Dùng khi:\n" +
    "   → Thêm sheet bản mới vào file\n" +
    "   → Công thức bị hỏng/lỗi\n\n" +
    "═══════════════════════════════════\n" +
    "Ghi chú: Dữ liệu nhập vào cột D-G (màu vàng)\n" +
    "trong các SHEET BẢN. BIỂU TỔNG tự động tổng hợp.";

  SpreadsheetApp.getUi().alert(msg);
}
