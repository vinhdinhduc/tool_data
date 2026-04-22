/**
 * ============================================================
 * TONGHOP.GS - CHỨC NĂNG TỔNG HỢP DỮ LIỆU
 * ============================================================
 * Gồm 2 chức năng chính:
 *
 * 1. tongHopBieuTong()
 *    Tổng hợp số liệu từ tất cả sheet bản/tiểu khu
 *    vào sheet BIỂU TỔNG.
 *    → Rebuild toàn bộ công thức tổng hợp trong BIỂU TỔNG.
 *
 * 2. tongHopToanXa()
 *    Cập nhật sheet BIỂU TỔNG TOÀN XÃ:
 *    số hộ và nhân khẩu theo từng bản.
 * ============================================================
 */

/**
 * CHỨC NĂNG CHÍNH (Cơ bản):
 * Tổng hợp dữ liệu từ tất cả sheet bản → BIỂU TỔNG
 *
 * ────────────────────────────────────────────────────────────
 * NGUYÊN LÝ HOẠT ĐỘNG:
 * ────────────────────────────────────────────────────────────
 * BIỂU TỔNG có các ô công thức dạng:
 *   D9 = ='BẢN NÀ LỐC'!D9 + 'BẢN NONG HEO'!D9 + ... + 'BẢN CÁT LÌNH'!D9
 *
 * Hàm này SCAN toàn bộ dòng trong BIỂU TỔNG:
 *   - Nếu ô (row, D|E|F|G) hiện có công thức tổng hợp:
 *     → Rebuild lại với ĐÚNG tên sheet hiện có + ĐÚNG số dòng hiện tại
 *   - Nếu ô trống (dòng tiêu đề, dòng đặc biệt):
 *     → Bỏ qua
 *
 * "TÍNH VỊ TRÍ THEO TỔNG SỐ DÒNG":
 *   Công thức dùng số dòng THỰC TẾ (biến `actualRow`),
 *   không phải số cố định → Đúng cả khi thêm/xóa dòng.
 * ────────────────────────────────────────────────────────────
 */
function tongHopBieuTong() {
  var ui = SpreadsheetApp.getUi();

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetTong = ss.getSheetByName(CONFIG.SHEET_BIEU_TONG);

    // Kiểm tra sheet tồn tại
    if (!sheetTong) {
      ui.alert(
        '❌ Không tìm thấy sheet "' +
          CONFIG.SHEET_BIEU_TONG +
          '"!\n' +
          "Kiểm tra lại tên sheet trong CONFIG.SHEET_BIEU_TONG",
      );
      return;
    }

    // Lấy danh sách sheet bản
    var banNames = layTenSheetBan(ss);
    if (banNames.length === 0) {
      ui.alert(
        "❌ Không tìm thấy sheet bản/tiểu khu nào!\n" +
          "Kiểm tra lại CONFIG.SHEETS_HE_THONG",
      );
      return;
    }

    var dataStartRow = CONFIG.DONG_DU_LIEU_BAT_DAU;
    var lastRow = sheetTong.getLastRow();

    if (lastRow < dataStartRow) {
      ui.alert("❌ BIỂU TỔNG không có dữ liệu (lastRow=" + lastRow + ")!");
      return;
    }

    var numRows = lastRow - dataStartRow + 1;
    var formulaStartCol = CONFIG.COT_DAU_KY;
    var formulaColCount = CONFIG.COT_T12 - CONFIG.COT_DAU_KY + 1; // D..G

    // ── BATCH READ: Chỉ đọc công thức vùng tổng hợp D:G ──
    var formulaRange = sheetTong.getRange(
      dataStartRow,
      formulaStartCol,
      numRows,
      formulaColCount,
    );
    var formulas = formulaRange.getFormulas(); // Mảng 2D [row][col], index từ 0

    var updatedCount = 0;

    // ── XỬ LÝ TỪNG DÒNG ──
    for (var i = 0; i < formulas.length; i++) {
      // Số dòng THỰC TẾ trong Google Sheets (không phải index mảng)
      // Đây chính là "tính vị trí theo tổng số dòng"
      var actualRow = dataStartRow + i;

      // Kiểm tra cột D, E, F, G
      for (var j = 0; j < formulaColCount; j++) {
        var colNumber = formulaStartCol + j;
        var currentFormula = formulas[i][j];

        // Chỉ rebuild nếu ô hiện tại là công thức tổng hợp
        if (laFormulaTongHop(currentFormula)) {
          // Xây dựng công thức mới với:
          //   - Sheet names: lấy từ file HIỆN TẠI (không cố định)
          //   - Row number: dùng actualRow THỰC TẾ (không cố định)
          var newFormula = xayDungCongThucTongHop(
            actualRow,
            colNumber,
            banNames,
          );

          // Chỉ ghi lại ô có thay đổi để tránh ghi đè dữ liệu không phải công thức.
          if (currentFormula !== newFormula) {
            sheetTong.getRange(actualRow, colNumber).setFormula(newFormula);
            updatedCount++;
          }
        }
        // Nếu ô trống hoặc công thức đặc biệt → giữ nguyên
      }

      // Cột H (Luỹ kế) - index 7: KHÔNG cần rebuild
      // Google Sheets tự động cập nhật "=D9+E9+F9+G9" khi insert/delete rows
    }

    // Thông báo kết quả
    ui.alert(
      "✅ TỔNG HỢP BIỂU TỔNG HOÀN TẤT!\n\n" +
        "📊 Ô công thức đã cập nhật: " +
        updatedCount +
        "\n" +
        "🏘️ Số bản/tiểu khu:         " +
        banNames.length +
        "\n" +
        "📋 Số dòng dữ liệu:          " +
        numRows,
    );
  } catch (e) {
    ui.alert(
      "❌ LỖI KHI TỔNG HỢP!\n\n" +
        "Chi tiết: " +
        e.message +
        "\n\n" +
        "Vui lòng kiểm tra cấu trúc file và Config.gs",
    );
    Logger.log("[tongHopBieuTong] ERROR: " + e.message + "\nStack: " + e.stack);
  }
}

/**
 * CHỨC NĂNG PHỤ:
 * Cập nhật BIỂU TỔNG TOÀN XÃ từ số liệu sheet bản
 *
 * ────────────────────────────────────────────────────────────
 * NGUYÊN LÝ:
 * ────────────────────────────────────────────────────────────
 * - BIỂU TỔNG TOÀN XÃ liệt kê 54 bản theo thứ tự (dòng 5-58)
 * - Mỗi dòng cần: Số hộ + Số nhân khẩu của bản đó
 * - Lấy từ sheet bản tương ứng (cột H = Luỹ kế)
 *
 * Tìm dòng "Tổng số hộ" ĐỘNG (timDong), không dùng số cố định.
 * Thứ tự sheet bản = thứ tự dòng trong BIỂU TỔNG TOÀN XÃ.
 * ────────────────────────────────────────────────────────────
 */
function tongHopToanXa() {
  var ui = SpreadsheetApp.getUi();

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetTongXa = ss.getSheetByName(CONFIG.SHEET_TONG_TOAN_XA);

    if (!sheetTongXa) {
      ui.alert('❌ Không tìm thấy sheet "' + CONFIG.SHEET_TONG_TOAN_XA + '"!');
      return;
    }

    var banSheets = laySheetBan(ss);
    if (banSheets.length === 0) {
      ui.alert("❌ Không tìm thấy sheet bản/tiểu khu nào!");
      return;
    }

    // ── TÌM DÒNG CÁC CHỈ TIÊU THEO NỘI DUNG (không hardcode số dòng) ──
    // Đây là nguyên lý "tính vị trí theo tổng số dòng":
    // Dùng hàm timDong() để tìm vị trí dựa trên NỘI DUNG cột B
    var firstSheet = banSheets[0];
    var dongTongSoHo = timDong(firstSheet, CONFIG.NOI_DUNG_TONG_SO_HO);
    var dongNhanKhau = timDong(firstSheet, CONFIG.NOI_DUNG_NHAN_KHAU);

    if (dongTongSoHo === -1) {
      ui.alert(
        '❌ Không tìm thấy chỉ tiêu "' +
          CONFIG.NOI_DUNG_TONG_SO_HO +
          '"!\n\n' +
          "Kiểm tra lại:\n" +
          "• Giá trị CONFIG.NOI_DUNG_TONG_SO_HO\n" +
          "• Nội dung cột B trong sheet bản",
      );
      return;
    }

    // ── ĐẶT CÔNG THỨC THAM CHIẾU CHO TỪNG BẢN ──
    var startRow = CONFIG.TONG_XA_DONG_BAN_BAT_DAU;
    var updatedCount = 0;

    for (var i = 0; i < banSheets.length; i++) {
      var sheet = banSheets[i];
      var targetRow = startRow + i; // Dòng tương ứng trong BIỂU TỔNG TOÀN XÃ

      // Dùng CÔNG THỨC (không phải giá trị) để tự động cập nhật khi bản thay đổi
      // Dạng: ='TÊN SHEET'!H9
      sheetTongXa
        .getRange(targetRow, CONFIG.TONG_XA_COT_SO_HO)
        .setFormula("='" + sheet.getName() + "'!H" + dongTongSoHo);

      if (dongNhanKhau !== -1) {
        sheetTongXa
          .getRange(targetRow, CONFIG.TONG_XA_COT_NHAN_KHAU)
          .setFormula("='" + sheet.getName() + "'!H" + dongNhanKhau);
      }

      updatedCount++;
    }

    // Thông báo kết quả
    ui.alert(
      "✅ CẬP NHẬT BIỂU TỔNG TOÀN XÃ HOÀN TẤT!\n\n" +
        "🏘️ Số bản đã cập nhật:      " +
        updatedCount +
        "\n" +
        '📍 Dòng "Tổng số hộ":        ' +
        dongTongSoHo +
        "\n" +
        '📍 Dòng "Số nhân khẩu":      ' +
        (dongNhanKhau !== -1 ? dongNhanKhau : "❌ Không tìm thấy") +
        "\n\n" +
        "Ghi chú: Cột Hộ nghèo / Hộ cận nghèo\n" +
        "cần nhập thủ công (không có trong báo cáo tháng).",
    );
  } catch (e) {
    ui.alert("❌ LỖI: " + e.message);
    Logger.log("[tongHopToanXa] ERROR: " + e.message);
  }
}
