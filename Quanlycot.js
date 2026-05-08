/**
 * ============================================================
 * QUANLYCOT.GS - QUẢN LÝ CỘT CHỈ TIÊU (TÍNH NĂNG NÂNG CAO)
 * ============================================================
 * Đây là phần NÂNG CAO mở rộng của đồ án:
 *   "Khi thêm 1 cột tháng tại BIỂU TỔNG thì các sheet bản
 *    cũng thêm cột tương ứng. Sau khi thêm cột thì vẫn giữ
 *    được chức năng tổng hợp."
 *
 * ────────────────────────────────────────────────────────────
 * NGUYÊN LÝ CỐT LÕI (tương tự QuanLyDong.gs):
 * ────────────────────────────────────────────────────────────
 *
 * THÊM CỘT (themCotDongBo):
 *   BƯỚC 1 – Chèn vào TẤT CẢ sheet bản trước.
 *     → Google Sheets tự động cập nhật công thức tham chiếu
 *       trong BIỂU TỔNG (các cột phía sau điểm chèn tự +1).
 *   BƯỚC 2 – Chèn vào BIỂU TỔNG.
 *   BƯỚC 3 – Đặt tiêu đề cột mới (dòng 6) cho tất cả sheet.
 *   BƯỚC 4 – Đặt công thức tổng hợp cho cột mới trong BIỂU TỔNG.
 *   BƯỚC 5 – Rebuild công thức Luỹ kế (cộng thêm cột mới).
 *
 * XÓA CỘT (xoaCotDongBo):
 *   BƯỚC 1 – Xóa từ TẤT CẢ sheet bản trước.
 *     → Google Sheets auto-adjust công thức trong BIỂU TỔNG.
 *     → Luỹ kế có thể hiện #REF! tạm thời.
 *   BƯỚC 2 – Xóa từ BIỂU TỔNG.
 *   BƯỚC 3 – Rebuild công thức Luỹ kế (loại bỏ cột đã xóa).
 *
 * ────────────────────────────────────────────────────────────
 * TÍNH VỊ TRÍ ĐỘNG (không dùng hằng số cố định):
 * ────────────────────────────────────────────────────────────
 * Thay vì dùng CONFIG.COT_LUY_KE = 8 (có thể lỗi thời sau khi
 * thêm cột), hệ thống TÌM cột Luỹ kế theo NỘI DUNG tiêu đề
 * → Dù thêm bao nhiêu cột tháng, hệ thống vẫn tìm đúng vị trí.
 * Đây là hàm layCotLuyKe() và timCot() bên dưới.
 * ────────────────────────────────────────────────────────────
 *
 *  GHI CHÚ VỀ Ô HỢP NHẤT (Merged Cells):
 * ────────────────────────────────────────────────────────────
 * Google Sheets tự động MỞ RỘNG ô hợp nhất khi chèn cột vào
 * GIỮA vùng hợp nhất. Tuy nhiên nếu ô hợp nhất KẾT THÚC trước
 * vị trí chèn (ví dụ: nhóm tiêu đề "Các tháng báo cáo" hợp
 * nhất T10-T12 nhưng cột mới chèn SAU T12), ô đó không mở rộng.
 * → Script sẽ THÔNG BÁO sau khi thêm cột để người dùng kiểm tra
 *   và điều chỉnh thủ công các ô tiêu đề nhóm (rows 1–5) nếu cần.
 * ────────────────────────────────────────────────────────────
 */

// ============================================================
// PHẦN 1: HÀM TIỆN ÍCH DÙNG CHUNG CHO QUẢN LÝ CỘT
// ============================================================

/**
 * Tìm số cột dựa theo NỘI DUNG của ô trong dòng tiêu đề.
 *
 * ★ Đây là "tìm vị trí CỘT theo nội dung" — tương tự timDong()
 *   nhưng theo chiều ngang thay vì chiều dọc. ★
 *
 * Thay vì dùng hằng số cố định (như COT_LUY_KE = 8),
 * hàm này TÌM KIẾM cột bằng cách đọc nội dung dòng tiêu đề.
 * → Sau khi thêm/xóa cột, hàm vẫn tìm đúng vị trí.
 *
 * @param {Sheet}  sheet      - Sheet cần tìm trong
 * @param {string} headerText - Nội dung tiêu đề cần tìm (khớp chính xác)
 * @param {number} [headerRow]- Dòng tiêu đề (mặc định: CONFIG.DONG_HEADER)
 * @returns {number} Số cột (1-based), hoặc -1 nếu không tìm thấy
 */
function timCot(sheet, headerText, headerRow) {
  headerRow = headerRow || CONFIG.DONG_HEADER;

  if (!headerText || typeof headerText !== "string") {
    Logger.log("[timCot] headerText rỗng hoặc không hợp lệ.");
    return -1;
  }

  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return -1;

  var values = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];

  for (var i = 0; i < values.length; i++) {
    var cellVal = values[i];
    if (cellVal && cellVal.toString().trim() === headerText.trim()) {
      return i + 1; // 1-based
    }
  }
  return -1;
}

/**
 * Lấy vị trí cột Luỹ kế trong sheet.
 *
 * Tìm theo các biến thể tên tiêu đề thường dùng.
 * Fallback về CONFIG.COT_LUY_KE nếu không tìm thấy.
 *
 * @param {Sheet} sheet - Sheet cần tìm
 * @returns {number} Số cột Luỹ kế (1-based)
 */
function layCotLuyKe(sheet) {
  // Thử lần lượt các biến thể tiêu đề phổ biến
  var candidates = [
    "Luỹ kế từ đầu năm",
    "Lũy kế từ đầu năm",
    "Luỹ kế",
    "Lũy kế",
    "Luy ke tu dau nam",
  ];
  for (var i = 0; i < candidates.length; i++) {
    var col = timCot(sheet, candidates[i]);
    if (col > 0) return col;
  }
  // Fallback: dùng hằng số trong CONFIG
  Logger.log(
    '[layCotLuyKe] Không tìm thấy tiêu đề Luỹ kế trong sheet "' +
      sheet.getName() +
      '". Dùng CONFIG.COT_LUY_KE = ' +
      CONFIG.COT_LUY_KE,
  );
  return CONFIG.COT_LUY_KE;
}

/**
 * Xây dựng công thức Luỹ kế ĐỘNG.
 *
 * Thay vì hardcode "=D+E+F+G", hàm này cộng TẤT CẢ cột
 * từ đầu kỳ (dauKyCol) đến ngay trước Luỹ kế (luyKeCol - 1).
 *
 * Ví dụ:
 *   dauKyCol=4, luyKeCol=8  → "=D9+E9+F9+G9"   (cấu trúc gốc)
 *   dauKyCol=4, luyKeCol=9  → "=D9+E9+F9+G9+H9" (sau khi thêm 1 cột)
 *   dauKyCol=4, luyKeCol=10 → "=D9+E9+F9+G9+H9+I9" (thêm 2 cột)
 *
 * ★ Đây chính là "tính vị trí theo tổng số cột" ★
 *
 * @param {number} row        - Số dòng thực tế
 * @param {number} dauKyCol   - Số cột đầu kỳ (= CONFIG.COT_DAU_KY = 4)
 * @param {number} luyKeCol   - Số cột Luỹ kế THỰC TẾ (tìm động)
 * @returns {string} Công thức luỹ kế động
 */
function xayDungCongThucLuyKeDong(row, dauKyCol, luyKeCol) {
  var parts = [];
  for (var c = dauKyCol; c < luyKeCol; c++) {
    parts.push(String.fromCharCode(64 + c) + row);
  }
  if (parts.length === 0) return "";
  return "=" + parts.join("+");
}

/**
 * Lấy chữ cột (A, B, C... Z, AA, AB...) từ số cột.
 * Hỗ trợ cột hai chữ cái (AA trở lên) để tương thích sheet lớn.
 *
 * @param {number} colNum - Số cột (1-based)
 * @returns {string} Chữ cột (ví dụ: 1→"A", 26→"Z", 27→"AA")
 */
function soSangChuCot(colNum) {
  var result = "";
  while (colNum > 0) {
    var remainder = (colNum - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    colNum = Math.floor((colNum - 1) / 26);
  }
  return result;
}

// ============================================================
// PHẦN 2: THÊM CỘT THÁNG MỚI
// ============================================================

/**
 * THÊM CỘT THÁNG MỚI đồng bộ vào BIỂU TỔNG và tất cả sheet bản.
 *
 * ────────────────────────────────────────────────────────────
 * QUY TẮC VỊ TRÍ CHÈN:
 * ────────────────────────────────────────────────────────────
 * Cột mới luôn được chèn VÀO TRƯỚC cột Luỹ kế.
 * Lý do: Đây là nơi tự nhiên nhất cho cột tháng báo cáo mới.
 * Sau khi chèn, Luỹ kế tự động dịch sang phải 1 cột.
 *
 * ────────────────────────────────────────────────────────────
 * VÍ DỤ:
 * ────────────────────────────────────────────────────────────
 * Trước: D(Đầu kỳ) | E(T10) | F(T11) | G(T12) | H(Luỹ kế) | I(Ghi chú)
 * Sau:   D(Đầu kỳ) | E(T10) | F(T11) | G(T12) | H(T1) [MỚI] | I(Luỹ kế) | J(Ghi chú)
 * Luỹ kế mới: =D+E+F+G+H (tự cộng thêm cột mới)
 */
function themCotDongBo() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetTong = ss.getSheetByName(CONFIG.SHEET_BIEU_TONG);

  if (!sheetTong) {
    ui.alert(' Không tìm thấy sheet "' + CONFIG.SHEET_BIEU_TONG + '"!');
    return;
  }

  // ── Tìm vị trí cột Luỹ kế ĐỘNG (không dùng hằng số) ──
  var cotLuyKe = layCotLuyKe(sheetTong);
  var cotDauKy = CONFIG.COT_DAU_KY; // Cột D = 4 (ít khi thay đổi)
  var soThangHT = cotLuyKe - cotDauKy; // Số cột dữ liệu hiện có (kể cả Đầu kỳ)

  // Cột mới chèn VÀO TRƯỚC vị trí Luỹ kế
  var insertBeforeCol = cotLuyKe; // chèn trước đây
  var newColNum = cotLuyKe; // sau khi chèn, cột mới ở đây
  var newLuyKeCol = cotLuyKe + 1; // Luỹ kế dịch sang phải 1

  // ── Hỏi tên tiêu đề cột mới ──
  var response = ui.prompt(
    "➕ Thêm cột tháng mới",
    "Cấu trúc hiện tại:\n" +
      "  Đầu kỳ (cột " +
      soSangChuCot(cotDauKy) +
      ") | " +
      "... " +
      (soThangHT - 1) +
      " tháng ... | " +
      "Luỹ kế (cột " +
      soSangChuCot(cotLuyKe) +
      ")\n\n" +
      "Cột mới sẽ chèn TRƯỚC Luỹ kế (cột " +
      soSangChuCot(insertBeforeCol) +
      ").\n\n" +
      "Nhập tên TIÊU ĐỀ cho cột mới:\n" +
      '(Ví dụ: "Tháng 1", "Tháng 2", "Tháng 9"...)',
    ui.ButtonSet.OK_CANCEL,
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  var tieuDeMoi = response.getResponseText().trim();
  if (!tieuDeMoi) {
    ui.alert(" Tên tiêu đề không được để trống!");
    return;
  }

  // ── Xác nhận trước khi thực hiện ──
  var banSheets = laySheetBan(ss);
  var confirm = ui.alert(
    " Xác nhận thêm cột",
    'Cột mới: "' +
      tieuDeMoi +
      '"\n' +
      "Vị trí:  Cột " +
      soSangChuCot(newColNum) +
      " (trước Luỹ kế)\n" +
      "Đồng bộ: BIỂU TỔNG + " +
      banSheets.length +
      " sheet bản\n\n" +
      " LƯU Ý: Kiểm tra ô hợp nhất ở các dòng tiêu đề nhóm\n" +
      "   (dòng 1–5) sau khi thêm cột nếu cần.\n\n" +
      "Tiếp tục?",
    ui.ButtonSet.YES_NO,
  );
  if (confirm !== ui.Button.YES) return;

  var banNames = banSheets.map(function (s) {
    return s.getName();
  });

  SpreadsheetApp.getActive().toast(
    "Đang chèn cột vào " + banSheets.length + " sheet bản...",
    " Đang xử lý",
    20,
  );

  try {
    // ════════════════════════════════════════════════════════
    // BƯỚC 1: Chèn cột vào TẤT CẢ sheet bản TRƯỚC
    // → Google Sheets tự động cập nhật tham chiếu công thức
    //   trong BIỂU TỔNG (='BẢN X'!G9 → ='BẢN X'!H9, v.v.)
    // ════════════════════════════════════════════════════════
    banSheets.forEach(function (sheet) {
      sheet.insertColumnBefore(insertBeforeCol);

      // Sao chép ĐỊNH DẠNG từ cột liền trái (cột cuối cùng của tháng cũ)
      // Giúp kế thừa: màu nền, font, border, căn lề, độ rộng cột...
      var refCol = insertBeforeCol - 1; // Cột T12 (hoặc cột dữ liệu cuối)
      sheet
        .getRange(1, refCol, sheet.getLastRow(), 1)
        .copyTo(
          sheet.getRange(1, newColNum, sheet.getLastRow(), 1),
          SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
          false,
        );

      // Xóa nội dung (giữ định dạng)
      sheet.getRange(1, newColNum, sheet.getLastRow(), 1).clearContent();

      // Đặt tiêu đề cho cột mới trong sheet bản
      sheet.getRange(CONFIG.DONG_HEADER, newColNum).setValue(tieuDeMoi);
    });

    // ════════════════════════════════════════════════════════
    // BƯỚC 2: Chèn cột vào BIỂU TỔNG
    // Lúc này tất cả công thức tham chiếu đã được auto-adjust
    // ════════════════════════════════════════════════════════
    sheetTong.insertColumnBefore(insertBeforeCol);

    // Sao chép định dạng từ cột T12
    sheetTong
      .getRange(1, insertBeforeCol - 1, sheetTong.getLastRow(), 1)
      .copyTo(
        sheetTong.getRange(1, newColNum, sheetTong.getLastRow(), 1),
        SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
        false,
      );
    sheetTong.getRange(1, newColNum, sheetTong.getLastRow(), 1).clearContent();

    // Đặt tiêu đề trong BIỂU TỔNG
    sheetTong.getRange(CONFIG.DONG_HEADER, newColNum).setValue(tieuDeMoi);

    // ════════════════════════════════════════════════════════
    // BƯỚC 3: Đặt công thức tổng hợp cho cột mới trong BIỂU TỔNG
    //
    // Đọc batch toàn bộ công thức cột tham chiếu để phát hiện
    // dòng nào có công thức tổng hợp (cần đặt công thức) và
    // dòng nào là tiêu đề/phần nhóm (bỏ qua).
    //
    // "Tính vị trí theo tổng số cột": dùng newColNum thực tế.
    // ════════════════════════════════════════════════════════
    var dataStart = CONFIG.DONG_DU_LIEU_BAT_DAU;
    var lastRow = sheetTong.getLastRow();
    var numDataRows = lastRow - dataStart + 1;

    if (numDataRows > 0) {
      // Đọc công thức từ cột tham chiếu (cột T12 cũ = newColNum - 1)
      // để xác định dòng nào là dòng dữ liệu (có công thức tổng hợp)
      var refFormulas = sheetTong
        .getRange(dataStart, newColNum - 1, numDataRows, 1)
        .getFormulas();

      // Xây dựng mảng công thức cho cột mới (batch write)
      var newColFormulas = refFormulas.map(function (rowArr, idx) {
        var actualRow = dataStart + idx;
        var refFormula = rowArr[0];
        // Chỉ đặt công thức nếu cột tham chiếu đã có công thức tổng hợp
        if (laFormulaTongHop(refFormula)) {
          return [xayDungCongThucTongHop(actualRow, newColNum, banNames)];
        }
        return [""]; // Dòng tiêu đề/phần → bỏ trống
      });

      sheetTong
        .getRange(dataStart, newColNum, numDataRows, 1)
        .setFormulas(newColFormulas);
    }

    // ════════════════════════════════════════════════════════
    // BƯỚC 4: Rebuild công thức Luỹ kế
    //
    // Sau khi chèn cột, Luỹ kế đã dịch sang cột newLuyKeCol.
    // Công thức cũ (=D+E+F+G) chưa cộng cột mới (H).
    // Cần rebuild thành =D+E+F+G+H (tính động qua xayDungCongThucLuyKeDong).
    //
    // Thực hiện cho CẢ BIỂU TỔNG và tất cả sheet bản.
    // ════════════════════════════════════════════════════════

    // -- 4a. Rebuild Luỹ kế trong BIỂU TỔNG --
    if (numDataRows > 0) {
      var luyKeFormulasOld = sheetTong
        .getRange(dataStart, newLuyKeCol, numDataRows, 1)
        .getFormulas();

      var luyKeFormulasNew = luyKeFormulasOld.map(function (rowArr, idx) {
        var actualRow = dataStart + idx;
        // Rebuild nếu ô đang có công thức (dù có thể đã sai sau khi dịch cột)
        if (rowArr[0] !== "") {
          return [xayDungCongThucLuyKeDong(actualRow, cotDauKy, newLuyKeCol)];
        }
        return [""]; // Dòng không có công thức → bỏ qua
      });

      sheetTong
        .getRange(dataStart, newLuyKeCol, numDataRows, 1)
        .setFormulas(luyKeFormulasNew);
    }

    // -- 4b. Rebuild Luỹ kế trong tất cả sheet bản --
    banSheets.forEach(function (sheet) {
      var banLastRow = sheet.getLastRow();
      var banNumRows = banLastRow - dataStart + 1;
      if (banNumRows <= 0) return;

      var banLuyKeOld = sheet
        .getRange(dataStart, newLuyKeCol, banNumRows, 1)
        .getFormulas();

      var banLuyKeNew = banLuyKeOld.map(function (rowArr, idx) {
        var actualRow = dataStart + idx;
        if (rowArr[0] !== "") {
          return [xayDungCongThucLuyKeDong(actualRow, cotDauKy, newLuyKeCol)];
        }
        return [""];
      });

      sheet
        .getRange(dataStart, newLuyKeCol, banNumRows, 1)
        .setFormulas(banLuyKeNew);
    });

    // ── THÔNG BÁO KẾT QUẢ ──
    ui.alert(
      " THÊM CỘT THÀNH CÔNG!\n\n" +
        " Cột mới:        " +
        soSangChuCot(newColNum) +
        ' ("' +
        tieuDeMoi +
        '")\n' +
        "📊 Luỹ kế:         Đã cập nhật ở cột " +
        soSangChuCot(newLuyKeCol) +
        "\n" +
        " Sheet đồng bộ:  BIỂU TỔNG + " +
        banSheets.length +
        " bản\n\n" +
        "⚡ VIỆC CẦN LÀM TIẾP THEO:\n" +
        "→ Vào từng sheet bản, nhập số liệu cột " +
        soSangChuCot(newColNum) +
        "\n" +
        '→ Chạy "📊 Tổng hợp BIỂU TỔNG" để cập nhật\n' +
        "→ Kiểm tra ô hợp nhất ở dòng tiêu đề nhóm (dòng 1–5)\n" +
        "   nếu có nhóm tiêu đề cần mở rộng thêm 1 cột!",
    );
  } catch (e) {
    ui.alert(" LỖI KHI THÊM CỘT!\n\nChi tiết: " + e.message);
    Logger.log("[themCotDongBo] ERROR: " + e.message + "\n" + e.stack);
  }
}

// ============================================================
// PHẦN 3: XÓA CỘT THÁNG
// ============================================================

/**
 * XÓA CỘT THÁNG đồng bộ khỏi BIỂU TỔNG và tất cả sheet bản.
 *
 * ────────────────────────────────────────────────────────────
 * NGUYÊN LÝ:
 * ────────────────────────────────────────────────────────────
 * BƯỚC 1 – Xóa từ TẤT CẢ sheet bản trước.
 *   → Google Sheets auto-adjust công thức tham chiếu.
 *   → Cột Luỹ kế trong BIỂU TỔNG có thể hiện #REF! tạm thời.
 *
 * BƯỚC 2 – Xóa từ BIỂU TỔNG.
 *   → Công thức tổng hợp của cột bị xóa cũng được xóa.
 *
 * BƯỚC 3 – Rebuild Luỹ kế.
 *   → Vì công thức cũ có thể tham chiếu cột đã xóa,
 *     cần rebuild toàn bộ.
 *
 * ────────────────────────────────────────────────────────────
 * GIỚI HẠN AN TOÀN:
 * ────────────────────────────────────────────────────────────
 * Chỉ cho phép xóa cột nằm trong vùng dữ liệu tháng.
 * KHÔNG cho phép xóa: A(STT), B(Nội dung), C(ĐVT),
 *                     D(Đầu kỳ), cột Luỹ kế, cột Ghi chú.
 */
function xoaCotDongBo() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetTong = ss.getSheetByName(CONFIG.SHEET_BIEU_TONG);

  if (!sheetTong) {
    ui.alert(' Không tìm thấy sheet "' + CONFIG.SHEET_BIEU_TONG + '"!');
    return;
  }

  var cotLuyKe = layCotLuyKe(sheetTong);
  var cotDauKy = CONFIG.COT_DAU_KY;
  var cotGhiChu = cotLuyKe + 1; // Ghi chú luôn nằm sau Luỹ kế

  // Phạm vi hợp lệ: từ sau Đầu kỳ (E trở đi) đến trước Luỹ kế
  var cotMinXoa = cotDauKy + 1; // Cột E (T10)
  var cotMaxXoa = cotLuyKe - 1; // Cột G (T12 hoặc cột tháng cuối)

  if (cotMinXoa > cotMaxXoa) {
    ui.alert(
      " Không có cột tháng nào để xóa!\n" +
        "Hiện chỉ có cột Đầu kỳ và Luỹ kế.\n" +
        "Cần ít nhất 1 cột tháng ở giữa để thực hiện xóa.",
    );
    return;
  }

  // Đọc tiêu đề các cột tháng để hiển thị cho người dùng
  var allHeaders = sheetTong
    .getRange(CONFIG.DONG_HEADER, cotMinXoa, 1, cotMaxXoa - cotMinXoa + 1)
    .getValues()[0];

  var danhSachCot = allHeaders
    .map(function (h, idx) {
      var colNum = cotMinXoa + idx;
      return "  Cột " + soSangChuCot(colNum) + ": " + (h || "(trống)");
    })
    .join("\n");

  // Hỏi người dùng cột cần xóa
  var response = ui.prompt(
    " Xóa cột tháng",
    " CẢNH BÁO: Thao tác này KHÔNG THỂ HOÀN TÁC!\n\n" +
      "Các cột tháng hiện có (có thể xóa):\n" +
      danhSachCot +
      "\n\n" +
      "Nhập CHỮ CỘT hoặc SỐ CỘT muốn xóa:\n" +
      '(Ví dụ: "E" hoặc "5")',
    ui.ButtonSet.OK_CANCEL,
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  var input = response.getResponseText().trim().toUpperCase();

  // Chuyển chữ cột → số cột
  var deleteCol;
  if (/^[A-Z]+$/.test(input)) {
    // Chữ cột (A, B, ... Z, AA, ...)
    deleteCol = 0;
    for (var i = 0; i < input.length; i++) {
      deleteCol = deleteCol * 26 + (input.charCodeAt(i) - 64);
    }
  } else if (/^\d+$/.test(input)) {
    deleteCol = parseInt(input);
  } else {
    ui.alert(
      " Giá trị không hợp lệ! Nhập chữ cột (E, F...) hoặc số cột (5, 6...).",
    );
    return;
  }

  // Kiểm tra phạm vi hợp lệ
  if (deleteCol < cotMinXoa || deleteCol > cotMaxXoa) {
    ui.alert(
      " Không thể xóa cột " +
        soSangChuCot(deleteCol) +
        "!\n\n" +
        "Chỉ được xóa cột trong phạm vi:\n" +
        "Cột " +
        soSangChuCot(cotMinXoa) +
        " đến cột " +
        soSangChuCot(cotMaxXoa) +
        "\n" +
        "(Các cột tháng báo cáo)",
    );
    return;
  }

  // Lấy tiêu đề cột cần xóa để hiển thị xác nhận
  var tieuDeCotXoa =
    sheetTong.getRange(CONFIG.DONG_HEADER, deleteCol).getValue() || "(trống)";

  // Xác nhận lần 2
  var banSheets = laySheetBan(ss);
  var confirm = ui.alert(
    " XÁC NHẬN XÓA CỘT " + soSangChuCot(deleteCol),
    'Cột sẽ bị xóa: "' +
      tieuDeCotXoa +
      '" (cột ' +
      soSangChuCot(deleteCol) +
      ")\n\n" +
      "Xóa khỏi BIỂU TỔNG VÀ tất cả " +
      banSheets.length +
      " sheet bản!\n\n" +
      "Bạn có chắc chắn không?",
    ui.ButtonSet.YES_NO,
  );
  if (confirm !== ui.Button.YES) {
    ui.alert("Đã hủy thao tác xóa.");
    return;
  }

  SpreadsheetApp.getActive().toast("Đang xóa cột...", " Đang xử lý", 15);

  try {
    // ════════════════════════════════════════════════════════
    // BƯỚC 1: Xóa từ TẤT CẢ sheet bản TRƯỚC
    // → Google Sheets auto-adjust tham chiếu công thức trong BIỂU TỔNG
    // ════════════════════════════════════════════════════════
    banSheets.forEach(function (sheet) {
      if (deleteCol <= sheet.getLastColumn()) {
        sheet.deleteColumn(deleteCol);
      }
    });

    // ════════════════════════════════════════════════════════
    // BƯỚC 2: Xóa từ BIỂU TỔNG
    // Cột này có thể đang hiển thị #REF! (do bước 1)
    // ════════════════════════════════════════════════════════
    if (deleteCol <= sheetTong.getLastColumn()) {
      sheetTong.deleteColumn(deleteCol);
    }

    // ════════════════════════════════════════════════════════
    // BƯỚC 3: Rebuild công thức Luỹ kế
    //
    // Sau khi xóa cột, vị trí Luỹ kế đã dịch sang trái 1.
    // Tìm lại vị trí Luỹ kế mới bằng layCotLuyKe().
    // Rebuild toàn bộ công thức Luỹ kế (loại bỏ cột đã xóa).
    // ════════════════════════════════════════════════════════
    var newLuyKeCol = layCotLuyKe(sheetTong); // Vị trí mới sau khi xóa

    var dataStart = CONFIG.DONG_DU_LIEU_BAT_DAU;
    var lastRow = sheetTong.getLastRow();
    var numDataRows = lastRow - dataStart + 1;

    if (numDataRows > 0) {
      // Rebuild Luỹ kế trong BIỂU TỔNG
      var luyKeFormulas = sheetTong
        .getRange(dataStart, newLuyKeCol, numDataRows, 1)
        .getFormulas();

      var rebuiltFormulas = luyKeFormulas.map(function (rowArr, idx) {
        var actualRow = dataStart + idx;
        // Rebuild nếu ô đang có nội dung (công thức hoặc giá trị)
        var cellVal = sheetTong.getRange(actualRow, newLuyKeCol).getValue();
        if (rowArr[0] !== "" || cellVal !== "") {
          return [xayDungCongThucLuyKeDong(actualRow, cotDauKy, newLuyKeCol)];
        }
        return [""];
      });

      sheetTong
        .getRange(dataStart, newLuyKeCol, numDataRows, 1)
        .setFormulas(rebuiltFormulas);

      // Rebuild Luỹ kế trong tất cả sheet bản
      banSheets.forEach(function (sheet) {
        var banLastRow = sheet.getLastRow();
        var banNumRows = banLastRow - dataStart + 1;
        if (banNumRows <= 0) return;

        var banLuyKe = sheet
          .getRange(dataStart, newLuyKeCol, banNumRows, 1)
          .getFormulas();

        var banRebuilt = banLuyKe.map(function (rowArr, idx) {
          var actualRow = dataStart + idx;
          if (rowArr[0] !== "") {
            return [xayDungCongThucLuyKeDong(actualRow, cotDauKy, newLuyKeCol)];
          }
          return [""];
        });

        sheet
          .getRange(dataStart, newLuyKeCol, banNumRows, 1)
          .setFormulas(banRebuilt);
      });
    }

    // ── THÔNG BÁO KẾT QUẢ ──
    ui.alert(
      " ĐÃ XÓA CỘT THÀNH CÔNG!\n\n" +
        '🗑️ Đã xóa cột "' +
        tieuDeCotXoa +
        '"\n' +
        "📊 Luỹ kế:        Đã rebuild ở cột " +
        soSangChuCot(newLuyKeCol) +
        "\n" +
        " Đồng bộ từ:    BIỂU TỔNG + " +
        banSheets.length +
        " sheet bản",
    );
  } catch (e) {
    ui.alert(" LỖI KHI XÓA CỘT!\n\nChi tiết: " + e.message);
    Logger.log("[xoaCotDongBo] ERROR: " + e.message + "\n" + e.stack);
  }
}

// ============================================================
// PHẦN 4: ĐỒNG BỘ TIÊU ĐỀ CỘT THỦ CÔNG
// ============================================================

/**
 * ĐỒNG BỘ TIÊU ĐỀ CỘT THỦ CÔNG
 * Sao chép toàn bộ dòng tiêu đề cột (dòng DONG_HEADER)
 * từ BIỂU TỔNG sang tất cả sheet bản.
 *
 * ────────────────────────────────────────────────────────────
 * KHI NÀO DÙNG?
 * ────────────────────────────────────────────────────────────
 * • Sau khi đổi tên tiêu đề cột trong BIỂU TỔNG
 * • Sau khi import dữ liệu hàng loạt không qua gõ tay
 * • Khi muốn đảm bảo tiêu đề 100% đồng nhất giữa các sheet
 * ────────────────────────────────────────────────────────────
 *
 * Phạm vi đồng bộ: từ cột D (Đầu kỳ) trở đi
 * (Giữ nguyên A, B, C vì đã được quản lý bởi dongBoNoiDungThuCong)
 */
function dongBoTieuDeCotThuCong() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetTong = ss.getSheetByName(CONFIG.SHEET_BIEU_TONG);

  if (!sheetTong) {
    ui.alert(' Không tìm thấy sheet "' + CONFIG.SHEET_BIEU_TONG + '"!');
    return;
  }

  var banSheets = laySheetBan(ss);
  if (banSheets.length === 0) {
    ui.alert(" Không tìm thấy sheet bản/tiểu khu nào!");
    return;
  }

  var confirm = ui.alert(
    " Đồng bộ tiêu đề cột sang tất cả sheet bản",
    "Thao tác này sẽ sao chép dòng tiêu đề cột (dòng " +
      CONFIG.DONG_HEADER +
      ")\n" +
      "từ cột D trở đi — từ BIỂU TỔNG → TẤT CẢ " +
      banSheets.length +
      " sheet bản.\n\n" +
      " Tiêu đề cột hiện có của các sheet bản SẼ BỊ GHI ĐÈ!\n\n" +
      "Bạn có muốn tiếp tục?",
    ui.ButtonSet.YES_NO,
  );
  if (confirm !== ui.Button.YES) {
    ui.alert("Đã hủy thao tác.");
    return;
  }

  try {
    var startCol = CONFIG.COT_DAU_KY; // Cột D
    var lastCol = sheetTong.getLastColumn();
    var numCols = lastCol - startCol + 1;

    if (numCols <= 0) {
      ui.alert(" BIỂU TỔNG không có cột dữ liệu (từ cột D trở đi)!");
      return;
    }

    // Đọc tiêu đề từ BIỂU TỔNG một lần
    var headerValues = sheetTong
      .getRange(CONFIG.DONG_HEADER, startCol, 1, numCols)
      .getValues();

    SpreadsheetApp.getActive().toast(
      "Đang đồng bộ tiêu đề sang " + banSheets.length + " sheet bản...",
      " Đang xử lý",
      10,
    );

    var updatedCount = 0;

    banSheets.forEach(function (sheet) {
      var banLastCol = sheet.getLastColumn();
      if (banLastCol < startCol) return;

      var writeCols = Math.min(numCols, banLastCol - startCol + 1);

      sheet
        .getRange(CONFIG.DONG_HEADER, startCol, 1, writeCols)
        .setValues([headerValues[0].slice(0, writeCols)]);

      updatedCount++;
    });

    ui.alert(
      " ĐỒNG BỘ TIÊU ĐỀ CỘT HOÀN TẤT!\n\n" +
        " Sheet bản đã cập nhật: " +
        updatedCount +
        "\n" +
        " Dòng tiêu đề:          Dòng " +
        CONFIG.DONG_HEADER +
        "\n" +
        " Cột đã đồng bộ:        " +
        soSangChuCot(startCol) +
        " → " +
        soSangChuCot(lastCol),
    );
  } catch (e) {
    ui.alert(" LỖI KHI ĐỒNG BỘ TIÊU ĐỀ CỘT!\n\nChi tiết: " + e.message);
    Logger.log("[dongBoTieuDeCotThuCong] ERROR: " + e.message + "\n" + e.stack);
  }
}
