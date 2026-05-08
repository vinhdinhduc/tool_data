/**
 * ============================================================
 * QUANLYDONG.GS - QUẢN LÝ DÒNG CHỈ TIÊU (TÍNH NĂNG NÂNG CAO)
 * ============================================================
 * Đây là phần NÂNG CAO của đồ án:
 *   "Khi thêm 1 dòng tại sheet tổng hợp thì các sheet khác
 *    cũng thêm dòng. Sau khi thêm dòng thì vẫn giữ được
 *    chức năng tổng hợp."
 *
 * ────────────────────────────────────────────────────────────
 * NGUYÊN LÝ CỐT LÕI:
 * ────────────────────────────────────────────────────────────
 * Khi chèn dòng ở vị trí `insertAfterRow`:
 *
 * BƯỚC 1: Chèn vào TẤT CẢ sheet bản trước.
 *   → Google Sheets tự động cập nhật công thức trong BIỂU TỔNG:
 *     ví dụ: '!D9' → '!D10' (cho các dòng phía sau điểm chèn)
 *
 * BƯỚC 2: Chèn vào BIỂU TỔNG.
 *   → Dòng mới trống, chưa có công thức.
 *
 * BƯỚC 3: Đặt công thức cho dòng mới.
 *   Dùng biến `newRow` (= insertAfterRow + 1) = vị trí THỰC TẾ.
 *   → Công thức tham chiếu đúng dòng mới trong tất cả sheet bản.
 *
 * ĐÂY CHÍNH LÀ "TÍNH VỊ TRÍ THEO TỔNG SỐ DÒNG":
 *   Biến `newRow` được tính từ số dòng thực tế,
 *   không dùng hằng số cố định.
 * ────────────────────────────────────────────────────────────
 */

/**
 * THÊM DÒNG CHỈ TIÊU MỚI đồng bộ vào BIỂU TỔNG và tất cả sheet bản
 */
function themDongDongBo() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetTong = ss.getSheetByName(CONFIG.SHEET_BIEU_TONG);
  var currentLastRow = sheetTong.getLastRow();

  // ── BƯỚC 0: Hỏi người dùng vị trí chèn ──
  var response = ui.prompt(
    "➕ Thêm dòng chỉ tiêu mới",
    "Nhập số dòng muốn chèn VÀO SAU:\n\n" +
      "┌─────────────────────────────────┐\n" +
      "│ Dòng " +
      CONFIG.DONG_HEADER +
      ": Tiêu đề (Stt, Nội dung...) │\n" +
      "│ Dòng " +
      CONFIG.DONG_DU_LIEU_BAT_DAU +
      ": PHẦN I THÔNG TIN CHUNG    │\n" +
      "│ Dòng cuối hiện tại: " +
      currentLastRow +
      "           │\n" +
      "└─────────────────────────────────┘\n\n" +
      "Ví dụ: Nhập 9 → Chèn dòng mới giữa dòng 9 và 10",
    ui.ButtonSet.OK_CANCEL,
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  // ── KIỂM TRA ĐẦU VÀO ──
  var insertAfterRow = parseInt(response.getResponseText().trim());

  if (isNaN(insertAfterRow)) {
    ui.alert(" Giá trị không hợp lệ! Vui lòng nhập một số nguyên.");
    return;
  }
  if (insertAfterRow < CONFIG.DONG_DU_LIEU_BAT_DAU - 1) {
    ui.alert(
      " Không thể chèn trước dòng " +
        CONFIG.DONG_DU_LIEU_BAT_DAU +
        "!\n" +
        "Chỉ chèn từ dòng " +
        (CONFIG.DONG_DU_LIEU_BAT_DAU - 1) +
        " trở đi.",
    );
    return;
  }
  if (insertAfterRow > currentLastRow) {
    ui.alert(
      " Dòng " +
        insertAfterRow +
        " vượt quá số dòng hiện có (" +
        currentLastRow +
        ")!",
    );
    return;
  }

  // ── TÍNH VỊ TRÍ DÒNG MỚI (không phải hằng số) ──
  var newRow = insertAfterRow + 1; // Vị trí thực tế của dòng mới sau khi chèn
  var banSheets = laySheetBan(ss);
  var banNames = banSheets.map(function (s) {
    return s.getName();
  });

  // Hiển thị tiến độ
  SpreadsheetApp.getActive().toast(
    "Đang chèn dòng vào " + banSheets.length + " sheet bản...",
    " Đang xử lý",
    10,
  );

  try {
    // ════════════════════════════════════════════════════════
    // BƯỚC 1: Chèn dòng vào TẤT CẢ sheet bản
    // Làm TRƯỚC để Google Sheets auto-adjust công thức tham chiếu
    // trong BIỂU TỔNG (các dòng phía dưới sẽ tự +1)
    // ════════════════════════════════════════════════════════
    banSheets.forEach(function (sheet) {
      // Chèn 1 dòng trống SAU dòng insertAfterRow
      sheet.insertRowAfter(insertAfterRow);

      // Sao chép ĐỊNH DẠNG từ dòng trên (màu nền, font, border...)
      // Dùng PASTE_FORMAT để KHÔNG copy giá trị/công thức
      sheet
        .getRange(insertAfterRow, 1, 1, CONFIG.COT_GHI_CHU)
        .copyTo(
          sheet.getRange(newRow, 1, 1, CONFIG.COT_GHI_CHU),
          SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
          false,
        );

      // Xóa nội dung (chỉ giữ định dạng)
      sheet.getRange(newRow, 1, 1, CONFIG.COT_GHI_CHU).clearContent();

      // Đặt công thức Luỹ kế cho dòng mới trong sheet bản
      // "Tính vị trí theo tổng số dòng": dùng newRow thực tế
      sheet
        .getRange(newRow, CONFIG.COT_LUY_KE)
        .setFormula(xayDungCongThucLuyKe(newRow));
    });

    // ════════════════════════════════════════════════════════
    // BƯỚC 2: Chèn dòng vào BIỂU TỔNG
    // Lúc này, Google Sheets đã auto-adjust công thức của các
    // dòng phía dưới trong BIỂU TỔNG (tham chiếu đến sheet bản)
    // ════════════════════════════════════════════════════════
    sheetTong.insertRowAfter(insertAfterRow);

    // Sao chép định dạng từ dòng trên
    sheetTong
      .getRange(insertAfterRow, 1, 1, CONFIG.COT_GHI_CHU)
      .copyTo(
        sheetTong.getRange(newRow, 1, 1, CONFIG.COT_GHI_CHU),
        SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
        false,
      );
    sheetTong.getRange(newRow, 1, 1, CONFIG.COT_GHI_CHU).clearContent();

    // ════════════════════════════════════════════════════════
    // BƯỚC 3: Đặt công thức tổng hợp cho dòng mới trong BIỂU TỔNG
    //
    // Dùng biến `newRow` để xây dựng công thức:
    //   D{newRow} = ='BẢN NÀ LỐC'!D{newRow} + 'BẢN NONG HEO'!D{newRow} + ...
    //
    // Đây là "tính vị trí theo tổng số dòng" - newRow được tính động,
    // không phải con số cố định
    // ════════════════════════════════════════════════════════
    for (var col = CONFIG.COT_DAU_KY; col <= CONFIG.COT_T12; col++) {
      sheetTong
        .getRange(newRow, col)
        .setFormula(xayDungCongThucTongHop(newRow, col, banNames));
    }

    // Đặt công thức Luỹ kế cho BIỂU TỔNG
    sheetTong
      .getRange(newRow, CONFIG.COT_LUY_KE)
      .setFormula(xayDungCongThucLuyKe(newRow));

    // ── THÔNG BÁO KẾT QUẢ ──
    ui.alert(
      " THÊM DÒNG THÀNH CÔNG!\n\n" +
        " Dòng mới tại vị trí:   " +
        newRow +
        "\n" +
        " Sheet đã đồng bộ:       " +
        banSheets.length +
        " bản + BIỂU TỔNG\n\n" +
        "⚡ VIỆC CẦN LÀM TIẾP THEO:\n" +
        "→ Vào BIỂU TỔNG, dòng " +
        newRow +
        "\n" +
        '→ Nhập "Nội dung" vào cột B\n' +
        '→ Nhập "Đơn vị tính" vào cột C\n' +
        "→ Các sheet bản đã sẵn sàng (nhập D-G)\n" +
        "→ Cột H (Luỹ kế) tự động tính!",
    );
  } catch (e) {
    ui.alert(" LỖI KHI THÊM DÒNG!\n\nChi tiết: " + e.message);
    Logger.log("[themDongDongBo] ERROR: " + e.message + "\n" + e.stack);
  }
}

/**
 * XÓA DÒNG CHỈ TIÊU đồng bộ khỏi BIỂU TỔNG và tất cả sheet bản
 *
 * ────────────────────────────────────────────────────────────
 * NGUYÊN LÝ:
 * ────────────────────────────────────────────────────────────
 * BƯỚC 1: Xóa từ sheet bản trước.
 *   → Google Sheets auto-adjust công thức tham chiếu trong BIỂU TỔNG
 *   → Dòng xóa trong BIỂU TỔNG sẽ hiện #REF! (tham chiếu đến dòng đã xóa)
 *
 * BƯỚC 2: Xóa từ BIỂU TỔNG.
 *   → Xóa luôn dòng #REF! ra khỏi BIỂU TỔNG.
 * ────────────────────────────────────────────────────────────
 */
function xoaDongDongBo() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetTong = ss.getSheetByName(CONFIG.SHEET_BIEU_TONG);
  var currentLastRow = sheetTong.getLastRow();

  // ── BƯỚC 0: Hỏi người dùng ──
  var response = ui.prompt(
    " Xóa dòng chỉ tiêu",
    "  CẢNH BÁO: Thao tác này KHÔNG THỂ HOÀN TÁC!\n\n" +
      "Nhập số dòng muốn xóa:\n" +
      "(Phạm vi hợp lệ: " +
      CONFIG.DONG_DU_LIEU_BAT_DAU +
      " - " +
      currentLastRow +
      ")",
    ui.ButtonSet.OK_CANCEL,
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  // ── KIỂM TRA ĐẦU VÀO ──
  var deleteRow = parseInt(response.getResponseText().trim());

  if (isNaN(deleteRow)) {
    ui.alert(" Giá trị không hợp lệ!");
    return;
  }
  if (deleteRow < CONFIG.DONG_DU_LIEU_BAT_DAU) {
    ui.alert(
      " Không thể xóa dòng tiêu đề hoặc header!\n" +
        "Chỉ xóa từ dòng " +
        CONFIG.DONG_DU_LIEU_BAT_DAU +
        " trở đi.",
    );
    return;
  }
  if (deleteRow > currentLastRow) {
    ui.alert(
      " Dòng " +
        deleteRow +
        " vượt quá số dòng hiện có (" +
        currentLastRow +
        ")!",
    );
    return;
  }

  // Lấy nội dung dòng để xác nhận
  var noiDungDong = sheetTong
    .getRange(deleteRow, CONFIG.COT_NOI_DUNG)
    .getValue();
  var dvtDong = sheetTong.getRange(deleteRow, CONFIG.COT_DVT).getValue();
  var moTaDong =
    (noiDungDong || "(trống)") + (dvtDong ? " [" + dvtDong + "]" : "");

  // ── XÁC NHẬN LẦN 2 ──
  var confirm = ui.alert(
    " XÁC NHẬN XÓA DÒNG " + deleteRow,
    "Dòng sẽ bị xóa:\n" +
      "────────────────────────\n" +
      '"' +
      moTaDong +
      '"\n' +
      "────────────────────────\n\n" +
      " Dòng này sẽ bị xóa khỏi BIỂU TỔNG\n" +
      "    VÀ tất cả " +
      laySheetBan(ss).length +
      " sheet bản!\n\n" +
      "Bạn có chắc chắn không?",
    ui.ButtonSet.YES_NO,
  );

  if (confirm !== ui.Button.YES) {
    ui.alert("Đã hủy thao tác xóa.");
    return;
  }

  var banSheets = laySheetBan(ss);

  try {
    // ════════════════════════════════════════════════════════
    // BƯỚC 1: Xóa từ TẤT CẢ sheet bản
    // → Google Sheets auto-adjust công thức tham chiếu trong BIỂU TỔNG
    // ════════════════════════════════════════════════════════
    banSheets.forEach(function (sheet) {
      if (deleteRow <= sheet.getLastRow()) {
        sheet.deleteRow(deleteRow);
      }
    });

    // ════════════════════════════════════════════════════════
    // BƯỚC 2: Xóa từ BIỂU TỔNG
    // Dòng này có thể đang hiển thị #REF! (do bước 1)
    // Xóa luôn để dọn sạch
    // ════════════════════════════════════════════════════════
    if (deleteRow <= sheetTong.getLastRow()) {
      sheetTong.deleteRow(deleteRow);
    }

    // ── THÔNG BÁO KẾT QUẢ ──
    ui.alert(
      " ĐÃ XÓA THÀNH CÔNG!\n\n" +
        "🗑️ Đã xóa dòng " +
        deleteRow +
        ': "' +
        moTaDong +
        '"\n' +
        " Đồng bộ từ " +
        (banSheets.length + 1) +
        " sheets",
    );
  } catch (e) {
    ui.alert(" LỖI KHI XÓA DÒNG!\n\nChi tiết: " + e.message);
    Logger.log("[xoaDongDongBo] ERROR: " + e.message + "\n" + e.stack);
  }
}
