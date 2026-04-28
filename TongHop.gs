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
    var luyKeFormulas = sheetTong
      .getRange(dataStartRow, CONFIG.COT_LUY_KE, numRows, 1)
      .getFormulas();

    var updatedCount = 0;

    // ── XỬ LÝ TỪNG DÒNG ──
    for (var i = 0; i < formulas.length; i++) {
      // Số dòng THỰC TẾ trong Google Sheets (không phải index mảng)
      // Đây chính là "tính vị trí theo tổng số dòng"
      var actualRow = dataStartRow + i;

      var canRebuildRow = laFormulaLuyKe(luyKeFormulas[i][0], actualRow);
      for (var k = 0; k < formulaColCount; k++) {
        if (laFormulaTongHop(formulas[i][k])) {
          canRebuildRow = true;
          break;
        }
      }

      if (!canRebuildRow) {
        continue;
      }

      // Kiểm tra cột D, E, F, G
      for (var j = 0; j < formulaColCount; j++) {
        var colNumber = formulaStartCol + j;
        var currentFormula = formulas[i][j];

        // Xây dựng công thức mới với:
        //   - Sheet names: lấy từ file HIỆN TẠI (không cố định)
        //   - Row number: dùng actualRow THỰC TẾ (không cố định)
        var newFormula = xayDungCongThucTongHop(actualRow, colNumber, banNames);

        if (currentFormula !== newFormula) {
          sheetTong.getRange(actualRow, colNumber).setFormula(newFormula);
          updatedCount++;
        }
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
    // Duyệt toàn bộ sheet bản để tránh lỗi khi sheet đầu tiên không phải sheet dữ liệu.
    var dongTongSoHo = -1;
    var dongNhanKhau = -1;
    var dongHoNgheo = -1;
    var dongHoCanNgheo = -1;
    for (var s = 0; s < banSheets.length; s++) {
      var tempDongTongSoHo = timDong(banSheets[s], CONFIG.NOI_DUNG_TONG_SO_HO);
      if (tempDongTongSoHo !== -1) {
        dongTongSoHo = tempDongTongSoHo;
      }

      var tempDongNhanKhau = timDong(banSheets[s], CONFIG.NOI_DUNG_NHAN_KHAU);
      if (tempDongNhanKhau !== -1) {
        dongNhanKhau = tempDongNhanKhau;
      }

      var tempDongHoNgheo = timDong(banSheets[s], CONFIG.NOI_DUNG_HO_NGHEO);
      if (tempDongHoNgheo !== -1) {
        dongHoNgheo = tempDongHoNgheo;
      }

      var tempDongHoCanNgheo = timDong(
        banSheets[s],
        CONFIG.NOI_DUNG_HO_CAN_NGHEO,
      );
      if (tempDongHoCanNgheo !== -1) {
        dongHoCanNgheo = tempDongHoCanNgheo;
      }

      if (
        dongTongSoHo !== -1 &&
        dongNhanKhau !== -1 &&
        dongHoNgheo !== -1 &&
        dongHoCanNgheo !== -1
      ) {
        break;
      }
    }

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

      if (dongHoNgheo !== -1) {
        sheetTongXa
          .getRange(targetRow, CONFIG.TONG_XA_COT_HO_NGHEO)
          .setFormula("='" + sheet.getName() + "'!H" + dongHoNgheo);
      }

      if (dongHoCanNgheo !== -1) {
        sheetTongXa
          .getRange(targetRow, CONFIG.TONG_XA_COT_CAN_NGHEO)
          .setFormula("='" + sheet.getName() + "'!H" + dongHoCanNgheo);
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

/**
 * Audit BIỂU TỔNG: phát hiện thiếu công thức và tham chiếu sai sheet.
 * Kết quả chi tiết được ghi ra sheet AUDIT_BIEU_TONG.
 */
function auditBieuTong() {
  var ui = SpreadsheetApp.getUi();

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetTong = ss.getSheetByName(CONFIG.SHEET_BIEU_TONG);
    var reportSheetName = "AUDIT_BIEU_TONG";

    if (!sheetTong) {
      ui.alert('❌ Không tìm thấy sheet "' + CONFIG.SHEET_BIEU_TONG + '"!');
      return;
    }

    var banNames = layTenSheetBan(ss);
    if (banNames.length === 0) {
      ui.alert("❌ Không tìm thấy sheet bản/tiểu khu nào để audit!");
      return;
    }

    var banNameMap = {};
    for (var b = 0; b < banNames.length; b++) {
      banNameMap[chuanHoaTenSheet_(banNames[b])] = true;
    }

    var dataStartRow = CONFIG.DONG_DU_LIEU_BAT_DAU;
    var lastRow = sheetTong.getLastRow();

    if (lastRow < dataStartRow) {
      ui.alert("ℹ️ BIỂU TỔNG chưa có dữ liệu để audit.");
      return;
    }

    var numRows = lastRow - dataStartRow + 1;
    var formulaStartCol = CONFIG.COT_DAU_KY;
    var formulaColCount = CONFIG.COT_T12 - CONFIG.COT_DAU_KY + 1;
    var colLetters = ["D", "E", "F", "G"];

    var formulas = sheetTong
      .getRange(dataStartRow, formulaStartCol, numRows, formulaColCount)
      .getFormulas();

    var luyKeFormulas = sheetTong
      .getRange(dataStartRow, CONFIG.COT_LUY_KE, numRows, 1)
      .getFormulas();

    var issues = [];

    for (var i = 0; i < numRows; i++) {
      var actualRow = dataStartRow + i;
      var rowHasFormulaTongHop = false;

      for (var k = 0; k < formulaColCount; k++) {
        if (laFormulaTongHop(formulas[i][k])) {
          rowHasFormulaTongHop = true;
          break;
        }
      }

      var canAuditRow =
        laFormulaLuyKe(luyKeFormulas[i][0], actualRow) || rowHasFormulaTongHop;
      if (!canAuditRow) {
        continue;
      }

      for (var j = 0; j < formulaColCount; j++) {
        var currentFormula = formulas[i][j];
        var expectedFormula = xayDungCongThucTongHop(
          actualRow,
          formulaStartCol + j,
          banNames,
        );
        var colLetter = colLetters[j];

        if (!currentFormula) {
          issues.push([
            actualRow,
            colLetter,
            "THIEU_CONG_THUC",
            "",
            expectedFormula,
            "Ô trống, chưa có công thức tổng hợp.",
          ]);
          continue;
        }

        if (currentFormula === expectedFormula) {
          continue;
        }

        var referencedSheets = trichXuatTenSheetTrongCongThuc_(currentFormula);
        var wrongSheets = referencedSheets.filter(function (sheetName) {
          return !banNameMap[chuanHoaTenSheet_(sheetName)];
        });

        var issueType = "CONG_THUC_KHONG_KHOP";
        var note = "Công thức khác mẫu kỳ vọng.";

        if (!laFormulaTongHop(currentFormula)) {
          issueType = "KHONG_PHAI_CONG_THUC_TONG_HOP";
          note = "Ô có thể đã bị ghi đè giá trị hoặc công thức khác.";
        } else if (wrongSheets.length > 0) {
          issueType = "THAM_CHIEU_SAI_SHEET";
          note = "Sheet sai: " + wrongSheets.join(", ");
        }

        issues.push([
          actualRow,
          colLetter,
          issueType,
          currentFormula,
          expectedFormula,
          note,
        ]);
      }
    }

    var reportSheet = ss.getSheetByName(reportSheetName);
    if (!reportSheet) {
      reportSheet = ss.insertSheet(reportSheetName);
    } else {
      reportSheet.clearContents();
    }

    var header = [
      "Thoi gian audit",
      "Sheet",
      "Dong",
      "Cot",
      "Loai loi",
      "Cong thuc hien tai",
      "Cong thuc ky vong",
      "Ghi chu",
    ];

    reportSheet.getRange(1, 1, 1, header.length).setValues([header]);

    var nowText = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "dd/MM/yyyy HH:mm:ss",
    );

    if (issues.length > 0) {
      var reportRows = issues.map(function (issue) {
        return [
          nowText,
          CONFIG.SHEET_BIEU_TONG,
          issue[0],
          issue[1],
          issue[2],
          congThucDangText_(issue[3]),
          congThucDangText_(issue[4]),
          issue[5],
        ];
      });

      reportSheet
        .getRange(2, 1, reportRows.length, header.length)
        .setValues(reportRows);
    }

    reportSheet.autoResizeColumns(1, header.length);

    ui.alert(
      "🩺 AUDIT BIỂU TỔNG HOÀN TẤT!\n\n" +
        "• Dòng đã quét: " +
        numRows +
        "\n" +
        "• Lỗi phát hiện: " +
        issues.length +
        "\n" +
        "• Báo cáo chi tiết: " +
        reportSheetName,
    );
  } catch (e) {
    ui.alert("❌ LỖI KHI AUDIT: " + e.message);
    Logger.log("[auditBieuTong] ERROR: " + e.message + "\n" + e.stack);
  }
}

/**
 * Sửa tự động các ô bị báo "THIEU_CONG_THUC" trong AUDIT_BIEU_TONG.
 * Đọc báo cáo audit, xác nhận rồi ghi lại công thức kỳ vọng vào BIỂU TỔNG.
 */
function suaTuDongTheoAudit() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var reportName = "AUDIT_BIEU_TONG";

  try {
    var reportSheet = ss.getSheetByName(reportName);
    if (!reportSheet) {
      ui.alert(
        '❌ Không tìm thấy báo cáo "' + reportName + '". Hãy chạy audit trước.',
      );
      return;
    }

    var lastRow = reportSheet.getLastRow();
    if (lastRow < 2) {
      ui.alert("ℹ️ Báo cáo audit không có lỗi để sửa.");
      return;
    }

    var rows = reportSheet.getRange(2, 1, lastRow - 1, 8).getValues();
    var fixes = [];

    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var loai = r[4]; // Loai loi
      if (loai !== "THIEU_CONG_THUC") continue;

      var dong = Number(r[2]); // Dong
      var cotLetter = String(r[3] || "").trim(); // Cot (D/E/F/G)
      var expected = r[6] || ""; // Cong thuc ky vong (may start with '\'')

      // Strip leading single-quote if present (audit stores formula as text)
      if (expected && expected.charAt(0) === "'") {
        expected = expected.substring(1);
      }

      if (!expected) continue;

      var colNumber = cotLetter.charCodeAt(0) - 64; // 'A'->1
      if (isNaN(dong) || !cotLetter || colNumber < 1) continue;

      fixes.push({ row: dong, col: colNumber, formula: expected });
    }

    if (fixes.length === 0) {
      ui.alert('ℹ️ Không tìm thấy lỗi "THIEU_CONG_THUC" trong báo cáo.');
      return;
    }

    var confirm = ui.alert(
      "⚡ Sửa tự động theo AUDIT",
      "Thao tác sẽ ghi " +
        fixes.length +
        " công thức vào BIỂU TỔNG (D:G).\nBạn có muốn tiếp tục?",
      ui.ButtonSet.YES_NO,
    );

    if (confirm !== ui.Button.YES) return;

    var sheetTong = ss.getSheetByName(CONFIG.SHEET_BIEU_TONG);
    if (!sheetTong) {
      ui.alert('❌ Không tìm thấy sheet "' + CONFIG.SHEET_BIEU_TONG + '"!');
      return;
    }

    var applied = 0;
    for (var j = 0; j < fixes.length; j++) {
      try {
        var f = fixes[j];
        sheetTong.getRange(f.row, f.col).setFormula(f.formula);
        applied++;
      } catch (err) {
        Logger.log(
          "[suaTuDongTheoAudit] Failed to set formula at " +
            f.row +
            "," +
            f.col +
            ": " +
            err.message,
        );
      }
    }

    ui.alert("✅ Sửa tự động hoàn tất. Ô đã ghi công thức: " + applied);
  } catch (e) {
    ui.alert("❌ LỖI khi sửa tự động: " + e.message);
    Logger.log("[suaTuDongTheoAudit] ERROR: " + e.message + "\n" + e.stack);
  }
}
