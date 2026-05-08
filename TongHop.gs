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
 *    số hộ, nhân khẩu, hộ nghèo, hộ cận nghèo theo từng bản.
 *
 * LỊCH SỬ SỬA LỖI:
 *   v1.1 - Sửa điều kiện canRebuildRow: chỉ rebuild khi dòng đã
 *          có ít nhất 1 công thức tổng hợp. Tránh ghi đè dữ liệu
 *          người dùng nhập trực tiếp.
 *        - Thêm rebuild cột H (Luỹ kế) khi bị mất.
 *        - suaTuDongTheoAudit() xử lý đủ 3 loại lỗi thay vì chỉ
 *          THIEU_CONG_THUC.
 * ============================================================
 */

/**
 * CHỨC NĂNG CHÍNH (Cơ bản):
 * Tổng hợp dữ liệu từ tất cả sheet bản → BIỂU TỔNG
 */
function tongHopBieuTong() {
  var ui = SpreadsheetApp.getUi();

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetTong = ss.getSheetByName(CONFIG.SHEET_BIEU_TONG);

    if (!sheetTong) {
      ui.alert(
        ' Không tìm thấy sheet "' +
          CONFIG.SHEET_BIEU_TONG +
          '"!\n' +
          "Kiểm tra lại tên sheet trong CONFIG.SHEET_BIEU_TONG",
      );
      return;
    }

    var banNames = layTenSheetBan(ss);
    if (banNames.length === 0) {
      ui.alert(
        " Không tìm thấy sheet bản/tiểu khu nào!\n" +
          "Kiểm tra lại CONFIG.SHEETS_HE_THONG",
      );
      return;
    }

    var dataStartRow = CONFIG.DONG_DU_LIEU_BAT_DAU;
    var lastRow = sheetTong.getLastRow();

    if (lastRow < dataStartRow) {
      ui.alert(" BIỂU TỔNG không có dữ liệu (lastRow=" + lastRow + ")!");
      return;
    }

    var numRows = lastRow - dataStartRow + 1;
    var formulaStartCol = CONFIG.COT_DAU_KY;
    var formulaColCount = CONFIG.COT_T12 - CONFIG.COT_DAU_KY + 1; // D..G

    var formulaRange = sheetTong.getRange(
      dataStartRow,
      formulaStartCol,
      numRows,
      formulaColCount,
    );
    var formulas = formulaRange.getFormulas();

    var luyKeFormulas = sheetTong
      .getRange(dataStartRow, CONFIG.COT_LUY_KE, numRows, 1)
      .getFormulas();

    var updatedCount = 0;

    for (var i = 0; i < formulas.length; i++) {
      var actualRow = dataStartRow + i;

      // [SỬA LỖI v1.1] Chỉ rebuild dòng nếu ÍT NHẤT 1 ô D-G đã có công thức tổng hợp.
      // Không dùng laFormulaLuyKe làm điều kiện khởi tạo để tránh ghi đè
      // dữ liệu người dùng nhập trực tiếp vào cột D-G.
      var canRebuildRow = false;
      for (var k = 0; k < formulaColCount; k++) {
        if (laFormulaTongHop(formulas[i][k])) {
          canRebuildRow = true;
          break;
        }
      }

      if (!canRebuildRow) {
        continue;
      }

      // Rebuild cột D-G
      for (var j = 0; j < formulaColCount; j++) {
        var colNumber = formulaStartCol + j;
        var currentFormula = formulas[i][j];
        var newFormula = xayDungCongThucTongHop(actualRow, colNumber, banNames);

        if (currentFormula !== newFormula) {
          sheetTong.getRange(actualRow, colNumber).setFormula(newFormula);
          updatedCount++;
        }
      }

      // [SỬA LỖI v1.1] Rebuild cột H (Luỹ kế) nếu bị mất hoặc sai.
      // Trước đây bị comment "KHÔNG cần rebuild" dẫn đến không phục hồi khi cột H bị xóa.
      var currentLuyKe = luyKeFormulas[i][0];
      var expectedLuyKe = xayDungCongThucLuyKe(actualRow);
      if (!laFormulaLuyKe(currentLuyKe, actualRow)) {
        sheetTong
          .getRange(actualRow, CONFIG.COT_LUY_KE)
          .setFormula(expectedLuyKe);
        updatedCount++;
      }
    }

    ui.alert(
      " TỔNG HỢP BIỂU TỔNG HOÀN TẤT!\n\n" +
        "📊 Ô công thức đã cập nhật: " +
        updatedCount +
        "\n" +
        " Số bản/tiểu khu:         " +
        banNames.length +
        "\n" +
        " Số dòng dữ liệu:          " +
        numRows,
    );
  } catch (e) {
    ui.alert(
      " LỖI KHI TỔNG HỢP!\n\n" +
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
 * Cập nhật BIỂU TỔNG TOÀN XÃ từ số liệu sheet bản.
 *
 * Tìm dòng "Tổng số hộ", "Tổng số nhân khẩu", "Hộ nghèo", "Hộ cận nghèo"
 * ĐỘNG theo nội dung, không hardcode số dòng.
 */
function tongHopToanXa() {
  var ui = SpreadsheetApp.getUi();

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetTongXa = ss.getSheetByName(CONFIG.SHEET_TONG_TOAN_XA);

    if (!sheetTongXa) {
      ui.alert(' Không tìm thấy sheet "' + CONFIG.SHEET_TONG_TOAN_XA + '"!');
      return;
    }

    var banSheets = laySheetBan(ss);
    if (banSheets.length === 0) {
      ui.alert(" Không tìm thấy sheet bản/tiểu khu nào!");
      return;
    }

    var dongTongSoHo = -1;
    var dongNhanKhau = -1;
    var dongHoNgheo = -1;
    var dongHoCanNgheo = -1;

    for (var s = 0; s < banSheets.length; s++) {
      if (dongTongSoHo === -1)
        dongTongSoHo = timDong(banSheets[s], CONFIG.NOI_DUNG_TONG_SO_HO);
      if (dongNhanKhau === -1)
        dongNhanKhau = timDong(banSheets[s], CONFIG.NOI_DUNG_NHAN_KHAU);
      if (dongHoNgheo === -1)
        dongHoNgheo = timDong(banSheets[s], CONFIG.NOI_DUNG_HO_NGHEO);
      if (dongHoCanNgheo === -1)
        dongHoCanNgheo = timDong(banSheets[s], CONFIG.NOI_DUNG_HO_CAN_NGHEO);

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
        ' Không tìm thấy chỉ tiêu "' +
          CONFIG.NOI_DUNG_TONG_SO_HO +
          '"!\n\n' +
          "Kiểm tra lại:\n" +
          "• Giá trị CONFIG.NOI_DUNG_TONG_SO_HO\n" +
          "• Nội dung cột B trong sheet bản",
      );
      return;
    }

    var startRow = CONFIG.TONG_XA_DONG_BAN_BAT_DAU;
    var updatedCount = 0;

    for (var i = 0; i < banSheets.length; i++) {
      var sheet = banSheets[i];
      var targetRow = startRow + i;

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

    // ════════════════════════════════════════════════════════
    // [SỬA LỖI v1.2] Cập nhật dòng "XÃ MƯỜNG LA" (row 4) = SUM toàn bộ bản
    // Trước đây bị thiếu khiến dòng tổng xã luôn trống.
    // ════════════════════════════════════════════════════════
    var dongXa = CONFIG.TONG_XA_DONG_XA;
    var dongBanCuoi = startRow + banSheets.length - 1;
    sheetTongXa
      .getRange(dongXa, CONFIG.TONG_XA_COT_SO_HO)
      .setFormula("=SUM(C" + startRow + ":C" + dongBanCuoi + ")");
    sheetTongXa
      .getRange(dongXa, CONFIG.TONG_XA_COT_NHAN_KHAU)
      .setFormula("=SUM(D" + startRow + ":D" + dongBanCuoi + ")");
    sheetTongXa
      .getRange(dongXa, CONFIG.TONG_XA_COT_HO_NGHEO)
      .setFormula("=SUM(E" + startRow + ":E" + dongBanCuoi + ")");
    sheetTongXa
      .getRange(dongXa, CONFIG.TONG_XA_COT_CAN_NGHEO)
      .setFormula("=SUM(F" + startRow + ":F" + dongBanCuoi + ")");

    var warnings = [];
    if (dongNhanKhau === -1)
      warnings.push('"' + CONFIG.NOI_DUNG_NHAN_KHAU + '"');
    if (dongHoNgheo === -1) warnings.push('"' + CONFIG.NOI_DUNG_HO_NGHEO + '"');
    if (dongHoCanNgheo === -1)
      warnings.push('"' + CONFIG.NOI_DUNG_HO_CAN_NGHEO + '"');

    var warnMsg =
      warnings.length > 0
        ? "\n\n Không tìm thấy chỉ tiêu:\n" + warnings.join("\n")
        : "";

    ui.alert(
      " CẬP NHẬT BIỂU TỔNG TOÀN XÃ HOÀN TẤT!\n\n" +
        " Bản/tiểu khu đã cập nhật: " +
        updatedCount +
        "\n" +
        " Tổng số hộ tại dòng:      " +
        dongTongSoHo +
        "\n" +
        " Nhân khẩu tại dòng:       " +
        dongNhanKhau +
        "\n" +
        " Hộ nghèo tại dòng:        " +
        dongHoNgheo +
        "\n" +
        " Hộ cận nghèo tại dòng:    " +
        dongHoCanNgheo +
        warnMsg,
    );
  } catch (e) {
    ui.alert(" LỖI KHI CẬP NHẬT TOÀN XÃ!\n\nChi tiết: " + e.message);
    Logger.log("[tongHopToanXa] ERROR: " + e.message + "\n" + e.stack);
  }
}

/**
 * AUDIT: Quét BIỂU TỔNG, phát hiện ô thiếu hoặc sai công thức.
 * Kết quả ghi ra sheet AUDIT_BIEU_TONG.
 */
function auditBieuTong() {
  var ui = SpreadsheetApp.getUi();

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetTong = ss.getSheetByName(CONFIG.SHEET_BIEU_TONG);
    var reportSheetName = "AUDIT_BIEU_TONG";

    if (!sheetTong) {
      ui.alert(' Không tìm thấy sheet "' + CONFIG.SHEET_BIEU_TONG + '"!');
      return;
    }

    var banNames = layTenSheetBan(ss);
    if (banNames.length === 0) {
      ui.alert(" Không tìm thấy sheet bản/tiểu khu nào để audit!");
      return;
    }

    var banNameMap = {};
    for (var b = 0; b < banNames.length; b++) {
      banNameMap[chuanHoaTenSheet_(banNames[b])] = true;
    }

    var dataStartRow = CONFIG.DONG_DU_LIEU_BAT_DAU;
    var lastRow = sheetTong.getLastRow();

    if (lastRow < dataStartRow) {
      ui.alert(" BIỂU TỔNG chưa có dữ liệu để audit.");
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

      if (!rowHasFormulaTongHop) continue;

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

        if (currentFormula === expectedFormula) continue;

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
    ui.alert(" LỖI KHI AUDIT: " + e.message);
    Logger.log("[auditBieuTong] ERROR: " + e.message + "\n" + e.stack);
  }
}

/**
 * Sửa tự động theo AUDIT_BIEU_TONG.
 *
 * [SỬA LỖI v1.1] Xử lý đủ 3 loại lỗi thay vì chỉ THIEU_CONG_THUC:
 *   - THIEU_CONG_THUC
 *   - CONG_THUC_KHONG_KHOP
 *   - THAM_CHIEU_SAI_SHEET
 *
 * KHONG_PHAI_CONG_THUC_TONG_HOP bị bỏ qua vì ô có thể chứa
 * công thức hợp lệ do người dùng cố ý nhập (ví dụ: =SUM(D36:D38)+D41).
 * Người dùng cần xem xét thủ công trước khi sửa.
 */
function suaTuDongTheoAudit() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var reportName = "AUDIT_BIEU_TONG";

  // Danh sách loại lỗi được tự động sửa
  var AUTO_FIX_TYPES = [
    "THIEU_CONG_THUC",
    "CONG_THUC_KHONG_KHOP",
    "THAM_CHIEU_SAI_SHEET",
  ];

  try {
    var reportSheet = ss.getSheetByName(reportName);
    if (!reportSheet) {
      ui.alert(
        ' Không tìm thấy báo cáo "' + reportName + '". Hãy chạy audit trước.',
      );
      return;
    }

    var lastRow = reportSheet.getLastRow();
    if (lastRow < 2) {
      ui.alert(" Báo cáo audit không có lỗi để sửa.");
      return;
    }

    var rows = reportSheet.getRange(2, 1, lastRow - 1, 8).getValues();
    var fixes = [];
    var skipped = 0;

    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var loai = r[4];
      var dong = Number(r[2]);
      var cotLetter = String(r[3] || "").trim();
      var expected = r[6] || "";

      if (AUTO_FIX_TYPES.indexOf(loai) === -1) {
        skipped++;
        continue;
      }

      if (expected && expected.charAt(0) === "'") {
        expected = expected.substring(1);
      }

      if (!expected) continue;

      var colNumber = cotLetter.charCodeAt(0) - 64;
      if (isNaN(dong) || !cotLetter || colNumber < 1) continue;

      fixes.push({ row: dong, col: colNumber, formula: expected, loai: loai });
    }

    if (fixes.length === 0) {
      ui.alert(
        " Không tìm thấy lỗi có thể sửa tự động.\n\n" +
          (skipped > 0
            ? skipped +
              " lỗi KHONG_PHAI_CONG_THUC_TONG_HOP cần xem xét thủ công."
            : ""),
      );
      return;
    }

    var confirmMsg =
      "Thao tác sẽ ghi " +
      fixes.length +
      " công thức vào BIỂU TỔNG (D:G)." +
      (skipped > 0
        ? "\n\n " +
          skipped +
          " lỗi KHONG_PHAI_CONG_THUC_TONG_HOP bị bỏ qua (cần xem xét thủ công)."
        : "") +
      "\n\nBạn có muốn tiếp tục?";

    var confirm = ui.alert(
      "⚡ Sửa tự động theo AUDIT",
      confirmMsg,
      ui.ButtonSet.YES_NO,
    );
    if (confirm !== ui.Button.YES) return;

    var sheetTong = ss.getSheetByName(CONFIG.SHEET_BIEU_TONG);
    if (!sheetTong) {
      ui.alert(' Không tìm thấy sheet "' + CONFIG.SHEET_BIEU_TONG + '"!');
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
          "[suaTuDongTheoAudit] Failed at row=" +
            fixes[j].row +
            " col=" +
            fixes[j].col +
            ": " +
            err.message,
        );
      }
    }

    ui.alert(" Sửa tự động hoàn tất. Ô đã ghi công thức: " + applied);
  } catch (e) {
    ui.alert(" LỖI khi sửa tự động: " + e.message);
    Logger.log("[suaTuDongTheoAudit] ERROR: " + e.message + "\n" + e.stack);
  }
}
