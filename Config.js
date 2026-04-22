/**
 * ============================================================
 * CONFIG.GS - CẤU HÌNH HỆ THỐNG
 * ============================================================
 * Đồ án: Xây dựng công cụ hỗ trợ tổng hợp dữ liệu cho chính quyền cấp xã
 * Nền tảng: Google Apps Script + Google Sheets
 * File: Biểu Báo cáo tháng các bản/tiểu khu - Xã Mường La
 *
 * HƯỚNG DẪN:
 * - Nếu cấu trúc sheet thay đổi, chỉnh sửa các giá trị ở đây
 * - KHÔNG thay đổi tên biến, chỉ thay đổi giá trị
 * ============================================================
 */

var CONFIG = {

  // ============================================================
  // TÊN CÁC SHEET HỆ THỐNG
  // ============================================================

  /** Sheet tổng hợp chỉ tiêu toàn xã (biểu báo cáo tháng) */
  SHEET_BIEU_TONG: "BIỂU TỔNG",

  /** Sheet tổng hợp số hộ/nhân khẩu từng bản */
  SHEET_TONG_TOAN_XA: "BIỂU TỔNG TOÀN XÃ",

  /** Sheet danh sách bản/tiểu khu (có dấu cách cuối tên) */
  SHEET_DS_BAN: "DS BẢN, TK ",

  /**
   * Danh sách TẤT CẢ sheet hệ thống (không phải sheet bản).
   * Script sẽ BỎ QUA các sheet này khi tổng hợp.
   * Nếu thêm sheet hệ thống mới → thêm tên vào đây.
   */
  SHEETS_HE_THONG: [
    "BIỂU TỔNG",
    "BIỂU TỔNG TOÀN XÃ",
    "DS BẢN, TK "
  ],


  // ============================================================
  // CẤU TRÚC DÒNG - BIỂU TỔNG & CÁC SHEET BẢN
  // (Cả BIỂU TỔNG và sheet bản có cùng cấu trúc dòng)
  // ============================================================

  /** Dòng chứa tiêu đề cột: Stt | Nội dung | Đơn vị tính | ... */
  DONG_HEADER: 6,

  /**
   * Dòng đầu tiên có nội dung dữ liệu.
   * Dòng 7 = "PHẦN I: THÔNG TIN CHUNG"
   */
  DONG_DU_LIEU_BAT_DAU: 7,


  // ============================================================
  // CẤU TRÚC CỘT - BIỂU TỔNG & CÁC SHEET BẢN
  // ============================================================

  COT_STT:      1,  // Cột A: Số thứ tự
  COT_NOI_DUNG: 2,  // Cột B: Nội dung / Tên chỉ tiêu
  COT_DVT:      3,  // Cột C: Đơn vị tính
  COT_DAU_KY:   4,  // Cột D: Số liệu đầu kỳ báo cáo
  COT_T10:      5,  // Cột E: Số liệu tháng 10
  COT_T11:      6,  // Cột F: Số liệu tháng 11
  COT_T12:      7,  // Cột G: Số liệu tháng 12
  COT_LUY_KE:   8,  // Cột H: Luỹ kế từ đầu năm (= D+E+F+G)
  COT_GHI_CHU:  9,  // Cột I: Ghi chú


  // ============================================================
  // CẤU TRÚC DÒNG/CỘT - BIỂU TỔNG TOÀN XÃ
  // ============================================================

  /** Dòng chứa tiêu đề cột trong BIỂU TỔNG TOÀN XÃ */
  TONG_XA_DONG_HEADER: 3,

  /** Dòng "XÃ MƯỜNG LA" */
  TONG_XA_DONG_XA: 4,

  /**
   * Dòng đầu tiên chứa dữ liệu bản (Bản Nà Lốc).
   * Thứ tự bản trong sheet này khớp với thứ tự sheet bản trong file.
   */
  TONG_XA_DONG_BAN_BAT_DAU: 5,

  TONG_XA_COT_STT:       1,  // Cột A: Stt
  TONG_XA_COT_TEN_BAN:   2,  // Cột B: Tên bản/tiểu khu
  TONG_XA_COT_SO_HO:     3,  // Cột C: Số hộ
  TONG_XA_COT_NHAN_KHAU: 4,  // Cột D: Số nhân khẩu
  TONG_XA_COT_HO_NGHEO:  5,  // Cột E: Hộ nghèo
  TONG_XA_COT_CAN_NGHEO: 6,  // Cột F: Hộ cận nghèo


  // ============================================================
  // CÁC CHỈ TIÊU QUAN TRỌNG (dùng để tìm dòng tự động)
  // Thay vì dùng số dòng cố định, script tìm theo NỘI DUNG cột B
  // → Đây chính là nguyên lý "tính vị trí theo tổng số dòng"
  // ============================================================

  /** Tên chỉ tiêu "Tổng số hộ" trong cột B */
  NOI_DUNG_TONG_SO_HO: "Tổng số hộ",

  /** Tên chỉ tiêu "Tổng số nhân khẩu" trong cột B */
  NOI_DUNG_NHAN_KHAU: "Tổng số nhân khẩu"

};
