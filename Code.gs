/**
 * NHẬT KÝ LỜI NHẮN – dán toàn bộ file này vào Google Apps Script
 * (Google Sheet > Tiện ích mở rộng > Apps Script).
 *
 * Lời nhắn được lưu vào trang tính "NhatKy" của chính Google Sheet này.
 * Chỉ ai nhập đúng MAT_KHAU trên web mới đọc / viết được.
 *
 * BẢN MỚI: mỗi thiệp có một "phòng" riêng (cột E "Phòng"), nên nhật ký của
 * khách này không lẫn với khách khác. Lời nhắn cũ (chưa có phòng) vẫn thuộc
 * phòng mặc định của bạn.
 */

// ====== ĐỔI MÃ BÍ MẬT Ở ĐÂY (hai người cùng dùng mã này) ======
const MAT_KHAU = "doi-ma-nay-di";

const TEN_TRANG = "NhatKy";
const TOI_DA_TRA_VE = 500; // chỉ gửi về web 500 lời nhắn gần nhất cho nhẹ

function doPost(e) {
  let d;
  try {
    d = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, loi: "du_lieu_loi" });
  }
  if (String(d.key || "") !== MAT_KHAU) return json_({ ok: false, loi: "sai_ma" });

  const phong = sach_(d.phong, 80); // "" = thiệp mặc định của bạn

  if (d.action === "list") return json_({ ok: true, list: docTatCa_(phong) });

  if (d.action === "add") {
    const loiNhan = sach_(d.loiNhan, 1000);
    if (!loiNhan) return json_({ ok: false, loi: "trong" });
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      trang_().appendRow([new Date(), sach_(d.nguoi, 40), sach_(d.tamTrang, 16), loiNhan, phong]);
    } finally {
      lock.releaseLock();
    }
    return json_({ ok: true });
  }

  return json_({ ok: false, loi: "khong_ro" });
}

// Mở link Web App trên trình duyệt để kiểm tra đã chạy chưa
function doGet() {
  return json_({ ok: true, ghiChu: "Nhật ký đang chạy ♡" });
}

function trang_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(TEN_TRANG);
  if (!sh) {
    sh = ss.insertSheet(TEN_TRANG);
    sh.appendRow(["Thời gian", "Người viết", "Tâm trạng", "Lời nhắn", "Phòng"]);
    sh.setFrozenRows(1);
  } else if (sh.getRange(1, 5).getValue() === "") {
    sh.getRange(1, 5).setValue("Phòng"); // trang tính cũ: bổ sung tiêu đề cột Phòng
  }
  return sh;
}

function docTatCa_(phong) {
  const sh = trang_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  // Đọc cả trang tính, chỉ giữ lời nhắn của đúng phòng, rồi lấy các lời nhắn mới nhất
  return sh.getRange(2, 1, last - 1, 5).getValues()
    .filter(r => r[3] !== "" && String(r[4] || "") === phong)
    .slice(-TOI_DA_TRA_VE)
    .map(r => ({
      t: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
      nguoi: String(r[1]),
      tamTrang: String(r[2]),
      loiNhan: String(r[3]),
    }));
}

// Cắt độ dài + chặn chữ bắt đầu bằng = + - @ bị Google Sheet hiểu thành công thức
function sach_(v, max) {
  let s = String(v == null ? "" : v).trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
