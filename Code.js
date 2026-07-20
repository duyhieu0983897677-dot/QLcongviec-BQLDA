// Hàm include() cho phép Index.html nhúng Style.html và các file Script.html
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// HangMuc nay CHỈ là dữ liệu tham chiếu (mã/tên/cây cha-con) để lọc & để Gói thầu gộp —
// không còn nhận Nhật ký tiến độ trực tiếp (xem CongViec bên dưới).
// stt (số thứ tự hiển thị) thêm ở CUỐI mỗi bảng — không chèn giữa để không xáo trộn cột của các
// sheet đã tồn tại. Cho phép sửa tay để tự sắp xếp lại thứ tự hiển thị (1->n), độc lập với mã.
const HANGMUC_HEADERS_ = ['maHangMuc', 'tenHangMuc', 'capDo', 'parentId', 'active', 'stt'];
const GOITHAU_HEADERS_ = ['maGoiThau', 'tenGoiThau', 'maHangMucList', 'stt'];
// CongViec: đơn vị công việc thật do Admin/Giám sát tự tạo — đây mới là nơi Nhật ký tiến độ trỏ vào.
// phanLoai ('ho_so'/'thi_cong') tự gán = phanMacDinh của người tạo tại thời điểm tạo, không sửa tay
// — trừ người có phanMacDinh='ca_hai' thì phải chọn tay lúc tạo (xem adminSaveCongViec).
// nguoiGiao_id/nguoiGiao_ten CHỈ có giá trị khi Admin tạo Công việc và giao cho người KHÁC (khác
// nguoiTao_id — người phụ trách); rỗng nghĩa là tự tạo/tự đảm nhận, không ai giao.
const CONGVIEC_HEADERS_ = ['maCongViec', 'tenCongViec', 'maGoiThau', 'maHangMucList', 'phanLoai',
  'nguoiTao_id', 'nguoiTao_ten', 'ngayBatDauKH', 'ngayKetThucKH', 'active', 'nguoiGiao_id', 'nguoiGiao_ten'];
// ThongBao: thông báo 1 chiều gửi cho 1 người (được giao việc / có người bình luận vào việc mình
// phụ trách). BinhLuanCongViec: bình luận 2 chiều append-only theo từng Công việc.
const THONGBAO_HEADERS_ = ['id', 'userId', 'loai', 'noiDung', 'maCongViec', 'nguoiGui_id', 'nguoiGui_ten', 'thoiGian', 'daDoc', 'active'];
// 'nguoiDuocTag' luôn APPEND ở CUỐI (sau 'active') — không chèn giữa, giữ đúng convention các cột
// mới thêm khác trong app (xem ghiChu/giamGiaDot).
const BINHLUAN_HEADERS_ = ['id', 'maCongViec', 'nguoiBinhLuan_id', 'nguoiBinhLuan_ten', 'noiDung', 'thoiGian', 'active', 'nguoiDuocTag'];
const NHATKY_HEADERS_ = ['logId', 'maCongViec', 'ngayBaoCao', 'phanTramNgay', 'ghiChu',
  'nguoiNhap_id', 'nguoiNhap_ten', 'thoiGianNhap', 'fileDinhKem', 'active'];
// PhepThang: sổ phép/tháng (phepDauThang = baseline đầu tháng, tự cộng 1 ngày/tháng — xem
// chotPhepDenThangHienTai_). ChamCongThang (định nghĩa ở phần CHẤM CÔNG & TĂNG CA phía dưới): chấm
// công + tăng ca CẢ THÁNG — "Cộng/Trừ phép tháng" luôn tính lại từ 2 sheet này, không lưu cột riêng.
const PHEP_HEADERS_ = ['userId', 'thang', 'phepDauThang', 'soNgayNghi', 'chotLuc'];
const THU_KY_CHUC_DANH_ = 'thư ký bqlda';

// ==== QUẢN LÝ HỢP ĐỒNG (Hợp đồng - BOQ - Phụ lục - Nghiệm thu - Thanh toán - Quyết toán) ====
// Mọi số liệu tổng hợp (khối lượng hiệu lực, đã nghiệm thu, đã thanh toán, Quyết toán...) đều TÍNH
// LẠI khi đọc từ các sheet gốc bên dưới, không lưu cột trùng lặp — cùng triết lý với luyKe của
// CongViec (xem getData()/getHopDongData()).
const HOPDONG_HEADERS_ = ['maHopDong', 'tenHopDong', 'maGoiThau', 'nhaThau', 'ngayKy',
  'tamUngHopDong', 'tamUngThiCong', 'thueVAT', 'giamGia', 'active'];
// BOQHangMuc: cây hạng mục trong 1 Hợp đồng, phân cấp qua stt (dạng "1.1.2") + cap (số).
// maPhuLucTao rỗng = thuộc BOQ gốc; có giá trị = do 1 Phụ lục thêm mới (loại MOI).
// 'ghiChu' luôn APPEND ở CUỐI (sau 'active') — không chèn giữa, vì adminXoaBOQItem() hardcode cột
// active là index 13 (sheet.getRange(i+1, 13)); thêm cột mới ở cuối giữ nguyên mọi index cũ, dòng
// dữ liệu cũ chưa có Ghi chú vẫn đọc được bình thường (ô trống).
const BOQ_HEADERS_ = ['maBOQ', 'maHopDong', 'stt', 'tenHangMuc', 'isHeader', 'donVi',
  'khoiLuongHopDong', 'donGia', 'cap', 'ngayBatDauKH', 'ngayKetThucKH', 'maPhuLucTao', 'active', 'ghiChu'];
const PHULUC_HEADERS_ = ['maPhuLuc', 'maHopDong', 'soHieu', 'ngayPhuLuc', 'ghiChu', 'active'];
// loaiThayDoi: MOI (thêm hạng mục mới, không cần khối lượng/giá riêng ở đây) / THAY_THE (đổi hẳn
// khối lượng hiệu lực = khoiLuongMoi) / DIEU_CHINH (cộng/trừ chênh lệch = khoiLuongDieuChinh).
const PHULUC_THAYDOI_HEADERS_ = ['id', 'maPhuLuc', 'maBOQ', 'loaiThayDoi', 'khoiLuongMoi',
  'khoiLuongDieuChinh', 'donGiaMoi', 'active'];
// NghiemThu: log khối lượng thi công đã xác nhận theo ngày, append-only cộng dồn theo maBOQ —
// giống hệt NhatKyTienDo (xem themNghiemThu()).
const NGHIEMTHU_HEADERS_ = ['maNghiemThu', 'maBOQ', 'ngayNghiemThu', 'khoiLuong', 'ghiChu',
  'nguoiNhap_id', 'nguoiNhap_ten', 'thoiGianNhap', 'active'];
// 'giamGiaDot' luôn APPEND ở CUỐI (sau 'active') — không chèn giữa, vì adminXoaDotThanhToan()
// hardcode cột active là index 11 (sheet.getRange(i+1, 11)); thêm ở cuối giữ nguyên mọi index cũ.
// Khác với hd.giamGia (giảm giá 1 LẦN cho cả hợp đồng, sửa ở tab BOQ) — giamGiaDot là khoản giảm
// giá riêng của TỪNG Đợt thanh toán (giống thu hồi tạm ứng/khấu trừ khác, cũng trừ vào Đợt đó).
const DOTTHANHTOAN_HEADERS_ = ['maDotThanhToan', 'maHopDong', 'tenDot', 'ngayThanhToan',
  'phanTramThanhToan', 'ghiChu', 'thuHoiTamUngHopDong', 'thuHoiTamUngThiCong', 'khauTruKhac',
  'ghiChuKhauTru', 'active', 'giamGiaDot'];
const DOTTHANHTOAN_CHITIET_HEADERS_ = ['id', 'maDotThanhToan', 'maBOQ', 'khoiLuong', 'active'];
// QuyetToan: chỉ lưu mốc xác nhận — mọi số liệu tổng hợp tính lại từ BOQ/PhuLuc/NghiemThu/ThanhToan
// mỗi lần đọc (xem tinhQuyetToanHopDong_()).
const QUYETTOAN_HEADERS_ = ['maQuyetToan', 'maHopDong', 'ngayQuyetToan', 'nguoiXacNhan_id',
  'nguoiXacNhan_ten', 'ghiChu', 'active'];

// Chạy 1 lần duy nhất sau khi tạo Sheet mới: tạo đủ các tab cần thiết
// và 1 tài khoản admin mặc định (admin / 123456) để đăng nhập lần đầu.
function thietLapBanDauSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss.getSheetByName('Data')) {
    ss.insertSheet('Data').getRange('A1').setValue('Dữ liệu hệ thống');
  }

  if (!ss.getSheetByName('DanhSachUser')) {
    const sh = ss.insertSheet('DanhSachUser');
    sh.getRange(1, 1, 2, 6).setValues([
      ['ID', 'MatKhau', 'VaiTro', 'HoTen', 'ChucDanh', 'PhanMacDinh'],
      ['admin', '123456', 'ADMIN', 'Ban QLDA', '', '']
    ]);
  }

  if (!ss.getSheetByName('NhatKy')) {
    ss.insertSheet('NhatKy').getRange(1, 1, 1, 4)
      .setValues([['Thời gian', 'Người dùng', 'Hành động', 'Chi tiết']]);
  }

  if (!ss.getSheetByName('HangMuc')) {
    ss.insertSheet('HangMuc').getRange(1, 1, 1, HANGMUC_HEADERS_.length).setValues([HANGMUC_HEADERS_]);
  }

  if (!ss.getSheetByName('GoiThau')) {
    ss.insertSheet('GoiThau').getRange(1, 1, 1, GOITHAU_HEADERS_.length).setValues([GOITHAU_HEADERS_]);
  }

  if (!ss.getSheetByName('CongViec')) {
    ss.insertSheet('CongViec').getRange(1, 1, 1, CONGVIEC_HEADERS_.length).setValues([CONGVIEC_HEADERS_]);
  }

  if (!ss.getSheetByName('NhatKyTienDo')) {
    ss.insertSheet('NhatKyTienDo').getRange(1, 1, 1, NHATKY_HEADERS_.length).setValues([NHATKY_HEADERS_]);
  }

  if (!ss.getSheetByName('PhepThang')) {
    ss.insertSheet('PhepThang').getRange(1, 1, 1, PHEP_HEADERS_.length).setValues([PHEP_HEADERS_]);
  }

  if (!ss.getSheetByName('ChamCongThang')) {
    ss.insertSheet('ChamCongThang').getRange(1, 1, 1, CHAMCONG_HEADERS_.length).setValues([CHAMCONG_HEADERS_]);
  }

  if (!ss.getSheetByName('ThongBao')) {
    ss.insertSheet('ThongBao').getRange(1, 1, 1, THONGBAO_HEADERS_.length).setValues([THONGBAO_HEADERS_]);
  }

  if (!ss.getSheetByName('BinhLuanCongViec')) {
    ss.insertSheet('BinhLuanCongViec').getRange(1, 1, 1, BINHLUAN_HEADERS_.length).setValues([BINHLUAN_HEADERS_]);
  }

  if (!ss.getSheetByName('HopDong')) {
    ss.insertSheet('HopDong').getRange(1, 1, 1, HOPDONG_HEADERS_.length).setValues([HOPDONG_HEADERS_]);
  }

  if (!ss.getSheetByName('BOQHangMuc')) {
    ss.insertSheet('BOQHangMuc').getRange(1, 1, 1, BOQ_HEADERS_.length).setValues([BOQ_HEADERS_]);
  }

  if (!ss.getSheetByName('PhuLucHopDong')) {
    ss.insertSheet('PhuLucHopDong').getRange(1, 1, 1, PHULUC_HEADERS_.length).setValues([PHULUC_HEADERS_]);
  }

  if (!ss.getSheetByName('PhuLucThayDoi')) {
    ss.insertSheet('PhuLucThayDoi').getRange(1, 1, 1, PHULUC_THAYDOI_HEADERS_.length).setValues([PHULUC_THAYDOI_HEADERS_]);
  }

  if (!ss.getSheetByName('NghiemThu')) {
    ss.insertSheet('NghiemThu').getRange(1, 1, 1, NGHIEMTHU_HEADERS_.length).setValues([NGHIEMTHU_HEADERS_]);
  }

  if (!ss.getSheetByName('DotThanhToan')) {
    ss.insertSheet('DotThanhToan').getRange(1, 1, 1, DOTTHANHTOAN_HEADERS_.length).setValues([DOTTHANHTOAN_HEADERS_]);
  }

  if (!ss.getSheetByName('DotThanhToanChiTiet')) {
    ss.insertSheet('DotThanhToanChiTiet').getRange(1, 1, 1, DOTTHANHTOAN_CHITIET_HEADERS_.length).setValues([DOTTHANHTOAN_CHITIET_HEADERS_]);
  }

  if (!ss.getSheetByName('QuyetToan')) {
    ss.insertSheet('QuyetToan').getRange(1, 1, 1, QUYETTOAN_HEADERS_.length).setValues([QUYETTOAN_HEADERS_]);
  }

  ss.getSheets().forEach(sh => {
    if (/^(Sheet1|Trang tính1)$/.test(sh.getName()) && ss.getSheets().length > 3) {
      ss.deleteSheet(sh);
    }
  });

  return 'Đã thiết lập xong: Data, DanhSachUser (admin/123456), NhatKy, HangMuc, GoiThau, CongViec, NhatKyTienDo, PhepThang, ChamCongThang, ThongBao, BinhLuanCongViec, HopDong, BOQHangMuc, PhuLucHopDong, PhuLucThayDoi, NghiemThu, DotThanhToan, DotThanhToanChiTiet, QuyetToan.';
}

// Chạy 1 lần khi cần xóa sạch toàn bộ Hạng mục + Công việc kiểu cũ để làm lại từ đầu.
// KHÔNG còn dùng cho luồng Nhật ký tiến độ (nay lưu ở HangMuc/CongViec/NhatKyTienDo) — giữ lại
// phòng khi cần dọn sạch dữ liệu blob cũ trong tab Data.
function resetAllTasks() {
  backupCurrentData();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
  sheet.clearContents();
  sheet.getRange('A1').setValue('Dữ liệu hệ thống');
  sheet.getRange('A2').setValue(JSON.stringify({ tasks: [] }));
  return 'Đã xóa toàn bộ Hạng mục và Công việc (đã backup dữ liệu cũ).';
}

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Quản Lý Công Việc - Ban QLDA')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// =======================================================
// ĐỌC DỮ LIỆU: HangMuc (cây tham chiếu) + GoiThau + CongViec + NhatKyTienDo (log cộng dồn)
// =======================================================

function isActiveVal_(v) {
  return v !== false && String(v).trim().toUpperCase() !== 'FALSE';
}

function formatDateCell_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v).trim();
}

// Cột "thang" (yyyy-MM) ghi bằng chuỗi nhưng Google Sheets tự nhận diện là ngày tháng và âm thầm
// đổi thành giá trị Date (thành ngày 1 của tháng đó) — nếu đọc lại bằng String() thường sẽ SAI
// hoàn toàn (không khớp được "2026-07"), khiến mọi so khớp userId+thang thất bại. Phải test
// instanceof Date như formatDateCell_ để đọc lại đúng.
function formatThangCell_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM');
  return String(v).trim();
}

// Cột giờ tăng ca (gioBatDauTC/gioKetThucTC) ghi bằng chuỗi 'HH:mm' nhưng Google Sheets tự nhận diện
// là giờ và âm thầm đổi thành giá trị Date (ngày 30/12/1899 — mốc epoch giờ của Sheets) — cùng lỗi
// với formatDateCell_/formatThangCell_ ở trên, đọc lại bằng String() sẽ ra cả chuỗi ngày tháng rác.
// Một số dòng cũ (nhập TRƯỚC khi có hàm này) đã lỡ bị đọc + ghi ngược lại đúng chuỗi rác đó dưới dạng
// text thường (không còn là Date nữa) — CỐ parse lại bằng new Date() không đáng tin (offset giờ lịch
// sử của Đông Dương lệch vài phút, ra giờ sai mà nhìn tưởng đúng), nên coi mọi giá trị không đúng
// đúng dạng 'H:mm' là dữ liệu hỏng, trả rỗng để giao diện coi như chưa nhập — vào lại popup Tăng ca
// nhập lại là Sheet ghi giá trị sạch, tự hết lỗi từ đó về sau.
function formatTimeCell_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm');
  const s = String(v).trim();
  return /^\d{1,2}:\d{2}$/.test(s) ? s : '';
}

function formatDateTimeCell_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
  return String(v).trim();
}

function readHangMucRows_(sheet) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, HANGMUC_HEADERS_.length).getValues();
  const result = [];
  values.forEach(r => {
    const maHangMuc = String(r[0] || '').trim();
    if (!maHangMuc) return;
    if (!isActiveVal_(r[4])) return;
    result.push({
      maHangMuc,
      tenHangMuc: String(r[1] || '').trim(),
      capDo: parseInt(r[2], 10) || 1,
      parentId: String(r[3] || '').trim(),
      // Dòng cũ chưa từng gán stt sẽ đọc ra '' -> mặc định 0, các dòng mới tạo có stt tăng dần
      // (xem adminSaveHangMuc) nên vẫn xếp SAU mọi dòng cũ khi sort tăng dần, đúng thứ tự hiện có.
      stt: parseInt(r[5], 10) || 0
    });
  });
  return result;
}

function readGoiThauRows_(sheet) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, GOITHAU_HEADERS_.length).getValues();
  const result = [];
  values.forEach(r => {
    const maGoiThau = String(r[0] || '').trim();
    if (!maGoiThau) return;
    result.push({
      maGoiThau,
      tenGoiThau: String(r[1] || '').trim(),
      maHangMucList: String(r[2] || '').split(',').map(s => s.trim()).filter(Boolean),
      stt: parseInt(r[3], 10) || 0
    });
  });
  return result;
}

// includeInactive: true để lấy cả Công việc đã bị Giám sát xóa (active=false) — dùng cho tab
// "Thi công / Tiến độ" hiển thị gạch ngang, KHÔNG dùng cho báo cáo tổng hợp (xem getBaoCaoTongHop).
function readCongViecRows_(sheet, includeInactive) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, CONGVIEC_HEADERS_.length).getValues();
  const result = [];
  values.forEach((r, idx) => {
    const maCongViec = String(r[0] || '').trim();
    if (!maCongViec) return;
    const active = isActiveVal_(r[9]);
    if (!active && !includeInactive) return;
    result.push({
      maCongViec,
      tenCongViec: String(r[1] || '').trim(),
      maGoiThau: String(r[2] || '').trim(),
      maHangMucList: String(r[3] || '').split(',').map(s => s.trim()).filter(Boolean),
      phanLoai: String(r[4] || '').trim(),
      nguoiTaoId: String(r[5] || '').trim(),
      nguoiTaoTen: String(r[6] || '').trim(),
      ngayBatDauKH: formatDateCell_(r[7]),
      ngayKetThucKH: formatDateCell_(r[8]),
      active,
      nguoiGiaoId: String(r[10] || '').trim(),
      nguoiGiaoTen: String(r[11] || '').trim(),
      rowIndex: idx + 2
    });
  });
  return result;
}

function readNhatKyRows_(sheet) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, NHATKY_HEADERS_.length).getValues();
  const result = [];
  values.forEach((r, idx) => {
    const logId = String(r[0] || '').trim();
    if (!logId) return;
    result.push({
      logId,
      maCongViec: String(r[1] || '').trim(),
      ngayBaoCao: formatDateCell_(r[2]),
      phanTramNgay: Number(r[3]) || 0,
      ghiChu: String(r[4] || '').trim(),
      nguoiNhapId: String(r[5] || '').trim(),
      nguoiNhapTen: String(r[6] || '').trim(),
      thoiGianNhap: formatDateTimeCell_(r[7]),
      fileDinhKem: String(r[8] || '').split(';').map(s => s.trim()).filter(Boolean),
      active: isActiveVal_(r[9]),
      rowIndex: idx + 2
    });
  });
  return result;
}

// Trả về { hangMuc, goiThau, congViec, users } — congViec đã tính sẵn luyKe/trangThaiMau/recentLogs
// (cộng dồn từ NhatKyTienDo theo maCongViec, 1 lần đọc sheet duy nhất — KHÔNG tách riêng hàm/lệnh
// gọi khác để lấy recentLogs nữa, tránh round-trip kép lên Apps Script làm chậm lần tải đầu).
// trangThaiMau là 1 trong 4 giá trị, xét theo hạn ngayKetThucKH (KHÔNG nội suy % kế hoạch theo
// thời gian đã trôi qua như trước):
// - pending    : chưa hoàn thành và chưa tới ngayBatDauKH (chưa tới thời điểm bắt đầu)
// - inprogress : chưa hoàn thành (luyKe<100), đã tới ngày bắt đầu và còn trong hạn
// - done       : đã hoàn thành (luyKe>=100) đúng hạn (hoặc chưa đặt hạn để so)
// - overdue    : hôm nay đã qua ngayKetThucKH và CHƯA hoàn thành, HOẶC đã hoàn thành nhưng ngày
//                hoàn thành thực tế (log đưa lũy kế chạm mốc 100% đầu tiên) trễ hơn ngayKetThucKH
//                (gộp chung "trễ nhưng xong" vào "trễ tiến độ", không tách riêng nữa)
// congViec ở đây gồm CẢ Công việc đã bị Giám sát xóa (active=false) — để tab "Thi công / Tiến độ"
// hiển thị gạch ngang; getBaoCaoTongHop() bên dưới tự lọc bỏ trước khi trả về báo cáo tổng hợp.
function getData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hangMuc = readHangMucRows_(ss.getSheetByName('HangMuc'));
  const goiThau = readGoiThauRows_(ss.getSheetByName('GoiThau'));
  const congViec = readCongViecRows_(ss.getSheetByName('CongViec'), true);
  const logs = readNhatKyRows_(ss.getSheetByName('NhatKyTienDo')).filter(l => l.active);
  const binhLuan = readBinhLuanRows_(ss.getSheetByName('BinhLuanCongViec'));

  const luyKeMap = {};
  const logsByCongViec = {};
  logs.forEach(l => {
    luyKeMap[l.maCongViec] = (luyKeMap[l.maCongViec] || 0) + l.phanTramNgay;
    (logsByCongViec[l.maCongViec] = logsByCongViec[l.maCongViec] || []).push(l);
  });
  const recentLogsByCongViec = {};
  Object.keys(logsByCongViec).forEach(id => {
    recentLogsByCongViec[id] = logsByCongViec[id].slice().sort((a, b) => b.logId.localeCompare(a.logId)).slice(0, 5);
    logsByCongViec[id].sort((a, b) => a.ngayBaoCao.localeCompare(b.ngayBaoCao));
  });

  // Gắn sẵn số lượng + 5 bình luận gần nhất cho mỗi Công việc (giống recentLogs ở trên) — để chuông
  // số trên tên Công việc + tooltip xem nhanh không cần round-trip getBinhLuanCongViec() riêng.
  const binhLuanByCongViec = {};
  binhLuan.forEach(c => { (binhLuanByCongViec[c.maCongViec] = binhLuanByCongViec[c.maCongViec] || []).push(c); });

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  congViec.forEach(cv => {
    const luyKe = Math.round((luyKeMap[cv.maCongViec] || 0) * 10) / 10;
    cv.luyKe = luyKe;
    cv.recentLogs = recentLogsByCongViec[cv.maCongViec] || [];
    const cmts = binhLuanByCongViec[cv.maCongViec] || [];
    cv.soBinhLuan = cmts.length;
    cv.recentComments = cmts.slice().sort((a, b) => b.id.localeCompare(a.id)).slice(0, 5);

    if (luyKe >= 100) {
      let ngayHoanThanhThucTe = null;
      let running = 0;
      const logsOfThis = logsByCongViec[cv.maCongViec] || [];
      for (let i = 0; i < logsOfThis.length; i++) {
        running += logsOfThis[i].phanTramNgay;
        if (running >= 99.999) { ngayHoanThanhThucTe = logsOfThis[i].ngayBaoCao; break; }
      }
      cv.trangThaiMau = (cv.ngayKetThucKH && ngayHoanThanhThucTe && ngayHoanThanhThucTe > cv.ngayKetThucKH) ? 'overdue' : 'done';
      if (cv.trangThaiMau === 'overdue') cv.soNgayTre = soNgayGiuaHaiNgayISO_(ngayHoanThanhThucTe, cv.ngayKetThucKH);
    } else if (cv.ngayBatDauKH && today < cv.ngayBatDauKH) {
      cv.trangThaiMau = 'pending';
    } else {
      cv.trangThaiMau = (cv.ngayKetThucKH && today > cv.ngayKetThucKH) ? 'overdue' : 'inprogress';
      if (cv.trangThaiMau === 'overdue') cv.soNgayTre = soNgayGiuaHaiNgayISO_(today, cv.ngayKetThucKH);
    }
  });

  return JSON.stringify({ hangMuc, goiThau, congViec, users: listSupervisors() });
}

// Giống getData() nhưng chỉ giữ Công việc còn hoạt động (active) — dùng cho màn Báo cáo tổng hợp.
// KHÔNG đọc lại NhatKyTienDo nữa (getData() ở trên đã gắn sẵn recentLogs), tránh round-trip kép.
function getBaoCaoTongHop() {
  const parsed = JSON.parse(getData());
  // Công việc đã bị Giám sát xóa KHÔNG được tính vào báo cáo tổng hợp/Tổng quan — chỉ hiển thị
  // (gạch ngang) ở tab "Thi công / Tiến độ" thông qua getData() ở trên.
  parsed.congViec = parsed.congViec.filter(cv => cv.active);
  return JSON.stringify(parsed);
}

function getNhatKyGanNhat(maCongViec, limit) {
  const logs = readNhatKyRows_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('NhatKyTienDo'))
    .filter(l => l.active && l.maCongViec === maCongViec);
  logs.sort((a, b) => b.logId.localeCompare(a.logId));
  return JSON.stringify(logs.slice(0, limit || 5));
}

// Đọc danh sách tài khoản (Giám sát/Admin) trực tiếp từ tab DanhSachUser — KHÔNG trả về mật khẩu.
function listSupervisors() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DanhSachUser');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    result.push({
      id: String(data[i][0]).trim(),
      role: String(data[i][2]).trim().toUpperCase(),
      name: String(data[i][3]).trim(),
      chucDanh: String(data[i][4] || '').trim(),
      phanMacDinh: String(data[i][5] || '').trim()
    });
  }
  return result;
}

// (Admin) Thêm mới hoặc sửa 1 tài khoản Giám sát trực tiếp vào tab DanhSachUser.
// chucDanh: nhãn hiển thị tự do (vd "QS", "QAQC", "Giám sát hiện trường").
// phanMacDinh: 'ho_so' hoặc 'thi_cong' — quyết định Công việc do người này tự tạo rơi vào phần nào;
// 'ca_hai' = làm cả 2 phần, phải tự chọn Phần cho từng Công việc lúc tạo (xem adminSaveCongViec).
function adminSaveSupervisor(adminId, adminPass, targetId, name, role, newPassword, chucDanh, phanMacDinh) {
  const adminUser = login(adminId, adminPass);
  if (adminUser.role !== 'ADMIN') throw new Error("Chỉ Quản trị mới có quyền quản lý tài khoản!");

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DanhSachUser');
  if (!sheet) throw new Error("Chưa cấu hình Tab DanhSachUser trên Google Sheet!");
  if (!sheet.getRange(1, 5).getValue()) sheet.getRange(1, 5).setValue('ChucDanh');
  if (!sheet.getRange(1, 6).getValue()) sheet.getRange(1, 6).setValue('PhanMacDinh');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(targetId).trim()) {
      sheet.getRange(i + 1, 3).setValue(role);
      sheet.getRange(i + 1, 4).setValue(name);
      sheet.getRange(i + 1, 5).setValue(chucDanh || '');
      sheet.getRange(i + 1, 6).setValue(phanMacDinh || '');
      if (newPassword) sheet.getRange(i + 1, 2).setValue(newPassword); // lưu thường, tự mã hóa ở lần đăng nhập đầu (xem login())
      logActivity(adminId, adminUser.name, "Sửa Giám sát", `Sửa tài khoản ${targetId}`);
      return "Success";
    }
  }

  if (!newPassword) throw new Error("Vui lòng đặt mật khẩu ban đầu cho tài khoản mới!");
  sheet.appendRow([targetId, newPassword, role, name, chucDanh || '', phanMacDinh || '']);
  logActivity(adminId, adminUser.name, "Thêm Giám sát", `Thêm tài khoản ${targetId}`);
  return "Success";
}

// (Admin) Xóa 1 tài khoản Giám sát khỏi tab DanhSachUser.
function adminDeleteSupervisor(adminId, adminPass, targetId) {
  const adminUser = login(adminId, adminPass);
  if (adminUser.role !== 'ADMIN') throw new Error("Chỉ Quản trị mới có quyền quản lý tài khoản!");

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DanhSachUser');
  if (!sheet) throw new Error("Chưa cấu hình Tab DanhSachUser trên Google Sheet!");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(targetId).trim()) {
      sheet.deleteRow(i + 1);
      logActivity(adminId, adminUser.name, "Xóa Giám sát", `Xóa tài khoản ${targetId}`);
      return "Success";
    }
  }
  throw new Error("Không tìm thấy tài khoản!");
}

// =======================================================
// QUẢN LÝ CÂY HẠNG MỤC (mã tham chiếu, chỉ Admin) + GÓI THẦU
// =======================================================

// node = { maHangMuc, tenHangMuc, capDo, parentId } — chỉ Admin được thêm/sửa/xoá.
// node: { maHangMuc (mã mới), oldMaHangMuc (mã cũ khi sửa — rỗng khi tạo mới), tenHangMuc, capDo, parentId }
function adminSaveHangMuc(userId, password, node) {
  const user = login(userId, password);
  if (user.role !== 'ADMIN') throw new Error('Chỉ Quản trị mới có quyền quản lý Hạng mục!');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('HangMuc');
  if (!sheet) throw new Error('Chưa cấu hình Tab HangMuc trên Google Sheet!');

  const maHangMuc = String(node.maHangMuc || '').trim();
  const oldMaHangMuc = String(node.oldMaHangMuc || '').trim();
  if (!maHangMuc) throw new Error('Vui lòng nhập Mã hạng mục!');

  const data = sheet.getDataRange().getValues();
  const isRename = oldMaHangMuc && oldMaHangMuc !== maHangMuc;

  if (isRename && data.some((r, i) => i > 0 && String(r[0]).trim() === maHangMuc)) {
    throw new Error(`Mã Hạng mục "${maHangMuc}" đã tồn tại, vui lòng chọn mã khác!`);
  }

  const keyToFind = oldMaHangMuc || maHangMuc;
  let foundRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === keyToFind) { foundRow = i + 1; break; }
  }

  let stt = parseInt(node.stt, 10);
  if (!stt) {
    let maxStt = 0;
    data.forEach((r, i) => { if (i > 0) { const s = parseInt(r[5], 10) || 0; if (s > maxStt) maxStt = s; } });
    stt = maxStt + 1;
  }
  const row = [maHangMuc, node.tenHangMuc || '', parseInt(node.capDo, 10) || 1, node.parentId || '', true, stt];

  if (foundRow > 0) {
    sheet.getRange(foundRow, 1, 1, row.length).setValues([row]);
    logActivity(userId, user.name, 'Sửa Hạng mục', `Sửa ${keyToFind}${isRename ? ' → ' + maHangMuc : ''} - ${node.tenHangMuc}`);
  } else {
    sheet.appendRow(row);
    logActivity(userId, user.name, 'Thêm Hạng mục', `Thêm ${maHangMuc} - ${node.tenHangMuc}`);
  }

  // Đổi mã thì phải cập nhật dây chuyền mọi nơi đang tham chiếu mã cũ, nếu không Hạng mục con/
  // Gói thầu/Công việc sẽ bị "mồ côi" (trỏ vào 1 mã không còn tồn tại).
  if (isRename) capNhatMaHangMucThamChieu_(ss, oldMaHangMuc, maHangMuc);
  return 'Success';
}

function capNhatMaHangMucThamChieu_(ss, oldMa, newMa) {
  const hmSheet = ss.getSheetByName('HangMuc');
  if (hmSheet) {
    const data = hmSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][3]).trim() === oldMa) hmSheet.getRange(i + 1, 4).setValue(newMa);
    }
  }

  const gtSheet = ss.getSheetByName('GoiThau');
  if (gtSheet) {
    const data = gtSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const list = String(data[i][2] || '').split(',').map(s => s.trim()).filter(Boolean);
      if (list.includes(oldMa)) {
        gtSheet.getRange(i + 1, 3).setValue(list.map(id => id === oldMa ? newMa : id).join(','));
      }
    }
  }

  const cvSheet = ss.getSheetByName('CongViec');
  if (cvSheet) {
    const data = cvSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const list = String(data[i][3] || '').split(',').map(s => s.trim()).filter(Boolean);
      if (list.includes(oldMa)) {
        cvSheet.getRange(i + 1, 4).setValue(list.map(id => id === oldMa ? newMa : id).join(','));
      }
    }
  }
}

// Soft delete (active=false) — không xoá dòng thật để giữ toàn vẹn tham chiếu với CongViec.
function adminDeleteHangMuc(userId, password, maHangMuc) {
  const user = login(userId, password);
  if (user.role !== 'ADMIN') throw new Error('Chỉ Quản trị mới có quyền quản lý Hạng mục!');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('HangMuc');
  if (!sheet) throw new Error('Chưa cấu hình Tab HangMuc trên Google Sheet!');
  const data = sheet.getDataRange().getValues();

  const hasChildren = data.some((r, i) => i > 0 && isActiveVal_(r[4]) && String(r[3]).trim() === String(maHangMuc).trim());
  if (hasChildren) throw new Error('Hạng mục này còn Hạng mục con, vui lòng xoá các Hạng mục con trước!');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(maHangMuc).trim()) {
      sheet.getRange(i + 1, 5).setValue(false);
      logActivity(userId, user.name, 'Xoá Hạng mục', `Xoá ${maHangMuc}`);
      return 'Success';
    }
  }
  throw new Error('Không tìm thấy Hạng mục!');
}

function adminSaveGoiThau(userId, password, goiThau) {
  const user = login(userId, password);
  if (user.role !== 'ADMIN') throw new Error('Chỉ Quản trị mới có quyền quản lý Gói thầu!');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GoiThau');
  if (!sheet) throw new Error('Chưa cấu hình Tab GoiThau trên Google Sheet!');

  const tenGoiThau = String(goiThau.tenGoiThau || '').trim();
  if (!tenGoiThau) throw new Error('Vui lòng nhập Tên gói thầu!');
  const maHangMucList = (goiThau.maHangMucList || []).join(',');

  const data = sheet.getDataRange().getValues();
  let maGoiThau = String(goiThau.maGoiThau || '').trim();
  let stt = parseInt(goiThau.stt, 10);

  if (maGoiThau) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === maGoiThau) {
        if (!stt) stt = parseInt(data[i][3], 10) || 0;
        sheet.getRange(i + 1, 2, 1, 3).setValues([[tenGoiThau, maHangMucList, stt]]);
        logActivity(userId, user.name, 'Sửa Gói thầu', tenGoiThau);
        return 'Success';
      }
    }
  }

  let maxSTT = 0;
  let maxDisplayStt = 0;
  for (let i = 1; i < data.length; i++) {
    const s = parseInt(String(data[i][0]).replace('GT', ''), 10);
    if (!isNaN(s) && s > maxSTT) maxSTT = s;
    const ds = parseInt(data[i][3], 10) || 0;
    if (ds > maxDisplayStt) maxDisplayStt = ds;
  }
  maGoiThau = 'GT' + (maxSTT + 1);
  if (!stt) stt = maxDisplayStt + 1;
  sheet.appendRow([maGoiThau, tenGoiThau, maHangMucList, stt]);
  logActivity(userId, user.name, 'Thêm Gói thầu', tenGoiThau);
  return 'Success';
}

function adminDeleteGoiThau(userId, password, maGoiThau) {
  const user = login(userId, password);
  if (user.role !== 'ADMIN') throw new Error('Chỉ Quản trị mới có quyền quản lý Gói thầu!');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GoiThau');
  if (!sheet) throw new Error('Chưa cấu hình Tab GoiThau trên Google Sheet!');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(maGoiThau).trim()) {
      sheet.deleteRow(i + 1);
      logActivity(userId, user.name, 'Xoá Gói thầu', maGoiThau);
      return 'Success';
    }
  }
  throw new Error('Không tìm thấy Gói thầu!');
}

// =======================================================
// CÔNG VIỆC — do Admin hoặc Giám sát tự tạo, gắn Gói thầu + Hạng mục (để lọc),
// Phần (hồ sơ/thi công) tự suy ra từ phanMacDinh của người tạo, không chọn tay.
// =======================================================

function sinhMaCongViec_(sheet, now) {
  const ngay = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd');
  const prefix = 'CV-' + ngay + '-';
  const lastRow = sheet.getLastRow();
  let maxSTT = 0;
  if (lastRow >= 2) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    ids.forEach(r => {
      const id = String(r[0] || '');
      if (id.indexOf(prefix) === 0) {
        const stt = parseInt(id.slice(prefix.length), 10);
        if (!isNaN(stt) && stt > maxSTT) maxSTT = stt;
      }
    });
  }
  return prefix + String(maxSTT + 1).padStart(3, '0');
}

// congViec = { maCongViec (rỗng nếu tạo mới), tenCongViec, maGoiThau, maHangMucList (mảng, chọn
// được 1 hoặc nhiều Hạng mục), ngayBatDauKH, ngayKetThucKH, nguoiDuocGiaoId (optional — CHỈ Admin,
// CHỈ lúc tạo mới, giao Công việc cho 1 Giám sát khác thay vì tự đảm nhận) }
function adminSaveCongViec(userId, password, congViec) {
  const user = login(userId, password);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CongViec');
  if (!sheet) throw new Error('Chưa cấu hình Tab CongViec trên Google Sheet!');

  const tenCongViec = String(congViec.tenCongViec || '').trim();
  const maGoiThau = String(congViec.maGoiThau || '').trim();
  const maHangMucJoined = (congViec.maHangMucList || []).map(s => String(s).trim()).filter(Boolean).join(',');
  if (!tenCongViec) throw new Error('Vui lòng nhập Tên công việc!');
  if (!maGoiThau) throw new Error('Vui lòng chọn Gói thầu!');
  if (!maHangMucJoined) throw new Error('Vui lòng chọn ít nhất 1 Hạng mục!');

  const maCongViec = String(congViec.maCongViec || '').trim();

  // Chỉ đọc cột A (mã) để dò dòng cần sửa thay vì tải nguyên sheet — sheet CongViec tích lũy nhiều
  // năm dữ liệu, đọc đủ 12 cột x toàn bộ dòng mỗi lần Lưu (kể cả khi Thêm mới, không dùng tới) là
  // nguyên nhân chính khiến thao tác Lưu bị chậm.
  const lastRow = sheet.getLastRow();
  if (maCongViec) {
    const ids = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, 1).getValues() : [];
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === maCongViec) {
        const rowIdx = i + 2;
        const rowData = sheet.getRange(rowIdx, 1, 1, CONGVIEC_HEADERS_.length).getValues()[0];
        if (user.role !== 'ADMIN' && String(rowData[5]).trim() !== user.supervisorId) {
          throw new Error('Bạn không có quyền sửa Công việc này!');
        }
        // Giữ nguyên phanLoai (cột E), người phụ trách (cột F/G) và người giao (cột K/L) — chỉ cho
        // sửa các field mô tả, không cho đổi tay người phụ trách sau khi đã tạo.
        sheet.getRange(rowIdx, 2, 1, 4).setValues([[tenCongViec, maGoiThau, maHangMucJoined, rowData[4]]]);
        sheet.getRange(rowIdx, 8, 1, 2).setValues([[congViec.ngayBatDauKH || '', congViec.ngayKetThucKH || '']]);
        logActivity(userId, user.name, 'Sửa Công việc', `Sửa ${maCongViec} - ${tenCongViec}`);
        return 'Success';
      }
    }
    throw new Error('Không tìm thấy Công việc!');
  }

  const allUsers = listSupervisors();
  let nguoiPhuTrachId = user.supervisorId;
  let nguoiPhuTrachTen = user.name;
  let nguoiGiaoId = '';
  let nguoiGiaoTen = '';

  const nguoiDuocGiaoId = String(congViec.nguoiDuocGiaoId || '').trim();
  if (nguoiDuocGiaoId && nguoiDuocGiaoId !== user.supervisorId) {
    if (user.role !== 'ADMIN') throw new Error('Chỉ Quản trị mới có quyền giao việc cho người khác!');
    const target = allUsers.find(u => u.id === nguoiDuocGiaoId);
    if (!target) throw new Error('Không tìm thấy nhân sự được giao!');
    nguoiPhuTrachId = target.id;
    nguoiPhuTrachTen = target.name;
    nguoiGiaoId = user.supervisorId;
    nguoiGiaoTen = user.name;
  }

  const nguoiPhuTrach = allUsers.find(u => u.id === nguoiPhuTrachId);
  const phanMacDinh = String((nguoiPhuTrach && nguoiPhuTrach.phanMacDinh) || '').trim();
  if (!phanMacDinh) {
    throw new Error(nguoiGiaoId
      ? `Nhân sự "${nguoiPhuTrachTen}" chưa được Quản trị gán Phần mặc định (Chức danh). Vui lòng gán trước khi giao việc.`
      : 'Tài khoản của bạn chưa được Quản trị gán Phần mặc định (Chức danh). Vui lòng liên hệ Quản trị trước khi tự tạo Công việc.');
  }

  // "Cả hai" (phanMacDinh='ca_hai') không tự suy ra được — bắt buộc client gửi kèm phanLoai
  // (xem dropdown "Phần" ở modal Thêm công việc hàng loạt, 4_ViecCuaToi.html).
  let phanLoai = phanMacDinh;
  if (phanMacDinh === 'ca_hai') {
    phanLoai = String(congViec.phanLoai || '').trim();
    if (phanLoai !== 'ho_so' && phanLoai !== 'thi_cong') {
      throw new Error(`Nhân sự "${nguoiPhuTrachTen}" được gán "Cả hai" — vui lòng chọn Phần (Hồ sơ/Thi công) cho công việc này!`);
    }
  }

  const now = new Date();
  const newId = sinhMaCongViec_(sheet, now);
  sheet.appendRow([
    newId, tenCongViec, maGoiThau, maHangMucJoined, phanLoai,
    nguoiPhuTrachId, nguoiPhuTrachTen,
    congViec.ngayBatDauKH || '', congViec.ngayKetThucKH || '', true,
    nguoiGiaoId, nguoiGiaoTen
  ]);
  logActivity(userId, user.name, 'Thêm Công việc', `Thêm ${newId} - ${tenCongViec}`);

  if (nguoiGiaoId) {
    guiThongBao_(nguoiPhuTrachId, 'giao_viec', `Bạn được giao công việc mới: "${tenCongViec}"`, newId, nguoiGiaoId, nguoiGiaoTen);
  }

  return newId;
}

// payloadArray: mảng { tenCongViec, maGoiThau, maHangMucList, nguoiDuocGiaoId, phanLoai,
// ngayBatDauKH, ngayKetThucKH } — GIỐNG HỆT tham số congViec của adminSaveCongViec() ở trên, chỉ
// khác là LUÔN tạo mới (modal "Thêm công việc hàng loạt" ở 4_ViecCuaToi.html không có sửa). Gộp
// login + đọc DanhSachUser + sinh mã + ghi Sheet thành ĐÚNG 1 lượt gọi từ client thay vì client gọi
// lặp adminSaveCongViec() từng dòng (N dòng = N round-trip riêng lên Apps Script, mỗi lượt round-trip
// tốn 1-3s do kiến trúc google.script.run — đây là nguyên nhân chính khiến "Thêm công việc" (luôn là
// bảng nhiều dòng) cảm giác lưu rất lâu). Trả về JSON mảng kết quả song song với payloadArray, mỗi
// phần tử { ok, maCongViec } hoặc { ok:false, error } — 1 dòng lỗi không làm hỏng các dòng còn lại.
function adminSaveCongViecHangLoat(userId, password, payloadArray) {
  const user = login(userId, password);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CongViec');
  if (!sheet) throw new Error('Chưa cấu hình Tab CongViec trên Google Sheet!');

  const allUsers = listSupervisors();
  const now = new Date();
  const nowStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
  const prefix = 'CV-' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd') + '-';

  const lastRow = sheet.getLastRow();
  let maxSTT = 0;
  if (lastRow >= 2) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    ids.forEach(r => {
      const id = String(r[0] || '');
      if (id.indexOf(prefix) === 0) {
        const stt = parseInt(id.slice(prefix.length), 10);
        if (!isNaN(stt) && stt > maxSTT) maxSTT = stt;
      }
    });
  }

  const results = [];
  const newRows = [];
  const logRows = [];
  const thongBaoList = [];

  (payloadArray || []).forEach(congViec => {
    try {
      const tenCongViec = String(congViec.tenCongViec || '').trim();
      const maGoiThau = String(congViec.maGoiThau || '').trim();
      const maHangMucJoined = (congViec.maHangMucList || []).map(s => String(s).trim()).filter(Boolean).join(',');
      if (!tenCongViec) throw new Error('Vui lòng nhập Tên công việc!');
      if (!maGoiThau) throw new Error('Vui lòng chọn Gói thầu!');
      if (!maHangMucJoined) throw new Error('Vui lòng chọn ít nhất 1 Hạng mục!');

      let nguoiPhuTrachId = user.supervisorId;
      let nguoiPhuTrachTen = user.name;
      let nguoiGiaoId = '';
      let nguoiGiaoTen = '';

      const nguoiDuocGiaoId = String(congViec.nguoiDuocGiaoId || '').trim();
      if (nguoiDuocGiaoId && nguoiDuocGiaoId !== user.supervisorId) {
        if (user.role !== 'ADMIN') throw new Error('Chỉ Quản trị mới có quyền giao việc cho người khác!');
        const target = allUsers.find(u => u.id === nguoiDuocGiaoId);
        if (!target) throw new Error('Không tìm thấy nhân sự được giao!');
        nguoiPhuTrachId = target.id;
        nguoiPhuTrachTen = target.name;
        nguoiGiaoId = user.supervisorId;
        nguoiGiaoTen = user.name;
      }

      const nguoiPhuTrach = allUsers.find(u => u.id === nguoiPhuTrachId);
      const phanMacDinh = String((nguoiPhuTrach && nguoiPhuTrach.phanMacDinh) || '').trim();
      if (!phanMacDinh) {
        throw new Error(nguoiGiaoId
          ? `Nhân sự "${nguoiPhuTrachTen}" chưa được Quản trị gán Phần mặc định (Chức danh). Vui lòng gán trước khi giao việc.`
          : 'Tài khoản của bạn chưa được Quản trị gán Phần mặc định (Chức danh). Vui lòng liên hệ Quản trị trước khi tự tạo Công việc.');
      }

      let phanLoai = phanMacDinh;
      if (phanMacDinh === 'ca_hai') {
        phanLoai = String(congViec.phanLoai || '').trim();
        if (phanLoai !== 'ho_so' && phanLoai !== 'thi_cong') {
          throw new Error(`Nhân sự "${nguoiPhuTrachTen}" được gán "Cả hai" — vui lòng chọn Phần (Hồ sơ/Thi công) cho công việc này!`);
        }
      }

      maxSTT++;
      const newId = prefix + String(maxSTT).padStart(3, '0');
      newRows.push([
        newId, tenCongViec, maGoiThau, maHangMucJoined, phanLoai,
        nguoiPhuTrachId, nguoiPhuTrachTen,
        congViec.ngayBatDauKH || '', congViec.ngayKetThucKH || '', true,
        nguoiGiaoId, nguoiGiaoTen
      ]);
      logRows.push([nowStr, `${user.name} (${userId})`, 'Thêm Công việc', `Thêm ${newId} - ${tenCongViec}`]);
      if (nguoiGiaoId) thongBaoList.push({ nguoiPhuTrachId, tenCongViec, newId, nguoiGiaoId, nguoiGiaoTen });
      results.push({ ok: true, maCongViec: newId });
    } catch (e) {
      results.push({ ok: false, error: e.message });
    }
  });

  if (newRows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, CONGVIEC_HEADERS_.length).setValues(newRows);
  }
  if (logRows.length) {
    const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('NhatKy');
    if (logSheet) logSheet.getRange(logSheet.getLastRow() + 1, 1, logRows.length, 4).setValues(logRows);
  }
  thongBaoList.forEach(tb => guiThongBao_(tb.nguoiPhuTrachId, 'giao_viec',
    `Bạn được giao công việc mới: "${tb.tenCongViec}"`, tb.newId, tb.nguoiGiaoId, tb.nguoiGiaoTen));

  return JSON.stringify(results);
}

function adminDeleteCongViec(userId, password, maCongViec) {
  const user = login(userId, password);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CongViec');
  if (!sheet) throw new Error('Chưa cấu hình Tab CongViec trên Google Sheet!');
  const lastRow = sheet.getLastRow();
  const ids = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, 1).getValues() : [];
  const maTrim = String(maCongViec).trim();

  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === maTrim) {
      const rowIdx = i + 2;
      if (user.role !== 'ADMIN') {
        const nguoiTaoId = String(sheet.getRange(rowIdx, 6).getValue()).trim();
        if (nguoiTaoId !== user.supervisorId) throw new Error('Bạn không có quyền xoá Công việc này!');
      }
      sheet.getRange(rowIdx, 10).setValue(false);
      logActivity(userId, user.name, 'Xoá Công việc', `Xoá ${maCongViec}`);
      return 'Success';
    }
  }
  throw new Error('Không tìm thấy Công việc!');
}

// =======================================================
// THÔNG BÁO & BÌNH LUẬN CÔNG VIỆC — báo cho nhân sự khi được giao việc mới hoặc có người bình
// luận vào việc mình phụ trách. KHÔNG có polling định kỳ (Apps Script không đẩy realtime được) —
// frontend chỉ hỏi lại sau mỗi lần reloadData() (bấm "Làm mới"/sau khi lưu/mở lại trang).
// =======================================================

function sinhMaTuTang_(sheet, prefix, now) {
  const ngay = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd');
  const fullPrefix = prefix + '-' + ngay + '-';
  const lastRow = sheet.getLastRow();
  let maxSTT = 0;
  if (lastRow >= 2) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    ids.forEach(r => {
      const id = String(r[0] || '');
      if (id.indexOf(fullPrefix) === 0) {
        const stt = parseInt(id.slice(fullPrefix.length), 10);
        if (!isNaN(stt) && stt > maxSTT) maxSTT = stt;
      }
    });
  }
  return fullPrefix + String(maxSTT + 1).padStart(3, '0');
}

// Gửi 1 thông báo cho 1 người — tự bỏ qua nếu không có người nhận hoặc người nhận chính là người
// gây ra hành động (không cần tự báo cho chính mình).
function guiThongBao_(userId, loai, noiDung, maCongViec, nguoiGuiId, nguoiGuiTen) {
  if (!userId || userId === nguoiGuiId) return;
  const sheet = layHoacTaoSheet_('ThongBao', THONGBAO_HEADERS_);
  const id = sinhMaTuTang_(sheet, 'TB', new Date());
  sheet.appendRow([id, userId, loai, noiDung, maCongViec || '', nguoiGuiId || '', nguoiGuiTen || '', new Date(), false, true]);
}

function readThongBaoRows_(sheet) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, THONGBAO_HEADERS_.length).getValues();
  const result = [];
  values.forEach((r, idx) => {
    const id = String(r[0] || '').trim();
    if (!id) return;
    if (!isActiveVal_(r[9])) return;
    result.push({
      id,
      userId: String(r[1] || '').trim(),
      loai: String(r[2] || '').trim(),
      noiDung: String(r[3] || '').trim(),
      maCongViec: String(r[4] || '').trim(),
      nguoiGuiId: String(r[5] || '').trim(),
      nguoiGuiTen: String(r[6] || '').trim(),
      thoiGian: formatDateTimeCell_(r[7]),
      daDoc: r[8] === true || String(r[8]).trim().toUpperCase() === 'TRUE',
      rowIndex: idx + 2
    });
  });
  return result;
}

// Trả về thông báo của đúng user gọi hàm, mới nhất trước.
function getThongBaoCuaToi(userId, password) {
  const user = login(userId, password);
  const sheet = layHoacTaoSheet_('ThongBao', THONGBAO_HEADERS_);
  const list = readThongBaoRows_(sheet).filter(t => t.userId === user.supervisorId);
  list.sort((a, b) => b.id.localeCompare(a.id));
  return JSON.stringify(list);
}

function danhDauDaDocThongBao(userId, password, thongBaoId) {
  const user = login(userId, password);
  const sheet = layHoacTaoSheet_('ThongBao', THONGBAO_HEADERS_);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(thongBaoId).trim()) {
      if (String(data[i][1]).trim() !== user.supervisorId) throw new Error('Thông báo này không thuộc về bạn!');
      sheet.getRange(i + 1, 9).setValue(true);
      return 'Success';
    }
  }
  throw new Error('Không tìm thấy thông báo!');
}

function readBinhLuanRows_(sheet) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, BINHLUAN_HEADERS_.length).getValues();
  const result = [];
  values.forEach(r => {
    const id = String(r[0] || '').trim();
    if (!id) return;
    if (!isActiveVal_(r[6])) return;
    result.push({
      id,
      maCongViec: String(r[1] || '').trim(),
      nguoiBinhLuanId: String(r[2] || '').trim(),
      nguoiBinhLuanTen: String(r[3] || '').trim(),
      noiDung: String(r[4] || '').trim(),
      thoiGian: formatDateTimeCell_(r[5]),
      nguoiDuocTag: String(r[7] || '').split(',').map(s => s.trim()).filter(Boolean)
    });
  });
  return result;
}

// Không cần đăng nhập để xem — nhất quán với getNhatKyGanNhat cũng đọc công khai.
function getBinhLuanCongViec(maCongViec) {
  const sheet = layHoacTaoSheet_('BinhLuanCongViec', BINHLUAN_HEADERS_);
  const list = readBinhLuanRows_(sheet).filter(c => c.maCongViec === String(maCongViec || '').trim());
  list.sort((a, b) => a.id.localeCompare(b.id));
  return JSON.stringify(list);
}

// Mọi nhân sự đã đăng nhập (không chỉ người phụ trách/Admin) đều bình luận được vào bất kỳ Công
// việc nào — để hỗ trợ phối hợp qua @tag; Sửa/Xóa Công việc vẫn chỉ dành cho người phụ trách/Admin
// (xem adminSaveCongViec/adminXoaCongViec), không đổi.
function themBinhLuanCongViec(userId, password, maCongViec, noiDung, taggedUserIds) {
  const user = login(userId, password);
  const noiDungTrim = String(noiDung || '').trim();
  if (!noiDungTrim) throw new Error('Vui lòng nhập nội dung bình luận!');

  const cvSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CongViec');
  if (!cvSheet) throw new Error('Chưa cấu hình Tab CongViec trên Google Sheet!');
  const cv = readCongViecRows_(cvSheet, true).find(c => c.maCongViec === String(maCongViec || '').trim());
  if (!cv) throw new Error('Không tìm thấy Công việc!');

  const allUsers = listSupervisors();
  const validIds = new Set(allUsers.map(u => u.id));
  const taggedIds = Array.from(new Set((taggedUserIds || []).map(t => String(t || '').trim()).filter(t => t && validIds.has(t) && t !== user.supervisorId)));

  const blSheet = layHoacTaoSheet_('BinhLuanCongViec', BINHLUAN_HEADERS_);
  const id = sinhMaTuTang_(blSheet, 'BL', new Date());
  blSheet.appendRow([id, cv.maCongViec, user.supervisorId, user.name, noiDungTrim, new Date(), true, taggedIds.join(',')]);
  logActivity(userId, user.name, 'Bình luận Công việc', `${cv.maCongViec}: ${noiDungTrim}`);

  // Người được @tag nhận thông báo riêng, rõ ràng hơn là "có người bình luận" chung chung.
  taggedIds.forEach(id2 => {
    guiThongBao_(id2, 'tag', `${user.name} đã nhắc bạn trong công việc "${cv.tenCongViec}": ${noiDungTrim}`, cv.maCongViec, user.supervisorId, user.name);
  });

  // Báo cho mọi người từng liên quan tới Công việc này (người phụ trách, người giao — nếu có, và
  // mọi người đã từng bình luận), trừ chính người vừa bình luận và trừ người đã nhận thông báo 'tag'
  // ở trên (tránh nhận 2 thông báo trùng ý cho cùng 1 bình luận).
  const nguoiLienQuan = new Set([cv.nguoiTaoId, cv.nguoiGiaoId].filter(Boolean));
  readBinhLuanRows_(blSheet).forEach(c => { if (c.maCongViec === cv.maCongViec) nguoiLienQuan.add(c.nguoiBinhLuanId); });
  nguoiLienQuan.delete(user.supervisorId);
  taggedIds.forEach(id2 => nguoiLienQuan.delete(id2));
  nguoiLienQuan.forEach(id2 => {
    guiThongBao_(id2, 'binh_luan', `${user.name} vừa bình luận vào "${cv.tenCongViec}"`, cv.maCongViec, user.supervisorId, user.name);
  });

  return 'Success';
}

// =======================================================
// NHẬT KÝ TIẾN ĐỘ — log cộng dồn % trên CÔNG VIỆC, append-only, validate + khoá chống ghi đè
// =======================================================

function sinhLogId_(logSheet, now) {
  const ngay = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd');
  const prefix = 'NK-' + ngay + '-';
  const lastRow = logSheet.getLastRow();
  let maxSTT = 0;
  if (lastRow >= 2) {
    const ids = logSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    ids.forEach(r => {
      const id = String(r[0] || '');
      if (id.indexOf(prefix) === 0) {
        const stt = parseInt(id.slice(prefix.length), 10);
        if (!isNaN(stt) && stt > maxSTT) maxSTT = stt;
      }
    });
  }
  return prefix + String(maxSTT + 1).padStart(3, '0');
}

// Validate mục 3 của spec + ghi 1 dòng NhatKyTienDo. PHẢI được gọi bên trong LockService.
// Số ngày nguyên giữa 2 chuỗi ngày 'yyyy-MM-dd' (ngaySau - ngayTruoc) — dùng tính "trễ bao nhiêu
// ngày" hiển thị ở cột Tình trạng (xem getData() bên dưới) và validate ngày báo cáo.
function soNgayGiuaHaiNgayISO_(ngaySau, ngayTruoc) {
  if (!ngaySau || !ngayTruoc) return 0;
  const d1 = new Date(ngaySau + 'T00:00:00Z');
  const d2 = new Date(ngayTruoc + 'T00:00:00Z');
  return Math.round((d1.getTime() - d2.getTime()) / 86400000);
}

// Khoá cứng ngày báo cáo = ĐÚNG hôm nay (không cho chọn lùi/tới ngày khác) và bắt buộc ghi chú lý
// do khi Công việc đã trễ hạn (today > ngayKetThucKH) — chặn tình trạng nhân sự khai lùi ngày báo
// cáo để "hô biến" một việc trễ thành đúng hạn. Validate ở ĐÂY (không chỉ ở client) để không thể
// lách qua bằng cách gọi thẳng API.
function validateVaGhiNhatKy_(user, maCongViec, ngayBaoCao, phanTram, ghiChu, fileDinhKemArr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cvSheet = ss.getSheetByName('CongViec');
  if (!cvSheet) throw new Error('Chưa cấu hình Tab CongViec trên Google Sheet!');
  const congViec = readCongViecRows_(cvSheet);
  const node = congViec.find(c => c.maCongViec === maCongViec);
  if (!node) throw new Error('Không tìm thấy Công việc!');

  const phanTramNum = parseFloat(phanTram);
  if (isNaN(phanTramNum)) throw new Error('% không hợp lệ!');

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (ngayBaoCao !== today) {
    throw new Error('Ngày báo cáo phải là ngày hôm nay — hệ thống khoá cứng, không cho chọn ngày khác để tránh khai lùi ngày báo cáo!');
  }
  const ghiChuTrim = String(ghiChu || '').trim();
  if (node.ngayKetThucKH && today > node.ngayKetThucKH && !ghiChuTrim) {
    throw new Error(`Công việc "${node.tenCongViec}" đã trễ hạn (${node.ngayKetThucKH}) — vui lòng ghi chú rõ lý do trễ trước khi gửi báo cáo!`);
  }

  const logSheet = ss.getSheetByName('NhatKyTienDo');
  if (!logSheet) throw new Error('Chưa cấu hình Tab NhatKyTienDo trên Google Sheet!');
  const logs = readNhatKyRows_(logSheet).filter(l => l.active && l.maCongViec === maCongViec);
  const luyKeHienTai = Math.round(logs.reduce((s, l) => s + l.phanTramNgay, 0) * 10) / 10;
  const tongMoi = Math.round((luyKeHienTai + phanTramNum) * 10) / 10;

  if (tongMoi > 100.001) {
    const conLai = Math.max(0, Math.round((100 - luyKeHienTai) * 10) / 10);
    throw new Error(`Công việc đã đạt ${luyKeHienTai}%, chỉ còn tối đa ${conLai}% để nhập`);
  }

  const now = new Date();
  const logId = sinhLogId_(logSheet, now);
  logSheet.appendRow([
    logId, maCongViec, ngayBaoCao || '', phanTramNum, ghiChuTrim,
    user.supervisorId, user.name,
    Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'),
    (fileDinhKemArr || []).join(';'), true
  ]);
  logActivity(user.supervisorId, user.name, 'Nhập nhật ký tiến độ', `${maCongViec}: +${phanTramNum}%`);
  return tongMoi;
}

function themNhatKyTienDo(userId, password, maCongViec, ngayBaoCao, phanTram, ghiChu, fileDinhKemArr) {
  const user = login(userId, password);
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    return validateVaGhiNhatKy_(user, maCongViec, ngayBaoCao, phanTram, ghiChu, fileDinhKemArr);
  } finally {
    lock.releaseLock();
  }
}

// payloadArray: [{ maCongViec, ngayBaoCao, phanTram, ghiChu, fileDinhKem }]
function themNhatKyTienDoHangLoat(userId, password, payloadArray) {
  const user = login(userId, password);
  const lock = LockService.getScriptLock();
  const results = [];
  try {
    lock.waitLock(30000);
    (payloadArray || []).forEach(p => {
      try {
        const luyKeMoi = validateVaGhiNhatKy_(user, p.maCongViec, p.ngayBaoCao, p.phanTram, p.ghiChu, p.fileDinhKem);
        results.push({ maCongViec: p.maCongViec, ok: true, luyKe: luyKeMoi });
      } catch (e) {
        results.push({ maCongViec: p.maCongViec, ok: false, error: e.message });
      }
    });
  } finally {
    lock.releaseLock();
  }
  return results;
}

// Lưu ảnh hiện trường vào 1 thư mục cố định trên Drive của tài khoản chạy script, trả về URL.
function uploadAnhHienTruong(userId, password, base64Data, tenFile) {
  login(userId, password);
  const FOLDER_NAME = 'NhatKyTienDo_AnhHienTruong';
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);

  const matches = String(base64Data).match(/^data:(.+);base64,(.*)$/);
  const contentType = matches ? matches[1] : 'image/jpeg';
  const rawBase64 = matches ? matches[2] : base64Data;
  const bytes = Utilities.base64Decode(rawBase64);
  const blob = Utilities.newBlob(bytes, contentType, tenFile || ('anh_' + new Date().getTime() + '.jpg'));

  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

// =======================================================
// SỔ PHÉP (baseline đầu tháng) — dùng chung bởi tab "Chấm công & Tăng ca" (xem phần CHẤM CÔNG &
// TĂNG CA bên dưới, nơi định nghĩa CHAMCONG_HEADERS_/tinhPhepChiTietTuChamCong_/assertQuyenChamCongTangCa_).
// Tab "Điểm danh & Phép" (Thứ 7/CN riêng, sổ DiemDanhCuoiTuan) đã bị GỘP vào tab Chấm công & Tăng ca
// để tránh 2 nơi chấm công trùng lặp — chotPhepDenThangHienTai_ dưới đây giờ tính "Cộng/Trừ phép
// tháng" từ ChamCongThang (đủ mọi ngày trong tháng) thay vì chỉ từ Thứ 7/CN như trước.
// =======================================================

function chuanHoaChucDanh_(s) {
  return String(s || '').trim().toLowerCase();
}

// Lấy sheet theo tên, tự tạo kèm header nếu chưa có — tránh phụ thuộc việc người dùng phải vào
// Apps Script editor chạy tay thietLapBanDauSheet() sau mỗi lần thêm sheet mới.
function layHoacTaoSheet_(tenSheet, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(tenSheet);
  if (!sheet) {
    sheet = ss.insertSheet(tenSheet);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sheet;
}

function readPhepThangRows_(sheet) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, PHEP_HEADERS_.length).getValues();
  const result = [];
  values.forEach(r => {
    const userId = String(r[0] || '').trim();
    const thang = formatThangCell_(r[1]);
    if (!userId || !thang) return;
    result.push({ userId, thang, phepDauThang: Number(r[2]) || 0, soNgayNghi: Number(r[3]) || 0 });
  });
  return result;
}

// 'yyyy-MM' + số tháng cộng thêm (có thể âm) -> 'yyyy-MM' mới, tự xử lý qua năm.
function thangCong_(thang, soThang) {
  const parts = String(thang).split('-');
  const d = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1 + soThang, 1);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM');
}

// Chốt ngầm mỗi khi có người mở tab: với mỗi nhân sự, tạo tuần tự các dòng PhepThang còn thiếu cho
// tới tháng hiện tại thật (không nhảy cóc nếu app bị bỏ quên nhiều tháng liền), theo công thức:
// TổngCộng(tháng cũ) = phepDauThang + Cộng/TrừPhépThángTừChấmCông - soNgayNghi (lịch sử cũ, nếu có);
// phepDauThang(tháng mới) = TổngCộng(tháng cũ) + 1 (1 phép baseline/tháng).
// tinhPhepChiTietTuChamCong_ định nghĩa ở phần CHẤM CÔNG & TĂNG CA bên dưới (cùng công thức hiển thị
// ở getChamCongThang, tránh lệch số giữa "dự kiến" trên giao diện và số THẬT được chốt ở đây).
// allUsers/ccByThang_ do getChamCongThang() đọc + nhóm sẵn truyền vào — KHÔNG tự đọc lại
// ChamCongThang ở đây nữa (trước đây hàm này đọc 1 lần, rồi getChamCongThang() đọc lại lần 2 NGUYÊN
// SHEET đó ngay sau khi hàm này chạy xong, lãng phí 1 lượt đọc trùng mỗi lần mở tab/đổi tháng).
// Trả về LUÔN mảng allPhep đã gồm cả các dòng vừa chốt thêm (nếu có) — trước đây hàm này chỉ ghi
// vào Sheet rồi thôi, buộc getChamCongThang() phải đọc lại NGUYÊN sheet PhepThang lần thứ 2 ngay sau
// đó để lấy đúng dữ liệu mới nhất, dù dữ liệu đó vốn đã có sẵn trong bộ nhớ ở đây rồi — mỗi lượt đọc
// sheet tốn ~0,4-1 giây CỐ ĐỊNH ở nền tảng Apps Script bất kể sheet nhỏ hay lớn (đã đo thực tế: 23
// nhân sự/317 dòng ChamCongThang vẫn mất 3,8 giây tổng cộng cho getChamCongThang, chủ yếu do CÁC LƯỢT
// ĐỌC SHEET RIÊNG LẺ, không phải do khối lượng dữ liệu) — bớt được 1 lượt đọc trùng là bớt được ~1
// giây mỗi lần mở tab/đổi tháng.
function chotPhepDenThangHienTai_(allUsers, ccByThang_) {
  const phepSheet = layHoacTaoSheet_('PhepThang', PHEP_HEADERS_);

  const thangHienTai = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  const allPhep = readPhepThangRows_(phepSheet);

  const rowsToAppend = [];
  allUsers.forEach(u => {
    const phepOfUser = allPhep.filter(p => p.userId === u.id).sort((a, b) => a.thang.localeCompare(b.thang));
    const lastRow = phepOfUser.length ? phepOfUser[phepOfUser.length - 1] : null;

    if (!lastRow) {
      // Chưa từng có dòng nào — chỉ bắt đầu từ tháng hiện tại, không truy hồi lùi các tháng trước.
      rowsToAppend.push([u.id, thangHienTai, 1, 0, new Date()]);
      allPhep.push({ userId: u.id, thang: thangHienTai, phepDauThang: 1, soNgayNghi: 0 });
      return;
    }

    let thangDangXet = lastRow.thang;
    let tongCongThangDo = lastRow.phepDauThang
      + tinhPhepChiTietTuChamCong_(ccByThang_.get(thangDangXet) || [], u.id, thangDangXet).congTruPhepThang
      - lastRow.soNgayNghi;

    while (thangDangXet < thangHienTai) {
      thangDangXet = thangCong_(thangDangXet, 1);
      const phepDauThangMoi = Math.round((tongCongThangDo + 1) * 10) / 10;
      rowsToAppend.push([u.id, thangDangXet, phepDauThangMoi, 0, new Date()]);
      allPhep.push({ userId: u.id, thang: thangDangXet, phepDauThang: phepDauThangMoi, soNgayNghi: 0 });
      tongCongThangDo = phepDauThangMoi + tinhPhepChiTietTuChamCong_(ccByThang_.get(thangDangXet) || [], u.id, thangDangXet).congTruPhepThang;
    }
  });

  if (rowsToAppend.length) {
    phepSheet.getRange(phepSheet.getLastRow() + 1, 1, rowsToAppend.length, PHEP_HEADERS_.length).setValues(rowsToAppend);
  }

  return allPhep;
}

// Chỉ Admin (không áp dụng cho Thư ký BQLDA) được sửa trực tiếp "Phép + bù" — dùng để nhập số dư
// phép sẵn có của dự án (đã theo dõi thủ công trên Google Sheet trước khi có tab này), hoặc chỉnh
// tay khi cần — sửa trực tiếp ở tab "Chấm công & Tăng ca" (ô "Phép đầu tháng").
function adminCapNhatPhepDauThang(userId, password, targetUserId, thang, phepDauThang) {
  const user = login(userId, password);
  if (user.role !== 'ADMIN') throw new Error('Chỉ Quản trị mới có quyền điều chỉnh trực tiếp "Phép + bù"!');

  const targetId = String(targetUserId || '').trim();
  const thangStr = String(thang || '').trim();
  if (!targetId || !thangStr) throw new Error('Thiếu thông tin nhân sự/tháng!');

  const phepDauThangNum = parseFloat(phepDauThang);
  if (isNaN(phepDauThangNum)) throw new Error('"Phép + bù" không hợp lệ!');

  const sheet = layHoacTaoSheet_('PhepThang', PHEP_HEADERS_);

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === targetId && formatThangCell_(data[i][1]) === thangStr) {
      sheet.getRange(i + 1, 3).setValue(phepDauThangNum);
      logActivity(userId, user.name, 'Sửa Phép + bù', `${targetId} - ${thangStr}: ${phepDauThangNum}`);
      return 'Success';
    }
  }

  sheet.appendRow([targetId, thangStr, phepDauThangNum, 0, new Date()]);
  logActivity(userId, user.name, 'Sửa Phép + bù', `${targetId} - ${thangStr}: ${phepDauThangNum}`);
  return 'Success';
}

// =======================================================
// CHẤM CÔNG & TĂNG CA THEO THÁNG — mỗi dòng = 1 (userId, ngày), gộp cả chấm công (buoi: 1 cả ngày /
// 0.5 nửa ngày / 'P' nghỉ phép) CHO MỌI NGÀY trong tháng và tăng ca (gioBatDauTC/gioKetThucTC dạng
// 'HH:mm', kết thúc <= bắt đầu ngầm hiểu là qua ngày hôm sau) trên CÙNG 1 dòng. Đây là tab DUY NHẤT
// chấm công (tab "Điểm danh & Phép"/DiemDanhCuoiTuan cũ, chỉ chấm được Thứ 7/CN, đã bị GỘP vào đây để
// khỏi trùng 2 nơi chấm công). Ai cũng xem được (giống getData()), chỉ Admin/Thư ký BQLDA
// (assertQuyenChamCongTangCa_) chấm công/sửa tăng ca được — riêng "Phép đầu tháng" chỉ Admin sửa
// được, qua adminCapNhatPhepDauThang() (định nghĩa ở phần SỔ PHÉP phía trên).
//
// Quy tắc quy đổi giờ tăng ca: tách theo mốc 22h trong ca đã nhập — phần trước 22h hệ số 1, phần từ
// 22h hệ số 2, cộng lại thành "giờ quy đổi" (xem tinhGioTangCa_).
//
// Quy tắc cộng/trừ Phép mỗi tháng (xem tinhPhepChiTietTuChamCong_, dùng chung cho cả hiển thị "dự
// kiến" ở getChamCongThang lẫn chốt số THẬT ở chotPhepDenThangHienTai_ khi sang tháng mới):
//   Cộng/Trừ phép tháng = (Công Thứ 7 + Công Chủ Nhật − Định mức Thứ 7) − Số ngày đánh dấu P
//   Định mức Thứ 7 = 0,5 công × số Thứ 7 trong tháng — nhưng Chủ Nhật được tính GỘP CHUNG 1 định mức
//   với Thứ 7 (đi làm Chủ Nhật thay Thứ 7 vẫn coi là đủ định mức, không bắt buộc đúng ngày Thứ 7),
//   phần vượt định mức cuối tuần (Thứ 7 + Chủ Nhật) mới cộng thêm vào Phép. Số ngày P CHỈ bị trừ khi
//   cả tháng đó có 0 công cuối tuần (không đi làm Thứ 7 lẫn Chủ Nhật buổi nào) — hễ có đi làm cuối
//   tuần (Thứ 7 hoặc Chủ Nhật, dù không dư định mức) thì mọi ngày P trong tháng đó được miễn, không
//   trừ phép (theo yêu cầu thực tế của BQLDA).
// =======================================================

const CHAMCONG_HEADERS_ = ['userId', 'ngay', 'buoi', 'gioBatDauTC', 'gioKetThucTC'];

function assertQuyenChamCongTangCa_(user) {
  if (user.role !== 'ADMIN' && chuanHoaChucDanh_(user.chucDanh) !== THU_KY_CHUC_DANH_) {
    throw new Error('Chỉ Quản trị hoặc Thư ký BQLDA mới có quyền cập nhật Chấm công/Tăng ca!');
  }
}

// { gioThuong, gioDem, quyDoi, quaNgay } — quaNgay=true nghĩa là gioKetThuc rơi vào ngày hôm sau.
function tinhGioTangCa_(gioBatDau, gioKetThuc) {
  if (!gioBatDau || !gioKetThuc) return { gioThuong: 0, gioDem: 0, quyDoi: 0, quaNgay: false };
  const toMin_ = s => {
    const p = String(s).split(':');
    return (parseInt(p[0], 10) || 0) * 60 + (parseInt(p[1], 10) || 0);
  };
  let s = toMin_(gioBatDau), e = toMin_(gioKetThuc);
  const quaNgay = e <= s;
  if (quaNgay) e += 1440;
  const boundary = 22 * 60;
  let gioThuong, gioDem;
  if (e <= boundary) { gioThuong = e - s; gioDem = 0; }
  else if (s >= boundary) { gioThuong = 0; gioDem = e - s; }
  else { gioThuong = boundary - s; gioDem = e - boundary; }
  gioThuong /= 60; gioDem /= 60;
  return { gioThuong, gioDem, quyDoi: gioThuong + gioDem * 2, quaNgay };
}

// { soNgay, saturdays: ['yyyy-MM-dd', ...], sundays: [...] } cho 1 tháng 'yyyy-MM'.
function layThu7CNTrongThang_(thang) {
  const parts = String(thang).split('-');
  const year = parseInt(parts[0], 10), month1 = parseInt(parts[1], 10);
  const soNgay = new Date(year, month1, 0).getDate();
  const saturdays = [], sundays = [];
  for (let day = 1; day <= soNgay; day++) {
    const d = new Date(year, month1 - 1, day);
    const iso = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (d.getDay() === 6) saturdays.push(iso);
    if (d.getDay() === 0) sundays.push(iso);
  }
  return { soNgay, saturdays, sundays };
}

function readChamCongRows_(sheet) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, CHAMCONG_HEADERS_.length).getValues();
  const result = [];
  values.forEach(r => {
    const userId = String(r[0] || '').trim();
    const ngay = formatDateCell_(r[1]);
    if (!userId || !ngay) return;
    const buoiRaw = r[2];
    let buoi = '';
    if (buoiRaw === 'P') buoi = 'P';
    else { const n = Number(buoiRaw); if (n === 1 || n === 0.5) buoi = n; }
    result.push({ userId, ngay, buoi, gioBatDauTC: formatTimeCell_(r[3]), gioKetThucTC: formatTimeCell_(r[4]) });
  });
  return result;
}

// Nhóm mảng đọc từ readChamCongRows_() theo 'yyyy-MM' — tra 1 tháng bất kỳ chỉ còn O(1) thay vì mỗi
// nơi cần dữ liệu 1 tháng phải .filter() lại NGUYÊN mảng toàn bộ lịch sử chấm công từ trước tới giờ
// (ngày càng dài theo thời gian dự án, vì ChamCongThang không bao giờ dọn/lưu trữ riêng theo tháng).
// Trước đây getChamCongThang() lặp allUsers.map() rồi bên trong mỗi vòng lặp lại .filter() nguyên
// allCC 1-2 lần (kể cả gọi qua chotPhepDenThangHienTai_) — với N nhân sự và M dòng lịch sử, tổng chi
// phí tỉ lệ O(N×M), càng chấm công nhiều tháng thì mở tab càng chậm dù chỉ xem đúng 1 tháng. Gom 1
// lần thành Map rồi tra theo tháng đưa chi phí về gần O(N + M).
function nhomChamCongTheoThang_(allCC) {
  const map = new Map();
  allCC.forEach(r => {
    const t = r.ngay.slice(0, 7);
    if (!map.has(t)) map.set(t, []);
    map.get(t).push(r);
  });
  return map;
}

// { quotaThu7, congThu7, congChuNhat, congCuoiTuan, soNgayP, phepTruDoP, congTruPhepThang } cho 1
// nhân sự/1 tháng — DÙNG CHUNG bởi getChamCongThang (hiển thị "dự kiến") và chotPhepDenThangHienTai_
// (chốt số THẬT khi sang tháng mới), để không lệch công thức giữa 2 nơi.
// congCuoiTuan = congThu7 + congChuNhat: Chủ Nhật được tính GỘP CHUNG 1 định mức với Thứ 7 (làm Chủ
// Nhật thay Thứ 7 vẫn đủ định mức bình thường, không bắt buộc phải đúng ngày Thứ 7) — phần vượt định
// mức cuối tuần mới cộng vào Phép; ngày P chỉ bị trừ khi CẢ THÁNG không đi làm cuối tuần buổi nào
// (Thứ 7 lẫn Chủ Nhật đều 0), không riêng gì Thứ 7 như trước.
function tinhPhepChiTietTuChamCong_(allCC, userId, thang) {
  const { saturdays, sundays } = layThu7CNTrongThang_(thang);
  const quotaThu7 = Math.round(saturdays.length * 0.5 * 10) / 10;

  const cc = {};
  allCC.filter(r => r.userId === userId && r.ngay.slice(0, 7) === thang && r.buoi !== '')
    .forEach(r => { cc[r.ngay] = r.buoi; });

  const soNum_ = v => (typeof v === 'number' ? v : 0);
  const congThu7 = Math.round(saturdays.reduce((s, d) => s + soNum_(cc[d]), 0) * 10) / 10;
  const congChuNhat = Math.round(sundays.reduce((s, d) => s + soNum_(cc[d]), 0) * 10) / 10;
  const congCuoiTuan = Math.round((congThu7 + congChuNhat) * 10) / 10;
  const soNgayP = Object.values(cc).filter(v => v === 'P').length;
  const phepTruDoP = congCuoiTuan > 0 ? 0 : soNgayP;
  const congTruPhepThang = Math.round((congCuoiTuan - quotaThu7 - phepTruDoP) * 10) / 10;
  // Công cuối tuần chỉ tính vào "công thực làm" TỐI ĐA bằng định mức (VD 4 Thứ 7 => tối đa 4 nửa =
  // 2 công); phần vượt định mức không cộng thêm vào công thực làm nữa mà đã được cộng sang Phép ở
  // congTruPhepThang phía trên rồi (tránh tính 2 lần).
  const congCuoiTuanThucLam = Math.min(congCuoiTuan, quotaThu7);

  return { quotaThu7, congThu7, congChuNhat, congCuoiTuan, congCuoiTuanThucLam, soNgayP, phepTruDoP, congTruPhepThang };
}

// thang: 'yyyy-MM', bỏ trống = tháng hiện tại. KHÔNG cần đăng nhập để xem (giống getData()).
//
// Đã đo thực tế (23 nhân sự, 317 dòng ChamCongThang, chỉ 1 tháng dữ liệu — rất nhỏ) mà vẫn mất ~3,8
// giây: phần lớn là do CÁC LƯỢT ĐỌC SHEET RIÊNG LẺ (mỗi lượt ~0,4-1 giây CỐ ĐỊNH ở nền tảng Apps
// Script, không phụ thuộc sheet nhỏ hay lớn), không phải do vòng lặp tính toán. Vì vậy tối ưu quan
// trọng nhất là GIẢM SỐ LƯỢT ĐỌC SHEET, không phải tối ưu thuật toán — hàm này chỉ còn đọc đúng 3
// sheet 1 lần mỗi sheet: DanhSachUser (qua listSupervisors), ChamCongThang, PhepThang (qua
// chotPhepDenThangHienTai_, đã trả thẳng dữ liệu mới nhất, KHÔNG đọc lại PhepThang lần 2 như trước).
function getChamCongThang(thang) {
  const thangXem = String(thang || '').trim() || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');

  const { soNgay, saturdays, sundays } = layThu7CNTrongThang_(thangXem);

  const allUsers = listSupervisors();
  // Đọc + nhóm theo tháng ĐÚNG 1 LẦN cho cả lượt gọi này — chotPhepDenThangHienTai_() và vòng lặp
  // allUsers.map() bên dưới đều tra theo tháng qua Map (O(1)/tháng) thay vì mỗi nơi tự .filter() lại
  // NGUYÊN mảng toàn bộ lịch sử chấm công (xem chú thích ở nhomChamCongTheoThang_()) — có lợi khi
  // lịch sử chấm công tích lũy nhiều tháng, dù đây không phải chi phí chính hiện tại (xem trên).
  const allCC = readChamCongRows_(layHoacTaoSheet_('ChamCongThang', CHAMCONG_HEADERS_));
  const ccByThang_ = nhomChamCongTheoThang_(allCC);
  // Trả thẳng allPhep đã gồm dòng vừa chốt thêm (nếu có) — KHÔNG đọc lại PhepThang lần 2.
  const allPhep = chotPhepDenThangHienTai_(allUsers, ccByThang_);

  const ccThang = ccByThang_.get(thangXem) || [];

  const rows = allUsers.map(u => {
    const phepRow = allPhep.find(p => p.userId === u.id && p.thang === thangXem);
    const phepDauThang = phepRow ? phepRow.phepDauThang : 0;

    const cc = {}, tc = {};
    ccThang.filter(r => r.userId === u.id).forEach(r => {
      if (r.buoi !== '') cc[r.ngay] = r.buoi;
      if (r.gioBatDauTC && r.gioKetThucTC) tc[r.ngay] = { start: r.gioBatDauTC, end: r.gioKetThucTC };
    });

    const chiTiet = tinhPhepChiTietTuChamCong_(ccThang, u.id, thangXem);
    // Công thực làm = công các ngày trong tuần + công cuối tuần đã bị chặn trần theo định mức (phần
    // cuối tuần vượt định mức được cộng sang Phép thay vì tính vào công thực làm, xem congCuoiTuanThucLam).
    const congNgayThuong = Object.entries(cc).reduce((s, [ngay, v]) => {
      if (saturdays.includes(ngay) || sundays.includes(ngay)) return s;
      return s + (typeof v === 'number' ? v : 0);
    }, 0);
    const congThucLam = congNgayThuong + chiTiet.congCuoiTuanThucLam;
    const tongGioTC = Object.values(tc).reduce((s, e) => s + tinhGioTangCa_(e.start, e.end).quyDoi, 0);

    return {
      userId: u.id, hoTen: u.name, phepDauThang,
      cc, tc,
      congThu7: chiTiet.congThu7,
      congChuNhat: chiTiet.congChuNhat,
      congCuoiTuan: chiTiet.congCuoiTuan,
      soNgayP: chiTiet.soNgayP,
      phepTruDoP: chiTiet.phepTruDoP,
      congTruPhepThang: chiTiet.congTruPhepThang,
      phepCuoiThang: Math.round((phepDauThang + chiTiet.congTruPhepThang) * 10) / 10,
      congThucLam: Math.round(congThucLam * 10) / 10,
      tongGioTC: Math.round(tongGioTC * 10) / 10,
      ngayCongTC: Math.round((tongGioTC / 8) * 100) / 100
    };
  });

  const quotaThu7 = Math.round(saturdays.length * 0.5 * 10) / 10;
  return JSON.stringify({ thang: thangXem, soNgay, saturdays, sundays, quotaThu7, rows });
}

// Gộp TOÀN BỘ thay đổi chấm công + tăng ca của 1 lượt sửa (nhiều ô, nhiều nhân sự, nhiều ngày) của
// tab "Chấm công & Tăng ca" thành ĐÚNG 1 lượt gọi từ client (nút "Lưu thay đổi" ở 15_ChamCong.html)
// thay vì gọi lặp lại từng ô bấm — trước đây mỗi lần bấm 1 ô là 1 round-trip + 1 lần load lại toàn
// bộ bảng, rất giật khi cần sửa nhiều ô liên tiếp.
// danhSachChamCong: mảng { targetUserId, ngay, buoi: '1'|'0.5'|'P'|'' }
// danhSachTangCa:   mảng { targetUserId, ngay, gioBatDau, gioKetThuc } — cả 2 rỗng nghĩa là xoá tăng ca
// Mỗi phần tử là TRẠNG THÁI CUỐI CÙNG mong muốn của đúng 1 ô (không phải log thao tác), nên áp dụng
// theo bất kỳ thứ tự nào đều ra cùng kết quả.
function capNhatChamCongHangLoat(userId, password, danhSachChamCong, danhSachTangCa) {
  const user = login(userId, password);
  assertQuyenChamCongTangCa_(user);

  const sheet = layHoacTaoSheet_('ChamCongThang', CHAMCONG_HEADERS_);
  const lastRow = sheet.getLastRow();
  const existing = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, CHAMCONG_HEADERS_.length).getValues() : [];

  // key "userId|ngay" -> trạng thái hiện tại (rowIdx=null nghĩa là chưa có dòng trong Sheet)
  const state = new Map();
  existing.forEach((r, i) => {
    const uid = String(r[0] || '').trim(), ngay = formatDateCell_(r[1]);
    if (!uid || !ngay) return;
    state.set(uid + '|' + ngay, { rowIdx: i + 2, userId: uid, ngay, buoi: r[2], gioBD: formatTimeCell_(r[3]), gioKT: formatTimeCell_(r[4]) });
  });
  function layHoacTao_(targetId, ngay) {
    const key = targetId + '|' + ngay;
    if (!state.has(key)) state.set(key, { rowIdx: null, userId: targetId, ngay, buoi: '', gioBD: '', gioKT: '' });
    return state.get(key);
  }

  (danhSachChamCong || []).forEach(item => {
    const targetId = String(item.targetUserId || '').trim();
    const ngayStr = String(item.ngay || '').trim();
    if (!targetId || !/^\d{4}-\d{2}-\d{2}$/.test(ngayStr)) return;
    const s = layHoacTao_(targetId, ngayStr);
    const buoiStr = String(item.buoi || '').trim();
    s.buoi = buoiStr === 'P' ? 'P' : (buoiStr === '1' ? 1 : (buoiStr === '0.5' ? 0.5 : ''));
  });

  (danhSachTangCa || []).forEach(item => {
    const targetId = String(item.targetUserId || '').trim();
    const ngayStr = String(item.ngay || '').trim();
    if (!targetId || !/^\d{4}-\d{2}-\d{2}$/.test(ngayStr)) return;
    const s = layHoacTao_(targetId, ngayStr);
    s.gioBD = String(item.gioBatDau || '').trim();
    s.gioKT = String(item.gioKetThuc || '').trim();
  });

  const rowsToUpdate = [], rowsToDelete = [], rowsToAppend = [];
  state.forEach(s => {
    const rong = s.buoi === '' && !s.gioBD && !s.gioKT;
    if (s.rowIdx) {
      if (rong) rowsToDelete.push(s.rowIdx);
      else rowsToUpdate.push({ rowIdx: s.rowIdx, values: [s.userId, s.ngay, s.buoi, s.gioBD, s.gioKT] });
    } else if (!rong) {
      rowsToAppend.push([s.userId, s.ngay, s.buoi, s.gioBD, s.gioKT]);
    }
  });

  rowsToUpdate.forEach(u => sheet.getRange(u.rowIdx, 1, 1, CHAMCONG_HEADERS_.length).setValues([u.values]));
  rowsToDelete.sort((a, b) => b - a).forEach(r => sheet.deleteRow(r)); // xoá từ dưới lên để khỏi lệch rowIdx
  if (rowsToAppend.length) sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, CHAMCONG_HEADERS_.length).setValues(rowsToAppend);

  logActivity(userId, user.name, 'Lưu Chấm công/Tăng ca hàng loạt',
    `${rowsToUpdate.length} dòng sửa, ${rowsToAppend.length} dòng thêm, ${rowsToDelete.length} dòng xoá`);
  return 'Success';
}

// Tiện ích chạy TAY 1 LẦN từ trình soạn thảo Apps Script (không gọi từ giao diện web) để nạp số liệu
// "Phép + bù Tháng 6" (đầu vào Phép tháng 7/2026) từ bảng Excel thủ công cũ của Thư ký BQLDA sang
// PhepThang — khớp theo email = userId (tài khoản trong DanhSachUser đang dùng email làm ID đăng
// nhập). Chạy lại nhiều lần vẫn an toàn (ghi đè đúng dòng đã có, không tạo trùng).
function adminSeedPhepThang7_2026() {
  const THANG = '2026-07';
  const DU_LIEU = [
    ['tranthemanh77@gmail.com', 5.5], ['hoang31xd@gmail.com', 0], ['trongpvin@gmail.com', 5.5],
    ['hohongvan93@gmail.com', 5], ['lequanghong1984@gmail.com', 0], ['longsaovietpy@gmail.com', 10],
    ['trongtoanxd2012@gmail.com', 7], ['buiquoclamtd@gmail.com', 7], ['tronghieushdk@gmail.com', 7],
    ['dongphuong3112@gmail.com', 0], ['congtay1610@gmail.com', 2], ['vlnam0481@gmail.com', 8.5],
    ['vantien.7799@gmail.com', 3], ['ktshongvu@gmail.com', 2], ['congchinh1603@gmail.com', 0],
    ['duyhieu0983897677@gmail.com', 1], ['thoaicosevcopy@gmail.com', 0], ['ntht.hoaithuong.287@gmail.com', 0],
    ['trungdang.const@gmail.com', 0], ['bienthiaivan@gmail.com', 0]
  ];

  const allUsers = listSupervisors();
  const sheet = layHoacTaoSheet_('PhepThang', PHEP_HEADERS_);
  const data = sheet.getDataRange().getValues();
  const boQua = [];

  DU_LIEU.forEach(([email, phepDauThang]) => {
    const u = allUsers.find(x => x.id.toLowerCase() === email.toLowerCase());
    if (!u) { boQua.push(email); return; }

    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === u.id && formatThangCell_(data[i][1]) === THANG) {
        sheet.getRange(i + 1, 3).setValue(phepDauThang);
        found = true;
        break;
      }
    }
    if (!found) sheet.appendRow([u.id, THANG, phepDauThang, 0, new Date()]);
  });

  const ketQua = `Đã seed Phép đầu tháng 7/2026 cho ${DU_LIEU.length - boQua.length}/${DU_LIEU.length} người. Bỏ qua (không khớp email nào trong DanhSachUser): ${boQua.join(', ') || '(không có)'}`;
  Logger.log(ketQua);
  return ketQua;
}

// =======================================================
// HỆ THỐNG MÃ HÓA & XÁC THỰC USER ĐA TÀI KHOẢN TỪ SHEET
// =======================================================

function hashPassword(text) {
  // KHÓA BÍ MẬT CỦA CÔNG TY (SALT) - đổi thành bất cứ gì bạn thích, chỉ cần không lộ ra ngoài.
  const SECRET_SALT = "BanQLDA_BaiXep_BaoMat_2026!@#";

  const saltedText = String(text).trim() + SECRET_SALT;

  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, saltedText, Utilities.Charset.UTF_8);
  let txtHash = '';
  for (let i = 0; i < rawHash.length; i++) {
    let hashVal = rawHash[i];
    if (hashVal < 0) hashVal += 256;
    if (hashVal.toString(16).length == 1) txtHash += '0';
    txtHash += hashVal.toString(16);
  }
  return txtHash;
}

function login(userId, password) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DanhSachUser');
  if (!sheet) throw new Error("Chưa cấu hình Tab DanhSachUser trên Google Sheet!");

  const data = sheet.getDataRange().getValues();
  const hashedInput = hashPassword(password);
  const plainInput = String(password).trim();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(userId).trim()) {
      const storedPass = String(data[i][1]).trim();

      if (storedPass === hashedInput || storedPass === plainInput) {
        // TỰ ĐỘNG BẢO MẬT: nếu pass trong Sheet vẫn là pass thường, lập tức mã hóa và lưu đè lại
        if (storedPass !== hashedInput) {
          sheet.getRange(i + 1, 2).setValue(hashedInput);
        }

        return {
          role: String(data[i][2]).trim().toUpperCase(),
          name: String(data[i][3]).trim(),
          supervisorId: String(data[i][0]).trim(),
          chucDanh: String(data[i][4] || '').trim(),
          phanMacDinh: String(data[i][5] || '').trim()
        };
      } else {
        throw new Error("Sai Tên đăng nhập (ID) hoặc Mật khẩu!");
      }
    }
  }
  throw new Error("Tài khoản không tồn tại trong hệ thống!");
}

function changeUserPassword(userId, oldPass, newPass) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('DanhSachUser');
  const data = sheet.getDataRange().getValues();

  const hashedOld = hashPassword(oldPass);
  const hashedNew = hashPassword(newPass);
  const plainOld = String(oldPass).trim();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(userId).trim()) {
      const storedPass = String(data[i][1]).trim();

      if (storedPass === hashedOld || storedPass === plainOld) {
        sheet.getRange(i + 1, 2).setValue(hashedNew);
        logActivity(userId, data[i][3], "Đổi mật khẩu", "Người dùng tự đổi mật khẩu cá nhân");
        return "Success";
      } else {
        throw new Error("Mật khẩu cũ không chính xác!");
      }
    }
  }
  throw new Error("Không tìm thấy tài khoản trong hệ thống!");
}

// =======================================================
// HÀM GHI NHẬT KÝ LƯU VẾT LỊCH SỬ (AUDIT TRAIL)
// =======================================================
function logActivity(userId, userName, action, details) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('NhatKy');
  if (!sheet) return;

  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
  sheet.appendRow([now, `${userName} (${userId})`, action, details]);
}

// =======================================================
// HỆ THỐNG BACKUP DỮ LIỆU
// =======================================================
function backupCurrentData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName('Data');
  if (!dataSheet) return;

  const lastRow = dataSheet.getLastRow();
  if (lastRow < 2) return;

  const now = new Date();
  const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
  const backupName = `Backup_${timestamp}`;

  const backupSheet = dataSheet.copyTo(ss);
  backupSheet.setName(backupName);
  backupSheet.hideSheet();

  cleanupOldBackups(ss, 15);
}

function cleanupOldBackups(ss, keepCount) {
  const backupSheets = ss.getSheets().filter(s => s.getName().indexOf('Backup_') === 0);
  backupSheets.sort((a, b) => a.getName().localeCompare(b.getName()));
  const excessCount = backupSheets.length - keepCount;
  for (let i = 0; i < excessCount; i++) {
    ss.deleteSheet(backupSheets[i]);
  }
}

// =======================================================
// HÀM LƯU DỮ LIỆU KIỂU CŨ (blob JSON trong tab Data) — KHÔNG còn dùng cho luồng
// Hạng mục/Gói thầu/Công việc/Nhật ký tiến độ (nay ghi trực tiếp qua adminSaveHangMuc/
// adminSaveCongViec/themNhatKyTienDo...), giữ lại phòng khi cần thao tác dữ liệu Data cũ.
// =======================================================
function saveData(jsonString, userId, password, actionDetail) {
  let user;
  try {
    user = login(userId, password);
  } catch (e) {
    throw new Error("Xác thực thất bại! Phiên đăng nhập đã hết hạn hoặc sai mật khẩu.");
  }

  backupCurrentData();

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
  sheet.clearContents();
  sheet.getRange('A1').setValue('Dữ liệu hệ thống');

  const CHUNK_SIZE = 45000;
  let chunks = [];
  for (let i = 0; i < jsonString.length; i += CHUNK_SIZE) {
    chunks.push([jsonString.substring(i, i + CHUNK_SIZE)]);
  }
  if (chunks.length > 0) {
    sheet.getRange(2, 1, chunks.length, 1).setValues(chunks);
  }

  logActivity(userId, user.name, "Lưu hệ thống", actionDetail || "Cập nhật dữ liệu");

  return "Success";
}

// =======================================================
// QUẢN LÝ HỢP ĐỒNG — Hợp đồng / BOQ / Phụ lục / Nghiệm thu / Thanh toán / Quyết toán.
// Chỉ ADMIN hoặc Giám sát có Chức danh "QS" được sửa (login() + kiemTraQuyenHopDong_) — người khác
// vẫn xem được qua getHopDongData() (không cần đăng nhập, giống getData()) nhưng số liệu tiền tệ bị
// ẩn ở giao diện (xem canXemGiaTriHopDong() ở 9_HopDong.html), chỉ hiện khối lượng.
// Toàn bộ số liệu tổng hợp tính lại khi đọc, không lưu cột trùng.
// =======================================================

function kiemTraQuyenHopDong_(userId, password) {
  const user = login(userId, password);
  if (user.role !== 'ADMIN' && chuanHoaChucDanh_(user.chucDanh) !== 'qs') {
    throw new Error('Chỉ Quản trị hoặc Giám sát QS mới có quyền thao tác dữ liệu Hợp đồng!');
  }
  return user;
}

// Quy ước làm tròn dùng thống nhất toàn module Hợp đồng: mọi KHỐI LƯỢNG làm tròn 3 chữ số thập
// phân (tránh sai số cộng dồn khi cộng nhiều lượt Nghiệm thu/điều chỉnh Phụ lục), mọi ĐƠN GIÁ/THÀNH
// TIỀN/GIÁ TRỊ tiền tệ làm tròn về số nguyên (đồng) — áp dụng cả khi ĐỌC (chuẩn hóa dữ liệu cũ/nhập
// Excel) lẫn khi GHI (chặn ngay từ đầu, không đợi tính lại mới đúng).
function round3_(n) { return Math.round((Number(n) || 0) * 1000) / 1000; }
function round0_(n) { return Math.round(Number(n) || 0); }

function readHopDongRows_(sheet) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, HOPDONG_HEADERS_.length).getValues();
  const result = [];
  values.forEach(r => {
    const maHopDong = String(r[0] || '').trim();
    if (!maHopDong) return;
    if (!isActiveVal_(r[9])) return;
    result.push({
      maHopDong, tenHopDong: String(r[1] || '').trim(), maGoiThau: String(r[2] || '').trim(),
      nhaThau: String(r[3] || '').trim(), ngayKy: formatDateCell_(r[4]),
      tamUngHopDong: round0_(r[5]), tamUngThiCong: round0_(r[6]),
      thueVAT: Number(r[7]) || 0, giamGia: round0_(r[8])
    });
  });
  return result;
}

function readBOQRows_(sheet) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, BOQ_HEADERS_.length).getValues();
  const result = [];
  values.forEach(r => {
    const maBOQ = String(r[0] || '').trim();
    if (!maBOQ) return;
    if (!isActiveVal_(r[12])) return;
    result.push({
      maBOQ, maHopDong: String(r[1] || '').trim(), stt: String(r[2] || '').trim(),
      tenHangMuc: String(r[3] || '').trim(),
      isHeader: r[4] === true || String(r[4]).trim().toUpperCase() === 'TRUE',
      donVi: String(r[5] || '').trim(), khoiLuongHopDong: round3_(r[6]),
      donGia: round0_(r[7]), cap: parseInt(r[8], 10) || 0,
      ngayBatDauKH: formatDateCell_(r[9]), ngayKetThucKH: formatDateCell_(r[10]),
      maPhuLucTao: String(r[11] || '').trim(), ghiChu: String(r[13] || '').trim()
    });
  });
  return result;
}

// includeInactive: true để lấy cả Phụ lục đã Vô hiệu hóa (active=false) — KHÔNG hard-delete Phụ
// lục bao giờ, chỉ tạm ngừng áp dụng (xem adminToggleTrangThaiPhuLuc), vẫn hiển thị để tra cứu.
function readPhuLucRows_(sheet, includeInactive) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, PHULUC_HEADERS_.length).getValues();
  const result = [];
  values.forEach(r => {
    const maPhuLuc = String(r[0] || '').trim();
    if (!maPhuLuc) return;
    const active = isActiveVal_(r[5]);
    if (!active && !includeInactive) return;
    result.push({
      maPhuLuc, maHopDong: String(r[1] || '').trim(), soHieu: String(r[2] || '').trim(),
      ngayPhuLuc: formatDateCell_(r[3]), ghiChu: String(r[4] || '').trim(), active
    });
  });
  return result;
}

function readPhuLucThayDoiRows_(sheet) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, PHULUC_THAYDOI_HEADERS_.length).getValues();
  const result = [];
  values.forEach(r => {
    const id = String(r[0] || '').trim();
    if (!id) return;
    if (!isActiveVal_(r[7])) return;
    result.push({
      id, maPhuLuc: String(r[1] || '').trim(), maBOQ: String(r[2] || '').trim(),
      loaiThayDoi: String(r[3] || '').trim(),
      khoiLuongMoi: (r[4] === '' || r[4] === null) ? null : round3_(r[4]),
      khoiLuongDieuChinh: (r[5] === '' || r[5] === null) ? null : round3_(r[5]),
      donGiaMoi: (r[6] === '' || r[6] === null) ? null : round0_(r[6])
    });
  });
  return result;
}

function readNghiemThuRows_(sheet) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, NGHIEMTHU_HEADERS_.length).getValues();
  const result = [];
  values.forEach(r => {
    const maNghiemThu = String(r[0] || '').trim();
    if (!maNghiemThu) return;
    if (!isActiveVal_(r[8])) return;
    result.push({
      maNghiemThu, maBOQ: String(r[1] || '').trim(), ngayNghiemThu: formatDateCell_(r[2]),
      khoiLuong: round3_(r[3]), ghiChu: String(r[4] || '').trim(),
      nguoiNhapId: String(r[5] || '').trim(), nguoiNhapTen: String(r[6] || '').trim(),
      thoiGianNhap: formatDateTimeCell_(r[7])
    });
  });
  return result;
}

function readDotThanhToanRows_(sheet) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, DOTTHANHTOAN_HEADERS_.length).getValues();
  const result = [];
  values.forEach(r => {
    const maDotThanhToan = String(r[0] || '').trim();
    if (!maDotThanhToan) return;
    if (!isActiveVal_(r[10])) return;
    result.push({
      maDotThanhToan, maHopDong: String(r[1] || '').trim(), tenDot: String(r[2] || '').trim(),
      ngayThanhToan: formatDateCell_(r[3]), phanTramThanhToan: Number(r[4]) || 0,
      ghiChu: String(r[5] || '').trim(), thuHoiTamUngHopDong: round0_(r[6]),
      thuHoiTamUngThiCong: round0_(r[7]), khauTruKhac: round0_(r[8]),
      ghiChuKhauTru: String(r[9] || '').trim(), giamGiaDot: round0_(r[11])
    });
  });
  return result;
}

function readDotThanhToanChiTietRows_(sheet) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, DOTTHANHTOAN_CHITIET_HEADERS_.length).getValues();
  const result = [];
  values.forEach(r => {
    const id = String(r[0] || '').trim();
    if (!id) return;
    if (!isActiveVal_(r[4])) return;
    result.push({
      id, maDotThanhToan: String(r[1] || '').trim(), maBOQ: String(r[2] || '').trim(),
      khoiLuong: round3_(r[3])
    });
  });
  return result;
}

function readQuyetToanRows_(sheet) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, QUYETTOAN_HEADERS_.length).getValues();
  const result = [];
  values.forEach(r => {
    const maQuyetToan = String(r[0] || '').trim();
    if (!maQuyetToan) return;
    if (!isActiveVal_(r[6])) return;
    result.push({
      maQuyetToan, maHopDong: String(r[1] || '').trim(), ngayQuyetToan: formatDateCell_(r[2]),
      nguoiXacNhanId: String(r[3] || '').trim(), nguoiXacNhanTen: String(r[4] || '').trim(),
      ghiChu: String(r[5] || '').trim()
    });
  });
  return result;
}

// Khối lượng hiệu lực từng hạng mục BOQ sau khi áp toàn bộ Phụ lục, theo đúng thứ tự tạo (id tăng
// dần = thứ tự thời gian, vì sinhMaTuTang_ sinh mã PREFIX-yyyyMMdd-NNN sắp được bằng localeCompare):
// THAY_THE ghi đè hẳn, DIEU_CHINH cộng/trừ chênh lệch, tính từ khối lượng hợp đồng gốc.
function tinhKhoiLuongHieuLuc_(boqItems, phuLucThayDoiList) {
  const map = {};
  boqItems.forEach(item => { map[item.maBOQ] = item.khoiLuongHopDong; });
  phuLucThayDoiList.slice().sort((a, b) => a.id.localeCompare(b.id)).forEach(ch => {
    if (!(ch.maBOQ in map)) return;
    if (ch.loaiThayDoi === 'THAY_THE' && ch.khoiLuongMoi !== null) {
      map[ch.maBOQ] = ch.khoiLuongMoi;
    } else if (ch.loaiThayDoi === 'DIEU_CHINH' && ch.khoiLuongDieuChinh !== null) {
      map[ch.maBOQ] = Math.round((map[ch.maBOQ] + ch.khoiLuongDieuChinh) * 1000) / 1000;
    }
  });
  return map;
}

// Đơn giá hiệu lực — Phụ lục có thể đổi giá (donGiaMoi), lấy giá mới nhất theo thứ tự tạo nếu có.
function tinhDonGiaHieuLuc_(boqItems, phuLucThayDoiList) {
  const map = {};
  boqItems.forEach(item => { map[item.maBOQ] = item.donGia; });
  phuLucThayDoiList.slice().sort((a, b) => a.id.localeCompare(b.id)).forEach(ch => {
    if (ch.donGiaMoi !== null && ch.maBOQ in map) map[ch.maBOQ] = ch.donGiaMoi;
  });
  return map;
}

// Số liệu Quyết toán 1 Hợp đồng — tự tính từ BOQ (đã gắn khoiLuongHieuLuc/donGiaHieuLuc), Đợt thanh
// toán + chi tiết. preloaded = { hopDong, boq, dotThanhToan, dotThanhToanChiTiet } (boq đã enrich).
function tinhQuyetToanHopDong_(maHopDong, preloaded) {
  const boqOfHD = preloaded.boq.filter(b => b.maHopDong === maHopDong && !b.isHeader);
  const hopDong = preloaded.hopDong.find(h => h.maHopDong === maHopDong);
  const giaTriSauPhuLuc = Math.round(boqOfHD.reduce((s, b) => s + b.khoiLuongHieuLuc * b.donGiaHieuLuc, 0));
  const giaTriTruocThue = Math.max(0, giaTriSauPhuLuc - ((hopDong && hopDong.giamGia) || 0));
  const vatRate = (hopDong && hopDong.thueVAT) || 0;
  const thueGTGT = Math.round(giaTriTruocThue * vatRate / 100);
  const giaTriSauThue = giaTriTruocThue + thueGTGT;

  const dotsOfHD = preloaded.dotThanhToan.filter(d => d.maHopDong === maHopDong);
  let tongDaThanhToan = 0;
  dotsOfHD.forEach(dot => {
    const chiTiet = preloaded.dotThanhToanChiTiet.filter(c => c.maDotThanhToan === dot.maDotThanhToan);
    let thucHien = 0;
    chiTiet.forEach(ct => {
      const boq = boqOfHD.find(b => b.maBOQ === ct.maBOQ);
      if (boq) thucHien += ct.khoiLuong * boq.donGiaHieuLuc;
    });
    const tamTinh = thucHien * (dot.phanTramThanhToan || 100) / 100;
    const thue = tamTinh * vatRate / 100;
    tongDaThanhToan += tamTinh + thue - dot.thuHoiTamUngHopDong - dot.thuHoiTamUngThiCong - dot.khauTruKhac - (dot.giamGiaDot || 0);
  });
  tongDaThanhToan = Math.round(tongDaThanhToan);

  const tongGiaTriHieuLuc = boqOfHD.reduce((s, b) => s + b.khoiLuongHieuLuc * b.donGiaHieuLuc, 0);
  const tongGiaTriDaNghiemThu = boqOfHD.reduce((s, b) => s + b.khoiLuongDaNghiemThu * b.donGiaHieuLuc, 0);
  const tyLeHoanThanhKhoiLuong = tongGiaTriHieuLuc > 0 ? Math.round((tongGiaTriDaNghiemThu / tongGiaTriHieuLuc) * 1000) / 10 : 0;

  return {
    giaTriSauPhuLuc, giaTriTruocThue, thueGTGT, giaTriSauThue,
    tongDaThanhToan, conPhaiThanhToan: giaTriSauThue - tongDaThanhToan, tyLeHoanThanhKhoiLuong
  };
}

// Trả về toàn bộ dữ liệu module Hợp đồng trong 1 round-trip (giống getData() cho Công việc) —
// KHÔNG cần đăng nhập để xem, nhất quán với getData()/getNhatKyGanNhat.
function getHopDongData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hopDong = readHopDongRows_(ss.getSheetByName('HopDong'));
  const boq = readBOQRows_(ss.getSheetByName('BOQHangMuc'));
  // Lấy cả Phụ lục đã Vô hiệu hóa để hiển thị (giữ lại tra cứu), nhưng CHỈ áp dụng thay đổi khối
  // lượng/đơn giá của Phụ lục CÒN HIỆU LỰC vào phần tính toán bên dưới — đúng rule app cũ
  // (recalcEffectiveQty/recalcCurrentPrices chỉ chạy qua addendums.filter(active !== false)).
  const phuLuc = readPhuLucRows_(ss.getSheetByName('PhuLucHopDong'), true);
  const phuLucThayDoi = readPhuLucThayDoiRows_(ss.getSheetByName('PhuLucThayDoi'));
  const nghiemThu = readNghiemThuRows_(ss.getSheetByName('NghiemThu'));
  const dotThanhToan = readDotThanhToanRows_(ss.getSheetByName('DotThanhToan'));
  const dotThanhToanChiTiet = readDotThanhToanChiTietRows_(ss.getSheetByName('DotThanhToanChiTiet'));
  const quyetToan = readQuyetToanRows_(ss.getSheetByName('QuyetToan'));

  const activePhuLucIds_ = new Set(phuLuc.filter(p => p.active).map(p => p.maPhuLuc));
  const phuLucThayDoiHieuLuc_ = phuLucThayDoi.filter(ch => activePhuLucIds_.has(ch.maPhuLuc));
  const effectiveQtyMap = tinhKhoiLuongHieuLuc_(boq, phuLucThayDoiHieuLuc_);
  const donGiaHieuLucMap = tinhDonGiaHieuLuc_(boq, phuLucThayDoiHieuLuc_);
  const executedByBOQ = {};
  nghiemThu.forEach(n => { executedByBOQ[n.maBOQ] = Math.round(((executedByBOQ[n.maBOQ] || 0) + n.khoiLuong) * 1000) / 1000; });
  // Khối lượng đã đưa vào thanh toán, lũy kế qua MỌI đợt (bảng lũy kế tab Thanh toán) — không cần
  // join qua dotThanhToan.maHopDong vì maBOQ vốn đã gắn cố định với đúng 1 hợp đồng.
  const paidByBOQ = {};
  dotThanhToanChiTiet.forEach(c => { paidByBOQ[c.maBOQ] = Math.round(((paidByBOQ[c.maBOQ] || 0) + c.khoiLuong) * 1000) / 1000; });

  boq.forEach(item => {
    item.khoiLuongHieuLuc = effectiveQtyMap[item.maBOQ] != null ? effectiveQtyMap[item.maBOQ] : item.khoiLuongHopDong;
    item.donGiaHieuLuc = donGiaHieuLucMap[item.maBOQ] != null ? donGiaHieuLucMap[item.maBOQ] : item.donGia;
    item.khoiLuongDaNghiemThu = executedByBOQ[item.maBOQ] || 0;
    item.khoiLuongDaThanhToan = paidByBOQ[item.maBOQ] || 0;
  });

  quyetToan.forEach(qt => {
    Object.assign(qt, tinhQuyetToanHopDong_(qt.maHopDong, { hopDong, boq, dotThanhToan, dotThanhToanChiTiet }));
  });

  return JSON.stringify({ hopDong, boq, phuLuc, phuLucThayDoi, nghiemThu, dotThanhToan, dotThanhToanChiTiet, quyetToan });
}

// Giống getHopDongData() nhưng CHỈ đọc/tính lại dữ liệu của ĐÚNG 1 Hợp đồng — dùng để tải lại sau
// khi Lưu ở BOQ/Phụ lục/Nghiệm thu/Thanh toán/Quyết toán/Tiến độ (mọi thao tác Lưu ở các tab đó đều
// biết chắc đang sửa Hợp đồng nào qua currentHopDongId). getHopDongData() ở trên vẫn phải giữ
// nguyên (đọc + tính lại TOÀN BỘ hệ thống) vì tab "Hợp đồng" cần tổng hợp giá trị/đã thanh toán của
// MỌI Hợp đồng cùng lúc — nhưng gọi lại nguyên hàm đó sau MỌI lần Lưu, kể cả khi chỉ sửa 1 Hợp đồng,
// khiến việc tính lại hiệu lực/quyết toán tốn công theo TOÀN BỘ dữ liệu tích lũy cả dự án (càng
// nhiều Hợp đồng/BOQ/Nghiệm thu qua các năm càng chậm dần), dù chỉ vừa sửa 1 dòng — đây chính là lý
// do "bấm Lưu xong load lâu". Cùng 8 lượt đọc sheet (chi phí đọc Google Sheets không đổi vì Sheets
// API luôn phải đọc nguyên cột), nhưng vòng lặp tính hiệu lực + payload JSON trả về client chỉ còn
// tỉ lệ với 1 Hợp đồng thay vì cả hệ thống. Trả kèm maHopDong để client biết chính xác đang gộp dữ
// liệu của Hợp đồng nào vào cache (xem applyLoadedHopDongChiTiet_ ở 0_TienIch.html).
function getHopDongChiTietMotHopDong(maHopDong) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const maHD = String(maHopDong || '').trim();

  const hopDongAll = readHopDongRows_(ss.getSheetByName('HopDong'));
  const hd = hopDongAll.find(h => h.maHopDong === maHD);
  const hopDong = hd ? [hd] : [];

  const boq = readBOQRows_(ss.getSheetByName('BOQHangMuc')).filter(b => b.maHopDong === maHD);
  const boqIds_ = new Set(boq.map(b => b.maBOQ));

  const phuLuc = readPhuLucRows_(ss.getSheetByName('PhuLucHopDong'), true).filter(p => p.maHopDong === maHD);
  const phuLucIds_ = new Set(phuLuc.map(p => p.maPhuLuc));
  const phuLucThayDoi = readPhuLucThayDoiRows_(ss.getSheetByName('PhuLucThayDoi')).filter(ch => phuLucIds_.has(ch.maPhuLuc));

  const nghiemThu = readNghiemThuRows_(ss.getSheetByName('NghiemThu')).filter(n => boqIds_.has(n.maBOQ));
  const dotThanhToan = readDotThanhToanRows_(ss.getSheetByName('DotThanhToan')).filter(d => d.maHopDong === maHD);
  const dotIds_ = new Set(dotThanhToan.map(d => d.maDotThanhToan));
  const dotThanhToanChiTiet = readDotThanhToanChiTietRows_(ss.getSheetByName('DotThanhToanChiTiet')).filter(c => dotIds_.has(c.maDotThanhToan));
  const quyetToan = readQuyetToanRows_(ss.getSheetByName('QuyetToan')).filter(q => q.maHopDong === maHD);

  const activePhuLucIds_ = new Set(phuLuc.filter(p => p.active).map(p => p.maPhuLuc));
  const phuLucThayDoiHieuLuc_ = phuLucThayDoi.filter(ch => activePhuLucIds_.has(ch.maPhuLuc));
  const effectiveQtyMap = tinhKhoiLuongHieuLuc_(boq, phuLucThayDoiHieuLuc_);
  const donGiaHieuLucMap = tinhDonGiaHieuLuc_(boq, phuLucThayDoiHieuLuc_);
  const executedByBOQ = {};
  nghiemThu.forEach(n => { executedByBOQ[n.maBOQ] = Math.round(((executedByBOQ[n.maBOQ] || 0) + n.khoiLuong) * 1000) / 1000; });
  const paidByBOQ = {};
  dotThanhToanChiTiet.forEach(c => { paidByBOQ[c.maBOQ] = Math.round(((paidByBOQ[c.maBOQ] || 0) + c.khoiLuong) * 1000) / 1000; });

  boq.forEach(item => {
    item.khoiLuongHieuLuc = effectiveQtyMap[item.maBOQ] != null ? effectiveQtyMap[item.maBOQ] : item.khoiLuongHopDong;
    item.donGiaHieuLuc = donGiaHieuLucMap[item.maBOQ] != null ? donGiaHieuLucMap[item.maBOQ] : item.donGia;
    item.khoiLuongDaNghiemThu = executedByBOQ[item.maBOQ] || 0;
    item.khoiLuongDaThanhToan = paidByBOQ[item.maBOQ] || 0;
  });

  quyetToan.forEach(qt => {
    Object.assign(qt, tinhQuyetToanHopDong_(qt.maHopDong, { hopDong, boq, dotThanhToan, dotThanhToanChiTiet }));
  });

  return JSON.stringify({ maHopDong: maHD, hopDong, boq, phuLuc, phuLucThayDoi, nghiemThu, dotThanhToan, dotThanhToanChiTiet, quyetToan });
}

// --- CRUD Hợp đồng ---
function adminSaveHopDong(userId, password, hopDong) {
  const user = kiemTraQuyenHopDong_(userId, password);
  const sheet = layHoacTaoSheet_('HopDong', HOPDONG_HEADERS_);
  const tenHopDong = String(hopDong.tenHopDong || '').trim();
  if (!tenHopDong) throw new Error('Vui lòng nhập Tên hợp đồng!');

  const maHopDong = String(hopDong.maHopDong || '').trim();
  const row = [
    '', tenHopDong, String(hopDong.maGoiThau || '').trim(), String(hopDong.nhaThau || '').trim(),
    hopDong.ngayKy || '', round0_(hopDong.tamUngHopDong), round0_(hopDong.tamUngThiCong),
    Number(hopDong.thueVAT) || 0, round0_(hopDong.giamGia), true
  ];

  if (maHopDong) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === maHopDong) {
        row[0] = maHopDong;
        sheet.getRange(i + 1, 1, 1, HOPDONG_HEADERS_.length).setValues([row]);
        logActivity(userId, user.name, 'Sửa Hợp đồng', `${maHopDong} - ${tenHopDong}`);
        return maHopDong;
      }
    }
    throw new Error('Không tìm thấy Hợp đồng!');
  }

  const newId = sinhMaTuTang_(sheet, 'HD', new Date());
  row[0] = newId;
  sheet.appendRow(row);
  logActivity(userId, user.name, 'Thêm Hợp đồng', `${newId} - ${tenHopDong}`);
  return newId;
}

function adminXoaHopDong(userId, password, maHopDong) {
  const user = kiemTraQuyenHopDong_(userId, password);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('HopDong');
  if (!sheet) throw new Error('Chưa cấu hình Tab HopDong trên Google Sheet!');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(maHopDong).trim()) {
      sheet.getRange(i + 1, 10).setValue(false);
      logActivity(userId, user.name, 'Xoá Hợp đồng', maHopDong);
      return 'Success';
    }
  }
  throw new Error('Không tìm thấy Hợp đồng!');
}

// --- CRUD BOQ ---
function adminSaveBOQItem(userId, password, item) {
  const user = kiemTraQuyenHopDong_(userId, password);
  const sheet = layHoacTaoSheet_('BOQHangMuc', BOQ_HEADERS_);
  const maHopDong = String(item.maHopDong || '').trim();
  if (!maHopDong) throw new Error('Thiếu mã Hợp đồng!');
  const tenHangMuc = String(item.tenHangMuc || '').trim();
  if (!tenHangMuc) throw new Error('Vui lòng nhập Tên hạng mục!');

  const maBOQ = String(item.maBOQ || '').trim();
  const row = [
    '', maHopDong, String(item.stt || '').trim(), tenHangMuc, !!item.isHeader,
    String(item.donVi || '').trim(), round3_(item.khoiLuongHopDong), round0_(item.donGia),
    parseInt(item.cap, 10) || 0, item.ngayBatDauKH || '', item.ngayKetThucKH || '',
    String(item.maPhuLucTao || '').trim(), true, String(item.ghiChu || '').trim()
  ];

  if (maBOQ) {
    // Chỉ đọc cột A (mã) để dò dòng thay vì tải nguyên sheet BOQHangMuc (gồm mọi Hợp đồng) — nhập
    // Excel hàng loạt gọi hàm này lặp lại cho từng dòng, đọc nguyên sheet mỗi lần khiến việc nhập
    // nhiều hạng mục cùng lúc rất chậm.
    const lastRow = sheet.getLastRow();
    const ids = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, 1).getValues() : [];
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === maBOQ) {
        const rowIdx = i + 2;
        row[0] = maBOQ;
        row[11] = String(sheet.getRange(rowIdx, 12).getValue() || '').trim(); // giữ nguyên maPhuLucTao gốc, không cho sửa tay
        sheet.getRange(rowIdx, 1, 1, BOQ_HEADERS_.length).setValues([row]);
        logActivity(userId, user.name, 'Sửa hạng mục BOQ', `${maBOQ} - ${tenHangMuc}`);
        return maBOQ;
      }
    }
    throw new Error('Không tìm thấy Hạng mục BOQ!');
  }

  const newId = sinhMaTuTang_(sheet, 'BOQ', new Date());
  row[0] = newId;
  sheet.appendRow(row);
  logActivity(userId, user.name, 'Thêm hạng mục BOQ', `${newId} - ${tenHangMuc}`);
  return newId;
}

// itemArray: mảng tham số item giống hệt adminSaveBOQItem() ở trên — có thể trộn lẫn Sửa (có maBOQ,
// dùng bởi import Excel Tiến độ — xem importTdtcFromExcel ở 14_TienDoThiCong.html) và Thêm mới (rỗng
// maBOQ, dùng bởi import Excel BOQ — xem importBOQFromExcel ở 10_BOQ.html) trong cùng 1 lượt gọi. Gộp
// thành 1 lần đọc cột A (dò mã) + 1 lần ghi hàng loạt thay vì client gọi lặp adminSaveBOQItem() từng
// dòng — file Excel nhiều dòng trước đây là N round-trip riêng lên Apps Script, rất chậm khi import.
// Trả về JSON mảng kết quả song song với itemArray, mỗi phần tử { ok, maBOQ } hoặc { ok:false, error }.
function adminSaveBOQItemHangLoat(userId, password, itemArray) {
  const user = kiemTraQuyenHopDong_(userId, password);
  const sheet = layHoacTaoSheet_('BOQHangMuc', BOQ_HEADERS_);
  const lastRow = sheet.getLastRow();
  const existing = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, 12).getValues() : [];
  const idIndex = new Map(); // maBOQ -> { rowIdx, maPhuLucTao }
  existing.forEach((r, i) => {
    const id = String(r[0] || '').trim();
    if (id) idIndex.set(id, { rowIdx: i + 2, maPhuLucTao: String(r[11] || '').trim() });
  });

  const now = new Date();
  const nowStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
  const fullPrefix = 'BOQ-' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd') + '-';
  let maxSTT = 0;
  existing.forEach(r => {
    const id = String(r[0] || '');
    if (id.indexOf(fullPrefix) === 0) {
      const stt = parseInt(id.slice(fullPrefix.length), 10);
      if (!isNaN(stt) && stt > maxSTT) maxSTT = stt;
    }
  });

  const results = [];
  const newRows = [];
  const logRows = [];

  (itemArray || []).forEach(item => {
    try {
      const maHopDong = String(item.maHopDong || '').trim();
      if (!maHopDong) throw new Error('Thiếu mã Hợp đồng!');
      const tenHangMuc = String(item.tenHangMuc || '').trim();
      if (!tenHangMuc) throw new Error('Vui lòng nhập Tên hạng mục!');

      const maBOQ = String(item.maBOQ || '').trim();
      const row = [
        '', maHopDong, String(item.stt || '').trim(), tenHangMuc, !!item.isHeader,
        String(item.donVi || '').trim(), round3_(item.khoiLuongHopDong), round0_(item.donGia),
        parseInt(item.cap, 10) || 0, item.ngayBatDauKH || '', item.ngayKetThucKH || '',
        String(item.maPhuLucTao || '').trim(), true, String(item.ghiChu || '').trim()
      ];

      if (maBOQ) {
        const found = idIndex.get(maBOQ);
        if (!found) throw new Error('Không tìm thấy Hạng mục BOQ!');
        row[0] = maBOQ;
        row[11] = found.maPhuLucTao; // giữ nguyên maPhuLucTao gốc, không cho sửa tay
        sheet.getRange(found.rowIdx, 1, 1, BOQ_HEADERS_.length).setValues([row]);
        logRows.push([nowStr, `${user.name} (${userId})`, 'Sửa hạng mục BOQ', `${maBOQ} - ${tenHangMuc}`]);
        results.push({ ok: true, maBOQ });
      } else {
        maxSTT++;
        const newId = fullPrefix + String(maxSTT).padStart(3, '0');
        row[0] = newId;
        newRows.push(row);
        logRows.push([nowStr, `${user.name} (${userId})`, 'Thêm hạng mục BOQ', `${newId} - ${tenHangMuc}`]);
        results.push({ ok: true, maBOQ: newId });
      }
    } catch (e) {
      results.push({ ok: false, error: e.message });
    }
  });

  if (newRows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, BOQ_HEADERS_.length).setValues(newRows);
  }
  if (logRows.length) {
    const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('NhatKy');
    if (logSheet) logSheet.getRange(logSheet.getLastRow() + 1, 1, logRows.length, 4).setValues(logRows);
  }

  return JSON.stringify(results);
}

function adminXoaBOQItem(userId, password, maBOQ) {
  const user = kiemTraQuyenHopDong_(userId, password);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BOQHangMuc');
  if (!sheet) throw new Error('Chưa cấu hình Tab BOQHangMuc trên Google Sheet!');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(maBOQ).trim()) {
      sheet.getRange(i + 1, 13).setValue(false);
      logActivity(userId, user.name, 'Xoá hạng mục BOQ', maBOQ);
      return 'Success';
    }
  }
  throw new Error('Không tìm thấy Hạng mục BOQ!');
}

// --- Phụ lục Hợp đồng ---
function adminTaoPhuLuc(userId, password, maHopDong, soHieu, ngayPhuLuc, ghiChu) {
  const user = kiemTraQuyenHopDong_(userId, password);
  const soHieuTrim = String(soHieu || '').trim();
  if (!soHieuTrim) throw new Error('Vui lòng nhập Số hiệu Phụ lục!');
  const sheet = layHoacTaoSheet_('PhuLucHopDong', PHULUC_HEADERS_);
  const newId = sinhMaTuTang_(sheet, 'PL', new Date());
  sheet.appendRow([newId, String(maHopDong || '').trim(), soHieuTrim, ngayPhuLuc || '', String(ghiChu || '').trim(), true]);
  logActivity(userId, user.name, 'Thêm Phụ lục Hợp đồng', `${newId} - ${soHieuTrim}`);
  return newId;
}

// Sửa thông tin (Số hiệu/Ngày/Ghi chú) 1 Phụ lục đã tạo — KHÔNG đổi các dòng "thay đổi" bên trong.
function adminSuaPhuLuc(userId, password, maPhuLuc, soHieu, ngayPhuLuc, ghiChu) {
  const user = kiemTraQuyenHopDong_(userId, password);
  const soHieuTrim = String(soHieu || '').trim();
  if (!soHieuTrim) throw new Error('Vui lòng nhập Số hiệu Phụ lục!');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PhuLucHopDong');
  if (!sheet) throw new Error('Chưa cấu hình Tab PhuLucHopDong trên Google Sheet!');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(maPhuLuc).trim()) {
      sheet.getRange(i + 1, 3, 1, 2).setValues([[soHieuTrim, ngayPhuLuc || '']]);
      sheet.getRange(i + 1, 5).setValue(String(ghiChu || '').trim());
      logActivity(userId, user.name, 'Sửa Phụ lục Hợp đồng', `${maPhuLuc} - ${soHieuTrim}`);
      return 'Success';
    }
  }
  throw new Error('Không tìm thấy Phụ lục!');
}

// Vô hiệu hóa/Khôi phục 1 Phụ lục — KHÔNG xóa hẳn (đúng rule app cũ, toggleAddendumActive): chỉ tạm
// ngừng/áp dụng lại toàn bộ thay đổi khối lượng/đơn giá của Phụ lục này (xem getHopDongData()),
// dữ liệu Phụ lục và các dòng thay đổi vẫn giữ nguyên để tra cứu lịch sử. Trả về trạng thái MỚI.
function adminToggleTrangThaiPhuLuc(userId, password, maPhuLuc) {
  const user = kiemTraQuyenHopDong_(userId, password);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PhuLucHopDong');
  if (!sheet) throw new Error('Chưa cấu hình Tab PhuLucHopDong trên Google Sheet!');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(maPhuLuc).trim()) {
      const dangActive = isActiveVal_(data[i][5]);
      sheet.getRange(i + 1, 6).setValue(!dangActive);
      logActivity(userId, user.name, dangActive ? 'Vô hiệu hóa Phụ lục' : 'Khôi phục Phụ lục', maPhuLuc);
      return !dangActive;
    }
  }
  throw new Error('Không tìm thấy Phụ lục!');
}

// thayDoi: { maBOQ, loaiThayDoi, khoiLuongMoi?, khoiLuongDieuChinh?, donGiaMoi?, boqMoi? } — boqMoi
// (object hạng mục mới) bắt buộc khi loaiThayDoi = 'MOI'.
function adminThemThayDoiPhuLuc(userId, password, maPhuLuc, thayDoi) {
  const user = kiemTraQuyenHopDong_(userId, password);
  const loai = String(thayDoi.loaiThayDoi || '').trim();
  if (['MOI', 'THAY_THE', 'DIEU_CHINH'].indexOf(loai) === -1) throw new Error('Loại thay đổi không hợp lệ!');

  let maBOQ = String(thayDoi.maBOQ || '').trim();
  if (loai === 'MOI') {
    if (!thayDoi.boqMoi) throw new Error('Thiếu thông tin Hạng mục mới!');
    const pl = readPhuLucRows_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PhuLucHopDong'), true).find(p => p.maPhuLuc === maPhuLuc);
    if (!pl) throw new Error('Không tìm thấy Phụ lục!');
    maBOQ = adminSaveBOQItem(userId, password, Object.assign({}, thayDoi.boqMoi, {
      maHopDong: pl.maHopDong, maPhuLucTao: maPhuLuc
    }));
  } else if (!maBOQ) {
    throw new Error('Vui lòng chọn Hạng mục BOQ cần thay đổi!');
  }

  const sheet = layHoacTaoSheet_('PhuLucThayDoi', PHULUC_THAYDOI_HEADERS_);
  const newId = sinhMaTuTang_(sheet, 'PLTD', new Date());
  sheet.appendRow([
    newId, maPhuLuc, maBOQ, loai,
    thayDoi.khoiLuongMoi != null ? round3_(thayDoi.khoiLuongMoi) : '',
    thayDoi.khoiLuongDieuChinh != null ? round3_(thayDoi.khoiLuongDieuChinh) : '',
    thayDoi.donGiaMoi != null ? round0_(thayDoi.donGiaMoi) : '', true
  ]);
  logActivity(userId, user.name, 'Thêm thay đổi Phụ lục', `${maPhuLuc}: ${loai} - ${maBOQ}`);
  return newId;
}

// --- Nghiệm thu (log khối lượng, append-only — copy mẫu validateVaGhiNhatKy_/themNhatKyTienDo) ---
function themNghiemThu(userId, password, maBOQ, ngayNghiemThu, khoiLuong, ghiChu) {
  const user = kiemTraQuyenHopDong_(userId, password);
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const boqSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BOQHangMuc');
    if (!boqSheet) throw new Error('Chưa cấu hình Tab BOQHangMuc trên Google Sheet!');
    const item = readBOQRows_(boqSheet).find(b => b.maBOQ === maBOQ);
    if (!item) throw new Error('Không tìm thấy Hạng mục BOQ!');

    const klNum = parseFloat(khoiLuong);
    if (isNaN(klNum) || klNum <= 0) throw new Error('Khối lượng không hợp lệ!');

    const sheet = layHoacTaoSheet_('NghiemThu', NGHIEMTHU_HEADERS_);
    const newId = sinhMaTuTang_(sheet, 'NT', new Date());
    sheet.appendRow([
      newId, maBOQ, ngayNghiemThu || '', round3_(klNum), String(ghiChu || '').trim(),
      user.supervisorId, user.name, new Date(), true
    ]);
    logActivity(userId, user.name, 'Nghiệm thu khối lượng', `${maBOQ}: +${klNum} ${item.donVi}`);
    return 'Success';
  } finally {
    lock.releaseLock();
  }
}

// Sửa lại khối lượng/ghi chú của ĐÚNG 1 dòng Nghiệm thu đã ghi (khác themNghiemThu() luôn tạo dòng
// mới) — dùng khi Giám sát/QS cần sửa lại số liệu đã nhập sai cho 1 ngày cụ thể (xem openNtEditModal
// ở 11_NghiemThu.html).
function adminSuaNghiemThu(userId, password, maNghiemThu, khoiLuong, ghiChu) {
  const user = kiemTraQuyenHopDong_(userId, password);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('NghiemThu');
  if (!sheet) throw new Error('Chưa cấu hình Tab NghiemThu trên Google Sheet!');
  const klNum = parseFloat(khoiLuong);
  if (isNaN(klNum) || klNum <= 0) throw new Error('Khối lượng không hợp lệ!');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(maNghiemThu).trim()) {
      sheet.getRange(i + 1, 4).setValue(round3_(klNum));
      sheet.getRange(i + 1, 5).setValue(String(ghiChu || '').trim());
      logActivity(userId, user.name, 'Sửa Nghiệm thu', maNghiemThu);
      return 'Success';
    }
  }
  throw new Error('Không tìm thấy dòng Nghiệm thu!');
}

function adminXoaNghiemThu(userId, password, maNghiemThu) {
  const user = kiemTraQuyenHopDong_(userId, password);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('NghiemThu');
  if (!sheet) throw new Error('Chưa cấu hình Tab NghiemThu trên Google Sheet!');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(maNghiemThu).trim()) {
      sheet.getRange(i + 1, 9).setValue(false);
      logActivity(userId, user.name, 'Xoá Nghiệm thu', maNghiemThu);
      return 'Success';
    }
  }
  throw new Error('Không tìm thấy dòng Nghiệm thu!');
}

// --- Thanh toán ---
// dotThanhToan: { maHopDong, tenDot, ngayThanhToan, phanTramThanhToan, ghiChu, thuHoiTamUngHopDong,
// thuHoiTamUngThiCong, khauTruKhac, ghiChuKhauTru, chiTiet: [{maBOQ, khoiLuong}] }
function adminTaoDotThanhToan(userId, password, dotThanhToan) {
  const user = kiemTraQuyenHopDong_(userId, password);
  const tenDot = String(dotThanhToan.tenDot || '').trim();
  if (!tenDot) throw new Error('Vui lòng nhập Tên/Số hiệu Đợt thanh toán!');
  const maHopDong = String(dotThanhToan.maHopDong || '').trim();
  if (!maHopDong) throw new Error('Thiếu mã Hợp đồng!');

  const sheet = layHoacTaoSheet_('DotThanhToan', DOTTHANHTOAN_HEADERS_);
  const newId = sinhMaTuTang_(sheet, 'TT', new Date());
  sheet.appendRow([
    newId, maHopDong, tenDot, dotThanhToan.ngayThanhToan || '',
    dotThanhToan.phanTramThanhToan != null ? Number(dotThanhToan.phanTramThanhToan) : 100,
    String(dotThanhToan.ghiChu || '').trim(), round0_(dotThanhToan.thuHoiTamUngHopDong),
    round0_(dotThanhToan.thuHoiTamUngThiCong), round0_(dotThanhToan.khauTruKhac),
    String(dotThanhToan.ghiChuKhauTru || '').trim(), true, round0_(dotThanhToan.giamGiaDot)
  ]);

  const chiTietSheet = layHoacTaoSheet_('DotThanhToanChiTiet', DOTTHANHTOAN_CHITIET_HEADERS_);
  (dotThanhToan.chiTiet || []).forEach(ct => {
    const klNum = parseFloat(ct.khoiLuong);
    if (!ct.maBOQ || isNaN(klNum) || klNum <= 0) return;
    const ctId = sinhMaTuTang_(chiTietSheet, 'TTCT', new Date());
    chiTietSheet.appendRow([ctId, newId, ct.maBOQ, round3_(klNum), true]);
  });

  logActivity(userId, user.name, 'Thêm Đợt thanh toán', `${newId} - ${tenDot}`);
  return newId;
}

// Chỉ cho sửa phần điều chỉnh (thu hồi tạm ứng/khấu trừ/ghi chú) sau khi tạo, giống app cũ
// (savePaymentAdjustment) — KHÔNG cho đổi lại khối lượng/% đã chốt của Đợt.
function adminSuaDieuChinhThanhToan(userId, password, maDotThanhToan, dieuChinh) {
  const user = kiemTraQuyenHopDong_(userId, password);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DotThanhToan');
  if (!sheet) throw new Error('Chưa cấu hình Tab DotThanhToan trên Google Sheet!');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(maDotThanhToan).trim()) {
      sheet.getRange(i + 1, 7, 1, 4).setValues([[
        round0_(dieuChinh.thuHoiTamUngHopDong), round0_(dieuChinh.thuHoiTamUngThiCong),
        round0_(dieuChinh.khauTruKhac), String(dieuChinh.ghiChuKhauTru || '').trim()
      ]]);
      // giamGiaDot nằm ở cột 12 (sau active, cột 11) — không liền kề 4 cột trên nên ghi riêng.
      sheet.getRange(i + 1, 12).setValue(round0_(dieuChinh.giamGiaDot));
      logActivity(userId, user.name, 'Sửa điều chỉnh Thanh toán', maDotThanhToan);
      return 'Success';
    }
  }
  throw new Error('Không tìm thấy Đợt thanh toán!');
}

function adminXoaDotThanhToan(userId, password, maDotThanhToan) {
  const user = kiemTraQuyenHopDong_(userId, password);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DotThanhToan');
  if (!sheet) throw new Error('Chưa cấu hình Tab DotThanhToan trên Google Sheet!');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(maDotThanhToan).trim()) {
      sheet.getRange(i + 1, 11).setValue(false);
      logActivity(userId, user.name, 'Xoá Đợt thanh toán', maDotThanhToan);
      return 'Success';
    }
  }
  throw new Error('Không tìm thấy Đợt thanh toán!');
}

// Gợi ý số Thu hồi tạm ứng cho 1 Đợt = tỷ lệ (Tạm ứng / Giá trị HĐ hiện hành) x Giá trị thực hiện kỳ
// này, không vượt số tạm ứng còn lại — port nguyên công thức suggestAdvanceRecovery() của app cũ.
// loaiTamUng: 'hopDong' | 'thiCong'. chiTietKyNay: [{maBOQ, khoiLuong}] (khối lượng dự kiến kỳ này).
function goiYThuHoiTamUng(maHopDong, chiTietKyNay, loaiTamUng) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hopDong = readHopDongRows_(ss.getSheetByName('HopDong')).find(h => h.maHopDong === maHopDong);
  if (!hopDong) return 0;
  const advanceAmount = loaiTamUng === 'thiCong' ? hopDong.tamUngThiCong : hopDong.tamUngHopDong;
  if (!advanceAmount) return 0;

  const boq = readBOQRows_(ss.getSheetByName('BOQHangMuc')).filter(b => b.maHopDong === maHopDong && !b.isHeader);
  const phuLucThayDoi = readPhuLucThayDoiRows_(ss.getSheetByName('PhuLucThayDoi'));
  const effMap = tinhKhoiLuongHieuLuc_(boq, phuLucThayDoi);
  const giaMap = tinhDonGiaHieuLuc_(boq, phuLucThayDoi);
  const contractValue = Math.max(0, Math.round(boq.reduce((s, b) => s + (effMap[b.maBOQ] || 0) * (giaMap[b.maBOQ] || 0), 0)) - (hopDong.giamGia || 0));
  if (contractValue <= 0) return 0;

  let thucHien = 0;
  (chiTietKyNay || []).forEach(ct => { thucHien += (ct.khoiLuong || 0) * (giaMap[ct.maBOQ] || 0); });

  const raw = Math.round(thucHien * (advanceAmount / contractValue));
  const dots = readDotThanhToanRows_(ss.getSheetByName('DotThanhToan')).filter(d => d.maHopDong === maHopDong);
  const daThuHoi = dots.reduce((s, d) => s + (loaiTamUng === 'thiCong' ? d.thuHoiTamUngThiCong : d.thuHoiTamUngHopDong), 0);
  return Math.min(raw, Math.max(0, advanceAmount - daThuHoi));
}

// --- Quyết toán (tự tính, chỉ lưu mốc xác nhận — xem tinhQuyetToanHopDong_()) ---
function adminTaoQuyetToan(userId, password, maHopDong, ghiChu) {
  const user = kiemTraQuyenHopDong_(userId, password);
  const maHD = String(maHopDong || '').trim();
  if (!maHD) throw new Error('Thiếu mã Hợp đồng!');

  const sheet = layHoacTaoSheet_('QuyetToan', QUYETTOAN_HEADERS_);
  if (readQuyetToanRows_(sheet).some(q => q.maHopDong === maHD)) throw new Error('Hợp đồng này đã có Quyết toán!');

  const newId = sinhMaTuTang_(sheet, 'QT', new Date());
  sheet.appendRow([newId, maHD, Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'), user.supervisorId, user.name, String(ghiChu || '').trim(), true]);
  logActivity(userId, user.name, 'Xác nhận Quyết toán Hợp đồng', maHD);
  return newId;
}

function adminXoaQuyetToan(userId, password, maQuyetToan) {
  const user = kiemTraQuyenHopDong_(userId, password);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('QuyetToan');
  if (!sheet) throw new Error('Chưa cấu hình Tab QuyetToan trên Google Sheet!');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(maQuyetToan).trim()) {
      sheet.getRange(i + 1, 7).setValue(false);
      logActivity(userId, user.name, 'Hủy Quyết toán', maQuyetToan);
      return 'Success';
    }
  }
  throw new Error('Không tìm thấy Quyết toán!');
}
