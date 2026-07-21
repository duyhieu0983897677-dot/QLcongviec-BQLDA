// Hàm include() cho phép Index.html nhúng Style.html và các file Script.html
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Trạng thái nhân sự (cột 'TrangThai' trong DanhSachUser, cột G/index 6) — chỉ Admin đặt được (xem
// adminSetTrangThaiNhanSu). Cột này thêm SAU khi nhiều tài khoản đã tồn tại nên các dòng cũ trống ô
// này — coi trống = TRANGTHAI_HOATDONG_ (đang hoạt động) để không khoá nhầm ai (xem listSupervisors/
// login). "Nghỉ việc": khoá đăng nhập hẳn (login() chặn), tên hiển thị gạch ngang, ẩn khỏi các danh
// sách CHỌN MỚI (giao việc, mời vào phòng chat...) — nhưng dữ liệu lịch sử vẫn giữ nguyên. "Thử việc":
// vẫn đăng nhập/chấm công bình thường nhưng KHÔNG được cộng 1 ngày phép baseline mỗi tháng (chỉ nhân
// sự chính thức — TRANGTHAI_HOATDONG_ — mới được, xem chotSoSachDenThangHienTai_).
const TRANGTHAI_HOATDONG_ = 'Đang hoạt động';
const TRANGTHAI_THUVIEC_ = 'Thử việc';
const TRANGTHAI_NGHIVIEC_ = 'Nghỉ việc';

// KhongChamCongChung (cột I/index 8 trong DanhSachUser) — cờ RIÊNG, KHÁC với TrangThai: dành cho nhân
// sự không tham gia chấm công/tăng ca chung với BQLDA (VD tài khoản dùng chung, lãnh đạo không chấm
// công...) — KHÔNG ảnh hưởng đăng nhập/các danh sách chọn khác, CHỈ đóng băng Phép/Bù/Số công tính
// lương ở tab Chấm công (xem adminSetKhongChamCongChung/chotSoSachDenThangHienTai_/getChamCongThang).

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
// PhepThang: sổ PHÉP/tháng (phepDauThang = baseline đầu tháng, tự cộng 1 ngày/tháng CHỈ cho nhân sự
// CHÍNH THỨC — xem chotSoSachDenThangHienTai_). Theo đúng cách BQLDA đang làm tay trên Excel (cột Phép
// T5/T6 TÁCH RIÊNG khỏi cột Bù T5/T6): Phép giờ CHỈ còn bị trừ bởi ngày nghỉ P (miễn trừ nếu tháng đó
// có phát sinh Bù) — phần dư/thiếu cuối tuần và tăng ca KHÔNG còn ảnh hưởng Phép nữa, mà dồn hết vào
// sổ BuThang bên dưới.
// BuThang: sổ BÙ/tháng (ngày nghỉ bù tích lũy từ 2 nguồn: (1) phần dư cuối tuần vượt định mức Thứ
// 7/CN, (2) giờ tăng ca quy đổi theo block 4h=0,5 công/8h=1 công — xem tinhQuyDoiTangCaCong_). Áp
// dụng cho CẢ nhân sự chính thức lẫn Thử việc (không phân biệt, khác với Phép). gioTCDauThang: số GIỜ
// tăng ca lẻ (chưa đủ 4h) mang từ tháng trước sang, lưu trong CHÍNH sổ Bù (không phải Phép) vì đây là
// nguồn nuôi số Bù, không phải Phép.
// ChamCongThang (định nghĩa ở phần CHẤM CÔNG & TĂNG CA phía dưới): chấm công + tăng ca CẢ THÁNG —
// "Cộng/Trừ phép tháng"/"Bù phát sinh tháng" luôn tính lại từ ChamCongThang, không lưu cột riêng.
const PHEP_HEADERS_ = ['userId', 'thang', 'phepDauThang', 'soNgayNghi', 'chotLuc'];
const BU_HEADERS_ = ['userId', 'thang', 'buDauThang', 'soNgayNghi', 'chotLuc', 'gioTCDauThang'];
const THU_KY_CHUC_DANH_ = 'thư ký bqlda';

// NgayLe: danh sách ngày nghỉ lễ (ngay 'yyyy-MM-dd', tenLe) do Admin tự nhập/xoá (VD Quốc khánh 2/9,
// nghỉ bù 3/9...) — CHỈ dùng để tính "Số ngày công chuẩn dự kiến" tham khảo hiển thị ở tab Chấm công
// (xem getChamCongThang/getDanhSachNgayLe), KHÔNG ảnh hưởng tới "Số công tính lương" của từng nhân sự
// (chấm công cá nhân vẫn tính y như hiện tại, không tự miễn trừ theo ngày lễ).
const NGAYLE_HEADERS_ = ['ngay', 'tenLe'];

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
    sh.getRange(1, 1, 2, 8).setValues([
      ['ID', 'MatKhau', 'VaiTro', 'HoTen', 'ChucDanh', 'PhanMacDinh', 'TrangThai', 'Salt'],
      ['admin', '123456', 'ADMIN', 'Ban QLDA', '', '', TRANGTHAI_HOATDONG_, '']
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

  if (!ss.getSheetByName('BuThang')) {
    ss.insertSheet('BuThang').getRange(1, 1, 1, BU_HEADERS_.length).setValues([BU_HEADERS_]);
  }

  if (!ss.getSheetByName('NgayLe')) {
    ss.insertSheet('NgayLe').getRange(1, 1, 1, NGAYLE_HEADERS_.length).setValues([NGAYLE_HEADERS_]);
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

  return 'Đã thiết lập xong: Data, DanhSachUser (admin/123456), NhatKy, HangMuc, GoiThau, CongViec, NhatKyTienDo, PhepThang, BuThang, NgayLe, ChamCongThang, ThongBao, BinhLuanCongViec, HopDong, BOQHangMuc, PhuLucHopDong, PhuLucThayDoi, NghiemThu, DotThanhToan, DotThanhToanChiTiet, QuyetToan.';
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
      phanMacDinh: String(data[i][5] || '').trim(),
      trangThai: String(data[i][6] || '').trim() || TRANGTHAI_HOATDONG_,
      khongChamCongChung: data[i][8] === true || String(data[i][8] || '').trim().toUpperCase() === 'TRUE'
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
  if (!sheet.getRange(1, 7).getValue()) sheet.getRange(1, 7).setValue('TrangThai');
  if (!sheet.getRange(1, 8).getValue()) sheet.getRange(1, 8).setValue('Salt');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(targetId).trim()) {
      sheet.getRange(i + 1, 3).setValue(role);
      sheet.getRange(i + 1, 4).setValue(name);
      sheet.getRange(i + 1, 5).setValue(chucDanh || '');
      sheet.getRange(i + 1, 6).setValue(phanMacDinh || '');
      if (newPassword) {
        sheet.getRange(i + 1, 2).setValue(newPassword); // lưu thường, tự mã hóa ở lần đăng nhập đầu (xem login())
        sheet.getRange(i + 1, 8).setValue(''); // xoá Salt cũ (nếu có) — kẻo login() so khớp nhầm mật khẩu thường với hash cũ
      }
      logActivity(adminId, adminUser.name, "Sửa Giám sát", `Sửa tài khoản ${targetId}`);
      return "Success";
    }
  }

  if (!newPassword) throw new Error("Vui lòng đặt mật khẩu ban đầu cho tài khoản mới!");
  sheet.appendRow([targetId, newPassword, role, name, chucDanh || '', phanMacDinh || '', TRANGTHAI_HOATDONG_]);
  logActivity(adminId, adminUser.name, "Thêm Giám sát", `Thêm tài khoản ${targetId}`);
  return "Success";
}

// (Admin) Đổi Trạng thái 1 nhân sự — 'Đang hoạt động'/'Thử việc'/'Nghỉ việc' (xem TRANGTHAI_*_ đầu
// file). Nghỉ việc = khoá đăng nhập hẳn (login() chặn), KHÔNG xoá tài khoản/dữ liệu lịch sử liên quan.
function adminSetTrangThaiNhanSu(adminId, adminPass, targetId, trangThai) {
  const adminUser = login(adminId, adminPass);
  if (adminUser.role !== 'ADMIN') throw new Error("Chỉ Quản trị mới có quyền đổi Trạng thái nhân sự!");

  const tt = String(trangThai || '').trim();
  if (tt !== TRANGTHAI_HOATDONG_ && tt !== TRANGTHAI_THUVIEC_ && tt !== TRANGTHAI_NGHIVIEC_) {
    throw new Error("Trạng thái không hợp lệ!");
  }
  if (String(targetId).trim() === String(adminId).trim() && tt === TRANGTHAI_NGHIVIEC_) {
    throw new Error("Không thể tự đặt chính mình thành Nghỉ việc!");
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DanhSachUser');
  if (!sheet) throw new Error("Chưa cấu hình Tab DanhSachUser trên Google Sheet!");
  if (!sheet.getRange(1, 7).getValue()) sheet.getRange(1, 7).setValue('TrangThai');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(targetId).trim()) {
      sheet.getRange(i + 1, 7).setValue(tt);
      logActivity(adminId, adminUser.name, "Đổi Trạng thái nhân sự", `${targetId} -> ${tt}`);
      return "Success";
    }
  }
  throw new Error("Không tìm thấy nhân sự!");
}

// (Admin) Đánh dấu 1 nhân sự "Không chấm công chung" (cột 'KhongChamCongChung' trong DanhSachUser,
// cột I/index 8) — dành cho nhân sự không tham gia chấm công/tăng ca chung với BQLDA (VD tài khoản
// dùng chung, cấp lãnh đạo không chấm công...). KHÔNG ảnh hưởng đăng nhập/các danh sách chọn khác
// (khác với TrangThai Nghỉ việc) — CHỈ đóng băng Phép/Bù/Số công tính lương ở tab Chấm công (xem
// chotSoSachDenThangHienTai_/getChamCongThang), dòng của họ hiện dấu "—" thay vì số.
function adminSetKhongChamCongChung(adminId, adminPass, targetId, giaTri) {
  const adminUser = login(adminId, adminPass);
  if (adminUser.role !== 'ADMIN') throw new Error("Chỉ Quản trị mới có quyền đặt \"Không chấm công chung\"!");

  const targetIdStr = String(targetId || '').trim();
  if (!targetIdStr) throw new Error("Thiếu thông tin nhân sự!");
  const coGiaTri = giaTri === true || giaTri === 'true' || giaTri === 1 || giaTri === '1';

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DanhSachUser');
  if (!sheet) throw new Error("Chưa cấu hình Tab DanhSachUser trên Google Sheet!");
  if (!sheet.getRange(1, 9).getValue()) sheet.getRange(1, 9).setValue('KhongChamCongChung');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === targetIdStr) {
      sheet.getRange(i + 1, 9).setValue(coGiaTri);
      logActivity(adminId, adminUser.name, "Đổi Không chấm công chung", `${targetIdStr}: ${coGiaTri}`);
      return "Success";
    }
  }
  throw new Error("Không tìm thấy nhân sự!");
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
// TĂNG CA bên dưới, nơi định nghĩa CHAMCONG_HEADERS_/tinhChiTietThangChoNguoi_/assertQuyenChamCongTangCa_).
// Tab "Điểm danh & Phép" (Thứ 7/CN riêng, sổ DiemDanhCuoiTuan) đã bị GỘP vào tab Chấm công & Tăng ca
// để tránh 2 nơi chấm công trùng lặp — chotSoSachDenThangHienTai_ dưới đây giờ tính "Cộng/Trừ phép
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

function readBuThangRows_(sheet) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, BU_HEADERS_.length).getValues();
  const result = [];
  values.forEach(r => {
    const userId = String(r[0] || '').trim();
    const thang = formatThangCell_(r[1]);
    if (!userId || !thang) return;
    result.push({
      userId, thang, buDauThang: Number(r[2]) || 0, soNgayNghi: Number(r[3]) || 0,
      gioTCDauThang: Number(r[5]) || 0
    });
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
// phepDauThang(tháng mới) = TổngCộng(tháng cũ) + baseline (1 nếu là nhân sự CHÍNH THỨC — TRANGTHAI_
// HOATDONG_ —, 0 nếu đang Thử việc; dùng ĐÚNG trạng thái hiện tại của tài khoản cho mọi tháng bù lại,
// vì sheet không lưu lịch sử trạng thái theo từng tháng). Nhân sự Nghỉ việc bị ĐÓNG BĂNG hoàn toàn —
// không tạo/chốt thêm dòng PhepThang/BuThang nào nữa kể từ khi phát hiện trạng thái Nghỉ việc.
// Bù (buDauThang, sổ RIÊNG BuThang) chốt CÙNG lúc với Phép trong 1 vòng lặp vì Phép giờ phụ thuộc Bù
// (miễn trừ ngày P nếu tháng đó có phát sinh Bù — xem tinhChiTietThangChoNguoi_). Bù áp dụng cho CẢ
// chính thức lẫn Thử việc (không có baseline riêng, chỉ cộng dồn từ cuối tuần dư + tăng ca quy đổi).
// tinhChiTietThangChoNguoi_ định nghĩa ở phần CHẤM CÔNG & TĂNG CA bên dưới (cùng công thức hiển thị ở
// getChamCongThang, tránh lệch số giữa "dự kiến" trên giao diện và số THẬT được chốt ở đây).
// allUsers/ccByThang_ do getChamCongThang() đọc + nhóm sẵn truyền vào — KHÔNG tự đọc lại
// ChamCongThang ở đây nữa (trước đây hàm này đọc 1 lần, rồi getChamCongThang() đọc lại lần 2 NGUYÊN
// SHEET đó ngay sau khi hàm này chạy xong, lãng phí 1 lượt đọc trùng mỗi lần mở tab/đổi tháng).
// Trả về LUÔN { allPhep, allBu } đã gồm cả các dòng vừa chốt thêm (nếu có) — trước đây hàm này chỉ ghi
// vào Sheet rồi thôi, buộc getChamCongThang() phải đọc lại NGUYÊN sheet lần thứ 2 ngay sau đó để lấy
// đúng dữ liệu mới nhất, dù dữ liệu đó vốn đã có sẵn trong bộ nhớ ở đây rồi — mỗi lượt đọc sheet tốn
// ~0,4-1 giây CỐ ĐỊNH ở nền tảng Apps Script bất kể sheet nhỏ hay lớn (đã đo thực tế: 23 nhân sự/317
// dòng ChamCongThang vẫn mất 3,8 giây tổng cộng cho getChamCongThang, chủ yếu do CÁC LƯỢT ĐỌC SHEET
// RIÊNG LẺ, không phải do khối lượng dữ liệu) — bớt được 1 lượt đọc trùng là bớt được ~1 giây mỗi lần
// mở tab/đổi tháng.
function chotSoSachDenThangHienTai_(allUsers, ccByThang_) {
  const phepSheet = layHoacTaoSheet_('PhepThang', PHEP_HEADERS_);
  const buSheet = layHoacTaoSheet_('BuThang', BU_HEADERS_);

  const thangHienTai = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  const allPhep = readPhepThangRows_(phepSheet);
  const allBu = readBuThangRows_(buSheet);

  const phepRowsToAppend = [];
  const buRowsToAppend = [];

  allUsers.forEach(u => {
    const trangThai = u.trangThai || TRANGTHAI_HOATDONG_;
    if (trangThai === TRANGTHAI_NGHIVIEC_) return; // đã nghỉ việc — đóng băng cả 2 sổ, không chốt thêm.
    if (u.khongChamCongChung) return; // không chấm công chung với BQLDA — đóng băng Phép/Bù tương tự.
    const baseline = trangThai === TRANGTHAI_HOATDONG_ ? 1 : 0; // Thử việc: không có phép baseline (Bù vẫn có).

    const phepOfUser = allPhep.filter(p => p.userId === u.id).sort((a, b) => a.thang.localeCompare(b.thang));
    const buOfUser = allBu.filter(p => p.userId === u.id).sort((a, b) => a.thang.localeCompare(b.thang));
    let lastPhep = phepOfUser.length ? phepOfUser[phepOfUser.length - 1] : null;
    let lastBu = buOfUser.length ? buOfUser[buOfUser.length - 1] : null;

    if (!lastPhep && !lastBu) {
      // Nhân sự hoàn toàn mới — chỉ bắt đầu từ tháng hiện tại cho cả 2 sổ, không truy hồi lùi.
      phepRowsToAppend.push([u.id, thangHienTai, baseline, 0, new Date()]);
      allPhep.push({ userId: u.id, thang: thangHienTai, phepDauThang: baseline, soNgayNghi: 0 });
      buRowsToAppend.push([u.id, thangHienTai, 0, 0, new Date(), 0]);
      allBu.push({ userId: u.id, thang: thangHienTai, buDauThang: 0, soNgayNghi: 0, gioTCDauThang: 0 });
      return;
    }

    // 1 trong 2 sổ chưa có dòng nào (VD sổ Bù mới thêm lần đầu cho nhân sự đã có lịch sử Phép từ
    // trước) — "mồi" bằng 0 tại ĐÚNG tháng của sổ kia để cả 2 đồng bộ tháng trước khi đi tiếp.
    if (!lastPhep) lastPhep = { thang: lastBu.thang, phepDauThang: 0, soNgayNghi: 0 };
    if (!lastBu) lastBu = { thang: lastPhep.thang, buDauThang: 0, soNgayNghi: 0, gioTCDauThang: 0 };

    let thangDangXet = lastPhep.thang < lastBu.thang ? lastPhep.thang : lastBu.thang;
    let tongPhepThangDo = lastPhep.phepDauThang - lastPhep.soNgayNghi;
    let tongBuThangDo = lastBu.buDauThang - lastBu.soNgayNghi;
    let gioTCDauThangDangXet = lastBu.gioTCDauThang;

    while (thangDangXet < thangHienTai) {
      const chiTiet = tinhChiTietThangChoNguoi_(ccByThang_.get(thangDangXet) || [], u.id, thangDangXet, tongBuThangDo, gioTCDauThangDangXet, tongPhepThangDo);
      // Chỉ áp dụng delta tháng này cho sổ nào ĐÃ tồn tại tới đúng thangDangXet — phòng trường hợp 1
      // trong 2 sổ có dữ liệu THẬT mới hơn (hiếm khi xảy ra vì luôn "mồi" bằng nhau ở trên).
      if (lastPhep.thang <= thangDangXet) tongPhepThangDo += chiTiet.congTruPhepThang;
      if (lastBu.thang <= thangDangXet) tongBuThangDo += chiTiet.buTangThangNay;

      thangDangXet = thangCong_(thangDangXet, 1);
      const phepDauThangMoi = Math.round((tongPhepThangDo + baseline) * 10) / 10;
      const buDauThangMoi = Math.round(tongBuThangDo * 10) / 10;

      phepRowsToAppend.push([u.id, thangDangXet, phepDauThangMoi, 0, new Date()]);
      allPhep.push({ userId: u.id, thang: thangDangXet, phepDauThang: phepDauThangMoi, soNgayNghi: 0 });
      buRowsToAppend.push([u.id, thangDangXet, buDauThangMoi, 0, new Date(), chiTiet.gioTCDu]);
      allBu.push({ userId: u.id, thang: thangDangXet, buDauThang: buDauThangMoi, soNgayNghi: 0, gioTCDauThang: chiTiet.gioTCDu });

      tongPhepThangDo = phepDauThangMoi;
      tongBuThangDo = buDauThangMoi;
      gioTCDauThangDangXet = chiTiet.gioTCDu;
    }
  });

  if (phepRowsToAppend.length) {
    phepSheet.getRange(phepSheet.getLastRow() + 1, 1, phepRowsToAppend.length, PHEP_HEADERS_.length).setValues(phepRowsToAppend);
  }
  if (buRowsToAppend.length) {
    buSheet.getRange(buSheet.getLastRow() + 1, 1, buRowsToAppend.length, BU_HEADERS_.length).setValues(buRowsToAppend);
  }

  return { allPhep, allBu };
}

// Chỉ Admin (không áp dụng cho Thư ký BQLDA) được sửa trực tiếp "Phép đầu tháng" — dùng để nhập số dư
// phép sẵn có của dự án (đã theo dõi thủ công trên Google Sheet trước khi có tab này), hoặc chỉnh
// tay khi cần — sửa trực tiếp ở tab "Chấm công & Tăng ca" (ô "Phép đầu tháng").
function adminCapNhatPhepDauThang(userId, password, targetUserId, thang, phepDauThang) {
  const user = login(userId, password);
  if (user.role !== 'ADMIN') throw new Error('Chỉ Quản trị mới có quyền điều chỉnh trực tiếp "Phép đầu tháng"!');

  const targetId = String(targetUserId || '').trim();
  const thangStr = String(thang || '').trim();
  if (!targetId || !thangStr) throw new Error('Thiếu thông tin nhân sự/tháng!');

  const phepDauThangNum = parseFloat(phepDauThang);
  if (isNaN(phepDauThangNum)) throw new Error('"Phép đầu tháng" không hợp lệ!');

  const sheet = layHoacTaoSheet_('PhepThang', PHEP_HEADERS_);

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === targetId && formatThangCell_(data[i][1]) === thangStr) {
      sheet.getRange(i + 1, 3).setValue(phepDauThangNum);
      logActivity(userId, user.name, 'Sửa Phép đầu tháng', `${targetId} - ${thangStr}: ${phepDauThangNum}`);
      return 'Success';
    }
  }

  sheet.appendRow([targetId, thangStr, phepDauThangNum, 0, new Date()]);
  logActivity(userId, user.name, 'Sửa Phép đầu tháng', `${targetId} - ${thangStr}: ${phepDauThangNum}`);
  return 'Success';
}

// Tương tự adminCapNhatPhepDauThang nhưng cho sổ Bù (BuThang) — sửa trực tiếp ở tab "Chấm công &
// Tăng ca" (ô "Bù đầu tháng").
function adminCapNhatBuDauThang(userId, password, targetUserId, thang, buDauThang) {
  const user = login(userId, password);
  if (user.role !== 'ADMIN') throw new Error('Chỉ Quản trị mới có quyền điều chỉnh trực tiếp "Bù đầu tháng"!');

  const targetId = String(targetUserId || '').trim();
  const thangStr = String(thang || '').trim();
  if (!targetId || !thangStr) throw new Error('Thiếu thông tin nhân sự/tháng!');

  const buDauThangNum = parseFloat(buDauThang);
  if (isNaN(buDauThangNum)) throw new Error('"Bù đầu tháng" không hợp lệ!');

  const sheet = layHoacTaoSheet_('BuThang', BU_HEADERS_);

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === targetId && formatThangCell_(data[i][1]) === thangStr) {
      sheet.getRange(i + 1, 3).setValue(buDauThangNum);
      logActivity(userId, user.name, 'Sửa Bù đầu tháng', `${targetId} - ${thangStr}: ${buDauThangNum}`);
      return 'Success';
    }
  }

  sheet.appendRow([targetId, thangStr, buDauThangNum, 0, new Date(), 0]);
  logActivity(userId, user.name, 'Sửa Bù đầu tháng', `${targetId} - ${thangStr}: ${buDauThangNum}`);
  return 'Success';
}

// =======================================================
// NGÀY LỄ — danh sách ngày nghỉ lễ do Admin tự nhập (Quốc khánh, Tết, nghỉ bù...), dùng để tính "Số
// ngày công chuẩn dự kiến" tham khảo ở tab Chấm công (xem getChamCongThang) — KHÔNG ảnh hưởng công
// thức "Số công tính lương" của từng nhân sự.
// =======================================================

function readNgayLeRows_(sheet) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, NGAYLE_HEADERS_.length).getValues();
  const result = [];
  values.forEach(r => {
    const ngay = formatDateCell_(r[0]);
    if (!ngay) return;
    result.push({ ngay, tenLe: String(r[1] || '').trim() });
  });
  return result;
}

// Ai cũng xem được (giống getData()), không cần đăng nhập.
function getDanhSachNgayLe() {
  const sheet = layHoacTaoSheet_('NgayLe', NGAYLE_HEADERS_);
  const list = readNgayLeRows_(sheet).sort((a, b) => a.ngay.localeCompare(b.ngay));
  return JSON.stringify(list);
}

// Chỉ Admin thêm/sửa được — trùng ngày đã có thì ghi đè tên lễ (không tạo trùng dòng).
function adminThemNgayLe(userId, password, ngay, tenLe) {
  const user = login(userId, password);
  if (user.role !== 'ADMIN') throw new Error('Chỉ Quản trị mới có quyền quản lý Ngày lễ!');

  const ngayStr = formatDateCell_(ngay) || String(ngay || '').trim();
  const tenLeStr = String(tenLe || '').trim();
  if (!ngayStr) throw new Error('Thiếu ngày lễ!');
  if (!tenLeStr) throw new Error('Thiếu tên ngày lễ!');

  const sheet = layHoacTaoSheet_('NgayLe', NGAYLE_HEADERS_);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (formatDateCell_(data[i][0]) === ngayStr) {
      sheet.getRange(i + 1, 2).setValue(tenLeStr);
      logActivity(userId, user.name, 'Sửa Ngày lễ', `${ngayStr}: ${tenLeStr}`);
      return 'Success';
    }
  }

  sheet.appendRow([ngayStr, tenLeStr]);
  logActivity(userId, user.name, 'Thêm Ngày lễ', `${ngayStr}: ${tenLeStr}`);
  return 'Success';
}

// Chỉ Admin xoá được.
function adminXoaNgayLe(userId, password, ngay) {
  const user = login(userId, password);
  if (user.role !== 'ADMIN') throw new Error('Chỉ Quản trị mới có quyền quản lý Ngày lễ!');

  const ngayStr = formatDateCell_(ngay) || String(ngay || '').trim();
  const sheet = layHoacTaoSheet_('NgayLe', NGAYLE_HEADERS_);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (formatDateCell_(data[i][0]) === ngayStr) {
      sheet.deleteRow(i + 1);
      logActivity(userId, user.name, 'Xoá Ngày lễ', ngayStr);
      return 'Success';
    }
  }
  throw new Error('Không tìm thấy ngày lễ này!');
}

// =======================================================
// CHẤM CÔNG & TĂNG CA THEO THÁNG — mỗi dòng = 1 (userId, ngày), gộp cả chấm công (buoi: 1 cả ngày /
// 0.5 nửa ngày / 'P' nghỉ phép) CHO MỌI NGÀY trong tháng và tăng ca (gioBatDauTC/gioKetThucTC dạng
// 'HH:mm', kết thúc <= bắt đầu ngầm hiểu là qua ngày hôm sau) trên CÙNG 1 dòng. Đây là tab DUY NHẤT
// chấm công (tab "Điểm danh & Phép"/DiemDanhCuoiTuan cũ, chỉ chấm được Thứ 7/CN, đã bị GỘP vào đây để
// khỏi trùng 2 nơi chấm công). Ai cũng xem được (giống getData()), chỉ Admin/Thư ký BQLDA
// (assertQuyenChamCongTangCa_) chấm công/sửa tăng ca được — riêng "Phép đầu tháng"/"Bù đầu tháng" chỉ
// Admin sửa được, qua adminCapNhatPhepDauThang()/adminCapNhatBuDauThang() (định nghĩa ở phần trên).
//
// Quy tắc quy đổi giờ tăng ca: tách theo mốc 22h trong ca đã nhập — phần trước 22h hệ số 1, phần từ
// 22h hệ số 2, cộng lại thành "giờ quy đổi" (xem tinhGioTangCa_).
//
// MÔ HÌNH PHÉP/BÙ (theo đúng cách BQLDA đang làm tay trên Excel — 2 sổ TÁCH BIỆT, xem
// tinhChiTietThangChoNguoi_, dùng chung cho cả hiển thị "dự kiến" ở getChamCongThang lẫn chốt số THẬT
// ở chotSoSachDenThangHienTai_ khi sang tháng mới):
//
//   SỔ BÙ (BuThang) — nhận CẢ 2 nguồn, áp dụng cho MỌI nhân sự (kể cả Thử việc, không riêng chính thức):
//     Bù phát sinh tháng = max(0, Công Thứ 7 + Công Chủ Nhật − Định mức Thứ 7) + Công tăng ca quy đổi
//     Định mức Thứ 7 = 0,5 công × số Thứ 7 trong tháng — Chủ Nhật tính GỘP CHUNG 1 định mức với Thứ 7
//     (đi làm Chủ Nhật thay Thứ 7 vẫn coi là đủ định mức). Làm THIẾU định mức KHÔNG bị trừ Bù (Bù chỉ
//     tăng, không âm) — phần thiếu chỉ ảnh hưởng tới việc miễn trừ ngày P bên dưới, không phạt gì thêm.
//     Tăng ca -> Công (xem tinhQuyDoiTangCaCong_): dùng GIỜ QUY ĐỔI (đã nhân hệ số đêm x2) — cứ đủ 4h
//     (cộng dồn giờ dư tháng trước + giờ phát sinh tháng này) chốt 0,5 công, đủ 8h chốt 1 công; giờ lẻ
//     CHƯA đủ 4h mang sang tháng sau cộng dồn tiếp (gioTCDauThang, lưu trong BuThang).
//
//   NGÀY NGHỈ PHÉP (P) — trả theo thứ tự THÁC NƯỚC (waterfall), ưu tiên Bù trước, không còn kiểu
//   "miễn trừ toàn bộ nếu tháng có phát sinh Bù" như trước:
//     1. Trừ vào Bù CUỐI THÁNG (đầu tháng + phát sinh tháng này) trước — coi như nghỉ phép CÓ LƯƠNG.
//     2. Bù không đủ thì trừ tiếp vào Phép ĐẦU THÁNG — vẫn CÓ LƯƠNG.
//     3. Cả 2 đều hết thì phần dư là nghỉ KHÔNG LƯƠNG (không cộng vào Số công tính lương).
//
//   SỔ PHÉP (PhepThang) — CHỈ còn phụ thuộc phần P không phủ được bởi Bù (bước 2 ở trên), KHÔNG còn
//   liên quan cuối tuần/tăng ca nữa. Baseline đầu tháng (+1): CHỈ áp dụng cho nhân sự CHÍNH THỨC
//   (TRANGTHAI_HOATDONG_) — Thử việc không có. Nhân sự Nghỉ việc bị đóng băng hoàn toàn CẢ 2 SỔ, không
//   cộng/trừ gì thêm (xem chotSoSachDenThangHienTai_).
//
//   Số công tính lương (getChamCongThang) = công ngày thường + cuối tuần (chặn ở định mức) + số ngày P
//   ĐƯỢC PHỦ bởi Bù/Phép (bước 1-2 ở trên, coi như đi làm đủ) — KHÔNG cộng gì từ phần Bù/tăng ca còn
//   dư (đã ở trong sổ Bù để nghỉ bù sau, không trả lương 2 lần) hay từ phần P không phủ được (nghỉ
//   không lương).
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

// Tổng GIỜ tăng ca quy đổi (đã nhân hệ số đêm x2 qua tinhGioTangCa_) của 1 nhân sự/1 tháng — DÙNG
// CHUNG bởi getChamCongThang (hiển thị "dự kiến") và chotSoSachDenThangHienTai_ (chốt số THẬT khi sang
// tháng mới), cùng nguyên tắc với tinhChiTietThangChoNguoi_ để không lệch công thức giữa 2 nơi.
function tinhTongGioTangCaThang_(allCC, userId, thang) {
  return allCC
    .filter(r => r.userId === userId && r.ngay.slice(0, 7) === thang && r.gioBatDauTC && r.gioKetThucTC)
    .reduce((s, r) => s + tinhGioTangCa_(r.gioBatDauTC, r.gioKetThucTC).quyDoi, 0);
}

// Quy đổi giờ tăng ca -> Công theo block 4h (=0,5 công): cứ đủ 4h chốt 0,5 công, đủ 8h chốt 1 công,
// giờ lẻ chưa đủ 4h KHÔNG quy đổi mà mang sang tháng sau cộng dồn tiếp (gioTCDauThang trong PhepThang
// — xem chotSoSachDenThangHienTai_). { gioKhaDung, soCong, gioDu }.
function tinhQuyDoiTangCaCong_(gioDauThang, gioPhatSinhThang) {
  const gioKhaDung = Math.round(((Number(gioDauThang) || 0) + (Number(gioPhatSinhThang) || 0)) * 100) / 100;
  const soBlock = Math.floor(gioKhaDung / 4 + 1e-9); // +epsilon: tránh lỗi làm tròn số thực (VD 3.999999)
  const soCong = Math.round(soBlock * 0.5 * 10) / 10;
  const gioDu = Math.round((gioKhaDung - soBlock * 4) * 100) / 100;
  return { gioKhaDung, soCong, gioDu };
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
// allCC 1-2 lần (kể cả gọi qua chotSoSachDenThangHienTai_) — với N nhân sự và M dòng lịch sử, tổng chi
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

// Tính TOÀN BỘ chi tiết Phép + Bù + tăng ca cho 1 nhân sự/1 tháng — DÙNG CHUNG bởi getChamCongThang
// (hiển thị "dự kiến") và chotSoSachDenThangHienTai_ (chốt số THẬT khi sang tháng mới), để không lệch
// công thức giữa 2 nơi. gioTCDauThangBu: giờ tăng ca lẻ mang từ tháng trước (đọc từ sổ Bù, KHÔNG phải
// sổ Phép — xem BU_HEADERS_). buDauThang/phepDauThang: số dư ĐẦU THÁNG của 2 sổ (trước khi trừ ngày
// P tháng này) — dùng để chạy "thác nước" (waterfall) trả ngày nghỉ phép bên dưới.
// congCuoiTuan = congThu7 + congChuNhat: Chủ Nhật được tính GỘP CHUNG 1 định mức với Thứ 7 (làm Chủ
// Nhật thay Thứ 7 vẫn đủ định mức bình thường, không bắt buộc phải đúng ngày Thứ 7).
// congCuoiTuanDu (phần VƯỢT định mức) + congTangCaQuyDoi (tăng ca quy đổi block 4h/8h) CỘNG THÀNH
// buPhatSinhTruocKhiTruP — dồn vào sổ BÙ (không phải Phép). Làm THIẾU định mức không bị trừ ở đâu cả
// (Bù chỉ tăng, không âm).
//
// Ngày nghỉ phép (P) được TRẢ theo thứ tự thác nước (waterfall), ưu tiên Bù trước:
//   1. Trừ vào Bù CUỐI THÁNG (= buDauThang + Bù phát sinh tháng này) trước — phần P này vẫn coi như đi
//      làm đủ công (nghỉ phép có lương, chỉ trừ số dư Bù).
//   2. Bù không đủ thì trừ tiếp vào Phép ĐẦU THÁNG — vẫn coi như có lương.
//   3. Cả Bù lẫn Phép đều hết thì phần dư ra là nghỉ KHÔNG LƯƠNG — không cộng vào Số công tính lương.
// soNgayPDuocPhu (= buDungChoP + phepDungChoP) được CỘNG THẲNG vào "Số công tính lương" ở
// getChamCongThang (coi như đi làm đủ, chỉ trừ sổ Bù/Phép thay vì trừ lương).
function tinhChiTietThangChoNguoi_(allCC, userId, thang, buDauThang, gioTCDauThangBu, phepDauThang) {
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

  // congCuoiTuanThucLam: phần cuối tuần được TÍNH VÀO LƯƠNG, tối đa bằng định mức (VD 4 Thứ 7 => tối
  // đa 4 nửa = 2 công) — dùng cho "Số công tính lương", KHÔNG dùng cho "Số công thực làm" (cột đó hiện
  // đúng số đã chấm, không chặn). congCuoiTuanDu: phần VƯỢT định mức, dồn sang Bù (không cộng thêm vào
  // lương kẻo trả 2 lần); làm thiếu thì congCuoiTuanDu = 0 (không trừ gì, không phạt kép).
  const congCuoiTuanThucLam = Math.min(congCuoiTuan, quotaThu7);
  const congCuoiTuanDu = Math.round(Math.max(0, congCuoiTuan - quotaThu7) * 10) / 10;

  const tongGioTC = tinhTongGioTangCaThang_(allCC, userId, thang);
  const quyDoiTC = tinhQuyDoiTangCaCong_(gioTCDauThangBu, tongGioTC);
  const buPhatSinhTruocKhiTruP = Math.round((congCuoiTuanDu + quyDoiTC.soCong) * 10) / 10;

  // Waterfall trả ngày P: Bù cuối tháng (đầu tháng + phát sinh tháng này) trước, rồi tới Phép đầu
  // tháng — Math.max(0, ...) phòng trường hợp số dư cũ đang âm (lịch sử trước khi có quy tắc này).
  const buKhaDung = Math.round((Math.max(0, buDauThang) + buPhatSinhTruocKhiTruP) * 10) / 10;
  const buDungChoP = Math.min(soNgayP, buKhaDung);
  const pConLaiSauBu = Math.round((soNgayP - buDungChoP) * 10) / 10;
  const phepDungChoP = Math.min(pConLaiSauBu, Math.max(0, phepDauThang));
  const pKhongDuocPhu = Math.round((pConLaiSauBu - phepDungChoP) * 10) / 10;
  const soNgayPDuocPhu = Math.round((soNgayP - pKhongDuocPhu) * 10) / 10;

  const congTruPhepThang = -phepDungChoP;
  const buTangThangNay = Math.round((buPhatSinhTruocKhiTruP - buDungChoP) * 10) / 10;

  return {
    quotaThu7, congThu7, congChuNhat, congCuoiTuan, congCuoiTuanThucLam, congCuoiTuanDu,
    soNgayP, buDungChoP, phepDungChoP, pKhongDuocPhu, soNgayPDuocPhu, congTruPhepThang,
    tongGioTC: Math.round(tongGioTC * 10) / 10, gioTCDu: quyDoiTC.gioDu,
    congTangCaQuyDoi: quyDoiTC.soCong, buTangThangNay
  };
}

// thang: 'yyyy-MM', bỏ trống = tháng hiện tại. KHÔNG cần đăng nhập để xem (giống getData()).
//
// Đã đo thực tế (23 nhân sự, 317 dòng ChamCongThang, chỉ 1 tháng dữ liệu — rất nhỏ) mà vẫn mất ~3,8
// giây: phần lớn là do CÁC LƯỢT ĐỌC SHEET RIÊNG LẺ (mỗi lượt ~0,4-1 giây CỐ ĐỊNH ở nền tảng Apps
// Script, không phụ thuộc sheet nhỏ hay lớn), không phải do vòng lặp tính toán. Vì vậy tối ưu quan
// trọng nhất là GIẢM SỐ LƯỢT ĐỌC SHEET, không phải tối ưu thuật toán — hàm này đọc DanhSachUser (qua
// listSupervisors), ChamCongThang, PhepThang + BuThang (qua chotSoSachDenThangHienTai_, đã trả thẳng
// dữ liệu mới nhất, KHÔNG đọc lại 2 sheet đó lần 2 như trước).
function getChamCongThang(thang) {
  const thangXem = String(thang || '').trim() || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');

  const { soNgay, saturdays, sundays } = layThu7CNTrongThang_(thangXem);

  const allUsers = listSupervisors();
  // Đọc + nhóm theo tháng ĐÚNG 1 LẦN cho cả lượt gọi này — chotSoSachDenThangHienTai_() và vòng lặp
  // allUsers.map() bên dưới đều tra theo tháng qua Map (O(1)/tháng) thay vì mỗi nơi tự .filter() lại
  // NGUYÊN mảng toàn bộ lịch sử chấm công (xem chú thích ở nhomChamCongTheoThang_()) — có lợi khi
  // lịch sử chấm công tích lũy nhiều tháng, dù đây không phải chi phí chính hiện tại (xem trên).
  const allCC = readChamCongRows_(layHoacTaoSheet_('ChamCongThang', CHAMCONG_HEADERS_));
  const ccByThang_ = nhomChamCongTheoThang_(allCC);
  // Trả thẳng allPhep/allBu đã gồm dòng vừa chốt thêm (nếu có) — KHÔNG đọc lại 2 sheet lần 2.
  const { allPhep, allBu } = chotSoSachDenThangHienTai_(allUsers, ccByThang_);

  const ccThang = ccByThang_.get(thangXem) || [];

  const rows = allUsers.map(u => {
    const phepRow = allPhep.find(p => p.userId === u.id && p.thang === thangXem);
    const buRow = allBu.find(p => p.userId === u.id && p.thang === thangXem);
    const phepDauThang = phepRow ? phepRow.phepDauThang : 0;
    const buDauThang = buRow ? buRow.buDauThang : 0;
    const gioTCDauThang = buRow ? buRow.gioTCDauThang : 0;

    const cc = {}, tc = {};
    ccThang.filter(r => r.userId === u.id).forEach(r => {
      if (r.buoi !== '') cc[r.ngay] = r.buoi;
      if (r.gioBatDauTC && r.gioKetThucTC) tc[r.ngay] = { start: r.gioBatDauTC, end: r.gioKetThucTC };
    });

    // Không chấm công chung với BQLDA — đóng băng hoàn toàn Phép/Bù/Số công tính lương (xem
    // chotSoSachDenThangHienTai_), client hiện dấu "—" thay vì số (các field null bên dưới).
    if (u.khongChamCongChung) {
      return {
        userId: u.id, hoTen: u.name, khongChamCongChung: true,
        phepDauThang: null, buDauThang: null, cc, tc,
        congThu7: 0, congChuNhat: 0, congCuoiTuan: 0, soNgayP: 0,
        buDungChoP: 0, phepDungChoP: 0, pKhongDuocPhu: 0, congTruPhepThang: 0,
        phepCuoiThang: null, buCuoiThang: null,
        congThucLam: null, congTinhLuong: null,
        tongGioTC: 0, gioTCDauThang: null, congTangCaQuyDoi: 0, gioTCConLaiDuKien: null,
        congCuoiTuanDu: 0, buTangThangNay: 0
      };
    }

    const chiTiet = tinhChiTietThangChoNguoi_(ccThang, u.id, thangXem, buDauThang, gioTCDauThang, phepDauThang);
    const congNgayThuong = Object.entries(cc).reduce((s, [ngay, v]) => {
      if (saturdays.includes(ngay) || sundays.includes(ngay)) return s;
      return s + (typeof v === 'number' ? v : 0);
    }, 0);
    // Số công thực làm = TOÀN BỘ công đã chấm (ngày thường + Thứ 7 + Chủ Nhật), KHÔNG chặn theo định
    // mức — thể hiện đúng số ngày/buổi đã thực sự đi làm trong tháng, thuần tuý ghi nhận chấm công.
    const congThucLam = congNgayThuong + chiTiet.congCuoiTuan;
    // Số công tính lương = cơ sở tính lương thực tế: cuối tuần LÀM DƯ định mức thì CHỈ tính tới đúng
    // định mức (phần dư đã dồn sang Bù ở trên, không cộng thêm vào đây kẻo trả 2 lần); LÀM THIẾU thì
    // KHÔNG bị trừ thêm ở cột này (không phạt kép). Tăng ca KHÔNG cộng vào lương — chỉ vào Bù để nghỉ
    // bù sau (khác với chi trả tiền lương). Ngày P được Bù/Phép "phủ" (soNgayPDuocPhu, xem
    // tinhChiTietThangChoNguoi_) coi như đi làm đủ công nên CỘNG THẲNG vào đây; phần không đủ Bù/Phép
    // để phủ (pKhongDuocPhu) là nghỉ không lương, không cộng gì.
    const congTinhLuong = congNgayThuong + chiTiet.congCuoiTuanThucLam + chiTiet.soNgayPDuocPhu;

    return {
      userId: u.id, hoTen: u.name, phepDauThang, buDauThang,
      cc, tc,
      congThu7: chiTiet.congThu7,
      congChuNhat: chiTiet.congChuNhat,
      congCuoiTuan: chiTiet.congCuoiTuan,
      soNgayP: chiTiet.soNgayP,
      buDungChoP: chiTiet.buDungChoP,
      phepDungChoP: chiTiet.phepDungChoP,
      pKhongDuocPhu: chiTiet.pKhongDuocPhu,
      congTruPhepThang: chiTiet.congTruPhepThang,
      phepCuoiThang: Math.round((phepDauThang + chiTiet.congTruPhepThang) * 10) / 10,
      buCuoiThang: Math.round((buDauThang + chiTiet.buTangThangNay) * 10) / 10,
      congThucLam: Math.round(congThucLam * 10) / 10,
      congTinhLuong: Math.round(congTinhLuong * 10) / 10,
      tongGioTC: chiTiet.tongGioTC,
      gioTCDauThang: Math.round(gioTCDauThang * 10) / 10,
      congTangCaQuyDoi: chiTiet.congTangCaQuyDoi,
      gioTCConLaiDuKien: chiTiet.gioTCDu,
      congCuoiTuanDu: chiTiet.congCuoiTuanDu,
      buTangThangNay: chiTiet.buTangThangNay
    };
  });

  const quotaThu7 = Math.round(saturdays.length * 0.5 * 10) / 10;

  // Số ngày công CHUẨN dự kiến cả tháng — CHỈ là con số tham khảo hiển thị ở dòng tiêu đề bảng, KHÔNG
  // ảnh hưởng "Số công tính lương" của từng nhân sự (xem NGAYLE_HEADERS_ ở đầu file):
  //   = (số ngày thường trong tháng − số ngày lễ rơi vào ngày thường) + định mức cuối tuần (quotaThu7)
  const soNgayThuongThang = soNgay - saturdays.length - sundays.length;
  const ngayLeThang = readNgayLeRows_(layHoacTaoSheet_('NgayLe', NGAYLE_HEADERS_))
    .filter(nl => nl.ngay.slice(0, 7) === thangXem && !saturdays.includes(nl.ngay) && !sundays.includes(nl.ngay));
  const soNgayCongChuan = Math.round((soNgayThuongThang - ngayLeThang.length + quotaThu7) * 10) / 10;

  return JSON.stringify({ thang: thangXem, soNgay, saturdays, sundays, quotaThu7, soNgayCongChuan, ngayLeThang, rows });
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

// Tiện ích chạy TAY 1 LẦN từ trình soạn thảo Apps Script (không gọi từ giao diện web) để TÁCH LẠI số
// liệu tháng 7/2026 đã seed ở adminSeedPhepThang7_2026() phía trên — số cũ ở đó là "Phép + Bù" GỘP
// CHUNG (theo cách làm CŨ trước khi tách sổ Bù riêng), nay cần tách đúng thành Phép/Bù độc lập theo
// bảng chấm công Excel gốc (cột "Tổng ngày Phép"/"Tổng ngày Bù" tháng 5+6/2026 của Thư ký BQLDA) —
// nếu KHÔNG chạy hàm này, PhepThang tháng 7/2026 sẽ vẫn giữ số GỘP CŨ (sai, vì lẫn cả phần Bù vào
// Phép) và BuThang sẽ bắt đầu từ 0 (mất phần Bù thực tế đã tích lũy). Chạy lại nhiều lần vẫn an toàn
// (ghi đè đúng dòng đã có, không tạo trùng). LƯU Ý: dòng "trongtoanxd2012@gmail.com" (Lê Trọng Toàn)
// chỉ xác nhận được TỔNG (Phép+Bù=7) khớp đúng, phần tách Phép=1/Bù=6 là suy đoán tốt nhất từ ảnh chụp
// — nên kiểm tra/sửa tay lại qua ô "Bù đầu tháng" ở tab Chấm công nếu số liệu chưa đúng.
function adminTachPhepBuThang7_2026() {
  const THANG = '2026-07';
  // [email, phepDauThang, buDauThang] — tách từ DU_LIEU gộp ở adminSeedPhepThang7_2026().
  const DU_LIEU = [
    ['tranthemanh77@gmail.com', 2.5, 3], ['hoang31xd@gmail.com', 0, 0], ['trongpvin@gmail.com', 3.5, 2],
    ['hohongvan93@gmail.com', 4.5, 0.5], ['lequanghong1984@gmail.com', 0, 0], ['longsaovietpy@gmail.com', 6, 4],
    ['trongtoanxd2012@gmail.com', 1, 6], ['buiquoclamtd@gmail.com', 6, 1], ['tronghieushdk@gmail.com', 6, 1],
    ['dongphuong3112@gmail.com', 0, 0], ['congtay1610@gmail.com', 1.5, 0.5], ['vlnam0481@gmail.com', 6, 2.5],
    ['vantien.7799@gmail.com', 0, 3], ['ktshongvu@gmail.com', 2, 0], ['congchinh1603@gmail.com', 0, 0],
    ['duyhieu0983897677@gmail.com', 1, 0], ['thoaicosevcopy@gmail.com', 0, 0], ['ntht.hoaithuong.287@gmail.com', 0, 0],
    ['trungdang.const@gmail.com', 0, 0], ['bienthiaivan@gmail.com', 0, 0]
  ];

  const allUsers = listSupervisors();
  const phepSheet = layHoacTaoSheet_('PhepThang', PHEP_HEADERS_);
  const buSheet = layHoacTaoSheet_('BuThang', BU_HEADERS_);
  const phepData = phepSheet.getDataRange().getValues();
  const buData = buSheet.getDataRange().getValues();
  const boQua = [];

  DU_LIEU.forEach(([email, phepDauThang, buDauThang]) => {
    const u = allUsers.find(x => x.id.toLowerCase() === email.toLowerCase());
    if (!u) { boQua.push(email); return; }

    let foundPhep = false;
    for (let i = 1; i < phepData.length; i++) {
      if (String(phepData[i][0]).trim() === u.id && formatThangCell_(phepData[i][1]) === THANG) {
        phepSheet.getRange(i + 1, 3).setValue(phepDauThang);
        foundPhep = true;
        break;
      }
    }
    if (!foundPhep) phepSheet.appendRow([u.id, THANG, phepDauThang, 0, new Date()]);

    let foundBu = false;
    for (let i = 1; i < buData.length; i++) {
      if (String(buData[i][0]).trim() === u.id && formatThangCell_(buData[i][1]) === THANG) {
        buSheet.getRange(i + 1, 3).setValue(buDauThang);
        foundBu = true;
        break;
      }
    }
    if (!foundBu) buSheet.appendRow([u.id, THANG, buDauThang, 0, new Date(), 0]);
  });

  const ketQua = `Đã tách lại Phép/Bù đầu tháng 7/2026 cho ${DU_LIEU.length - boQua.length}/${DU_LIEU.length} người. Bỏ qua (không khớp email nào trong DanhSachUser): ${boQua.join(', ') || '(không có)'}`;
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

// hashPassword() (trên) dùng 1 "muối" (salt) CỐ ĐỊNH DÙNG CHUNG cho mọi tài khoản — 2 người trùng
// mật khẩu ra cùng 1 hash, và chỉ băm ĐÚNG 1 vòng. hashPasswordV2_ khắc phục cả 2: salt RIÊNG cho
// từng tài khoản (cột 'Salt' trong DanhSachUser, sinh ngẫu nhiên) + băm LẶP LẠI HASH_SO_VONG_ vòng
// (kiểu PBKDF2 thủ công, Apps Script không có sẵn hàm PBKDF2/bcrypt/scrypt thật) — khiến dò mật khẩu
// ngoại tuyến (nếu dữ liệu Sheet từng bị lộ) tốn CPU hơn hẳn so với chỉ 1 vòng SHA-256. Vẫn tự nâng
// cấp trong suốt (không cần người dùng làm gì) — xem login(): tài khoản nào chưa có Salt (mật khẩu
// thường HOẶC hash kiểu cũ v1) sẽ được ghi lại theo scheme mới ngay lần đăng nhập đúng tiếp theo.
const HASH_SO_VONG_ = 1000;

function taoSaltMoi_() {
  return Utilities.getUuid() + Utilities.getUuid();
}

function hashPasswordV2_(text, salt) {
  let hash = String(text).trim() + '|' + String(salt);
  for (let i = 0; i < HASH_SO_VONG_; i++) {
    const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, hash, Utilities.Charset.UTF_8);
    let hex = '';
    for (let j = 0; j < raw.length; j++) {
      let v = raw[j];
      if (v < 0) v += 256;
      hex += (v < 16 ? '0' : '') + v.toString(16);
    }
    hash = hex;
  }
  return hash;
}

// Giới hạn thử sai mật khẩu liên tục — chặn tạm 1 tài khoản sau nhiều lần đăng nhập sai liên tiếp,
// giảm rủi ro dò mật khẩu tự động (brute-force). Dùng CacheService (tự hết hạn sau LOGIN_KHOA_PHUT_
// phút, không cần dọn tay, không tốn thêm sheet) — đếm theo ĐÚNG userId đang gõ (kể cả tài khoản
// không tồn tại), vì login() được MỌI hàm backend khác gọi lại để xác thực (không chỉ nút Đăng nhập),
// nên cơ chế này tự động bảo vệ toàn bộ hệ thống, không riêng màn hình đăng nhập.
const LOGIN_SO_LAN_TOI_DA_ = 5;
const LOGIN_KHOA_PHUT_ = 15;

function loginKeyCache_(userId) {
  return 'LOGIN_FAIL_' + String(userId).trim().toLowerCase();
}

function loginKiemTraKhoa_(userId) {
  const soLan = parseInt(CacheService.getScriptCache().get(loginKeyCache_(userId)), 10) || 0;
  if (soLan >= LOGIN_SO_LAN_TOI_DA_) {
    throw new Error(`Tài khoản tạm khoá do đăng nhập sai quá ${LOGIN_SO_LAN_TOI_DA_} lần liên tiếp. Vui lòng thử lại sau ${LOGIN_KHOA_PHUT_} phút.`);
  }
}

function loginGhiNhanSai_(userId) {
  const cache = CacheService.getScriptCache();
  const key = loginKeyCache_(userId);
  const soLan = (parseInt(cache.get(key), 10) || 0) + 1;
  cache.put(key, String(soLan), LOGIN_KHOA_PHUT_ * 60);
}

function login(userId, password) {
  loginKiemTraKhoa_(userId);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DanhSachUser');
  if (!sheet) throw new Error("Chưa cấu hình Tab DanhSachUser trên Google Sheet!");

  const data = sheet.getDataRange().getValues();
  const hashedInput = hashPassword(password);
  const plainInput = String(password).trim();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(userId).trim()) {
      const storedPass = String(data[i][1]).trim();
      const storedSalt = String(data[i][7] || '').trim();

      // Có Salt riêng (cột H) -> đã ở scheme mới (hashPasswordV2_, salt riêng + băm nhiều vòng).
      // Chưa có Salt -> scheme cũ (hashPassword() 1 salt chung) HOẶC vẫn còn mật khẩu thường.
      const khop = storedSalt
        ? storedPass === hashPasswordV2_(plainInput, storedSalt)
        : (storedPass === hashedInput || storedPass === plainInput);

      if (khop) {
        const trangThai = String(data[i][6] || '').trim() || TRANGTHAI_HOATDONG_;
        if (trangThai === TRANGTHAI_NGHIVIEC_) {
          throw new Error("Tài khoản đã Nghỉ việc, không thể đăng nhập. Liên hệ Quản trị nếu có nhầm lẫn.");
        }

        // TỰ ĐỘNG BẢO MẬT: nâng cấp lên scheme mới (salt riêng + băm nhiều vòng) ngay khi chưa có Salt
        // — dù đang là mật khẩu thường hay hash kiểu cũ, người dùng không cần làm gì.
        if (!storedSalt) {
          const saltMoi = taoSaltMoi_();
          sheet.getRange(i + 1, 2).setValue(hashPasswordV2_(plainInput, saltMoi));
          sheet.getRange(i + 1, 8).setValue(saltMoi);
        }

        CacheService.getScriptCache().remove(loginKeyCache_(userId)); // đăng nhập đúng — xoá bộ đếm sai
        return {
          role: String(data[i][2]).trim().toUpperCase(),
          name: String(data[i][3]).trim(),
          supervisorId: String(data[i][0]).trim(),
          chucDanh: String(data[i][4] || '').trim(),
          phanMacDinh: String(data[i][5] || '').trim(),
          trangThai: trangThai
        };
      } else {
        loginGhiNhanSai_(userId);
        throw new Error("Sai Tên đăng nhập (ID) hoặc Mật khẩu!");
      }
    }
  }
  loginGhiNhanSai_(userId);
  throw new Error("Tài khoản không tồn tại trong hệ thống!");
}

// Dùng chung bộ đếm khoá tạm với login() (loginKiemTraKhoa_/loginGhiNhanSai_) — đây cũng là 1 điểm
// có thể bị dò mật khẩu cũ (đổi được mật khẩu = chiếm được tài khoản), cần bảo vệ như nhau.
function changeUserPassword(userId, oldPass, newPass) {
  loginKiemTraKhoa_(userId);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('DanhSachUser');
  const data = sheet.getDataRange().getValues();

  const hashedOld = hashPassword(oldPass);
  const plainOld = String(oldPass).trim();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(userId).trim()) {
      const storedPass = String(data[i][1]).trim();
      const storedSalt = String(data[i][7] || '').trim();

      const khop = storedSalt
        ? storedPass === hashPasswordV2_(plainOld, storedSalt)
        : (storedPass === hashedOld || storedPass === plainOld);

      if (khop) {
        // Mật khẩu MỚI luôn ghi theo scheme mới (salt riêng + băm nhiều vòng), bất kể mật khẩu cũ
        // đang ở scheme nào.
        const saltMoi = taoSaltMoi_();
        sheet.getRange(i + 1, 2).setValue(hashPasswordV2_(String(newPass).trim(), saltMoi));
        sheet.getRange(i + 1, 8).setValue(saltMoi);
        CacheService.getScriptCache().remove(loginKeyCache_(userId));
        logActivity(userId, data[i][3], "Đổi mật khẩu", "Người dùng tự đổi mật khẩu cá nhân");
        return "Success";
      } else {
        loginGhiNhanSai_(userId);
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

// =======================================================
// CHAT NỘI BỘ (phòng chat riêng theo nhóm nhân sự tự chọn)
// =======================================================
//
// 4 sheet: ChatPhong (danh sách phòng), ChatThanhVien (ai được ở trong phòng nào — CHỈ thành viên
// mới thấy phòng tồn tại, kể cả Admin cũng bị loại trừ nếu không được mời), ChatTinNhan (nội dung đã
// MÃ HÓA tại nơi lưu trữ — mở Sheet trực tiếp chỉ thấy chuỗi mã hoá, xem giaiMaChat_/maHoaChat_ —
// LƯU Ý: đây là mã hoá bảo vệ khỏi xem trực tiếp trong Sheet, KHÔNG phải end-to-end thực sự chống lại
// người cố tình đọc code Apps Script để lấy khoá, vì nền tảng này không có hạ tầng quản lý khoá riêng
// cho từng người dùng), ChatDaDoc (mốc thời gian mỗi người đã đọc đến đâu trong 1 phòng — dùng tính
// số tin nhắn CHƯA đọc). Tin nhắn tự xoá theo số ngày do CHỦ PHÒNG tự đặt qua trigger hàng ngày
// donDepTinNhanChatQuaHan_() (xem chatDamBaoTrigger_(), tự cài đặt khi ai đó tạo phòng đầu tiên,
// không cần Admin chạy tay 1 hàm setup riêng).
//
// Real-time: Apps Script không có WebSocket — client tự hỏi lại server theo chu kỳ (polling, xem
// 16_ChatNoiBo.html), không phải tức thời tuyệt đối.

const CHAT_PHONG_HEADERS_ = ['maPhong', 'tenPhong', 'chuPhong', 'soNgayLuu', 'ngayTao'];
const CHAT_THANHVIEN_HEADERS_ = ['maPhong', 'userId', 'ngayThamGia'];
const CHAT_TINNHAN_HEADERS_ = ['maTinNhan', 'maPhong', 'userId', 'noiDungMaHoa', 'thoiGian'];
const CHAT_DADOC_HEADERS_ = ['maPhong', 'userId', 'thoiGianDaDoc'];

// Khoá RIÊNG cho từng phòng (không phải 1 khoá dùng chung cho mọi phòng như trước) — lưu ở Script
// Properties (chỉ ai có quyền SỬA Apps Script mới đọc được, không nằm trong Sheet). Lộ khoá 1 phòng
// (nếu có) không kéo theo lộ nội dung MỌI phòng khác như thiết kế cũ.
function chatKhoaPhong_(maPhong) {
  const props = PropertiesService.getScriptProperties();
  const propKey = 'CHAT_KHOA_' + String(maPhong).trim();
  let khoa = props.getProperty(propKey);
  if (!khoa) {
    khoa = Utilities.getUuid() + Utilities.getUuid() + Utilities.getUuid();
    props.setProperty(propKey, khoa);
  }
  return khoa;
}

// Sinh "dòng khoá" (keystream) dài soByte byte bằng HMAC-SHA256 nối tiếp theo từng khối (kiểu counter
// mode) — thay cho khoá lặp lại XOR kiểu Vigenère cũ (dễ lộ quy luật lặp khi văn bản dài hơn độ dài
// khoá). HMAC là hàm mật mã THẬT có sẵn trong Apps Script (Utilities.computeHmacSha256Signature),
// không phải tự chế — nhưng bản chất vẫn là mã hoá ĐỐI XỨNG (cùng 1 khoá mã hoá lẫn giải mã), vẫn
// KHÔNG phải end-to-end thực sự, xem chú thích ở đầu khối CHAT NỘI BỘ.
function chatKeystream_(khoa, soByte) {
  const out = [];
  let counter = 0;
  while (out.length < soByte) {
    const block = Utilities.computeHmacSha256Signature(String(counter), khoa, Utilities.Charset.UTF_8);
    for (let i = 0; i < block.length && out.length < soByte; i++) out.push(block[i]);
    counter++;
  }
  return out;
}

function maHoaChat_(text, maPhong) {
  const textBytes = Utilities.newBlob(String(text == null ? '' : text)).getBytes();
  const keystream = chatKeystream_(chatKhoaPhong_(maPhong), textBytes.length);
  const out = textBytes.map((b, i) => b ^ keystream[i]);
  return Utilities.base64Encode(out);
}

function giaiMaChat_(encoded, maPhong) {
  if (!encoded) return '';
  try {
    const bytes = Utilities.base64Decode(String(encoded));
    const keystream = chatKeystream_(chatKhoaPhong_(maPhong), bytes.length);
    const out = bytes.map((b, i) => b ^ keystream[i]);
    return Utilities.newBlob(out).getDataAsString();
  } catch (e) {
    return '(Không đọc được nội dung)';
  }
}

function readChatPhongRows_(sheet) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, CHAT_PHONG_HEADERS_.length).getValues()
    .filter(r => r[0])
    .map(r => ({ maPhong: String(r[0]).trim(), tenPhong: String(r[1] || ''), chuPhong: String(r[2] || '').trim(), soNgayLuu: Number(r[3]) || 30, ngayTao: r[4] }));
}

function readChatThanhVienRows_(sheet) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, CHAT_THANHVIEN_HEADERS_.length).getValues()
    .filter(r => r[0] && r[1])
    .map(r => ({ maPhong: String(r[0]).trim(), userId: String(r[1]).trim(), ngayThamGia: r[2] }));
}

function readChatTinNhanRows_(sheet) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, CHAT_TINNHAN_HEADERS_.length).getValues()
    .filter(r => r[0] && r[1])
    .map((r, i) => ({ maTinNhan: String(r[0]).trim(), maPhong: String(r[1]).trim(), userId: String(r[2]).trim(), noiDungMaHoa: String(r[3] || ''), thoiGian: r[4] instanceof Date ? r[4] : new Date(r[4]) }));
}

function readChatDaDocRows_(sheet) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, CHAT_DADOC_HEADERS_.length).getValues()
    .filter(r => r[0] && r[1])
    .map((r, i) => ({ maPhong: String(r[0]).trim(), userId: String(r[1]).trim(), thoiGianDaDoc: r[2] instanceof Date ? r[2] : new Date(r[2]), rowIdx: i + 2 }));
}

// Idempotent — tự kiểm tra đã có trigger hàng ngày dọn tin nhắn quá hạn chưa, chưa có thì tạo (chạy
// lúc 2h sáng, ít ảnh hưởng người dùng). Gọi mỗi khi tạo phòng mới thay vì bắt Admin chạy tay 1 hàm
// setup riêng — rẻ (chỉ đọc danh sách trigger, không tốn gì nếu đã tồn tại).
function chatDamBaoTrigger_() {
  const daCo = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'donDepTinNhanChatQuaHan_');
  if (!daCo) {
    ScriptApp.newTrigger('donDepTinNhanChatQuaHan_').timeBased().everyDays(1).atHour(2).create();
  }
}

// Hàm PUBLIC (không có dấu "_" cuối tên) để CHẠY TAY 1 LẦN trong Apps Script Editor (chọn hàm này ở
// dropdown "Run" trên thanh công cụ rồi bấm Run) — xin quyền quản lý Trigger cho Chat nội bộ. Các hàm
// kết thúc bằng "_" (chatDamBaoTrigger_, donDepTinNhanChatQuaHan_...) KHÔNG hiện trong dropdown đó vì
// Apps Script coi đó là quy ước đặt tên hàm nội bộ/riêng tư — đây là hàm "vỏ" public để chạy được.
function thietLapChatNoiBo() {
  chatDamBaoTrigger_();
  return 'Đã thiết lập xong: trigger tự xoá tin nhắn Chat nội bộ quá hạn (chạy hàng ngày lúc 2h sáng).';
}

function chatXacThucThanhVien_(maPhong, userId, allThanhVien) {
  if (!allThanhVien.some(tv => tv.maPhong === maPhong && tv.userId === userId)) {
    throw new Error("Bạn không có quyền xem/thao tác phòng chat này!");
  }
}

// thanhVienIds: mảng userId được mời (không cần tự thêm userId của người tạo, hàm tự thêm).
function taoPhongChat(userId, password, tenPhong, thanhVienIds, soNgayLuu) {
  const user = login(userId, password);
  const ten = String(tenPhong || '').trim();
  if (!ten) throw new Error("Vui lòng nhập tên phòng chat!");

  let soNgay = Math.round(Number(soNgayLuu));
  if (!soNgay || soNgay < 1) soNgay = 30;
  if (soNgay > 365) soNgay = 365;

  const phongSheet = layHoacTaoSheet_('ChatPhong', CHAT_PHONG_HEADERS_);
  const tvSheet = layHoacTaoSheet_('ChatThanhVien', CHAT_THANHVIEN_HEADERS_);

  const maPhong = Utilities.getUuid();
  const now = new Date();
  phongSheet.appendRow([maPhong, ten, userId, soNgay, now]);

  const memberSet = new Set((Array.isArray(thanhVienIds) ? thanhVienIds : []).map(id => String(id).trim()).filter(Boolean));
  memberSet.add(String(userId).trim());
  const rows = Array.from(memberSet).map(uid => [maPhong, uid, now]);
  tvSheet.getRange(tvSheet.getLastRow() + 1, 1, rows.length, CHAT_THANHVIEN_HEADERS_.length).setValues(rows);

  chatDamBaoTrigger_();
  logActivity(userId, user.name, "Tạo phòng chat", ten);
  return maPhong;
}

// Danh sách phòng CỦA CHÍNH NGƯỜI GỌI (chỉ phòng mình là thành viên) + tin nhắn cuối + số chưa đọc.
function getDanhSachPhongChat(userId, password) {
  login(userId, password);
  const uid = String(userId).trim();

  const allPhong = readChatPhongRows_(layHoacTaoSheet_('ChatPhong', CHAT_PHONG_HEADERS_));
  const allThanhVien = readChatThanhVienRows_(layHoacTaoSheet_('ChatThanhVien', CHAT_THANHVIEN_HEADERS_));
  const allTinNhan = readChatTinNhanRows_(layHoacTaoSheet_('ChatTinNhan', CHAT_TINNHAN_HEADERS_));
  const allDaDoc = readChatDaDocRows_(layHoacTaoSheet_('ChatDaDoc', CHAT_DADOC_HEADERS_));

  const maPhongCuaToi = new Set(allThanhVien.filter(tv => tv.userId === uid).map(tv => tv.maPhong));
  const tinNhanTheoPhong_ = new Map();
  allTinNhan.forEach(tn => {
    if (!tinNhanTheoPhong_.has(tn.maPhong)) tinNhanTheoPhong_.set(tn.maPhong, []);
    tinNhanTheoPhong_.get(tn.maPhong).push(tn);
  });
  const daDocMap_ = new Map(); // "maPhong|userId" -> Date
  allDaDoc.forEach(d => daDocMap_.set(d.maPhong + '|' + d.userId, d.thoiGianDaDoc));

  const rows = allPhong.filter(p => maPhongCuaToi.has(p.maPhong)).map(p => {
    const tinPhong = (tinNhanTheoPhong_.get(p.maPhong) || []).slice().sort((a, b) => a.thoiGian - b.thoiGian);
    const tinCuoi = tinPhong[tinPhong.length - 1];
    const daDocLuc = daDocMap_.get(p.maPhong + '|' + uid) || null;
    const soChuaDoc = tinPhong.filter(tn => tn.userId !== uid && (!daDocLuc || tn.thoiGian > daDocLuc)).length;
    const soThanhVien = allThanhVien.filter(tv => tv.maPhong === p.maPhong).length;

    return {
      maPhong: p.maPhong, tenPhong: p.tenPhong, chuPhong: p.chuPhong, laChuPhong: p.chuPhong === uid,
      soNgayLuu: p.soNgayLuu, soThanhVien,
      tinCuoi: tinCuoi ? { noiDung: giaiMaChat_(tinCuoi.noiDungMaHoa, p.maPhong), userId: tinCuoi.userId, thoiGian: tinCuoi.thoiGian } : null,
      thoiGianHoatDong: tinCuoi ? tinCuoi.thoiGian : p.ngayTao,
      soChuaDoc
    };
  }).sort((a, b) => new Date(b.thoiGianHoatDong) - new Date(a.thoiGianHoatDong));

  return JSON.stringify(rows);
}

// Mở 1 phòng: xác thực thành viên, trả tin nhắn đã giải mã + đánh dấu đã đọc đến hiện tại.
function getTinNhanPhongChat(userId, password, maPhong) {
  login(userId, password);
  const uid = String(userId).trim();
  const maP = String(maPhong || '').trim();

  const allThanhVien = readChatThanhVienRows_(layHoacTaoSheet_('ChatThanhVien', CHAT_THANHVIEN_HEADERS_));
  chatXacThucThanhVien_(maP, uid, allThanhVien);

  const allPhong = readChatPhongRows_(layHoacTaoSheet_('ChatPhong', CHAT_PHONG_HEADERS_));
  const phong = allPhong.find(p => p.maPhong === maP);
  if (!phong) throw new Error("Phòng chat không tồn tại (có thể đã bị xoá)!");

  const allTinNhan = readChatTinNhanRows_(layHoacTaoSheet_('ChatTinNhan', CHAT_TINNHAN_HEADERS_));
  const allUsers = listSupervisors();
  const tenTheoId_ = {};
  allUsers.forEach(u => { tenTheoId_[u.id] = u.name; });

  const tinNhan = allTinNhan.filter(tn => tn.maPhong === maP)
    .sort((a, b) => a.thoiGian - b.thoiGian)
    .map(tn => ({ userId: tn.userId, tenNguoiGui: tenTheoId_[tn.userId] || tn.userId, noiDung: giaiMaChat_(tn.noiDungMaHoa, maP), thoiGian: tn.thoiGian }));

  const thanhVienPhong = allThanhVien.filter(tv => tv.maPhong === maP).map(tv => tenTheoId_[tv.userId] || tv.userId);

  // Đánh dấu đã đọc đến hiện tại — upsert 1 dòng ChatDaDoc.
  const daDocSheet = layHoacTaoSheet_('ChatDaDoc', CHAT_DADOC_HEADERS_);
  const allDaDoc = readChatDaDocRows_(daDocSheet);
  const hienTai = new Date();
  const dong = allDaDoc.find(d => d.maPhong === maP && d.userId === uid);
  if (dong) daDocSheet.getRange(dong.rowIdx, 3).setValue(hienTai);
  else daDocSheet.appendRow([maP, uid, hienTai]);

  return JSON.stringify({
    maPhong: maP, tenPhong: phong.tenPhong, chuPhong: phong.chuPhong, laChuPhong: phong.chuPhong === uid,
    soNgayLuu: phong.soNgayLuu, thanhVien: thanhVienPhong, tinNhan
  });
}

function guiTinNhanChat(userId, password, maPhong, noiDung) {
  login(userId, password);
  const uid = String(userId).trim();
  const maP = String(maPhong || '').trim();
  const noi = String(noiDung || '').trim();
  if (!noi) throw new Error("Nội dung tin nhắn không được để trống!");
  if (noi.length > 2000) throw new Error("Tin nhắn quá dài (tối đa 2000 ký tự)!");

  const allThanhVien = readChatThanhVienRows_(layHoacTaoSheet_('ChatThanhVien', CHAT_THANHVIEN_HEADERS_));
  chatXacThucThanhVien_(maP, uid, allThanhVien);

  const tnSheet = layHoacTaoSheet_('ChatTinNhan', CHAT_TINNHAN_HEADERS_);
  tnSheet.appendRow([Utilities.getUuid(), maP, uid, maHoaChat_(noi, maP), new Date()]);
  return "Success";
}

function capNhatSoNgayLuuPhongChat(userId, password, maPhong, soNgayLuu) {
  login(userId, password);
  const uid = String(userId).trim();
  const maP = String(maPhong || '').trim();
  let soNgay = Math.round(Number(soNgayLuu));
  if (!soNgay || soNgay < 1) soNgay = 30;
  if (soNgay > 365) soNgay = 365;

  const phongSheet = layHoacTaoSheet_('ChatPhong', CHAT_PHONG_HEADERS_);
  const data = phongSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === maP) {
      if (String(data[i][2]).trim() !== uid) throw new Error("Chỉ chủ phòng mới đổi được số ngày lưu trữ!");
      phongSheet.getRange(i + 1, 4).setValue(soNgay);
      return "Success";
    }
  }
  throw new Error("Phòng chat không tồn tại!");
}

// Thành viên thường rời phòng — chủ phòng KHÔNG rời được (phải Xóa phòng nếu muốn kết thúc hẳn),
// tránh tình trạng phòng còn thành viên mà mất chủ, không ai đổi được số ngày lưu trữ/xoá phòng nữa.
function roiPhongChat(userId, password, maPhong) {
  login(userId, password);
  const uid = String(userId).trim();
  const maP = String(maPhong || '').trim();

  const phongSheet = layHoacTaoSheet_('ChatPhong', CHAT_PHONG_HEADERS_);
  const phong = readChatPhongRows_(phongSheet).find(p => p.maPhong === maP);
  if (phong && phong.chuPhong === uid) throw new Error("Chủ phòng không thể rời phòng — hãy Xóa phòng nếu muốn kết thúc.");

  const tvSheet = layHoacTaoSheet_('ChatThanhVien', CHAT_THANHVIEN_HEADERS_);
  const data = tvSheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim() === maP && String(data[i][1]).trim() === uid) {
      tvSheet.deleteRow(i + 1);
      return "Success";
    }
  }
  throw new Error("Bạn không ở trong phòng chat này!");
}

// Chủ phòng xoá hẳn phòng — dọn sạch cả 4 sheet liên quan (Phòng/Thành viên/Tin nhắn/Đã đọc).
function xoaPhongChat(userId, password, maPhong) {
  login(userId, password);
  const uid = String(userId).trim();
  const maP = String(maPhong || '').trim();

  const phongSheet = layHoacTaoSheet_('ChatPhong', CHAT_PHONG_HEADERS_);
  const phongData = phongSheet.getDataRange().getValues();
  let found = false;
  for (let i = phongData.length - 1; i >= 1; i--) {
    if (String(phongData[i][0]).trim() === maP) {
      if (String(phongData[i][2]).trim() !== uid) throw new Error("Chỉ chủ phòng mới xoá được phòng chat!");
      phongSheet.deleteRow(i + 1);
      found = true;
      break;
    }
  }
  if (!found) throw new Error("Phòng chat không tồn tại!");

  [['ChatThanhVien', CHAT_THANHVIEN_HEADERS_], ['ChatTinNhan', CHAT_TINNHAN_HEADERS_], ['ChatDaDoc', CHAT_DADOC_HEADERS_]].forEach(([tenSheet, headers]) => {
    const sh = layHoacTaoSheet_(tenSheet, headers);
    const data = sh.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]).trim() === maP) sh.deleteRow(i + 1);
    }
  });
  return "Success";
}

// Trigger hàng ngày (xem chatDamBaoTrigger_()) — xoá tin nhắn quá hạn theo số ngày CHỦ PHÒNG tự đặt
// cho từng phòng. Đọc hết 1 lần, tính danh sách còn giữ lại, ghi đè nguyên vùng dữ liệu 1 lượt thay
// vì xoá từng dòng lẻ tẻ (rẻ hơn khi số dòng cần xoá lớn).
function donDepTinNhanChatQuaHan_() {
  const phongSheet = layHoacTaoSheet_('ChatPhong', CHAT_PHONG_HEADERS_);
  const tnSheet = layHoacTaoSheet_('ChatTinNhan', CHAT_TINNHAN_HEADERS_);

  const soNgayTheoPhong_ = {};
  readChatPhongRows_(phongSheet).forEach(p => { soNgayTheoPhong_[p.maPhong] = p.soNgayLuu; });

  const allTinNhan = readChatTinNhanRows_(tnSheet);
  if (!allTinNhan.length) return;

  const now = new Date();
  const conGiu = allTinNhan.filter(tn => {
    const soNgay = soNgayTheoPhong_[tn.maPhong];
    if (soNgay == null) return false; // phòng đã bị xoá — dọn theo luôn
    const hanCuoi = new Date(tn.thoiGian.getTime() + soNgay * 24 * 60 * 60 * 1000);
    return hanCuoi > now;
  });

  if (conGiu.length === allTinNhan.length) return; // không có gì quá hạn

  const lastRow = tnSheet.getLastRow();
  if (lastRow >= 2) tnSheet.getRange(2, 1, lastRow - 1, CHAT_TINNHAN_HEADERS_.length).clearContent();
  if (conGiu.length) {
    const rows = conGiu.map(tn => [tn.maTinNhan, tn.maPhong, tn.userId, tn.noiDungMaHoa, tn.thoiGian]);
    tnSheet.getRange(2, 1, rows.length, CHAT_TINNHAN_HEADERS_.length).setValues(rows);
  }
}
