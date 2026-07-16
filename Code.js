// Hàm include() cho phép Index.html nhúng Style.html và các file Script.html
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// HangMuc nay CHỈ là dữ liệu tham chiếu (mã/tên/cây cha-con) để lọc & để Gói thầu gộp —
// không còn nhận Nhật ký tiến độ trực tiếp (xem CongViec bên dưới).
const HANGMUC_HEADERS_ = ['maHangMuc', 'tenHangMuc', 'capDo', 'parentId', 'active'];
const GOITHAU_HEADERS_ = ['maGoiThau', 'tenGoiThau', 'maHangMucList'];
// CongViec: đơn vị công việc thật do Admin/Giám sát tự tạo — đây mới là nơi Nhật ký tiến độ trỏ vào.
// phanLoai ('ho_so'/'thi_cong') tự gán = phanMacDinh của người tạo tại thời điểm tạo, không sửa tay.
// nguoiGiao_id/nguoiGiao_ten CHỈ có giá trị khi Admin tạo Công việc và giao cho người KHÁC (khác
// nguoiTao_id — người phụ trách); rỗng nghĩa là tự tạo/tự đảm nhận, không ai giao.
const CONGVIEC_HEADERS_ = ['maCongViec', 'tenCongViec', 'maGoiThau', 'maHangMucList', 'phanLoai',
  'nguoiTao_id', 'nguoiTao_ten', 'ngayBatDauKH', 'ngayKetThucKH', 'active', 'nguoiGiao_id', 'nguoiGiao_ten'];
// ThongBao: thông báo 1 chiều gửi cho 1 người (được giao việc / có người bình luận vào việc mình
// phụ trách). BinhLuanCongViec: bình luận 2 chiều append-only theo từng Công việc.
const THONGBAO_HEADERS_ = ['id', 'userId', 'loai', 'noiDung', 'maCongViec', 'nguoiGui_id', 'nguoiGui_ten', 'thoiGian', 'daDoc', 'active'];
const BINHLUAN_HEADERS_ = ['id', 'maCongViec', 'nguoiBinhLuan_id', 'nguoiBinhLuan_ten', 'noiDung', 'thoiGian', 'active'];
const NHATKY_HEADERS_ = ['logId', 'maCongViec', 'ngayBaoCao', 'phanTramNgay', 'ghiChu',
  'nguoiNhap_id', 'nguoiNhap_ten', 'thoiGianNhap', 'fileDinhKem', 'active'];
// DiemDanhCuoiTuan: trạng thái điểm danh Thứ 7/CN hiện tại (upsert theo userId+ngay, KHÔNG phải
// log lịch sử — bỏ tick thì xoá hẳn dòng). PhepThang: sổ phép/tháng, "Cộng phép tháng"/"TỔNG CỘNG"
// luôn tính lại từ 2 sheet này, không lưu cột riêng (xem getDiemDanhPhep()).
const DIEMDANH_HEADERS_ = ['userId', 'ngay', 'buoi'];
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
const BOQ_HEADERS_ = ['maBOQ', 'maHopDong', 'stt', 'tenHangMuc', 'isHeader', 'donVi',
  'khoiLuongHopDong', 'donGia', 'cap', 'ngayBatDauKH', 'ngayKetThucKH', 'maPhuLucTao', 'active'];
const PHULUC_HEADERS_ = ['maPhuLuc', 'maHopDong', 'soHieu', 'ngayPhuLuc', 'ghiChu', 'active'];
// loaiThayDoi: MOI (thêm hạng mục mới, không cần khối lượng/giá riêng ở đây) / THAY_THE (đổi hẳn
// khối lượng hiệu lực = khoiLuongMoi) / DIEU_CHINH (cộng/trừ chênh lệch = khoiLuongDieuChinh).
const PHULUC_THAYDOI_HEADERS_ = ['id', 'maPhuLuc', 'maBOQ', 'loaiThayDoi', 'khoiLuongMoi',
  'khoiLuongDieuChinh', 'donGiaMoi', 'active'];
// NghiemThu: log khối lượng thi công đã xác nhận theo ngày, append-only cộng dồn theo maBOQ —
// giống hệt NhatKyTienDo (xem themNghiemThu()).
const NGHIEMTHU_HEADERS_ = ['maNghiemThu', 'maBOQ', 'ngayNghiemThu', 'khoiLuong', 'ghiChu',
  'nguoiNhap_id', 'nguoiNhap_ten', 'thoiGianNhap', 'active'];
const DOTTHANHTOAN_HEADERS_ = ['maDotThanhToan', 'maHopDong', 'tenDot', 'ngayThanhToan',
  'phanTramThanhToan', 'ghiChu', 'thuHoiTamUngHopDong', 'thuHoiTamUngThiCong', 'khauTruKhac',
  'ghiChuKhauTru', 'active'];
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

  if (!ss.getSheetByName('DiemDanhCuoiTuan')) {
    ss.insertSheet('DiemDanhCuoiTuan').getRange(1, 1, 1, DIEMDANH_HEADERS_.length).setValues([DIEMDANH_HEADERS_]);
  }

  if (!ss.getSheetByName('PhepThang')) {
    ss.insertSheet('PhepThang').getRange(1, 1, 1, PHEP_HEADERS_.length).setValues([PHEP_HEADERS_]);
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

  return 'Đã thiết lập xong: Data, DanhSachUser (admin/123456), NhatKy, HangMuc, GoiThau, CongViec, NhatKyTienDo, DiemDanhCuoiTuan, PhepThang, ThongBao, BinhLuanCongViec, HopDong, BOQHangMuc, PhuLucHopDong, PhuLucThayDoi, NghiemThu, DotThanhToan, DotThanhToanChiTiet, QuyetToan.';
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
      parentId: String(r[3] || '').trim()
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
      maHangMucList: String(r[2] || '').split(',').map(s => s.trim()).filter(Boolean)
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
// trangThaiMau là 1 trong 5 giá trị, xét theo hạn ngayKetThucKH (KHÔNG nội suy % kế hoạch theo
// thời gian đã trôi qua như trước):
// - pending    : chưa hoàn thành và chưa tới ngayBatDauKH (chưa tới thời điểm bắt đầu)
// - inprogress : chưa hoàn thành (luyKe<100), đã tới ngày bắt đầu và còn trong hạn
// - done       : đã hoàn thành (luyKe>=100) đúng hạn (hoặc chưa đặt hạn để so)
// - late_done  : đã hoàn thành nhưng ngày hoàn thành thực tế (log đưa lũy kế chạm mốc 100%
//                đầu tiên) trễ hơn ngayKetThucKH
// - overdue    : chưa hoàn thành và hôm nay đã qua ngayKetThucKH
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
      cv.trangThaiMau = (cv.ngayKetThucKH && ngayHoanThanhThucTe && ngayHoanThanhThucTe > cv.ngayKetThucKH) ? 'late_done' : 'done';
    } else if (cv.ngayBatDauKH && today < cv.ngayBatDauKH) {
      cv.trangThaiMau = 'pending';
    } else {
      cv.trangThaiMau = (cv.ngayKetThucKH && today > cv.ngayKetThucKH) ? 'overdue' : 'inprogress';
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
// phanMacDinh: 'ho_so' hoặc 'thi_cong' — quyết định Công việc do người này tự tạo rơi vào phần nào.
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

  const row = [maHangMuc, node.tenHangMuc || '', parseInt(node.capDo, 10) || 1, node.parentId || '', true];

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

  if (maGoiThau) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === maGoiThau) {
        sheet.getRange(i + 1, 2, 1, 2).setValues([[tenGoiThau, maHangMucList]]);
        logActivity(userId, user.name, 'Sửa Gói thầu', tenGoiThau);
        return 'Success';
      }
    }
  }

  let maxSTT = 0;
  for (let i = 1; i < data.length; i++) {
    const stt = parseInt(String(data[i][0]).replace('GT', ''), 10);
    if (!isNaN(stt) && stt > maxSTT) maxSTT = stt;
  }
  maGoiThau = 'GT' + (maxSTT + 1);
  sheet.appendRow([maGoiThau, tenGoiThau, maHangMucList]);
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

  const data = sheet.getDataRange().getValues();
  const maCongViec = String(congViec.maCongViec || '').trim();

  if (maCongViec) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === maCongViec) {
        if (user.role !== 'ADMIN' && String(data[i][5]).trim() !== user.supervisorId) {
          throw new Error('Bạn không có quyền sửa Công việc này!');
        }
        // Giữ nguyên phanLoai (cột E), người phụ trách (cột F/G) và người giao (cột K/L) — chỉ cho
        // sửa các field mô tả, không cho đổi tay người phụ trách sau khi đã tạo.
        sheet.getRange(i + 1, 2, 1, 4).setValues([[tenCongViec, maGoiThau, maHangMucJoined, data[i][4]]]);
        sheet.getRange(i + 1, 8, 1, 2).setValues([[congViec.ngayBatDauKH || '', congViec.ngayKetThucKH || '']]);
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

  const now = new Date();
  const newId = sinhMaCongViec_(sheet, now);
  sheet.appendRow([
    newId, tenCongViec, maGoiThau, maHangMucJoined, phanMacDinh,
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

function adminDeleteCongViec(userId, password, maCongViec) {
  const user = login(userId, password);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CongViec');
  if (!sheet) throw new Error('Chưa cấu hình Tab CongViec trên Google Sheet!');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(maCongViec).trim()) {
      if (user.role !== 'ADMIN' && String(data[i][5]).trim() !== user.supervisorId) {
        throw new Error('Bạn không có quyền xoá Công việc này!');
      }
      sheet.getRange(i + 1, 10).setValue(false);
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
      thoiGian: formatDateTimeCell_(r[5])
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

function themBinhLuanCongViec(userId, password, maCongViec, noiDung) {
  const user = login(userId, password);
  const noiDungTrim = String(noiDung || '').trim();
  if (!noiDungTrim) throw new Error('Vui lòng nhập nội dung bình luận!');

  const cvSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CongViec');
  if (!cvSheet) throw new Error('Chưa cấu hình Tab CongViec trên Google Sheet!');
  const cv = readCongViecRows_(cvSheet, true).find(c => c.maCongViec === String(maCongViec || '').trim());
  if (!cv) throw new Error('Không tìm thấy Công việc!');
  if (user.role !== 'ADMIN' && cv.nguoiTaoId !== user.supervisorId) {
    throw new Error('Bạn không có quyền bình luận vào Công việc này!');
  }

  const blSheet = layHoacTaoSheet_('BinhLuanCongViec', BINHLUAN_HEADERS_);
  const id = sinhMaTuTang_(blSheet, 'BL', new Date());
  blSheet.appendRow([id, cv.maCongViec, user.supervisorId, user.name, noiDungTrim, new Date(), true]);
  logActivity(userId, user.name, 'Bình luận Công việc', `${cv.maCongViec}: ${noiDungTrim}`);

  // Báo cho mọi người từng liên quan tới Công việc này (người phụ trách, người giao — nếu có, và
  // mọi người đã từng bình luận), trừ chính người vừa bình luận.
  const nguoiLienQuan = new Set([cv.nguoiTaoId, cv.nguoiGiaoId].filter(Boolean));
  readBinhLuanRows_(blSheet).forEach(c => { if (c.maCongViec === cv.maCongViec) nguoiLienQuan.add(c.nguoiBinhLuanId); });
  nguoiLienQuan.delete(user.supervisorId);
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
function validateVaGhiNhatKy_(user, maCongViec, ngayBaoCao, phanTram, ghiChu, fileDinhKemArr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cvSheet = ss.getSheetByName('CongViec');
  if (!cvSheet) throw new Error('Chưa cấu hình Tab CongViec trên Google Sheet!');
  const congViec = readCongViecRows_(cvSheet);
  const node = congViec.find(c => c.maCongViec === maCongViec);
  if (!node) throw new Error('Không tìm thấy Công việc!');

  const phanTramNum = parseFloat(phanTram);
  if (isNaN(phanTramNum)) throw new Error('% không hợp lệ!');

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
    logId, maCongViec, ngayBaoCao || '', phanTramNum, ghiChu || '',
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
// ĐIỂM DANH CUỐI TUẦN & SỔ PHÉP — chỉ Admin/Thư ký BQLDA sửa được, ai cũng xem được (giống
// getData() không cần đăng nhập). "Cộng phép tháng"/"TỔNG CỘNG" luôn tính lại từ DiemDanhCuoiTuan +
// PhepThang mỗi lần đọc, không lưu cột riêng.
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

function assertQuyenDiemDanh_(user) {
  if (user.role !== 'ADMIN' && chuanHoaChucDanh_(user.chucDanh) !== THU_KY_CHUC_DANH_) {
    throw new Error('Chỉ Quản trị hoặc Thư ký BQLDA mới có quyền cập nhật Điểm danh/Phép!');
  }
}

function readDiemDanhRows_(sheet) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, DIEMDANH_HEADERS_.length).getValues();
  const result = [];
  values.forEach(r => {
    const userId = String(r[0] || '').trim();
    const ngay = formatDateCell_(r[1]);
    if (!userId || !ngay) return;
    result.push({ userId, ngay, buoi: Number(r[2]) || 0 });
  });
  return result;
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

// Gộp các ngày Thứ 7/CN liền kề trong tháng thành từng nhóm "Tuần" (kể cả nhóm mồ côi chỉ có 1
// ngày ở đầu/cuối tháng, khi Thứ 7 hoặc CN rơi sang tháng khác).
function layDanhSachThu7CN_(thang) {
  const parts = String(thang).split('-');
  const year = parseInt(parts[0], 10), month1 = parseInt(parts[1], 10);
  const soNgay = new Date(year, month1, 0).getDate();
  const ddmm = iso => iso.slice(8, 10) + '/' + iso.slice(5, 7);

  const groups = [];
  let openGroup = null;
  for (let day = 1; day <= soNgay; day++) {
    const d = new Date(year, month1 - 1, day);
    const dow = d.getDay(); // 0 = Chủ nhật, 6 = Thứ 7
    if (dow !== 0 && dow !== 6) continue;
    const iso = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (dow === 6) {
      openGroup = { ngayThu7: iso, ngayCN: null };
      groups.push(openGroup);
    } else if (openGroup && !openGroup.ngayCN) {
      openGroup.ngayCN = iso;
      openGroup = null;
    } else {
      groups.push({ ngayThu7: null, ngayCN: iso });
      openGroup = null;
    }
  }

  return groups.map((g, idx) => ({
    tuan: idx + 1,
    ngayThu7: g.ngayThu7,
    ngayCN: g.ngayCN,
    nhan: 'Tuần ' + (idx + 1) + ' (' + [g.ngayThu7, g.ngayCN].filter(Boolean).map(ddmm).join('-') + ')'
  }));
}

function tongCongPhepThangTuDiemDanh_(allDiemDanh, userId, thang) {
  return Math.round(allDiemDanh
    .filter(d => d.userId === userId && d.ngay.slice(0, 7) === thang)
    .reduce((s, d) => s + d.buoi, 0) * 10) / 10;
}

// 'yyyy-MM' + số tháng cộng thêm (có thể âm) -> 'yyyy-MM' mới, tự xử lý qua năm.
function thangCong_(thang, soThang) {
  const parts = String(thang).split('-');
  const d = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1 + soThang, 1);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM');
}

// Chốt ngầm mỗi khi có người mở tab: với mỗi nhân sự, tạo tuần tự các dòng PhepThang còn thiếu cho
// tới tháng hiện tại thật (không nhảy cóc nếu app bị bỏ quên nhiều tháng liền), theo công thức:
// TổngCộng(tháng cũ) = phepDauThang + CộngPhépThángTừĐiểmDanh - soNgayNghi;
// phepDauThang(tháng mới) = TổngCộng(tháng cũ) + 1 (1 phép baseline/tháng).
function chotPhepDenThangHienTai_() {
  const phepSheet = layHoacTaoSheet_('PhepThang', PHEP_HEADERS_);
  const ddSheet = layHoacTaoSheet_('DiemDanhCuoiTuan', DIEMDANH_HEADERS_);

  const thangHienTai = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  const allUsers = listSupervisors();
  const allPhep = readPhepThangRows_(phepSheet);
  const allDiemDanh = readDiemDanhRows_(ddSheet);

  const rowsToAppend = [];
  allUsers.forEach(u => {
    const phepOfUser = allPhep.filter(p => p.userId === u.id).sort((a, b) => a.thang.localeCompare(b.thang));
    const lastRow = phepOfUser.length ? phepOfUser[phepOfUser.length - 1] : null;

    if (!lastRow) {
      // Chưa từng có dòng nào — chỉ bắt đầu từ tháng hiện tại, không truy hồi lùi các tháng trước.
      rowsToAppend.push([u.id, thangHienTai, 1, 0, new Date()]);
      return;
    }

    let thangDangXet = lastRow.thang;
    let tongCongThangDo = lastRow.phepDauThang
      + tongCongPhepThangTuDiemDanh_(allDiemDanh, u.id, thangDangXet)
      - lastRow.soNgayNghi;

    while (thangDangXet < thangHienTai) {
      thangDangXet = thangCong_(thangDangXet, 1);
      const phepDauThangMoi = Math.round((tongCongThangDo + 1) * 10) / 10;
      rowsToAppend.push([u.id, thangDangXet, phepDauThangMoi, 0, new Date()]);
      tongCongThangDo = phepDauThangMoi + tongCongPhepThangTuDiemDanh_(allDiemDanh, u.id, thangDangXet);
    }
  });

  if (rowsToAppend.length) {
    phepSheet.getRange(phepSheet.getLastRow() + 1, 1, rowsToAppend.length, PHEP_HEADERS_.length).setValues(rowsToAppend);
  }
}

// thang: 'yyyy-MM', bỏ trống = tháng hiện tại. KHÔNG cần đăng nhập để xem (giống getData()).
function getDiemDanhPhep(thang) {
  const thangXem = String(thang || '').trim() || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');

  chotPhepDenThangHienTai_();

  const allUsers = listSupervisors();
  const allPhep = readPhepThangRows_(layHoacTaoSheet_('PhepThang', PHEP_HEADERS_));
  const allDiemDanh = readDiemDanhRows_(layHoacTaoSheet_('DiemDanhCuoiTuan', DIEMDANH_HEADERS_));

  const rows = allUsers.map(u => {
    const phepRow = allPhep.find(p => p.userId === u.id && p.thang === thangXem);
    const phepDauThang = phepRow ? phepRow.phepDauThang : 0;
    const soNgayNghi = phepRow ? phepRow.soNgayNghi : 0;
    const congPhepThang = tongCongPhepThangTuDiemDanh_(allDiemDanh, u.id, thangXem);
    const diemDanh = {};
    allDiemDanh.filter(d => d.userId === u.id && d.ngay.slice(0, 7) === thangXem)
      .forEach(d => { diemDanh[d.ngay] = d.buoi; });

    return {
      userId: u.id,
      hoTen: u.name,
      phepDauThang,
      soNgayNghi,
      congPhepThang,
      tongCong: Math.round((phepDauThang + congPhepThang - soNgayNghi) * 10) / 10,
      diemDanh
    };
  });

  return JSON.stringify({ thang: thangXem, tuanList: layDanhSachThu7CN_(thangXem), rows });
}

function capNhatDiemDanh(userId, password, targetUserId, ngay, buoi) {
  const user = login(userId, password);
  assertQuyenDiemDanh_(user);

  const ngayStr = String(ngay || '').trim();
  const parts = ngayStr.split('-');
  if (parts.length !== 3) throw new Error('Ngày không hợp lệ!');
  const d = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
  const dow = d.getDay();
  if (dow !== 0 && dow !== 6) throw new Error('Chỉ điểm danh được Thứ 7 hoặc Chủ nhật!');

  const targetId = String(targetUserId || '').trim();
  if (!targetId) throw new Error('Thiếu nhân sự cần điểm danh!');

  const sheet = layHoacTaoSheet_('DiemDanhCuoiTuan', DIEMDANH_HEADERS_);

  const data = sheet.getDataRange().getValues();
  let foundRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === targetId && formatDateCell_(data[i][1]) === ngayStr) {
      foundRow = i + 1;
      break;
    }
  }

  const buoiNum = parseFloat(buoi);
  if (!buoi || isNaN(buoiNum) || buoiNum <= 0) {
    if (foundRow > 0) sheet.deleteRow(foundRow);
  } else {
    if (buoiNum !== 1 && buoiNum !== 0.5) throw new Error('Buổi điểm danh chỉ nhận giá trị 1 hoặc 0.5!');
    if (foundRow > 0) sheet.getRange(foundRow, 3).setValue(buoiNum);
    else sheet.appendRow([targetId, ngayStr, buoiNum]);
  }

  logActivity(userId, user.name, 'Điểm danh cuối tuần', `${targetId} - ${ngayStr}: ${buoi || 'xoá'}`);
  return 'Success';
}

function capNhatSoNgayNghi(userId, password, targetUserId, thang, soNgayNghi) {
  const user = login(userId, password);
  assertQuyenDiemDanh_(user);

  const targetId = String(targetUserId || '').trim();
  const thangStr = String(thang || '').trim();
  if (!targetId || !thangStr) throw new Error('Thiếu thông tin nhân sự/tháng!');

  const soNgayNghiNum = parseFloat(soNgayNghi);
  if (isNaN(soNgayNghiNum) || soNgayNghiNum < 0) throw new Error('Số ngày nghỉ không hợp lệ!');

  const sheet = layHoacTaoSheet_('PhepThang', PHEP_HEADERS_);

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === targetId && formatThangCell_(data[i][1]) === thangStr) {
      sheet.getRange(i + 1, 4).setValue(soNgayNghiNum);
      logActivity(userId, user.name, 'Sửa Số ngày nghỉ', `${targetId} - ${thangStr}: ${soNgayNghiNum}`);
      return 'Success';
    }
  }

  sheet.appendRow([targetId, thangStr, 0, soNgayNghiNum, new Date()]);
  logActivity(userId, user.name, 'Sửa Số ngày nghỉ', `${targetId} - ${thangStr}: ${soNgayNghiNum}`);
  return 'Success';
}

// Chỉ Admin (không áp dụng cho Thư ký BQLDA) được sửa trực tiếp "Phép + bù" — dùng để nhập số dư
// phép sẵn có của dự án (đã theo dõi thủ công trên Google Sheet trước khi có tab này), khác với
// capNhatSoNgayNghi ở trên là việc vận hành hàng ngày.
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
      maPhuLucTao: String(r[11] || '').trim()
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
      ghiChuKhauTru: String(r[9] || '').trim()
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
    tongDaThanhToan += tamTinh + thue - dot.thuHoiTamUngHopDong - dot.thuHoiTamUngThiCong - dot.khauTruKhac;
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

  boq.forEach(item => {
    item.khoiLuongHieuLuc = effectiveQtyMap[item.maBOQ] != null ? effectiveQtyMap[item.maBOQ] : item.khoiLuongHopDong;
    item.donGiaHieuLuc = donGiaHieuLucMap[item.maBOQ] != null ? donGiaHieuLucMap[item.maBOQ] : item.donGia;
    item.khoiLuongDaNghiemThu = executedByBOQ[item.maBOQ] || 0;
  });

  quyetToan.forEach(qt => {
    Object.assign(qt, tinhQuyetToanHopDong_(qt.maHopDong, { hopDong, boq, dotThanhToan, dotThanhToanChiTiet }));
  });

  return JSON.stringify({ hopDong, boq, phuLuc, phuLucThayDoi, nghiemThu, dotThanhToan, dotThanhToanChiTiet, quyetToan });
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
    String(item.maPhuLucTao || '').trim(), true
  ];

  if (maBOQ) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === maBOQ) {
        row[0] = maBOQ;
        row[11] = String(data[i][11] || '').trim(); // giữ nguyên maPhuLucTao gốc, không cho sửa tay
        sheet.getRange(i + 1, 1, 1, BOQ_HEADERS_.length).setValues([row]);
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
    String(dotThanhToan.ghiChuKhauTru || '').trim(), true
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
