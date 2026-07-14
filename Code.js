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
const CONGVIEC_HEADERS_ = ['maCongViec', 'tenCongViec', 'maGoiThau', 'maHangMuc', 'phanLoai',
  'nguoiTao_id', 'nguoiTao_ten', 'ngayBatDauKH', 'ngayKetThucKH', 'active'];
const NHATKY_HEADERS_ = ['logId', 'maCongViec', 'ngayBaoCao', 'phanTramNgay', 'ghiChu',
  'nguoiNhap_id', 'nguoiNhap_ten', 'thoiGianNhap', 'fileDinhKem', 'active'];

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

  ss.getSheets().forEach(sh => {
    if (/^(Sheet1|Trang tính1)$/.test(sh.getName()) && ss.getSheets().length > 3) {
      ss.deleteSheet(sh);
    }
  });

  return 'Đã thiết lập xong: Data, DanhSachUser (admin/123456), NhatKy, HangMuc, GoiThau, CongViec, NhatKyTienDo.';
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

function readCongViecRows_(sheet) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, CONGVIEC_HEADERS_.length).getValues();
  const result = [];
  values.forEach((r, idx) => {
    const maCongViec = String(r[0] || '').trim();
    if (!maCongViec) return;
    if (!isActiveVal_(r[9])) return;
    result.push({
      maCongViec,
      tenCongViec: String(r[1] || '').trim(),
      maGoiThau: String(r[2] || '').trim(),
      maHangMuc: String(r[3] || '').trim(),
      phanLoai: String(r[4] || '').trim(),
      nguoiTaoId: String(r[5] || '').trim(),
      nguoiTaoTen: String(r[6] || '').trim(),
      ngayBatDauKH: formatDateCell_(r[7]),
      ngayKetThucKH: formatDateCell_(r[8]),
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

function daysBetween_(isoStart, isoEnd) {
  const a = new Date(isoStart), b = new Date(isoEnd);
  return Math.round((b - a) / 86400000);
}

// Trả về { hangMuc, goiThau, congViec, users } — congViec đã tính sẵn luyKe/phanTramKeHoach/
// chenhLech/trangThaiMau (cộng dồn từ NhatKyTienDo theo maCongViec).
function getData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hangMuc = readHangMucRows_(ss.getSheetByName('HangMuc'));
  const goiThau = readGoiThauRows_(ss.getSheetByName('GoiThau'));
  const congViec = readCongViecRows_(ss.getSheetByName('CongViec'));
  const logs = readNhatKyRows_(ss.getSheetByName('NhatKyTienDo')).filter(l => l.active);

  const luyKeMap = {};
  logs.forEach(l => { luyKeMap[l.maCongViec] = (luyKeMap[l.maCongViec] || 0) + l.phanTramNgay; });

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  congViec.forEach(cv => {
    const luyKe = Math.round((luyKeMap[cv.maCongViec] || 0) * 10) / 10;
    cv.luyKe = luyKe;
    let keHoach = 0;
    if (cv.ngayBatDauKH && cv.ngayKetThucKH) {
      const tongNgay = daysBetween_(cv.ngayBatDauKH, cv.ngayKetThucKH);
      if (tongNgay > 0) {
        const soNgayDaQua = daysBetween_(cv.ngayBatDauKH, today);
        keHoach = Math.max(0, Math.min(100, (soNgayDaQua / tongNgay) * 100));
      }
    }
    cv.phanTramKeHoach = Math.round(keHoach * 10) / 10;
    cv.chenhLech = Math.round((luyKe - cv.phanTramKeHoach) * 10) / 10;
    if (luyKe >= 100) cv.trangThaiMau = 'done';
    else if (cv.chenhLech <= -15) cv.trangThaiMau = 'bad';
    else if (cv.chenhLech < 0) cv.trangThaiMau = 'late';
    else cv.trangThaiMau = 'ok';
  });

  return JSON.stringify({ hangMuc, goiThau, congViec, users: listSupervisors() });
}

// Giống getData() nhưng gắn kèm 5 log gần nhất cho mỗi Công việc — dùng cho màn Báo cáo tổng hợp
// mở rộng dòng xem chi tiết mà không cần round-trip riêng.
function getBaoCaoTongHop() {
  const parsed = JSON.parse(getData());
  const logs = readNhatKyRows_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('NhatKyTienDo')).filter(l => l.active);
  const logsByCongViec = {};
  logs.forEach(l => { (logsByCongViec[l.maCongViec] = logsByCongViec[l.maCongViec] || []).push(l); });

  parsed.congViec.forEach(cv => {
    const list = (logsByCongViec[cv.maCongViec] || []).slice().sort((a, b) => b.logId.localeCompare(a.logId));
    cv.recentLogs = list.slice(0, 5);
  });
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
function adminSaveHangMuc(userId, password, node) {
  const user = login(userId, password);
  if (user.role !== 'ADMIN') throw new Error('Chỉ Quản trị mới có quyền quản lý Hạng mục!');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('HangMuc');
  if (!sheet) throw new Error('Chưa cấu hình Tab HangMuc trên Google Sheet!');

  const maHangMuc = String(node.maHangMuc || '').trim();
  if (!maHangMuc) throw new Error('Vui lòng nhập Mã hạng mục!');

  const data = sheet.getDataRange().getValues();
  let foundRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === maHangMuc) { foundRow = i + 1; break; }
  }

  const row = [maHangMuc, node.tenHangMuc || '', parseInt(node.capDo, 10) || 1, node.parentId || '', true];

  if (foundRow > 0) {
    sheet.getRange(foundRow, 1, 1, row.length).setValues([row]);
    logActivity(userId, user.name, 'Sửa Hạng mục', `Sửa ${maHangMuc} - ${node.tenHangMuc}`);
  } else {
    sheet.appendRow(row);
    logActivity(userId, user.name, 'Thêm Hạng mục', `Thêm ${maHangMuc} - ${node.tenHangMuc}`);
  }
  return 'Success';
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

// congViec = { maCongViec (rỗng nếu tạo mới), tenCongViec, maGoiThau, maHangMuc, ngayBatDauKH, ngayKetThucKH }
function adminSaveCongViec(userId, password, congViec) {
  const user = login(userId, password);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CongViec');
  if (!sheet) throw new Error('Chưa cấu hình Tab CongViec trên Google Sheet!');

  const tenCongViec = String(congViec.tenCongViec || '').trim();
  const maGoiThau = String(congViec.maGoiThau || '').trim();
  const maHangMuc = String(congViec.maHangMuc || '').trim();
  if (!tenCongViec) throw new Error('Vui lòng nhập Tên công việc!');
  if (!maGoiThau) throw new Error('Vui lòng chọn Gói thầu!');
  if (!maHangMuc) throw new Error('Vui lòng chọn Hạng mục!');

  const data = sheet.getDataRange().getValues();
  const maCongViec = String(congViec.maCongViec || '').trim();

  if (maCongViec) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === maCongViec) {
        if (user.role !== 'ADMIN' && String(data[i][5]).trim() !== user.supervisorId) {
          throw new Error('Bạn không có quyền sửa Công việc này!');
        }
        // Giữ nguyên phanLoai (cột E) và người tạo (cột F/G) — chỉ cho sửa các field mô tả.
        sheet.getRange(i + 1, 2, 1, 4).setValues([[tenCongViec, maGoiThau, maHangMuc, data[i][4]]]);
        sheet.getRange(i + 1, 8, 1, 2).setValues([[congViec.ngayBatDauKH || '', congViec.ngayKetThucKH || '']]);
        logActivity(userId, user.name, 'Sửa Công việc', `Sửa ${maCongViec} - ${tenCongViec}`);
        return 'Success';
      }
    }
    throw new Error('Không tìm thấy Công việc!');
  }

  const phanMacDinh = String(user.phanMacDinh || '').trim()
    || String((listSupervisors().find(u => u.id === user.supervisorId) || {}).phanMacDinh || '').trim();
  if (!phanMacDinh) {
    throw new Error('Tài khoản của bạn chưa được Quản trị gán Phần mặc định (Chức danh). Vui lòng liên hệ Quản trị trước khi tự tạo Công việc.');
  }

  const now = new Date();
  const newId = sinhMaCongViec_(sheet, now);
  sheet.appendRow([
    newId, tenCongViec, maGoiThau, maHangMuc, phanMacDinh,
    user.supervisorId, user.name,
    congViec.ngayBatDauKH || '', congViec.ngayKetThucKH || '', true
  ]);
  logActivity(userId, user.name, 'Thêm Công việc', `Thêm ${newId} - ${tenCongViec}`);
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
