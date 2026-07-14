// Hàm include() cho phép Index.html nhúng Style.html và các file Script.html
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Chạy 1 lần duy nhất sau khi tạo Sheet mới: tạo đủ 3 tab Data/DanhSachUser/NhatKy
// và 1 tài khoản admin mặc định (admin / 123456) để đăng nhập lần đầu.
function thietLapBanDauSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss.getSheetByName('Data')) {
    ss.insertSheet('Data').getRange('A1').setValue('Dữ liệu hệ thống');
  }

  if (!ss.getSheetByName('DanhSachUser')) {
    const sh = ss.insertSheet('DanhSachUser');
    sh.getRange(1, 1, 2, 4).setValues([
      ['ID', 'MatKhau', 'VaiTro', 'HoTen'],
      ['admin', '123456', 'ADMIN', 'Ban QLDA']
    ]);
  }

  if (!ss.getSheetByName('NhatKy')) {
    ss.insertSheet('NhatKy').getRange(1, 1, 1, 4)
      .setValues([['Thời gian', 'Người dùng', 'Hành động', 'Chi tiết']]);
  }

  ss.getSheets().forEach(sh => {
    if (/^(Sheet1|Trang tính1)$/.test(sh.getName()) && ss.getSheets().length > 3) {
      ss.deleteSheet(sh);
    }
  });

  return 'Đã thiết lập xong: Data, DanhSachUser (admin/123456), NhatKy.';
}

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Quản Lý Công Việc - Ban QLDA')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function getData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
  const lastRow = sheet.getLastRow();
  let fullString = "";
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < values.length; i++) fullString += values[i][0];
  }
  const parsed = fullString ? JSON.parse(fullString) : { tasks: [] };
  // Danh sách Giám sát LUÔN lấy trực tiếp từ tab DanhSachUser (nguồn xác thực thật),
  // không lấy từ JSON blob — tránh 2 nguồn dữ liệu user lệch nhau.
  parsed.users = listSupervisors();
  return JSON.stringify(parsed);
}

// Đọc danh sách tài khoản (Giám sát/Admin) trực tiếp từ tab DanhSachUser — KHÔNG trả về mật khẩu.
function listSupervisors() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DanhSachUser');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    result.push({ id: String(data[i][0]).trim(), role: String(data[i][2]).trim().toUpperCase(), name: String(data[i][3]).trim() });
  }
  return result;
}

// (Admin) Thêm mới hoặc sửa 1 tài khoản Giám sát trực tiếp vào tab DanhSachUser.
function adminSaveSupervisor(adminId, adminPass, targetId, name, role, newPassword) {
  const adminUser = login(adminId, adminPass);
  if (adminUser.role !== 'ADMIN') throw new Error("Chỉ Quản trị mới có quyền quản lý tài khoản!");

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DanhSachUser');
  if (!sheet) throw new Error("Chưa cấu hình Tab DanhSachUser trên Google Sheet!");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(targetId).trim()) {
      sheet.getRange(i + 1, 3).setValue(role);
      sheet.getRange(i + 1, 4).setValue(name);
      if (newPassword) sheet.getRange(i + 1, 2).setValue(newPassword); // lưu thường, tự mã hóa ở lần đăng nhập đầu (xem login())
      logActivity(adminId, adminUser.name, "Sửa Giám sát", `Sửa tài khoản ${targetId}`);
      return "Success";
    }
  }

  if (!newPassword) throw new Error("Vui lòng đặt mật khẩu ban đầu cho tài khoản mới!");
  sheet.appendRow([targetId, newPassword, role, name]);
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
          supervisorId: String(data[i][0]).trim()
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
// HÀM LƯU DỮ LIỆU CHÍNH TỪ GIAO DIỆN
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
