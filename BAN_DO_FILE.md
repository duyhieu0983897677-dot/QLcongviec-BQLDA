# BẢN ĐỒ FILE — App Quản Lý Công Việc (Ban QLDA Bãi Xép)

> Tài liệu này KHÔNG phải là code, chỉ để bạn tra cứu "muốn sửa gì thì mở file nào".

## Cấu trúc file trong Apps Script

| # | Tên file | Loại | Vai trò |
|---|---|---|---|
| 1 | `Code.js` | Server | `doGet()`, `include()`, đọc/ghi Hạng mục (mã tham chiếu) + Gói thầu + Công việc + Nhật ký tiến độ, đăng nhập, quản lý tài khoản Giám sát, backup, nhật ký thao tác |
| 2 | `Index.html` | HTML | Khung giao diện: header, sidebar, các Tab, các Modal |
| 3 | `Style.html` | CSS | Màu sắc (tông "công trường": nền kem, chữ đen, accent cam an toàn), layout, responsive, gồm cả CSS module Nhật ký tiến độ (lớp `ntd-`/`wtd-`/`bc-`) |
| 4 | `0_TienIch.html` | JS | Nền tảng dùng chung: chuyển tab, tải dữ liệu (`getData()`/`reloadData()`), `gsRun()` bọc `google.script.run` thành Promise, định dạng ngày/số |
| 5 | `1_DangNhap.html` | JS | Đăng nhập, đổi mật khẩu, xử lý lỗi backend dùng chung (`handleBackendError`) |
| 6 | `2_CongViec.html` | JS | 3 modal: Hạng mục (mã/tên tham chiếu, chỉ Admin), Gói thầu (chỉ Admin), **Công việc** (Admin/Giám sát tự tạo — Gói thầu → Hạng mục → Tên → Ngày) |
| 7 | `3_DanhSach.html` | JS | Tab "Danh sách": bảng **Công việc** nhóm theo Gói thầu → (I. Phần hồ sơ / II. Phần triển khai thi công) kèm lũy kế/trạng thái; bên dưới 2 khối admin-only quản lý Hạng mục & Gói thầu |
| 8 | `4_ViecCuaToi.html` | JS | Tab "Nhật ký tiến độ" (id tab vẫn `viecuatoi`): bản mobile (wizard 1 Công việc/lần) + bản desktop (nhập hàng loạt), ghi qua `themNhatKyTienDo`/`themNhatKyTienDoHangLoat`, có nút "Thêm công việc" |
| 9 | `5_Dashboard.html` | JS | Tab "Tổng quan": báo cáo tổng hợp tự sinh từ Nhật ký tiến độ (`getBaoCaoTongHop()`) theo **Công việc** — KPI, bảng tiến độ có vạch kế hoạch, lọc theo Gói thầu/Phần/GS, mở rộng xem log |
| 10 | `6_NguoiDung.html` | JS | (Admin) Quản lý tài khoản Giám sát — gán **Chức danh + Phần mặc định** (quyết định Công việc do người này tự tạo rơi vào I. Phần hồ sơ hay II. Phần triển khai thi công) |

## Tra cứu nhanh: "Tôi muốn sửa..." → "Mở file..."

| Tôi muốn... | Mở file |
|---|---|
| Đổi màu sắc, font chữ | `Style.html` |
| Đổi bố cục trang, thêm/bớt 1 Tab mới | `Index.html` |
| Đổi ngưỡng phân loại trễ nhẹ/nghiêm trọng, công thức % kế hoạch | `Code.js` (hàm `getData()`) |
| Đổi cách đăng nhập, đổi mật khẩu | `1_DangNhap.html` + `Code.js` |
| Thêm/sửa/xóa 1 Hạng mục (mã tham chiếu, chỉ Admin) | `2_CongViec.html` + `Code.js` (`adminSaveHangMuc`/`adminDeleteHangMuc`) |
| Thêm/sửa/xóa 1 Gói thầu | `2_CongViec.html` + `Code.js` (`adminSaveGoiThau`/`adminDeleteGoiThau`) |
| Thêm/sửa/xóa 1 Công việc (đơn vị thật nhận %) | `2_CongViec.html` + `Code.js` (`adminSaveCongViec`/`adminDeleteCongViec`) |
| Đổi cách hiển thị / lọc danh sách Công việc | `3_DanhSach.html` |
| Đổi giao diện nhập Nhật ký tiến độ (mobile/desktop) | `4_ViecCuaToi.html` |
| Đổi logic validate % (chặn vượt 100%, khoá chống ghi đè) | `Code.js` (hàm `validateVaGhiNhatKy_`, dùng `LockService`) |
| Đổi KPI / bảng ở Tổng quan | `5_Dashboard.html` |
| Đổi quy tắc Chức danh → Phần (hồ sơ/thi công) | `6_NguoiDung.html` (modal) — Phần lưu trực tiếp trên từng tài khoản, không suy luận tự động theo tên chức danh |
| Thêm/sửa/xóa tài khoản Giám sát | `6_NguoiDung.html` + `Code.js` (hàm `adminSaveSupervisor`/`adminDeleteSupervisor`) |

## Cấu trúc dữ liệu (Google Sheet)

- **`HangMuc`**: mã/tên tham chiếu để lọc & để Gói thầu gộp — `maHangMuc | tenHangMuc | capDo
  (1=nhóm lớn/2=đầu mục/3=chi tiết, chỉ để tổ chức cây) | parentId | active`. KHÔNG còn nhận Nhật
  ký tiến độ trực tiếp, chỉ Admin được thêm/sửa/xoá. Xoá = soft delete (`active=false`).
- **`GoiThau`**: `maGoiThau | tenGoiThau | maHangMucList` (danh sách mã Hạng mục gộp — mọi cấp,
  cách nhau dấu phẩy). Mỗi Gói thầu luôn có 2 phần cố định ở tầng hiển thị: **I. Phần hồ sơ** /
  **II. Phần triển khai thi công** (không lưu thành node riêng, suy ra từ `CongViec.phanLoai`).
- **`CongViec`**: đơn vị công việc THẬT, nơi Nhật ký tiến độ trỏ vào — `maCongViec (tự sinh
  CV-yyyyMMdd-NNN) | tenCongViec | maGoiThau | maHangMuc | phanLoai (ho_so/thi_cong, tự gán theo
  phanMacDinh của người tạo — không chọn tay) | nguoiTao_id | nguoiTao_ten | ngayBatDauKH |
  ngayKetThucKH | active`. Admin hoặc chính người tạo mới được sửa/xoá (soft delete).
- **`NhatKyTienDo`**: log **append-only**, nguồn sự thật duy nhất của % tiến độ — `logId |
  maCongViec | ngayBaoCao | phanTramNgay | ghiChu | nguoiNhap_id | nguoiNhap_ten | thoiGianNhap |
  fileDinhKem | active`. Không bao giờ sửa/xoá dòng gốc — sửa số liệu bằng cách thêm dòng mới (có
  thể âm). `getData()` tự tính lũy kế (SUM theo `maCongViec`), % kế hoạch, chênh lệch, trạng thái
  màu mỗi lần đọc.
- **`Data` (tab Sheet)**: dữ liệu kiểu cũ (blob JSON `{tasks:[...]}`) — **không còn được ghi**, giữ
  lại nhưng không dùng cho luồng hiện tại.
- **`DanhSachUser` (tab Sheet)**: NGUỒN THẬT của tài khoản đăng nhập — cột A=ID đăng nhập, B=Mật
  khẩu (tự mã hóa), C=Vai trò (`ADMIN`/`SUPERVISOR`), D=Họ tên hiển thị, E=Chức danh (nhãn tự do),
  F=Phần mặc định (`ho_so`/`thi_cong`).
- **`NhatKy` (tab Sheet)**: nhật ký thao tác hệ thống (ai, làm gì, khi nào) — khác với `NhatKyTienDo`
  (nhật ký báo cáo % thi công).
- **`Backup_...` (tab ẩn)**: bản sao tab `Data` trước mỗi lần ghi kiểu cũ, tự động giữ 15 bản gần nhất.
- Ảnh hiện trường lưu ở thư mục Drive **`NhatKyTienDo_AnhHienTruong`** của tài khoản đang chạy script.

## Thiết lập ban đầu (1 lần duy nhất)

1. Tạo 1 Google Sheet mới, đặt tên tuỳ ý.
2. `Extensions > Apps Script`, dán đủ các file trong thư mục này vào (tên file phải khớp: `Code.js`,
   `Index.html`, `Style.html`, `0_TienIch.html`... `6_NguoiDung.html`).
3. Chạy hàm `thietLapBanDauSheet()` 1 lần từ trình soạn thảo Apps Script — tự tạo đủ các tab
   `Data`, `DanhSachUser` (kèm tài khoản `admin/123456`), `NhatKy`, `HangMuc`, `GoiThau`, `CongViec`,
   `NhatKyTienDo`.
4. `Deploy > New deployment > Web app`, Execute as "Me", Who has access "Anyone" (hoặc theo nhu cầu
   bảo mật của bạn).
5. Đăng nhập lần đầu bằng tài khoản Admin:
   - Vào "Quản lý Giám sát" thêm các tài khoản thật, gán **Chức danh + Phần mặc định** cho từng
     người (vd QS/QAQC → I. Phần hồ sơ, Giám sát hiện trường → II. Phần triển khai thi công).
   - Vào tab "Danh sách" dựng cây **Hạng mục** tham chiếu (mã/tên) và tạo **Gói thầu** (chọn các
     Hạng mục thuộc gói thầu đó).
   - Sau đó mỗi người tự vào "Nhật ký tiến độ" bấm "Thêm công việc" (chọn Gói thầu → Hạng mục → đặt
     tên → ngày) rồi mới nhập % được.

## Nguyên tắc khi nhờ AI (Claude) sửa code

1. Tra bảng trên để biết đúng 1 file cần sửa.
2. Gửi đúng file đó kèm mô tả muốn sửa gì.
3. Nếu không chắc file nào, cứ mô tả vấn đề — Claude sẽ tự xác định.
4. Sau khi nhận file mới: **Dán đè vào đúng file → Ctrl+S → Deploy → New version → Triển khai.**

## Thứ tự include bắt buộc trong `Index.html`

```
0_TienIch.html → 1_DangNhap.html → 2_CongViec.html → 3_DanhSach.html → 4_ViecCuaToi.html → 5_Dashboard.html → 6_NguoiDung.html
```

**Không tự ý đổi thứ tự** các dòng `<?!= include(...); ?>` này, nếu không một số hàm có thể báo lỗi "chưa được định nghĩa".
