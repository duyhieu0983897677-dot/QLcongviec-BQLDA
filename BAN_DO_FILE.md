# BẢN ĐỒ FILE — App Quản Lý Công Việc (Ban QLDA Bãi Xép)

> Tài liệu này KHÔNG phải là code, chỉ để bạn tra cứu "muốn sửa gì thì mở file nào".

## Cấu trúc file trong Apps Script

| # | Tên file | Loại | Vai trò |
|---|---|---|---|
| 1 | `Code.js` | Server | `doGet()`, `include()`, đọc/ghi dữ liệu, đăng nhập, quản lý tài khoản Giám sát, backup, nhật ký |
| 2 | `Index.html` | HTML | Khung giao diện: header, sidebar, các Tab, các Modal |
| 3 | `Style.html` | CSS | Màu sắc, layout, responsive (kế thừa gam màu app Hợp đồng) |
| 4 | `0_TienIch.html` | JS | Nền tảng dùng chung: chuyển tab, tải dữ liệu, định dạng ngày/số, tính trạng thái công việc |
| 5 | `1_DangNhap.html` | JS | Đăng nhập, đổi mật khẩu, lưu dữ liệu |
| 6 | `2_CongViec.html` | JS | Thêm/Sửa/Xóa 1 đầu việc (Modal) |
| 7 | `3_DanhSach.html` | JS | Bảng cây Nhóm → Việc, lọc theo Giám sát/trạng thái, tìm kiếm |
| 8 | `4_ViecCuaToi.html` | JS | Màn hình mobile để Giám sát tự cập nhật % tiến độ + ghi chú |
| 9 | `5_Dashboard.html` | JS | KPI tổng quan, việc trễ tiến độ, khối lượng theo Giám sát |
| 10 | `6_NguoiDung.html` | JS | (Admin) Quản lý tài khoản Giám sát |

## Tra cứu nhanh: "Tôi muốn sửa..." → "Mở file..."

| Tôi muốn... | Mở file |
|---|---|
| Đổi màu sắc, font chữ | `Style.html` |
| Đổi bố cục trang, thêm/bớt 1 Tab mới | `Index.html` |
| Đổi công thức tính trạng thái (Trễ/Hoàn thành/Đang làm) | `0_TienIch.html` (hàm `getTaskStatus`) |
| Đổi cách đăng nhập, đổi mật khẩu | `1_DangNhap.html` + `Code.js` |
| Sửa lỗi "không lưu được dữ liệu" | `Code.js` (hàm `saveData`) + `1_DangNhap.html` |
| Thêm/sửa/xóa 1 công việc | `2_CongViec.html` |
| Đổi cách hiển thị / lọc Danh sách công việc | `3_DanhSach.html` |
| Đổi giao diện "Việc của tôi" (mobile) | `4_ViecCuaToi.html` |
| Đổi KPI / bảng ở Tổng quan | `5_Dashboard.html` |
| Thêm/sửa/xóa tài khoản Giám sát | `6_NguoiDung.html` + `Code.js` (hàm `adminSaveSupervisor`/`adminDeleteSupervisor`) |

## Cấu trúc dữ liệu

- **`Data` (tab Sheet)**: lưu 1 chuỗi JSON lớn dạng `{ "tasks": [...] }` — mỗi phần tử `tasks[]` là 1 dòng Nhóm (`isHeader:true`) hoặc 1 dòng Việc (`isHeader:false`, có `supervisorName`, `startDate`, `endDate`, `percent`, `note`, `dailyLogs[]`).
- **`DanhSachUser` (tab Sheet)**: NGUỒN THẬT của tài khoản đăng nhập — cột A=ID đăng nhập, B=Mật khẩu (tự mã hóa), C=Vai trò (`ADMIN`/`SUPERVISOR`), D=Họ tên hiển thị. **Không sửa cột B bằng tay** trừ khi muốn đặt mật khẩu thường (hệ thống sẽ tự mã hóa ở lần đăng nhập đầu).
- **`NhatKy` (tab Sheet)**: nhật ký thao tác (ai, làm gì, khi nào).
- **`Backup_...` (tab ẩn)**: bản sao tab `Data` trước mỗi lần lưu, tự động giữ 15 bản gần nhất.

## Thiết lập ban đầu (1 lần duy nhất)

1. Tạo 1 Google Sheet mới, đặt tên tuỳ ý.
2. Trong Sheet đó, tạo 3 tab: `Data` (để trống), `DanhSachUser` (điền sẵn 1 dòng Admin đầu tiên: ví dụ `admin | 123456 | ADMIN | Ban QLDA`), `NhatKy` (để trống).
3. `Extensions > Apps Script`, dán đủ các file trong thư mục này vào (tên file phải khớp: `Code.js`, `Index.html`, `Style.html`, `0_TienIch.html`... `6_NguoiDung.html`).
4. `Deploy > New deployment > Web app`, Execute as "Me", Who has access "Anyone" (hoặc theo nhu cầu bảo mật của bạn).
5. Đăng nhập lần đầu bằng tài khoản Admin đã tạo ở bước 2, vào "Quản lý Giám sát" để thêm các Giám sát thật (Mr. Lam, Mr. Hồng...).
6. Vào tab "Danh sách công việc" > (sẽ trống ban đầu) — nếu muốn nạp sẵn ~1829 công việc từ Sheet kế hoạch cũ, dán nội dung file `data_source/seed_data.json` vào ô A2 của tab `Data` (dạng `{"tasks": [...]}`) trước khi mở app lần đầu.

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
