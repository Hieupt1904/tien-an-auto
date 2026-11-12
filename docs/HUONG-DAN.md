# Hướng dẫn sử dụng (bản ngắn, cho người mới)

## 1. Cài một lần

* Tải Python 3 tại https://www.python.org/downloads/windows/
* Khi cài, **tích ô “Add python.exe to PATH”** rồi bấm Install.
* Chép thư mục `TienAnAuto` vào máy (ví dụ `D:\TienAnAuto`).

## 2. Chạy

* Nháy đúp **`Chay (Windows).bat`**.
* Cửa sổ có 4 nút:

```
┌──────────────────────────────────────────────┐
│ Tài khoản: (tên đăng nhập)                   │
│  ┌────────────────────────────────────────┐  │
│  │        ▶  Cập nhật dữ liệu             │  │  ← bấm nút này mỗi lần cần số
│  └────────────────────────────────────────┘  │
│  [██████████░░░░░░░░]  62%                   │
│  Đang lấy dữ liệu ngày 11/11 ...             │
│  [📂 Mở thư mục báo cáo] [⚙ Cài đặt] [📜 Nhật ký] │
│  ────────────────────────────────────────    │
│  14:20 Dùng lại phiên đăng nhập...           │
└──────────────────────────────────────────────┘
```

## 3. Lần đầu

1. Bấm **⚙ Cài đặt**
2. Nhập **Tên đăng nhập**, **Mật khẩu** (của qlmn.vn)
3. (Không bắt buộc) nhập **Tháng / Năm** muốn lấy — để trống = tháng hiện tại
4. Bấm **Lưu**

## 4. Lấy số liệu

1. Bấm **▶ Cập nhật dữ liệu**
2. Đợi khoảng 10–60 giây (tuỳ số ngày). Thanh % chạy dần.
3. Hiện hộp thoại “Đã cập nhật xong …” → bấm OK
4. Bấm **📂 Mở thư mục báo cáo** → mở file `Quyet toan tien an Thang 11-2026.xls`

## 5. Kết quả trông như thế nào

* Cột **Diễn giải** = tên mặt hàng (lấy đúng tên trên hệ thống)
* Mỗi ngày có phiếu = 2 cột: **Số lượng** và **Thành tiền**
* Cột **TỔNG** ở cuối
* Dòng cuối bảng = tổng tiền mỗi ngày
* Font Times New Roman, số dạng `#,##0`, đường kẻ, độ rộng cột: **giống hệt file mẫu**

## 6. Khi có sự cố

| Hiện tượng | Xử lý |
|---|---|
| Hộp thoại đỏ “Không cập nhật được” | Bấm **📜 Nhật ký**, mở file `.log` mới nhất, xem dòng cuối |
| Báo “Tên đăng nhập hoặc mật khẩu” | Bấm **⚙ Cài đặt** nhập lại |
| Báo “Tháng này chưa có dữ liệu” | Vào qlmn.vn kiểm tra đã kê phiếu chợ chưa |
| Máy chậm / mạng yếu | Bấm lại nút Cập nhật, không cần tắt app |

## 7. Tạo file .exe (không cần Python trên máy con)

Trên một máy có Python + mạng, chạy `tools\tao_exe.bat`.
Kết quả: thư mục `dist\TienAn` — chép nguyên thư mục đó sang máy khác là bấm
`TienAn.exe` được ngay (không cần cài gì thêm).
