# Tiền ăn – tự động cập nhật

Ứng dụng một cửa sổ, bốn nút: tự đăng nhập **qlmn.vn**, lấy số liệu **phiếu kê chợ**
của cả tháng và in ra file Excel **Quyết toán tiền ăn** đúng mẫu nhà trường đang dùng.

* Không cần cài Excel, không cần `pip install` gì.
* Chỉ cần **Python 3.8+** (đã có sẵn trên máy Windows mới, hoặc tải ở python.org).
* Định dạng file ra (font Times New Roman, số `#,##0`, đường kẻ, độ rộng cột,
  trộn ô, cột ngày, dòng tổng) **lấy nguyên từ file mẫu**.

## Chạy

1. Chép cả thư mục này vào máy, ví dụ `D:\TienAnAuto`.
2. Nháy đúp **`Chay (Windows).bat`**.
3. **⚙ Cài đặt** → nhập tên đăng nhập + mật khẩu qlmn.vn → **Lưu**.
4. Bấm **▶ CẬP NHẬT DỮ LIỆU** — đợi 10–60 giây.
5. **📂 Mở thư mục báo cáo** → mở `Quyet toan tien an Thang MM-YYYY.xls`.

Lần thứ hai bấm lại sẽ nhanh hơn vì chương trình dùng lại phiên đăng nhập (~90 phút).

## Bốn nút

| Nút | Việc |
|---|---|
| **▶ CẬP NHẬT DỮ LIỆU** | Đăng nhập, lấy số liệu cả tháng, ghi file `.xls` |
| **📂 Mở thư mục báo cáo** | Mở thư mục chứa file kết quả và nhật ký |
| **⚙ Cài đặt** | Tài khoản, tháng/năm, kho, khoảng ngày |
| **📜 Nhật ký** | Mở file log từng lần chạy |

## File ra trông thế nào

```
THANH TOÁN TIỀN ĂN THÁNG 11
────────────────────────────────────────────────────
Diễn giải │ Ngày 02-11 │ Ngày 03-11 │ … │ TỔNG
          │ SL │ Tiền  │ SL │ Tiền  │ … │ SL │ Tiền
Thịt lợn… │ 12.8│1792000│ 11 │1540000│ … │…  │ …
Nước mắm… │  1  │ 60000 │0.61│ 36600 │ … │…  │ …
…
          │    │       │    │       │ … │ TỔNG TIỀN MỖI NGÀY
```

* Dòng = mặt hàng (tự sắp: xuất hiện nhiều ngày trước, tiền lớn trước).
* Cột = ngày có phiếu trong tháng (đúng thứ tự ngày trong file mẫu).

## Cấu trúc mã

## Kiểm chứng số liệu (không đoán)

```
python tests/run_all.py                                  # 4 bài, không cần mạng
python tools/verify_report.py "Báo cáo/....xls"          # đối chiếu từng cột với API
```

* `test_coords.py` — kiểm tra **toạ độ ô**: dữ liệu đưa vào cột 5/6 phải ra đúng cột 5/6
  (MULRK chỉ mã hoá được các cột liền nhau, đứt cột là phải tách), kể cả ô rời rạc và cột TỔNG.
* `test_verify_snapshot.py` — 202 ô + tổng từng ngày khớp snapshot API đã lưu.
* `tools/verify_report.py` — mở file .xls, gọi lại API từng ngày, so từng mặt hàng và tổng tiền
  của cột; in `OK` / `LECH` cho từng ngày.

## Cấu trúc mã

```
src/_http.py     HTTP thuần urllib (cookie, gzip) – không cần requests
src/_ocr.py      đọc mã bảo vệ bằng so khớp mẫu – không cần PIL/OpenCV
src/_api.py      đăng nhập + gọi API qlmn.vn, tự đăng nhập lại khi hết phiên
src/_session.py  nhớ phiên đăng nhập ~90 phút
src/_pull.py     chuẩn hoá phiếu kê chợ theo ngày
src/_report.py   ghép ngày → dòng mặt hàng → cột
src/_xlsio.py    dựng file .xls từ khuôn mẫu (giữ định dạng gốc)
src/_xls.py      đọc .xls (BIFF8/OLE2) để lấy định dạng và kiểm tra
src/_ole.py      đóng gói stream BIFF8 thành file .xls
src/_log.py      nhật ký
src/app/main.py  giao diện Tkinter
tools/build_template.py   sinh lại src/template.json từ file mẫu mới
tools/verify_report.py    đối chiếu file báo cáo với API từng ngày
tools/save_snapshot.py    lưu số liệu 1 tháng để hồi quy
tools/tao_exe.bat         đóng gói .exe (PyInstaller)
tests/                    kiểm tra không cần mạng
```

Đổi mẫu Excel khác: thay file mẫu rồi chạy
`python tools/build_template.py "file_mau.xls" src/template.json`.

## Sự cố

| Hiện tượng | Xử lý |
|---|---|
| “Không cập nhật được” | Mở **📜 Nhật ký**, xem dòng cuối |
| Sai tài khoản / mật khẩu | **⚙ Cài đặt** nhập lại |
| Tháng chưa có dữ liệu | Vào qlmn.vn kiểm tra đã kê phiếu chợ |
| Mã bảo vệ sai nhiều lần | Cứ để chạy, chương trình tự thử lại (tối đa 80 lần, ~0,4 s/lần) |
