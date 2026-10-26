# OCARCARE Monitor

Báo cáo vận hành hằng ngày chương trình OCARCARE (VPBank × OPES × BH PVI), luỹ kế từ 21/09/2026.
Link: https://pvi-aviationdivision.github.io/OCARCARE_MONITOR/ (có mật khẩu).

## Cấu trúc thư mục `D:\5. OCARCARE Report\`

```
DuLieu\                  ← CHỖ DUY NHẤT thả file (không đẩy lên GitHub)
  input\                 báo cáo BH PVI hằng ngày — lấy file MỚI NHẤT
  mapping\               file mapping điểm bán — lấy file mới nhất
  input_opes\            file OPES có Mã DAO — giữ tất cả, script gộp lại (dùng đối soát)
  input_nhadong\         file "Báo cáo Nhà đồng" OPES gửi hằng tháng — lấy file mới nhất
Monitor\                 ← repo GitHub OCARCARE_MONITOR
  update.bat             chạy hằng ngày (đến 01/11/2026 cập nhật luôn trang Sales Contest)
  xem_thu.bat            tạo số liệu + mở trang trên máy, không đẩy lên
  matkhau.txt            mật khẩu xem trang (chỉ nằm trên máy, KHÔNG đẩy lên)
  tools\build_monitor.py đọc Excel → docs\data.enc.js (đã mã hoá)
  docs\                  trang web (GitHub Pages)
```

## Cập nhật hằng ngày
1. Lưu báo cáo BH PVI vào `DuLieu\input\` (tên có ngày, VD `Báo cáo 8h ngày 27102026 VPBank.xlsx`).
2. Nhấp đúp `Monitor\update.bat`.

Hằng tháng: lưu file "Báo cáo Nhà đồng" của OPES vào `DuLieu\input_nhadong\` (tên có ngày chốt).

## Bảo mật
- Số liệu được mã hoá AES-256 bằng mật khẩu trong `matkhau.txt` trước khi đẩy lên; repo chỉ chứa chuỗi đã mã hoá.
- Đổi mật khẩu: sửa `matkhau.txt` → chạy `update.bat` → báo mật khẩu mới cho các bên.
- Trang chỉ chứa số tổng hợp theo điểm bán và số hợp đồng (phục vụ đối soát); không có tên khách hàng, SĐT, biển số, tên cán bộ bán.

## Quy ước số liệu
- Phí = TỔNG PHÍ BH chưa VAT (100%) theo báo cáo BH PVI; phần BH PVI = 40%.
- Tỷ trọng nhà đồng theo file OPES (đã loại dòng trùng; tự quy về chưa VAT nếu cột phí của OPES đang gồm VAT).
- Đối soát: so số hợp đồng PL-… của OPES với cột "Số đơn BH từ đối tác" của báo cáo BH PVI.
