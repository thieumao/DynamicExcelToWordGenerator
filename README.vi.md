# DocGen – Sinh tài liệu Word động từ Excel

Ứng dụng desktop chạy **offline** (ưu tiên Windows, phát triển trên macOS). Ứng dụng đọc file Excel nhiều sheet của khách hàng, tra cứu mã/ID qua các sheet tham chiếu, rồi điền dữ liệu vào các file Word mẫu. Kết quả có thể là từng tài liệu riêng hoặc một tài liệu gộp. Mọi khác biệt giữa các khách hàng nằm trong **profile** (file YAML), nên đổi khách hàng không cần sửa code.

## Dành cho ngành Xây dựng (bản phát hành đầu tiên)

Dùng được ngay cả khi không rành máy tính:
- **Bộ cài Windows** `DocGen-Setup-…-x64.exe`: bấm đúp → *Next*. Không cần quyền admin, không cần Internet, không cần cài Office.
- **Trang chủ có 3 nút lớn**:
  1. *Tạo bộ hồ sơ mới*;
  2. *Dùng bộ hồ sơ đã lưu*;
  3. *Dùng thử hồ sơ mẫu*: hồ sơ nghiệm thu công việc xây dựng (Phiếu yêu cầu và Biên bản nghiệm thu) và nghiệm thu vật liệu đầu vào, dùng chung một file Excel.
- **Tạo hồ sơ nhanh** trong 4 bước:
  1. Chọn file Excel.
  2. Chọn các file Word mẫu.
  3. Kiểm tra chỗ cần điền: ứng dụng **tự ghép**, màu xanh là đã ghép, màu vàng là chưa ghép, có giá trị ví dụ.
  4. Tạo tài liệu.

  Không phải cấu hình gì.
- **Thông tin công trình nhập một lần** (sheet "tên mục | nội dung"), dùng cho mọi biên bản. Họ tên, chức vụ được tra theo mã (GS01, KT01…).
- Ngày, giờ kiểu biên bản: `ngày 02 tháng 04 năm 2024`, `08 giờ 00`.
- Mỗi công việc một file Word, kèm **một file tổng hợp để in** tất cả.
- Mọi thông báo lỗi bằng tiếng Việt, chỉ rõ sheet, dòng, cột và cách sửa.
- Các trang cấu hình chi tiết chỉ hiện khi bật *Chế độ nâng cao*.

Hướng dẫn từng bước: [docs/HUONG_DAN_SU_DUNG.md](docs/HUONG_DAN_SU_DUNG.md) (phần A).

## Tính năng chính

**Đọc Excel**
- Tên sheet, vị trí dòng tiêu đề (tự dò hoặc cố định), tên cột hay vị trí cột đều có thể khác nhau tùy khách hàng.
- Đọc được `.xlsx`, `.xlsm`, `.xls`, `.csv`. File CSV có thể ở dạng UTF-8 hoặc Windows-1258 (dấu tiếng Việt tổ hợp).

**Tra cứu và ánh xạ dữ liệu**
- Quan hệ giữa các sheet: khóa ghép, tra cứu nhiều cấp, quan hệ một-nhiều (bảng lặp), tổng hợp (sum/count/join…).
- Ánh xạ theo điều kiện, biểu thức an toàn (sandbox), đọc số thành chữ tiếng Việt, định dạng ngày/tiền kiểu Việt Nam.
- Tra cứu không khớp hoặc trùng khóa đều được báo rõ sheet/dòng/cột. Hệ thống không bao giờ tự chọn bừa một giá trị.
- Ô công thức chưa có kết quả được tính sẵn trong file sẽ không được dùng âm thầm.

**Word**
- Hỗ trợ 3 kiểu đánh dấu: `{{ ... }}` (Jinja), `<<...>>` tự định nghĩa, và MERGEFIELD.
- Dùng được **file Word có sẵn của khách hàng** mà không cần đánh dấu. Ứng dụng tự tìm các chỗ trống sau nhãn (`Họ và tên: ......`), ô bảng trống cạnh ô nhãn, ô nhập liệu (content control), form field, ô đánh dấu, danh sách chọn và khung ảnh. Chỗ nào được gắn với dữ liệu thì được điền; chỗ không gắn (chữ ký, ô dành cho cán bộ) giữ nguyên.
- Người dùng tự chọn file Word ngay trong ứng dụng: nút *Nhập…* ở trang Template Word, hoặc *Thêm file Word…* ở trang Sinh tài liệu.
- Gộp nhiều template thành một tài liệu mà vẫn giữ hướng trang, lề, header/footer riêng của từng phần.

**Sinh tài liệu**
- 5 chế độ: A một tài liệu, B hàng loạt, C theo nhóm, D gộp, E hàng loạt gộp.
- Đặt tên file theo mẫu, không ghi đè file đã có, ghi file an toàn (không để lại file dở).
- Có báo cáo audit dạng JSON và CSV.

**Giao diện**
- Tiếng Việt (mặc định) và tiếng Anh. Mọi thông báo lỗi, cảnh báo và gợi ý sửa đều có tiếng Việt, kèm chi tiết sheet, dòng, cột.
- Thiết kế mapping trực quan, xem trước kèm truy vết từng giá trị về ô Excel gốc.
- Có thanh tiến trình, cho phép hủy giữa chừng, UI không bị treo.

## Chạy thử (dành cho lập trình viên)

```bash
uv sync --extra gui
uv run docgen samples samples          # tạo lại dữ liệu mẫu (customer_a/b/c, xay_dung_nghiem_thu)
uv run docgen-gui                      # mở ứng dụng desktop
uv run docgen generate --profile samples/xay_dung_nghiem_thu/profile.yaml \
  --excel samples/xay_dung_nghiem_thu/data/HoSoNghiemThu.xlsx --out output/nghiem_thu
```

Người dùng Windows nhận bộ cài `DocGen-Setup-…-x64.exe` (hoặc thư mục `DocGen.exe` chạy được ngay), không cần cài Python. Cả hai được build và kiểm tra tự động bằng GitHub Actions; xem [docs/BUILD_WINDOWS.md](docs/BUILD_WINDOWS.md).

## Tài liệu

- Hướng dẫn sử dụng và soạn template bằng tiếng Việt: [docs/HUONG_DAN_SU_DUNG.md](docs/HUONG_DAN_SU_DUNG.md).
- Tài liệu kỹ thuật (tiếng Anh): xem danh sách trong [README.md](README.md).
