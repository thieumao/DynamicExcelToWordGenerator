# DocGen – Sinh tài liệu Word động từ Excel

Ứng dụng desktop chạy **offline** (ưu tiên Windows, phát triển trên macOS). Ứng dụng đọc file Excel nhiều sheet của khách hàng, tra cứu mã/ID qua các sheet tham chiếu, rồi điền dữ liệu vào các file Word mẫu. Kết quả có thể là từng tài liệu riêng hoặc một tài liệu gộp. Mọi khác biệt giữa các khách hàng nằm trong **profile** (file YAML), nên đổi khách hàng không cần sửa code.

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
- Dùng được **file Word có sẵn của khách hàng** mà không cần đánh dấu. Ứng dụng tự tìm các chỗ trống sau nhãn (`Họ và tên: ......`), ô bảng trống cạnh ô nhãn, ô nhập liệu (content control) và form field. Chỗ nào được gắn với dữ liệu thì được điền; chỗ không gắn (chữ ký, ô dành cho cán bộ) giữ nguyên.
- Người dùng tự chọn file Word ngay trong ứng dụng: nút *Nhập…* ở trang Template Word, hoặc *Thêm file Word…* ở trang Sinh tài liệu.
- Gộp nhiều template thành một tài liệu mà vẫn giữ hướng trang, lề, header/footer riêng của từng phần.

**Sinh tài liệu**
- 5 chế độ: A một tài liệu, B hàng loạt, C theo nhóm, D gộp, E hàng loạt gộp.
- Đặt tên file theo mẫu, không ghi đè file đã có, ghi file an toàn (không để lại file dở).
- Có báo cáo audit dạng JSON và CSV.

**Giao diện**
- Tiếng Việt và tiếng Anh.
- Thiết kế mapping trực quan, xem trước kèm truy vết từng giá trị về ô Excel gốc.
- Có thanh tiến trình, cho phép hủy giữa chừng, UI không bị treo.

## Chạy thử (dành cho lập trình viên)

```bash
uv sync --extra gui
uv run docgen samples samples          # tạo lại dữ liệu mẫu customer_a, customer_b, customer_c
uv run docgen-gui                      # mở ứng dụng desktop
uv run docgen generate --profile samples/customer_a/profile.yaml \
  --excel samples/customer_a/data/HR.xlsx --out output/customer_a
```

Người dùng Windows nhận một thư mục chạy được ngay (`DocGen.exe`), không cần cài Python. Bản này được build tự động bằng GitHub Actions; xem [docs/BUILD_WINDOWS.md](docs/BUILD_WINDOWS.md).

## Tài liệu

- Hướng dẫn sử dụng và soạn template bằng tiếng Việt: [docs/HUONG_DAN_SU_DUNG.md](docs/HUONG_DAN_SU_DUNG.md).
- Tài liệu kỹ thuật (tiếng Anh): xem danh sách trong [README.md](README.md).
