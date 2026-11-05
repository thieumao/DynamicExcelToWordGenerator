# Hướng dẫn sử dụng DocGen

## 1. Quy trình làm việc với một khách hàng mới

1. **Profile khách hàng → Tạo mới.** Nhập tên khách hàng. Mỗi khách hàng có một profile riêng.
2. **Cấu hình Excel**
   - Chọn file Excel rồi bấm *Đọc file Excel*. Ứng dụng tự liệt kê các sheet và đoán dòng tiêu đề (dòng được tô xanh trong bảng xem trước).
   - Chọn từng sheet cần dùng và bấm *Thêm sheet đã chọn vào profile*.
   - Đặt sheet chứa dữ liệu chính (mỗi dòng ứng với một hồ sơ hoặc hợp đồng…) làm *sheet chính*.
   - Kiểm tra lại bảng cột:
     - **ID** là tên nội bộ, không dấu, ví dụ `ma_nv`.
     - **Tiêu đề cột** là chữ đúng như trong Excel.
     - **Tên khác** ghi các cách viết khác, cách nhau bằng dấu phẩy, ví dụ `Mã NV, Ma NV`.
     - **Kiểu** là kiểu dữ liệu: date, decimal…
   - Bấm *Áp dụng thay đổi sheet*.
3. **Quan hệ giữa các sheet.** Khai báo cách nối mã: sheet nguồn, sheet đích, cặp cột khóa. Nếu cần khóa ghép (ví dụ chi nhánh + mã phòng) thì thêm nhiều cặp. Bấm *Thử tất cả quan hệ* để xem số khóa và số khóa bị trùng.
4. **Thiết kế mapping.** Tạo các trường mà file Word sẽ dùng, ví dụ `nhan_vien.phong_ban`. Có các cách lấy giá trị:
   - *Cột của sheet này*: lấy thẳng từ sheet chính.
   - *Tra cứu ở sheet khác*: chọn lần lượt các quan hệ rồi chọn cột trả về. Dòng chữ xanh bên dưới cho thấy toàn bộ đường tra cứu, ví dụ `Nhân viên.Mã PB → Phòng ban.Mã PB → Phòng ban.Tên phòng ban`.
   - *Biểu thức*, *Theo điều kiện*, *Giá trị đầu tiên không rỗng*, *Giá trị cố định*.

   Thêm các bước biến đổi nếu cần: viết hoa, định dạng ngày, tiền, đọc số thành chữ… Bấm *Thử mapping* với một dòng Excel để xem giá trị và nguồn gốc của nó.
5. **Template Word**
   - Chọn file Word của bạn (`.docx`) bằng nút *Nhập…* ở trang này, hoặc nút **Thêm file Word…** ở trang *Sinh tài liệu*. Ứng dụng sao chép file vào profile và không bao giờ sửa file gốc. File `.doc` cũ phải được lưu lại thành `.docx` bằng Word trước.
   - Chọn kiểu đánh dấu (xem mục 2) và kiểm tra bảng *Placeholder phát hiện được*. Bảng này liệt kê cả các **chỗ nhập liệu có sẵn** trong file (xem mục 2a).
   - Chỗ nào *chưa gắn* thì gắn với một trường ở cột *Gắn với trường*.
   - Sắp thứ tự các template bằng nút ↑ ↓.
   - Nếu template có hướng trang hoặc header/footer riêng, chọn "Khi gộp: giữ nguyên".
6. **Xem trước.** Kiểm tra dữ liệu, các vấn đề và danh sách file sẽ được tạo.
7. **Sinh tài liệu**
   - Chọn thư mục kết quả và chế độ sinh.
   - Bấm *Kiểm tra* trước, rồi *Sinh tài liệu*. Có thể bấm *Hủy* bất cứ lúc nào: tài liệu đang làm sẽ hoàn tất, các tài liệu sau sẽ không được tạo.
8. Nhấn **Ctrl+S** để lưu profile. Mỗi lần lưu, phiên bản cũ được giữ trong thư mục `history`.

## 2. Cách đánh dấu chỗ cần điền trong Word

| Kiểu | Cách viết trong Word | Khi nào dùng |
|---|---|---|
| Jinja (khuyên dùng) | `{{ nhan_vien.ho_ten }}` | Mọi trường hợp, đầy đủ tính năng |
| Ký hiệu tùy chỉnh | `<<Họ tên>>` (ký hiệu mở/đóng tự chọn) | File mẫu khách hàng đã có sẵn ký hiệu |
| MERGEFIELD | Field Mail Merge `«Ho_ten»` | File mẫu đã làm bằng Mail Merge của Word |

Với ký hiệu tùy chỉnh và MERGEFIELD, tên trong ký hiệu được gắn với trường dữ liệu ở trang *Template Word*. Nếu tên trùng với tên trường thì ứng dụng tự gắn và có thông báo để bạn kiểm tra.

### 2a. File Word có sẵn chưa đánh dấu gì (biểu mẫu, phiếu, hợp đồng mẫu)

Không cần sửa file Word của khách hàng. Ứng dụng tự tìm 3 loại chỗ nhập sau, kể cả trong header và footer:

| Loại chỗ nhập | Ví dụ trong file Word | Tên hiển thị trong ứng dụng |
|---|---|---|
| **Chỗ trống sau nhãn** | `Họ và tên: ..................`, `Sinh ngày ...... tháng ...... năm ........`, `Địa chỉ: ________`, gạch chân khoảng trắng; ô bảng trống nằm bên phải ô nhãn (`Số CCCD │ [trống]`) | Chữ của nhãn: `Họ và tên`, `tháng`; nhãn lặp lại lần sau sẽ là `tháng (2)` |
| **Ô nhập liệu (content control)** | Tạo bằng Developer → Controls: ô chữ, ô ngày, danh sách chọn, ô đánh dấu ☐ | *Tag* của ô, nếu không có tag thì dùng *Title* |
| **Form field kiểu cũ** | Ô chữ / ô đánh dấu của Legacy Forms | Tên bookmark của field |

Cách dùng:
1. Mở trang **Template Word** và chọn file. Bảng hiện từng chỗ nhập kèm loại, trạng thái (*sẽ được điền* / *giữ nguyên*) và **nội dung xung quanh**, để bạn biết `tháng` nào là tháng sinh, `tháng` nào là ngày ký.
2. Ở cột *Gắn với trường*, chọn trường dữ liệu cho từng chỗ cần điền. Bảng được quét lại ngay và trạng thái đổi thành *sẽ được điền*.
3. Chỗ trống **chỉ được điền khi bạn đã gắn**. Dòng chữ ký, "Hà Nội, ngày ... tháng ... năm ...", ô "dành cho cán bộ"… mà bạn không gắn thì được giữ nguyên như file gốc.
4. Ô nhập liệu và form field có tag hoặc tên trùng với tên trường thì được tự gắn. Ứng dụng có thông báo để bạn kiểm tra.
5. Ô đánh dấu nhận giá trị kiểu `Có/Không`, `x`, `1/0`, `TRUE/FALSE`: có giá trị thì thành ☒, để trống hoặc "Không" thì thành ☐.

Nếu chỗ trống ngắn quá không được nhận ra, hoặc dấu chấm trong câu bị nhận nhầm, chỉnh ô *số dấu chấm/gạch tối thiểu* (mặc định 5) rồi bấm *Áp dụng và quét lại*. Có thể tắt từng loại ở *Nhận diện thêm*.

Định dạng chữ của nhãn và chỗ trống được giữ nguyên. Ô nhập liệu vẫn là ô nhập liệu trong file kết quả, nên người dùng có thể sửa tiếp trong Word.

**Chưa hỗ trợ:** danh sách chọn kiểu cũ (FORMDROPDOWN) và ô ảnh (picture content control). Hai loại này được báo lại và giữ nguyên. Bảng kiểu danh sách (một dòng tiêu đề, nhiều dòng trống) cần dùng `{%tr for %}` (xem bên dưới).

### Bảng có số dòng thay đổi (mỗi dự án một dòng)

Tạo bảng 3 dòng trong Word:

| | | |
|---|---|---|
| `{%tr for p in nhan_vien.du_an %}` | | |
| `{{ p.ten }}` | `{{ p.vai_tro }}` | `{{ p.so_gio }}` |
| `{%tr endfor %}` | | |

Dòng 1 và dòng 3 sẽ biến mất khi sinh tài liệu. Dòng 2 được lặp lại cho mỗi dự án.

### Đoạn văn chỉ hiện khi có điều kiện

Mỗi thẻ phải nằm **một mình trên một đoạn**:

```
{%p if nhan_vien.du_an %}
... nội dung chỉ hiện khi có dự án ...
{%p else %}
Nhân viên chưa tham gia dự án nào.
{%p endif %}
```

### Danh sách gạch đầu dòng lặp

```
{%p for p in nhan_vien.du_an %}
{{ p.ten }} – {{ p.vai_tro }}        ← định dạng đoạn này là List Bullet / List Number
{%p endfor %}
```

### Định dạng ngay trong Word

- `{{ hop_dong.gia_tri | format_currency }}` cho ra `125.000.000 ₫`.
- `{{ hop_dong.gia_tri | number_to_words_vi(suffix="đồng") }}` cho ra `Một trăm hai mươi lăm triệu đồng`.
- `{{ ngay | format_date("%d/%m/%Y") }}`, `{{ ten | upper }}`, `{{ ghi_chu | default("—", true) }}` (hiện "—" khi rỗng).

### Dữ liệu chung của lần sinh

- `{{ report.generated_date }}` là ngày lập.
- `{{ report.profile.name }}` là tên profile.
- `{{ unit.index }}` / `{{ unit.count }}` là số thứ tự / tổng số tài liệu.

### Lưu ý khi soạn template

- Định dạng chữ (font, đậm, màu) của placeholder chính là định dạng của giá trị được điền vào.
- Không cần lo Word tách placeholder thành nhiều đoạn chữ nhỏ (khi sửa chính tả, đổi định dạng): ứng dụng xử lý được.
- Ở chế độ "giữ nguyên", header/footer, hướng trang ngang/dọc và lề của từng template được giữ khi gộp.
- Mục lục (TOC) được Word cập nhật khi mở file; Word sẽ hỏi cập nhật trường.
- **Không hỗ trợ:** vùng MERGEFIELD `TableStart`/`TableEnd` (hãy dùng `{%tr for %}`), và MERGEFIELD lồng trong trường IF.

## 3. Các chế độ sinh tài liệu

| Chế độ | Ý nghĩa |
|---|---|
| A | Một dòng Excel → một tài liệu (nhập số dòng ở ô *Bản ghi*) |
| B | Mỗi dòng → các tài liệu riêng theo từng template |
| C | Nhóm các dòng liên quan (ví dụ theo phòng ban) → một tài liệu cho mỗi nhóm. Profile cần `group_by`; template dùng `{% for r in group.records %}` |
| D | Tất cả các dòng → một tài liệu gộp |
| E | Mỗi dòng (hoặc nhóm) → một tài liệu gộp từ nhiều template |

Tên file lấy theo mẫu trong profile, ví dụ `HoSo_{nhan_vien.ma}.docx`. Dấu `/` trong mẫu sẽ tạo thư mục con, ví dụ `{hop_dong.chi_nhanh}/HD_{hop_dong.so}.docx`. File đã có **không bị ghi đè**: file mới được thêm số thứ tự, ví dụ `(2)`.

## 4. Khi có lỗi

Trang **Nhật ký và lỗi** cho biết với mỗi lỗi: sheet, dòng, cột, quy tắc, placeholder và gợi ý cách sửa.

| Lỗi thường gặp | Nguyên nhân và cách xử lý |
|---|---|
| Không tìm thấy dòng nào trong sheet đích khớp khóa | Mã ở sheet chính không có trong sheet tra cứu. Sửa dữ liệu, hoặc đặt `on_missing: default` và giá trị mặc định. |
| Khóa xuất hiện nhiều lần | Sheet tra cứu có mã trùng. Xóa dòng trùng, dùng khóa ghép, hoặc chọn rõ `match: first`. |
| Ô công thức chưa có kết quả | File Excel được tạo bằng phần mềm khác nên chưa có kết quả tính sẵn. Mở file bằng Excel, để Excel tính lại, rồi **Lưu**. |
| Placeholder không có trong dữ liệu | Tên trong Word không khớp với trường nào. Sửa tên hoặc gắn placeholder ở trang Template Word. |
| File đang được mở | Hãy đóng file trong Word. Ứng dụng không ghi đè file đang mở. |

Sau mỗi lần sinh, thư mục kết quả có thêm 2 file:
- `docgen-audit-<mã>.json`: thông tin lần chạy, đầu vào kèm mã băm, số lượng, lỗi, danh sách file.
- `docgen-report-<mã>.csv`: mỗi tài liệu một dòng, mở được bằng Excel.

## 5. Bảo mật

- Ứng dụng chạy hoàn toàn trên máy: không gửi dữ liệu lên mạng, không thu thập thông tin sử dụng.
- File Excel và Word gốc chỉ được đọc, không bao giờ bị sửa.
- File log chỉ ghi thông tin lần chạy (số lượng, thời gian, tên file), không ghi nội dung ô dữ liệu.
