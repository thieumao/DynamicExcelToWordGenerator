# Hướng dẫn sử dụng DocGen

DocGen tạo hồ sơ, biên bản, phiếu… bằng Word từ dữ liệu trong Excel. Bạn nhập dữ liệu một lần trong Excel; mỗi dòng sẽ thành một bộ tài liệu Word đã điền sẵn, sẵn sàng để in.

- **Phần A** dành cho mọi người dùng, không cần biết IT. Bản đầu tiên tập trung cho **hồ sơ nghiệm thu ngành Xây dựng**.
- **Phần B** dành cho người cấu hình các trường hợp phức tạp (chế độ nâng cao).

---

# Phần A – Dùng hằng ngày

## A1. Cài đặt

1. Tải file **`DocGen-Setup-…-x64.exe`** (người quản lý gửi cho bạn, hoặc tải ở mục *Actions → Build Windows → Artifacts* trên GitHub).
2. Bấm đúp vào file và bấm *Next* cho đến hết.
   - **Không cần quyền quản trị (admin)**: ứng dụng được cài cho tài khoản của bạn.
   - Có thể tích ô tạo biểu tượng **DocGen** trên màn hình Desktop.
3. Nếu Windows hiện cảnh báo màu xanh "Windows protected your PC", bấm **More info → Run anyway**. Cảnh báo này xuất hiện vì bộ cài chưa được ký số.
4. Mở DocGen từ menu Start hoặc biểu tượng trên Desktop.
5. Khi cần gỡ: *Start → DocGen → Gỡ cài đặt DocGen*. Dữ liệu và tài liệu bạn đã tạo vẫn được giữ lại.

Ứng dụng chạy hoàn toàn trên máy, **không cần Internet**, không cần cài Microsoft Office.

## A2. Thử ngay trong 1 phút

1. Mở DocGen. Ở **Trang chủ**, bấm **3. Dùng thử hồ sơ mẫu**.
2. Ứng dụng tạo thư mục `Tài liệu (Documents)\DocGen\Mẫu - Hồ sơ nghiệm thu` gồm:
   - `data\HoSoNghiemThu.xlsx`: file Excel mẫu (thông tin công trình, danh mục công việc nghiệm thu, nhân sự);
   - `templates\`: hai mẫu Word *Phiếu yêu cầu nghiệm thu* và *Biên bản nghiệm thu công việc xây dựng*.
3. Ứng dụng tạo luôn 6 bộ hồ sơ (mỗi công việc một file) và **một file tổng hợp để in**, lưu trong thư mục `Kết quả`. Bấm **Mở thư mục kết quả** để xem.
4. Mở file Excel và các mẫu Word ở trên để xem cách nhập dữ liệu và cách đánh dấu chỗ cần điền. Bạn có thể sửa thử file Excel rồi tạo lại.

## A3. Chuẩn bị file Excel

Một file Excel thường có 3 loại sheet. Tên sheet và tên cột do bạn đặt, ứng dụng tự nhận ra.

**1. Sheet danh sách: mỗi dòng là một bộ tài liệu** (ví dụ mỗi công việc cần nghiệm thu là một biên bản). Dòng tiêu đề có thể nằm ở dòng 1, hoặc dòng 3 nếu phía trên có tên bảng.

| STT | Số biên bản | Tên công việc | Vị trí | Ngày nghiệm thu | Giờ bắt đầu | Mã giám sát | Mã kỹ thuật |
|---|---|---|---|---|---|---|---|
| 1 | 01/NTCV | Lắp dựng cốt thép cột tầng 1 | Trục A–D / 1–5 | 02/04/2024 | 8:00 | GS01 | KT01 |
| 2 | 02/NTCV | Lắp dựng ván khuôn cột tầng 1 | Trục A–D / 1–5 | 03/04/2024 | 8:00 | GS01 | KT01 |

**2. Sheet thông tin chung: nhập một lần, dùng cho mọi tài liệu.** Gồm hai cột *tên mục | nội dung*:

| | |
|---|---|
| Tên công trình | Nhà làm việc 5 tầng |
| Địa điểm xây dựng | Số 12 đường Lê Văn Lương, Hà Nội |
| Chủ đầu tư | Công ty Cổ phần Đầu tư DEF |
| Nhà thầu thi công | Công ty Cổ phần Xây dựng ABC |

**3. Sheet tra cứu**, ví dụ *Nhân sự*. Sheet danh sách chỉ cần ghi mã (`GS01`); họ tên và chức vụ được lấy từ sheet này.

| Mã | Họ và tên | Chức vụ |
|---|---|---|
| GS01 | Nguyễn Văn Hùng | Kỹ sư tư vấn giám sát |
| KT01 | Trần Minh Tuấn | Cán bộ kỹ thuật thi công |

Mẹo:
- Ngày nên nhập đúng kiểu ngày của Excel (ví dụ `02/04/2024`), giờ nhập kiểu `8:00`.
- Ảnh (ví dụ ảnh 3x4): đặt file ảnh trong một thư mục cạnh file Excel, rồi ghi tên file vào cột, ví dụ `anh/NV01.png`.
- Ô đánh dấu (☒/☐): ghi `Có`/`Không` hoặc `x`.
- Không cần xóa các dòng tiêu đề, ghi chú phía trên bảng.

## A4. Chuẩn bị file Word mẫu

Dùng chính biểu mẫu Word của đơn vị bạn. Có hai cách đánh dấu chỗ cần điền; một file có thể dùng cả hai.

**Cách 1 (dễ nhất): gõ tên thông tin giữa `<<` và `>>`** ở đúng chỗ cần điền, ví dụ:

```
Công trình: <<Tên công trình>>
Đối tượng nghiệm thu: <<Tên công việc>> – Vị trí: <<Vị trí>>
Đại diện tư vấn giám sát: Ông/Bà <<Họ tên giám sát>> – Chức vụ: <<Chức vụ giám sát>>
Bắt đầu: <<Giờ bắt đầu>> <<Ngày nghiệm thu bằng chữ>>
```

- Nên đặt tên giống tên cột trong Excel. Ứng dụng tự ghép, không phân biệt chữ hoa/thường hay có dấu/không dấu.
- Thông tin tra theo mã thì ghép tên cột với vai trò: `<<Họ tên giám sát>>`, `<<Chức vụ kỹ thuật>>`.
- Ngày tháng:
  - `<<Ngày nghiệm thu>>` → `02/04/2024`;
  - `<<Ngày nghiệm thu bằng chữ>>` → `ngày 02 tháng 04 năm 2024`.
- Giờ: chọn dạng `08 giờ 00` ở bước ghép (xem A5).
- Định dạng chữ (font, đậm, nghiêng) của `<<…>>` là định dạng của nội dung được điền vào.

**Cách 2: giữ nguyên biểu mẫu có sẵn.** Ứng dụng tự tìm các chỗ sau, kể cả trong header và footer:

| Loại chỗ nhập | Ví dụ trong file Word |
|---|---|
| Chỗ trống sau nhãn | `Họ và tên: ..................`, `Sinh ngày ...... tháng ...... năm ........`, `Địa chỉ: ________`; ô bảng trống nằm bên phải ô nhãn |
| Ô nhập liệu của Word | Tạo bằng *Developer → Controls*: ô chữ, ô ngày, danh sách chọn, ô đánh dấu ☐, khung ảnh |
| Form field kiểu cũ | Ô chữ, ô đánh dấu, danh sách chọn của *Legacy Forms* |

- Chỗ trống **chỉ được điền khi bạn chọn dữ liệu cho nó**. Dòng chữ ký, "Hà Nội, ngày … tháng … năm …", ô "Ý kiến khác"… mà bạn không chọn thì được giữ nguyên để viết tay.
- Danh sách chọn (ví dụ Giới tính: Nam / Nữ): ứng dụng chọn mục trùng với giá trị trong Excel. Nếu Excel ghi giá trị không có trong danh sách, ứng dụng báo lỗi cho dòng đó chứ không chọn bừa.
- Khung ảnh được điền bằng ảnh, co giãn vừa khung và giữ đúng tỉ lệ.

File Word gốc **không bao giờ bị sửa**: ứng dụng làm việc trên bản sao.

## A5. Tạo bộ hồ sơ mới (4 bước)

Ở **Trang chủ**, bấm **1. Tạo bộ hồ sơ mới** (hoặc mục *Tạo hồ sơ nhanh* bên trái):

1. **Chọn file Excel.** Ứng dụng đọc file và cho biết sheet nào là danh sách, sheet nào là thông tin chung, sheet nào là bảng tra cứu. Nếu đoán sai, đổi lại ở cột *Dùng làm*.
2. **Chọn file Word mẫu.** Có thể chọn nhiều file, ví dụ phiếu yêu cầu và biên bản. Dùng nút ↑ ↓ để sắp thứ tự trang khi gộp.
3. **Kiểm tra chỗ cần điền.** Bảng liệt kê mọi chỗ cần điền tìm được trong các file Word:
   - **Màu xanh**: đã tự ghép với dữ liệu. Cột *Ví dụ* cho thấy giá trị lấy từ dòng đầu tiên của Excel.
   - **Màu vàng**: chưa ghép. Nếu cần điền thì chọn giá trị ở cột *Lấy từ Excel*. Ví dụ, với `Sinh ngày … tháng … năm …` thì chọn *Ngày sinh – ngày*, *– tháng*, *– năm*.
   - Chỗ màu vàng không chọn gì thì được giữ nguyên (chỗ trống) hoặc để trống (`<<…>>`).
4. **Tạo tài liệu.** Đặt tên bộ hồ sơ (mặc định là tên công trình), chọn:
   - *Mỗi dòng Excel → 1 file Word* (gồm tất cả các mẫu), hoặc *mỗi mẫu một file riêng*;
   - *Tạo thêm 1 file tổng hợp tất cả để in*;
   - thư mục lưu kết quả.

   Bấm **Lưu và tạo tài liệu ✔**. Ứng dụng chuyển sang trang *Tạo tài liệu* và hiện tiến trình. Xong, bấm **Mở thư mục kết quả**.

Tên file kết quả được đánh số theo dòng, ví dụ `01 - Lắp dựng cốt thép cột tầng 1.docx`. File đã có **không bị ghi đè**; file mới được thêm số, ví dụ `(2)`.

## A6. Những lần sau

Bộ hồ sơ đã được lưu lại. Khi có thêm công việc cần nghiệm thu:

1. Thêm dòng mới vào file Excel rồi lưu lại.
2. Ở **Trang chủ**, chọn bộ hồ sơ trong danh sách *Bộ hồ sơ đã lưu* và bấm **2. Dùng bộ hồ sơ đã lưu** (hoặc bấm đúp vào tên).
3. Ở trang **Tạo tài liệu**, kiểm tra *File Excel đầu vào* và *Thư mục kết quả*, rồi bấm **Tạo tài liệu**.
   - Chỉ cần tạo một vài dòng thì ghi số dòng Excel vào ô *Bản ghi*, ví dụ `5,7-9`.

## A7. Khi có lỗi

Trang **Tạo tài liệu** (và trang **Lỗi và cảnh báo**) liệt kê mọi lỗi bằng tiếng Việt, kèm sheet, dòng, cột và gợi ý cách sửa. Tài liệu của các dòng không lỗi vẫn được tạo bình thường.

| Thông báo thường gặp | Cách xử lý |
|---|---|
| Không tìm thấy dòng nào trong sheet đích khớp khóa … | Mã (ví dụ `GS03`) ở sheet danh sách chưa có trong sheet tra cứu (*Nhân sự*). Thêm người đó vào sheet *Nhân sự* hoặc sửa mã. |
| Sheet thông tin chung … chưa có dữ liệu | Điền thông tin công trình vào sheet thông tin chung. |
| Giá trị '…' không có trong danh sách chọn … | Sửa giá trị trong Excel cho trùng một mục của danh sách chọn trong file Word. |
| Không ghi được file kết quả / file đang mở | Đóng file đó trong Word rồi bấm *Tạo tài liệu* lại. |
| Ô công thức chưa có kết quả được lưu | Mở file Excel bằng Excel, đợi tính xong, bấm **Lưu**, rồi tạo lại. |
| Không chèn được ảnh … | Kiểm tra tên file ảnh trong Excel và thư mục ảnh cạnh file Excel. |

Sau mỗi lần tạo, thư mục kết quả có thêm:
- `docgen-report-<mã>.csv`: mở bằng Excel, mỗi tài liệu một dòng, có lỗi đầu tiên và gợi ý sửa;
- `docgen-audit-<mã>.json`: nhật ký chi tiết của lần chạy.

## A8. Bảo mật

- Ứng dụng chạy hoàn toàn trên máy: không gửi dữ liệu lên mạng, không thu thập thông tin sử dụng.
- File Excel và Word gốc chỉ được đọc, không bao giờ bị sửa.
- File log chỉ ghi thông tin lần chạy (số lượng, thời gian, tên file), không ghi nội dung ô dữ liệu.

---

# Phần B – Dành cho người cấu hình (chế độ nâng cao)

Bật **Cài đặt → Chế độ nâng cao** để hiện các trang *Cấu hình Excel*, *Tra cứu giữa các sheet*, *Ghép dữ liệu*, *File Word mẫu* và *Xem trước*. Các trang này dùng cho trường hợp phức tạp: khóa ghép nhiều cột, bảng lặp, điều kiện, biểu thức, sinh theo nhóm…

Bộ hồ sơ tạo bằng *Tạo hồ sơ nhanh* cũng mở và chỉnh được ở các trang này.

## B1. Cấu hình chi tiết một bộ hồ sơ


1. **Bộ hồ sơ đã lưu → Tạo mới** (ở chế độ nâng cao). Nhập tên. Mỗi khách hàng hoặc mỗi loại hồ sơ có một bộ riêng.
2. **Cấu hình Excel**
   - Chọn file Excel rồi bấm *Đọc file Excel*. Ứng dụng tự liệt kê các sheet và đoán dòng tiêu đề (dòng được tô xanh trong bảng xem trước).
   - Chọn từng sheet cần dùng và bấm *Thêm sheet đã chọn vào bộ hồ sơ*.
   - Đặt sheet chứa dữ liệu chính (mỗi dòng ứng với một hồ sơ hoặc hợp đồng…) làm *sheet chính*.
   - Kiểm tra lại bảng cột:
     - **ID** là tên nội bộ, không dấu, ví dụ `ma_nv`.
     - **Tiêu đề cột** là chữ đúng như trong Excel.
     - **Tên khác** ghi các cách viết khác, cách nhau bằng dấu phẩy, ví dụ `Mã NV, Ma NV`.
     - **Kiểu** là kiểu dữ liệu: date, decimal…
   - Bấm *Áp dụng thay đổi sheet*.
3. **Tra cứu giữa các sheet.** Khai báo cách nối mã: sheet nguồn, sheet đích, cặp cột khóa. Nếu cần khóa ghép (ví dụ chi nhánh + mã phòng) thì thêm nhiều cặp. Bấm *Thử tất cả quan hệ* để xem số khóa và số khóa bị trùng.
4. **Ghép dữ liệu.** Tạo các trường mà file Word sẽ dùng, ví dụ `nhan_vien.phong_ban`. Có các cách lấy giá trị:
   - *Cột của sheet này*: lấy thẳng từ sheet chính.
   - *Tra cứu ở sheet khác*: chọn lần lượt các quan hệ rồi chọn cột trả về. Dòng chữ xanh bên dưới cho thấy toàn bộ đường tra cứu, ví dụ `Nhân viên.Mã PB → Phòng ban.Mã PB → Phòng ban.Tên phòng ban`.
   - *Biểu thức*, *Theo điều kiện*, *Giá trị đầu tiên không rỗng*, *Giá trị cố định*.

   Thêm các bước biến đổi nếu cần: viết hoa, định dạng ngày, tiền, đọc số thành chữ… Bấm *Thử ghép dữ liệu* với một dòng Excel để xem giá trị và nguồn gốc của nó.
5. **File Word mẫu**
   - Chọn file Word của bạn (`.docx`) bằng nút *Nhập…* ở trang này, hoặc nút **Thêm file Word…** ở trang *Tạo tài liệu*. Ứng dụng sao chép file vào bộ hồ sơ và không bao giờ sửa file gốc. File `.doc` cũ phải được lưu lại thành `.docx` bằng Word trước.
   - Chọn kiểu đánh dấu (xem mục 2) và kiểm tra bảng *Placeholder phát hiện được*. Bảng này liệt kê cả các **chỗ nhập liệu có sẵn** trong file (xem mục 2a).
   - Chỗ nào *chưa gắn* thì gắn với một trường ở cột *Gắn với trường*.
   - Sắp thứ tự các template bằng nút ↑ ↓.
   - Nếu template có hướng trang hoặc header/footer riêng, chọn "Khi gộp: giữ nguyên".
6. **Xem trước.** Kiểm tra dữ liệu, các vấn đề và danh sách file sẽ được tạo.
7. **Tạo tài liệu**
   - Chọn thư mục kết quả và chế độ sinh.
   - Bấm *Kiểm tra* trước, rồi *Tạo tài liệu*. Có thể bấm *Hủy* bất cứ lúc nào: tài liệu đang làm sẽ hoàn tất, các tài liệu sau sẽ không được tạo.
8. Nhấn **Ctrl+S** để lưu bộ hồ sơ. Mỗi lần lưu, phiên bản cũ được giữ trong thư mục `history`.


## B2. Các kiểu đánh dấu trong Word

| Kiểu | Cách viết trong Word | Khi nào dùng |
|---|---|---|
| Jinja (khuyên dùng) | `{{ nhan_vien.ho_ten }}` | Mọi trường hợp, đầy đủ tính năng |
| Ký hiệu tùy chỉnh | `<<Họ tên>>` (ký hiệu mở/đóng tự chọn) | File mẫu khách hàng đã có sẵn ký hiệu |
| MERGEFIELD | Field Mail Merge `«Ho_ten»` | File mẫu đã làm bằng Mail Merge của Word |

Với ký hiệu tùy chỉnh và MERGEFIELD, tên trong ký hiệu được gắn với trường dữ liệu ở trang *File Word mẫu*. Nếu tên trùng với tên trường thì ứng dụng tự gắn và có thông báo để bạn kiểm tra.

Biểu mẫu có chỗ trống, ô nhập liệu và form field (mục A4) được gắn dữ liệu ở trang **File Word mẫu**:
- Cột *Gắn với trường* chọn dữ liệu cho từng chỗ.
- *Nhận diện thêm* bật/tắt từng loại chỗ cần điền.
- Nếu chỗ trống ngắn quá không được nhận ra, hoặc dấu chấm trong câu bị nhận nhầm, chỉnh ô *số dấu chấm/gạch tối thiểu* (mặc định 5) rồi bấm *Áp dụng và quét lại*.
- Bảng kiểu danh sách (một dòng tiêu đề, nhiều dòng trống) không được nhận là chỗ trống; hãy dùng `{%tr for %}` (xem bên dưới).

### Bảng có số dòng thay đổi (mỗi dự án một dòng)

Tạo bảng 3 dòng trong Word:

| | | |
|---|---|---|
| `{%tr for p in nhan_vien.du_an %}` | | |
| `{{ p.ten }}` | `{{ p.vai_tro }}` | `{{ p.so_gio }}` |
| `{%tr endfor %}` | | |

Dòng 1 và dòng 3 sẽ biến mất khi tạo tài liệu. Dòng 2 được lặp lại cho mỗi dự án.

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
- `{{ report.profile.name }}` là tên bộ hồ sơ.
- `{{ unit.index }}` / `{{ unit.count }}` là số thứ tự / tổng số tài liệu.

### Lưu ý khi soạn template

- Định dạng chữ (font, đậm, màu) của placeholder chính là định dạng của giá trị được điền vào.
- Không cần lo Word tách placeholder thành nhiều đoạn chữ nhỏ (khi sửa chính tả, đổi định dạng): ứng dụng xử lý được.
- Ở chế độ "giữ nguyên", header/footer, hướng trang ngang/dọc và lề của từng template được giữ khi gộp.
- Mục lục (TOC) được Word cập nhật khi mở file; Word sẽ hỏi cập nhật trường.
- **Không hỗ trợ:** vùng MERGEFIELD `TableStart`/`TableEnd` (hãy dùng `{%tr for %}`), và MERGEFIELD lồng trong trường IF.


## B3. Các chế độ tạo tài liệu

| Chế độ | Ý nghĩa |
|---|---|
| A | Một dòng Excel → một tài liệu (nhập số dòng ở ô *Bản ghi*) |
| B | Mỗi dòng → các tài liệu riêng theo từng template |
| C | Nhóm các dòng liên quan (ví dụ theo phòng ban) → một tài liệu cho mỗi nhóm. Bộ hồ sơ cần `group_by`; template dùng `{% for r in group.records %}` |
| D | Tất cả các dòng → một tài liệu gộp |
| E | Mỗi dòng (hoặc nhóm) → một tài liệu gộp từ nhiều template |

Tên file lấy theo mẫu trong bộ hồ sơ, ví dụ `HoSo_{nhan_vien.ma}.docx`. Dấu `/` trong mẫu sẽ tạo thư mục con, ví dụ `{hop_dong.chi_nhanh}/HD_{hop_dong.so}.docx`. File đã có **không bị ghi đè**: file mới được thêm số thứ tự, ví dụ `(2)`.


File tổng hợp để in (`output.combined_filename` trong bộ hồ sơ, hoặc ô *Tạo thêm 1 file tổng hợp tất cả để in*) gộp mọi tài liệu đã tạo theo thứ tự thành một file.

## B4. Lỗi cấu hình thường gặp

Trang **Lỗi và cảnh báo** cho biết với mỗi lỗi: sheet, dòng, cột, quy tắc, placeholder và gợi ý cách sửa.

| Lỗi thường gặp | Nguyên nhân và cách xử lý |
|---|---|
| Không tìm thấy dòng nào trong sheet đích khớp khóa | Mã ở sheet chính không có trong sheet tra cứu. Sửa dữ liệu, hoặc đặt `on_missing: default` và giá trị mặc định. |
| Khóa xuất hiện nhiều lần | Sheet tra cứu có mã trùng. Xóa dòng trùng, dùng khóa ghép, hoặc chọn rõ `match: first`. |
| Ô công thức chưa có kết quả | File Excel được tạo bằng phần mềm khác nên chưa có kết quả tính sẵn. Mở file bằng Excel, để Excel tính lại, rồi **Lưu**. |
| Placeholder không có trong dữ liệu | Tên trong Word không khớp với trường nào. Sửa tên hoặc gắn placeholder ở trang File Word mẫu. |
| File đang được mở | Hãy đóng file trong Word. Ứng dụng không ghi đè file đang mở. |

Sau mỗi lần sinh, thư mục kết quả có thêm 2 file:
- `docgen-audit-<mã>.json`: thông tin lần chạy, đầu vào kèm mã băm, số lượng, lỗi, danh sách file.
- `docgen-report-<mã>.csv`: mỗi tài liệu một dòng, mở được bằng Excel.

