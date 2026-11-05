"""Reproducible sample customers.

* **customer_a** – the structure from the requirements (Employees + Departments
  + Positions + Managers + Projects), Jinja placeholders, three templates
  (cover, landscape details with a repeating table, appendix).
* **customer_b** – a different customer: Vietnamese sheet names and headers,
  a title block above the header (header on row 4), a composite key
  (branch + department code), ``<<...>>`` markers and MERGEFIELDs.

Both run through the same code; only the profile differs.
"""

from __future__ import annotations

import datetime as dt
import random
from pathlib import Path

from docx import Document
from docx.document import Document as DocxDocument
from docx.enum.section import WD_ORIENT
from docx.shared import Mm, Pt

from docgen.testing.docx_builder import add_merge_field, add_page_number_field, tiny_png
from docgen.testing.xlsx_builder import build_xlsx

FAMILY = ["Nguyễn", "Trần", "Lê", "Phạm", "Hoàng", "Vũ", "Đặng", "Bùi", "Đỗ", "Hồ"]
MIDDLE = ["Văn", "Thị", "Minh", "Ngọc", "Đức", "Thu", "Hữu", "Thanh"]
GIVEN = [
    "Anh",
    "Bình",
    "Châu",
    "Dũng",
    "Hà",
    "Hùng",
    "Lan",
    "Long",
    "Mai",
    "Nam",
    "Phương",
    "Quân",
    "Sơn",
    "Thảo",
    "Trang",
    "Tuấn",
    "Vy",
    "Yến",
]


def vn_name(rng: random.Random) -> str:
    return f"{rng.choice(FAMILY)} {rng.choice(MIDDLE)} {rng.choice(GIVEN)}"


# ===================================================================== A
CUSTOMER_A_PROFILE = """\
schema_version: 1
profile:
  id: customer_a
  name: "Customer A – Hồ sơ nhân sự"
  version: 1
  description: Employee dossier generated from the HR workbook.
formatting:
  date_format: "%d/%m/%Y"
  decimal_separator: ","
  thousands_separator: "."
  currency_symbol: "₫"
excel:
  primary: emp
  formula_policy: cached_required
  sheets:
    - id: emp
      match: {name: Employees, aliases: ["Nhân viên"]}
      header: {mode: auto}
      key: [employee_id]
      columns:
        - {id: employee_id, header: EmployeeID, type: string}
        - {id: full_name, header: FullName}
        - {id: department_id, header: DepartmentID, type: string}
        - {id: position_id, header: PositionID, type: string}
        - {id: start_date, header: StartDate, type: date}
        - {id: salary, header: Salary, type: decimal}
        - {id: grade, header: Grade, type: integer}
        - {id: email, header: Email, required: false}
    - id: dept
      match: {name: Departments}
      columns:
        - {id: department_id, header: DepartmentID, type: string}
        - {id: name, header: DepartmentName}
        - {id: manager_id, header: ManagerID, type: string}
    - id: pos
      match: {name: Positions}
      columns:
        - {id: position_id, header: PositionID, type: string}
        - {id: name, header: PositionName}
    - id: mgr
      match: {name: Managers}
      columns:
        - {id: manager_id, header: ManagerID, type: string}
        - {id: name, header: ManagerName}
    - id: proj
      match: {name: Projects}
      columns:
        - {id: employee_id, header: EmployeeID, type: string}
        - {id: project, header: ProjectName}
        - {id: role, header: Role}
        - {id: hours, header: Hours, type: decimal}
relationships:
  - {id: emp_dept, from: emp, to: dept, keys: [{from: department_id, to: department_id}],
     cardinality: many_to_one, on_missing: error, on_duplicate: error}
  - {id: emp_pos, from: emp, to: pos, keys: [{from: position_id, to: position_id}],
     cardinality: many_to_one, on_missing: default, default: "(chưa xác định)"}
  - {id: dept_mgr, from: dept, to: mgr, keys: [{from: manager_id, to: manager_id}],
     cardinality: many_to_one}
  - {id: emp_proj, from: emp, to: proj, keys: [{from: employee_id, to: employee_id}],
     cardinality: one_to_many, sort: [{by: hours, order: desc}]}
fields:
  - {name: employee.code, source: employee_id, required: true}
  - {name: employee.full_name, source: full_name, transforms: [trim, title_case], required: true}
  - {name: employee.email, source: email, default: "—"}
  - {name: employee.department.name, path: emp_dept.name}
  - {name: employee.department.manager, path: emp_dept.dept_mgr.name}
  - {name: employee.position.name, path: emp_pos.name}
  - {name: employee.start_date, source: start_date, transforms: [format_date]}
  - {name: employee.salary_text, source: salary, transforms: [format_currency]}
  - {name: employee.salary_words, source: salary, transforms: [{number_to_words_vi: {suffix: "đồng"}}]}
  - name: employee.level
    when:
      - {if: "emp.grade >= 5", value: "Chuyên viên cao cấp"}
      - {if: "emp.grade >= 3", value: "Chuyên viên"}
      - {else: "Nhân viên"}
  - {name: employee.display, expr: "employee.full_name ~ ' (' ~ employee.code ~ ')'"}
collections:
  - name: employee.projects
    relationship: emp_proj
    fields:
      - {name: name, source: project}
      - {name: role, source: role, default: "Thành viên"}
      - {name: hours, source: hours, transforms: [{format_number: {decimals: 1}}]}
aggregations:
  - {name: employee.total_hours, relationship: emp_proj, op: sum, of: hours,
     transforms: [{format_number: {decimals: 1}}], default: "0"}
  - {name: employee.project_count, relationship: emp_proj, op: count}
validation:
  - {field: employee.email, rule: regex, pattern: "^(—|[^@\\\\s]+@[^@\\\\s]+)$", severity: warning}
  - {field: employee.code, rule: unique}
templates:
  - {id: cover, file: templates/Cover.docx, section: {mode: preserve}}
  - {id: details, file: templates/EmployeeDetails.docx, section: {mode: preserve}}
  - id: appendix
    file: templates/Appendix.docx
    include_if: "employee.project_count > 0"
    section: {mode: preserve}
merge:
  enabled: true
  order: [cover, details, appendix]
generation:
  unit: record
  output: merged
  sort: [{by: employee.code}]
output:
  merged_filename: "HoSo_{employee.code}_{employee.full_name}.docx"
  individual_filename: "{template.id}_{employee.code}.docx"
  on_conflict: suffix
"""


def customer_a_sheets(n_employees: int = 8, seed: int = 7) -> dict[str, list[list[object]]]:
    rng = random.Random(seed)
    depts = [
        ("D01", "Phòng Kế toán", "M01"),
        ("D02", "Phòng Công nghệ", "M02"),
        ("D03", "Phòng Nhân sự", "M03"),
        ("D04", "Phòng Kinh doanh", "M01"),
    ]
    positions = [("P1", "Nhân viên"), ("P2", "Trưởng nhóm"), ("P3", "Trưởng phòng")]
    managers = [("M01", "Phạm Quốc Bảo"), ("M02", "Lê Thị Ánh Tuyết"), ("M03", "Đỗ Minh Khoa")]
    projects_pool = [
        "Chuyển đổi số",
        "ERP 2025",
        "Kiểm toán nội bộ",
        "Tuyển dụng Q3",
        "Mở rộng miền Trung",
        "Cổng thông tin",
    ]
    employees: list[list[object]] = [
        ["DANH SÁCH NHÂN VIÊN"],
        [],
        ["EmployeeID", "FullName", "DepartmentID", "PositionID", "StartDate", "Salary", "Grade", "Email"],
    ]
    projects: list[list[object]] = [["EmployeeID", "ProjectName", "Role", "Hours"]]
    for i in range(1, n_employees + 1):
        code = f"E{i:03d}"
        dept = depts[(i - 1) % len(depts)][0]
        pos = positions[(i - 1) % len(positions)][0]
        employees.append(
            [
                code,
                "  " + vn_name(rng).lower() + " ",
                dept,
                pos,
                dt.datetime(2015 + i % 9, 1 + i % 12, 1 + i % 27),
                rng.randrange(8, 60) * 1_000_000 + rng.choice([0, 500_000]),
                1 + i % 6,
                f"{code.lower()}@example.vn" if i % 4 else None,
            ]
        )
        for p in rng.sample(projects_pool, k=i % 3):
            projects.append(
                [code, p, rng.choice(["Trưởng dự án", "Thành viên", None]), rng.choice([12, 24.5, 40, 80])]
            )
    return {
        "Employees": employees,
        "Departments": [["DepartmentID", "DepartmentName", "ManagerID"], *[list[object](d) for d in depts]],
        "Positions": [["PositionID", "PositionName"], *[list[object](p) for p in positions]],
        "Managers": [["ManagerID", "ManagerName"], *[list[object](m) for m in managers]],
        "Projects": projects,
    }


def _style_document(doc: DocxDocument) -> None:
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)


def _cover_template(path: Path) -> None:
    doc = Document()
    _style_document(doc)
    sec = doc.sections[0]
    sec.header.paragraphs[0].text = "HỒ SƠ NHÂN SỰ – {{ report.profile.name }}"
    fpar = sec.footer.paragraphs[0]
    fpar.text = "Mã hồ sơ {{ employee.code }} – Trang "
    add_page_number_field(fpar)
    doc.add_heading("HỒ SƠ NHÂN VIÊN", level=0)
    doc.add_paragraph("Họ và tên: {{ employee.full_name }}")
    doc.add_paragraph("Phòng ban: {{ employee.department.name }}")
    doc.add_paragraph("Ngày lập: {{ report.generated_date }}")
    doc.add_picture(__import__("io").BytesIO(tiny_png((20, 90, 160), (16, 16))), width=Mm(25))
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(path))


def _details_template(path: Path) -> None:
    doc = Document()
    _style_document(doc)
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = sec.page_height, sec.page_width
    sec.header.paragraphs[0].text = "CHI TIẾT – {{ employee.display }}"
    doc.add_heading("Thông tin chi tiết", level=1)
    info = doc.add_table(rows=0, cols=2)
    info.style = "Table Grid"
    for label, value in [
        ("Mã nhân viên", "{{ employee.code }}"),
        ("Họ tên", "{{ employee.full_name }}"),
        ("Chức vụ", "{{ employee.position.name }}"),
        ("Cấp bậc", "{{ employee.level }}"),
        ("Phòng ban", "{{ employee.department.name }}"),
        ("Quản lý phòng", "{{ employee.department.manager }}"),
        ("Ngày vào làm", "{{ employee.start_date }}"),
        ("Lương", "{{ employee.salary_text }} ({{ employee.salary_words }})"),
        ("Email", "{{ employee.email }}"),
    ]:
        cells = info.add_row().cells
        cells[0].text, cells[1].text = label, value
    doc.add_heading("Dự án tham gia", level=2)
    doc.add_paragraph("{%p if employee.projects %}")
    table = doc.add_table(rows=1, cols=3)
    table.style = "Table Grid"
    for cell, text in zip(table.rows[0].cells, ["Dự án", "Vai trò", "Số giờ"], strict=True):
        cell.text = text
    table.add_row().cells[0].text = "{%tr for p in employee.projects %}"
    row = table.add_row().cells
    row[0].text, row[1].text, row[2].text = "{{ p.name }}", "{{ p.role }}", "{{ p.hours }}"
    table.add_row().cells[0].text = "{%tr endfor %}"
    doc.add_paragraph("Tổng số giờ: {{ employee.total_hours }}")
    doc.add_paragraph("{%p else %}")
    doc.add_paragraph("Nhân viên chưa tham gia dự án nào.")
    doc.add_paragraph("{%p endif %}")
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(path))


def _appendix_template(path: Path) -> None:
    doc = Document()
    _style_document(doc)
    doc.sections[0].header.paragraphs[0].text = "PHỤ LỤC"
    doc.add_heading("Phụ lục: danh sách dự án ({{ employee.project_count }})", level=1)
    doc.add_paragraph("{%p for p in employee.projects %}")
    doc.add_paragraph("{{ p.name }} – {{ p.role }}", style="List Number")
    doc.add_paragraph("{%p endfor %}")
    doc.add_paragraph("Ghi chú: ký tự đặc biệt & < > được giữ nguyên.")
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(path))


def write_customer_a(directory: Path, *, n_employees: int = 8) -> tuple[Path, Path]:
    """Write profile.yaml, templates/*.docx and data/HR.xlsx; return (profile, workbook)."""
    directory.mkdir(parents=True, exist_ok=True)
    profile = directory / "profile.yaml"
    profile.write_text(CUSTOMER_A_PROFILE, encoding="utf-8")
    _cover_template(directory / "templates" / "Cover.docx")
    _details_template(directory / "templates" / "EmployeeDetails.docx")
    _appendix_template(directory / "templates" / "Appendix.docx")
    workbook = build_xlsx(directory / "data" / "HR.xlsx", customer_a_sheets(n_employees))
    return profile, workbook


# ===================================================================== B
CUSTOMER_B_PROFILE = """\
schema_version: 1
profile:
  id: customer_b
  name: "Customer B – Hợp đồng theo chi nhánh"
  version: 1
excel:
  primary: hd
  header_matching: {casefold: true, strip_accents: true}
  sheets:
    - id: hd
      match: {name: "Hợp đồng", aliases: ["HopDong"]}
      header: {mode: auto, scan_rows: 10}
      columns:
        - {id: so_hd, header: "Số HĐ", type: string}
        - {id: khach_hang, header: "Khách hàng"}
        - {id: chi_nhanh, header: "Chi nhánh", type: string}
        - {id: ma_pb, header: "Mã PB", type: string}
        - {id: gia_tri, header: "Giá trị", type: decimal}
        - {id: ngay_ky, header: "Ngày ký", type: date}
    - id: pb
      match: {name: "Phòng ban"}
      columns:
        - {id: chi_nhanh, header: "Chi nhánh", type: string}
        - {id: ma_pb, header: "Mã PB", type: string}
        - {id: ten_pb, header: "Tên phòng ban"}
        - {id: truong_phong, header: "Trưởng phòng"}
relationships:
  - id: hd_pb
    from: hd
    to: pb
    keys: [{from: chi_nhanh, to: chi_nhanh}, {from: ma_pb, to: ma_pb}]
    cardinality: many_to_one
fields:
  - {name: hop_dong.so, source: so_hd}
  - {name: hop_dong.khach_hang, source: khach_hang, transforms: [trim]}
  - {name: hop_dong.chi_nhanh, source: chi_nhanh}
  - {name: hop_dong.phong_ban, path: hd_pb.ten_pb}
  - {name: hop_dong.truong_phong, path: hd_pb.truong_phong}
  - {name: hop_dong.gia_tri, source: gia_tri, transforms: [format_currency]}
  - {name: hop_dong.gia_tri_chu, source: gia_tri, transforms: [{number_to_words_vi: {suffix: "đồng"}}]}
  - {name: hop_dong.ngay_ky, source: ngay_ky, transforms: [format_date]}
templates:
  - id: hop_dong
    file: templates/HopDong.docx
    marker: {style: custom, open: "<<", close: ">>"}
    bindings:
      "Số HĐ": hop_dong.so
      "Khách hàng": hop_dong.khach_hang
      "Giá trị": hop_dong.gia_tri
      "Bằng chữ": hop_dong.gia_tri_chu
    section: {mode: preserve}
  - id: xac_nhan
    file: templates/XacNhan.docx
    marker: {style: mergefield}
    section: {mode: preserve}
generation:
  unit: record
  output: both
output:
  merged_filename: "{hop_dong.chi_nhanh}/HD_{hop_dong.so}.docx"
  individual_filename: "{hop_dong.chi_nhanh}/{template.id}_{hop_dong.so}.docx"
"""


def customer_b_sheets() -> dict[str, list[list[object]]]:
    return {
        "Hợp đồng": [
            ["CÔNG TY TNHH ABC"],
            ["Báo cáo hợp đồng quý 3"],
            [],
            ["Số HĐ", "Khách hàng", "Chi nhánh", "Mã PB", "Giá trị", "Ngày ký"],
            ["HD-001", "Công ty Ánh Dương ", "HN", "KD", 125_000_000, dt.datetime(2024, 7, 2)],
            ["HD-002", "Cty Bình Minh", "HCM", "KD", 98_500_000, dt.datetime(2024, 8, 15)],
            ["HD-003", "Tập đoàn Sao Mai & Cộng sự", "HN", "KT", 1_250_000_000, dt.datetime(2024, 9, 1)],
        ],
        "Phòng ban": [
            ["Chi nhánh", "Mã PB", "Tên phòng ban", "Trưởng phòng"],
            ["HN", "KD", "Kinh doanh Hà Nội", "Ngô Văn Lực"],
            ["HCM", "KD", "Kinh doanh TP.HCM", "Trịnh Thu Hà"],
            ["HN", "KT", "Kế toán Hà Nội", "Phan Thanh Tâm"],
        ],
    }


def _contract_template(path: Path) -> None:
    doc = Document()
    _style_document(doc)
    doc.sections[0].header.paragraphs[0].text = "HỢP ĐỒNG SỐ <<Số HĐ>>"
    doc.add_heading("HỢP ĐỒNG DỊCH VỤ", level=0)
    par = doc.add_paragraph("Bên B: ")
    # marker split over several runs with different formatting, as Word often does
    r1 = par.add_run("<<Khách")
    r1.bold = True
    par.add_run(" hàng>>")
    doc.add_paragraph("Giá trị hợp đồng: <<Giá trị>> (Bằng chữ: <<Bằng chữ>>).")
    doc.add_paragraph("Phòng phụ trách: {{ hop_dong.phong_ban }} – Trưởng phòng {{ hop_dong.truong_phong }}")
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(path))


def _confirmation_template(path: Path) -> None:
    doc = Document()
    _style_document(doc)
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = sec.page_height, sec.page_width
    sec.header.paragraphs[0].text = "XÁC NHẬN"
    doc.add_heading("Biên bản xác nhận", level=1)
    p = doc.add_paragraph("Hợp đồng: ")
    add_merge_field(p, "hop_dong.so", bold=True)
    p = doc.add_paragraph("Ngày ký: ")
    add_merge_field(p, "ngay_ky", simple=True)
    p = doc.add_paragraph("Chi nhánh: ")
    add_merge_field(p, "chi_nhanh")
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(path))


def write_customer_b(directory: Path) -> tuple[Path, Path]:
    directory.mkdir(parents=True, exist_ok=True)
    profile = directory / "profile.yaml"
    profile.write_text(CUSTOMER_B_PROFILE, encoding="utf-8")
    _contract_template(directory / "templates" / "HopDong.docx")
    _confirmation_template(directory / "templates" / "XacNhan.docx")
    workbook = build_xlsx(directory / "data" / "HopDong.xlsx", customer_b_sheets())
    return profile, workbook


# ===================================================================== C
CUSTOMER_C_PROFILE = """\
schema_version: 1
profile:
  id: customer_c
  name: "Customer C – Phiếu thông tin nhân viên (biểu mẫu Word có sẵn)"
  version: 1
  description: >
    The customer's Word file is an existing form without placeholders: blanks after labels,
    content controls, legacy form fields and a label/value table. Bindings map each input to
    an Excel field; unbound inputs (signature, date of signing) stay empty.
excel:
  primary: ds
  sheets:
    - id: ds
      match: {name: "DS nhân viên"}
      columns:
        - {id: ma, header: "Mã NV", type: string}
        - {id: ho_ten, header: "Họ và tên"}
        - {id: ngay_sinh, header: "Ngày sinh", type: date}
        - {id: noi_sinh, header: "Nơi sinh"}
        - {id: cccd, header: "Số CCCD", type: string}
        - {id: dien_thoai, header: "Điện thoại", type: string}
        - {id: email, header: "Email", required: false}
        - {id: ma_pb, header: "Mã phòng", type: string}
        - {id: ma_cv, header: "Mã chức vụ", type: string}
        - {id: ngay_vao, header: "Ngày vào làm", type: date}
        - {id: dang_vien, header: "Đảng viên"}
        - {id: ket_hon, header: "Đã kết hôn"}
        - {id: mst, header: "MST", type: string, required: false}
    - id: pb
      match: {name: "Phòng ban"}
      columns:
        - {id: ma, header: "Mã phòng", type: string}
        - {id: ten, header: "Tên phòng"}
    - id: cv
      match: {name: "Chức vụ"}
      columns:
        - {id: ma, header: "Mã chức vụ", type: string}
        - {id: ten, header: "Tên chức vụ"}
relationships:
  - {id: ds_pb, from: ds, to: pb, keys: [{from: ma_pb, to: ma}]}
  - {id: ds_cv, from: ds, to: cv, keys: [{from: ma_cv, to: ma}]}
fields:
  - {name: nhan_vien.ma, source: ma}
  - {name: nhan_vien.ho_ten, source: ho_ten, transforms: [trim, title_case]}
  - {name: nhan_vien.sinh_ngay, source: ngay_sinh, transforms: [{format_date: {format: "%d"}}]}
  - {name: nhan_vien.sinh_thang, source: ngay_sinh, transforms: [{format_date: {format: "%m"}}]}
  - {name: nhan_vien.sinh_nam, source: ngay_sinh, transforms: [{format_date: {format: "%Y"}}]}
  - {name: nhan_vien.noi_sinh, source: noi_sinh}
  - {name: nhan_vien.cccd, source: cccd}
  - {name: nhan_vien.dien_thoai, source: dien_thoai}
  - {name: nhan_vien.email, source: email, default: "—"}
  - {name: nhan_vien.phong_ban, path: ds_pb.ten}
  - {name: nhan_vien.chuc_vu, path: ds_cv.ten}
  - {name: nhan_vien.ngay_vao, source: ngay_vao, transforms: [format_date]}
  - {name: nhan_vien.dang_vien, source: dang_vien}
  - {name: nhan_vien.ket_hon, source: ket_hon}
  - {name: nhan_vien.mst, source: mst, default: "(chưa có)"}
templates:
  - id: phieu
    file: templates/PhieuThongTin.docx
    missing_placeholder: error
    bindings:
      "Họ và tên": nhan_vien.ho_ten
      "Sinh ngày": nhan_vien.sinh_ngay
      "tháng": nhan_vien.sinh_thang
      "năm": nhan_vien.sinh_nam
      "tại": nhan_vien.noi_sinh
      "Số CCCD": nhan_vien.cccd
      "Điện thoại": nhan_vien.dien_thoai
      "Email": nhan_vien.email
      phong_ban: nhan_vien.phong_ban
      "Chức vụ": nhan_vien.chuc_vu
      NgayVao: nhan_vien.ngay_vao
      dang_vien: nhan_vien.dang_vien
      MST: nhan_vien.mst
      KetHon: nhan_vien.ket_hon
generation:
  unit: record
  output: separate
output:
  individual_filename: "PhieuTT_{nhan_vien.ma}.docx"
"""


def customer_c_sheets() -> dict[str, list[list[object]]]:
    return {
        "DS nhân viên": [
            [
                "Mã NV",
                "Họ và tên",
                "Ngày sinh",
                "Nơi sinh",
                "Số CCCD",
                "Điện thoại",
                "Email",
                "Mã phòng",
                "Mã chức vụ",
                "Ngày vào làm",
                "Đảng viên",
                "Đã kết hôn",
                "MST",
            ],
            [
                "NV01",
                "nguyễn văn ánh",
                dt.datetime(1990, 9, 5),
                "Hà Nội",
                "001090001234",
                "0912 345 678",
                "anh@example.vn",
                "KT",
                "TP",
                dt.datetime(2020, 3, 1),
                "Có",
                "Không",
                "0101234567",
            ],
            [
                "NV02",
                "trần thị bình",
                dt.datetime(1995, 12, 24),
                "Đà Nẵng",
                "048195002345",
                "0905 111 222",
                None,
                "IT",
                "NV",
                dt.datetime(2022, 7, 15),
                "Không",
                "Có",
                None,
            ],
        ],
        "Phòng ban": [["Mã phòng", "Tên phòng"], ["KT", "Phòng Kế toán"], ["IT", "Phòng Công nghệ"]],
        "Chức vụ": [["Mã chức vụ", "Tên chức vụ"], ["TP", "Trưởng phòng"], ["NV", "Nhân viên"]],
    }


def _form_template(path: Path) -> None:
    """A customer form as people make them in Word: no placeholders, just inputs."""
    from docgen.testing.docx_builder import (
        add_block_content_control,
        add_content_control,
        add_form_checkbox,
        add_form_text,
    )

    doc = Document()
    _style_document(doc)
    doc.sections[0].header.paragraphs[0].text = "CÔNG TY TNHH ABC – Biểu mẫu NS-01"
    doc.add_heading("PHIẾU THÔNG TIN NHÂN VIÊN", level=1)
    doc.add_paragraph("Họ và tên: ................................................")
    doc.add_paragraph("Sinh ngày ........ tháng ........ năm ............, tại: ____________________")
    p = doc.add_paragraph("Phòng ban: ")
    add_content_control(p, "phong_ban", alias="Phòng ban")
    p = doc.add_paragraph("Chức vụ: ")
    add_content_control(p, None, alias="Chức vụ", kind="dropdown", items=("Nhân viên", "Trưởng phòng"))
    p = doc.add_paragraph("Ngày vào làm: ")
    add_content_control(p, "NgayVao", kind="date", placeholder="Chọn ngày.")
    p = doc.add_paragraph("Đảng viên: ")
    add_content_control(p, "dang_vien", kind="checkbox")
    p = doc.add_paragraph("Mã số thuế cá nhân: ")
    add_form_text(p, "MST")
    p = doc.add_paragraph("Đã kết hôn ")
    add_form_checkbox(p, "KetHon")
    table = doc.add_table(rows=2, cols=4)
    table.style = "Table Grid"
    table.cell(0, 0).text = "Số CCCD"
    table.cell(0, 2).text = "Điện thoại"
    table.cell(1, 0).text = "Email"
    doc.add_paragraph("Ghi chú của phòng nhân sự:")
    add_block_content_control(doc, "ghi_chu", alias="Ghi chú")
    doc.add_paragraph("Hà Nội, ngày ...... tháng ...... năm ........")
    doc.add_paragraph("Người khai ký tên: ..........................................")
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(path))


def write_customer_c(directory: Path) -> tuple[Path, Path]:
    directory.mkdir(parents=True, exist_ok=True)
    profile = directory / "profile.yaml"
    profile.write_text(CUSTOMER_C_PROFILE, encoding="utf-8")
    _form_template(directory / "templates" / "PhieuThongTin.docx")
    workbook = build_xlsx(directory / "data" / "NhanVien.xlsx", customer_c_sheets())
    return profile, workbook
