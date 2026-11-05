"""Construction starter kit: acceptance records ("hồ sơ nghiệm thu công việc xây dựng").

A realistic, ready-to-run example for construction companies, the first target customers:

* Excel ``HoSoNghiemThu.xlsx``:
  - ``Thông tin công trình`` – project details entered once (label | value);
  - ``Danh mục nghiệm thu`` – one row per work item to accept (title rows above the headers);
  - ``Nhân sự`` – people referenced by code (supervisor / contractor engineer).
* Word templates with ``<<Tên trường>>`` markers – the simplest way for non-IT users to mark
  what to fill: type the name of the information between ``<<`` and ``>>``:
  - ``PhieuYeuCauNghiemThu.docx`` – the contractor's request for acceptance;
  - ``BienBanNghiemThu.docx`` – the acceptance record of the work item.
* One Word file per work item, containing the request and the record, plus one combined file
  with every work item, ready to print.

The forms are a reference layout only; users should check them against current regulations
and their own project requirements.
"""

from __future__ import annotations

import datetime as dt
from pathlib import Path

from docx import Document
from docx.document import Document as DocxDocument
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from docx.text.paragraph import Paragraph

from docgen.testing.xlsx_builder import build_xlsx

KIT_ID = "ho_so_nghiem_thu"
MATERIALS_KIT_ID = "nghiem_thu_vat_lieu"
MATERIALS_PROFILE_FILE = "profile_vat_lieu.yaml"

PROFILE = """\
schema_version: 1
profile:
  id: ho_so_nghiem_thu
  name: "Hồ sơ nghiệm thu công việc xây dựng (mẫu)"
  version: 1
  description: >
    Mỗi dòng của sheet 'Danh mục nghiệm thu' tạo một file Word gồm Phiếu yêu cầu nghiệm thu
    và Biên bản nghiệm thu công việc xây dựng. Thông tin công trình chỉ nhập một lần.
excel:
  primary: dm
  sheets:
    - id: dm
      match: {name: "Danh mục nghiệm thu"}
      columns:
        - {id: stt, header: "STT"}
        - {id: so_bien_ban, header: "Số biên bản", type: string}
        - {id: ten_cong_viec, header: "Tên công việc"}
        - {id: vi_tri, header: "Vị trí"}
        - {id: ban_ve_so, header: "Bản vẽ số", required: false}
        - {id: tieu_chuan, header: "Tiêu chuẩn áp dụng", required: false}
        - {id: ngay_nghiem_thu, header: "Ngày nghiệm thu", type: date}
        - {id: gio_bat_dau, header: "Giờ bắt đầu"}
        - {id: gio_ket_thuc, header: "Giờ kết thúc"}
        - {id: ma_giam_sat, header: "Mã giám sát", type: string}
        - {id: ma_ky_thuat, header: "Mã kỹ thuật", type: string}
        - {id: ket_qua, header: "Kết quả", required: false}
    - id: ct
      match: {name: "Thông tin công trình"}
      layout: key_value
      columns:
        - {id: ten_cong_trinh, header: "Tên công trình"}
        - {id: hang_muc, header: "Hạng mục", required: false}
        - {id: dia_diem, header: "Địa điểm xây dựng"}
        - {id: chu_dau_tu, header: "Chủ đầu tư"}
        - {id: tu_van_giam_sat, header: "Đơn vị tư vấn giám sát"}
        - {id: nha_thau, header: "Nhà thầu thi công"}
        - {id: hop_dong, header: "Hợp đồng thi công số", required: false}
        - {id: noi_lap, header: "Nơi lập biên bản", required: false}
    - id: ns
      match: {name: "Nhân sự"}
      columns:
        - {id: ma, header: "Mã", type: string}
        - {id: ho_ten, header: "Họ và tên"}
        - {id: chuc_vu, header: "Chức vụ"}
relationships:
  - {id: r_ct, from: dm, to: ct}   # every row uses the project details
  - {id: gs, from: dm, to: ns, keys: [{from: ma_giam_sat, to: ma}]}
  - {id: kt, from: dm, to: ns, keys: [{from: ma_ky_thuat, to: ma}]}
fields:
  # project details (sheet 'Thông tin công trình')
  - {name: cong_trinh.ten_cong_trinh, path: r_ct.ten_cong_trinh}
  - {name: cong_trinh.hang_muc, path: r_ct.hang_muc, default: ""}
  - {name: cong_trinh.dia_diem_xay_dung, path: r_ct.dia_diem}
  - {name: cong_trinh.chu_dau_tu, path: r_ct.chu_dau_tu}
  - {name: cong_trinh.don_vi_tu_van_giam_sat, path: r_ct.tu_van_giam_sat}
  - {name: cong_trinh.nha_thau_thi_cong, path: r_ct.nha_thau}
  - {name: cong_trinh.hop_dong_thi_cong_so, path: r_ct.hop_dong, default: "……"}
  - {name: cong_trinh.noi_lap_bien_ban, path: r_ct.noi_lap, default: "……"}
  # the work item (sheet 'Danh mục nghiệm thu')
  - {name: nghiem_thu.stt, source: stt}
  - {name: nghiem_thu.so_bien_ban, source: so_bien_ban}
  - {name: nghiem_thu.ten_cong_viec, source: ten_cong_viec}
  - {name: nghiem_thu.vi_tri, source: vi_tri}
  - {name: nghiem_thu.ban_ve_so, source: ban_ve_so, default: "……"}
  - {name: nghiem_thu.tieu_chuan_ap_dung, source: tieu_chuan, default: "……"}
  - {name: nghiem_thu.ket_qua, source: ket_qua, default: "Đạt"}
  - {name: nghiem_thu.ngay_nghiem_thu, source: ngay_nghiem_thu, transforms: [format_date]}
  - name: nghiem_thu.ngay_nghiem_thu_bang_chu
    source: ngay_nghiem_thu
    transforms: [{format_date: {format: "ngày %d tháng %m năm %Y"}}]
  - name: nghiem_thu.ngay_yeu_cau_bang_chu
    source: ngay_nghiem_thu
    transforms: [{add_days: {days: -1}}, {format_date: {format: "ngày %d tháng %m năm %Y"}}]
  - {name: nghiem_thu.gio_bat_dau, source: gio_bat_dau, transforms: [{format_time: {format: "%H giờ %M"}}]}
  - {name: nghiem_thu.gio_ket_thuc, source: gio_ket_thuc, transforms: [{format_time: {format: "%H giờ %M"}}]}
  # people (sheet 'Nhân sự', found by their code)
  - {name: nghiem_thu.ho_ten_giam_sat, path: gs.ho_ten}
  - {name: nghiem_thu.chuc_vu_giam_sat, path: gs.chuc_vu}
  - {name: nghiem_thu.ho_ten_ky_thuat, path: kt.ho_ten}
  - {name: nghiem_thu.chuc_vu_ky_thuat, path: kt.chuc_vu}
templates:
  - id: phieu_yeu_cau
    file: templates/PhieuYeuCauNghiemThu.docx
    marker: {style: custom, open: "<<", close: ">>"}
  - id: bien_ban
    file: templates/BienBanNghiemThu.docx
    marker: {style: custom, open: "<<", close: ">>"}
merge:
  order: [phieu_yeu_cau, bien_ban]
generation:
  unit: record
  output: merged
output:
  merged_filename: "{nghiem_thu.stt} - {nghiem_thu.ten_cong_viec}.docx"
  individual_filename: "{template.id} - {nghiem_thu.stt}.docx"
  combined_filename: "Tổng hợp hồ sơ nghiệm thu.docx"   # every document in one file, to print
"""


MATERIALS_PROFILE = (
    """\
schema_version: 1
profile:
  id: nghiem_thu_vat_lieu
  name: "Nghiệm thu vật liệu đầu vào (mẫu)"
  version: 1
  description: >
    Mỗi dòng của sheet 'Vật liệu đầu vào' tạo một Biên bản nghiệm thu vật liệu đầu vào.
    Dùng chung file Excel với hồ sơ nghiệm thu công việc (thông tin công trình, nhân sự).
excel:
  primary: vl
  sheets:
    - id: vl
      match: {name: "Vật liệu đầu vào"}
      columns:
        - {id: stt, header: "STT"}
        - {id: so_bien_ban, header: "Số biên bản", type: string}
        - {id: ten_vat_lieu, header: "Tên vật liệu"}
        - {id: quy_cach, header: "Quy cách chủng loại", required: false}
        - {id: nha_san_xuat, header: "Nhà sản xuất", required: false}
        - {id: khoi_luong, header: "Khối lượng", required: false}
        - {id: don_vi, header: "Đơn vị", required: false}
        - {id: hang_muc_su_dung, header: "Hạng mục sử dụng", required: false}
        - {id: chung_chi, header: "Chứng chỉ chất lượng", required: false}
        - {id: thi_nghiem, header: "Kết quả thí nghiệm", required: false}
        - {id: ngay_nghiem_thu, header: "Ngày nghiệm thu", type: date}
        - {id: ma_giam_sat, header: "Mã giám sát", type: string}
        - {id: ma_ky_thuat, header: "Mã kỹ thuật", type: string}
        - {id: ket_qua, header: "Kết quả", required: false}
"""
    + PROFILE[PROFILE.index("    - id: ct\n") : PROFILE.index("fields:\n")].replace("from: dm,", "from: vl,")
    + """fields:
"""
    + PROFILE[PROFILE.index("  # project details") : PROFILE.index("  # the work item")]
    + """  # the material (sheet 'Vật liệu đầu vào')
  - {name: vat_lieu.stt, source: stt}
  - {name: vat_lieu.so_bien_ban, source: so_bien_ban}
  - {name: vat_lieu.ten_vat_lieu, source: ten_vat_lieu}
  - {name: vat_lieu.quy_cach_chung_loai, source: quy_cach, default: "……"}
  - {name: vat_lieu.nha_san_xuat, source: nha_san_xuat, default: "……"}
  - {name: vat_lieu.khoi_luong, source: khoi_luong, default: "……"}
  - {name: vat_lieu.don_vi, source: don_vi, default: ""}
  - {name: vat_lieu.hang_muc_su_dung, source: hang_muc_su_dung, default: "……"}
  - {name: vat_lieu.chung_chi_chat_luong, source: chung_chi, default: "……"}
  - {name: vat_lieu.ket_qua_thi_nghiem, source: thi_nghiem, default: "……"}
  - {name: vat_lieu.ket_qua, source: ket_qua, default: "Đạt"}
  - name: vat_lieu.ngay_nghiem_thu_bang_chu
    source: ngay_nghiem_thu
    transforms: [{format_date: {format: "ngày %d tháng %m năm %Y"}}]
  - {name: vat_lieu.ho_ten_giam_sat, path: gs.ho_ten}
  - {name: vat_lieu.chuc_vu_giam_sat, path: gs.chuc_vu}
  - {name: vat_lieu.ho_ten_ky_thuat, path: kt.ho_ten}
  - {name: vat_lieu.chuc_vu_ky_thuat, path: kt.chuc_vu}
templates:
  - id: bien_ban_vat_lieu
    file: templates/BienBanNghiemThuVatLieu.docx
    marker: {style: custom, open: "<<", close: ">>"}
generation:
  unit: record
  output: merged
output:
  merged_filename: "VL {vat_lieu.stt} - {vat_lieu.ten_vat_lieu}.docx"
  individual_filename: "{template.id} - {vat_lieu.stt}.docx"
  combined_filename: "Tổng hợp nghiệm thu vật liệu.docx"
"""
)


def sheets() -> dict[str, list[list[object]]]:
    def day(d: int) -> dt.datetime:
        return dt.datetime(2024, 4, d)

    std_bt = "TCVN 4453:1995"
    return {
        "Thông tin công trình": [
            ["THÔNG TIN CHUNG CỦA CÔNG TRÌNH (nhập một lần, dùng cho mọi biên bản)"],
            [],
            ["Tên công trình", "Nhà làm việc 5 tầng – Công ty CP Đầu tư DEF"],
            ["Hạng mục", "Phần thân"],
            ["Địa điểm xây dựng", "Số 12 đường Lê Văn Lương, quận Thanh Xuân, Hà Nội"],
            ["Chủ đầu tư", "Công ty Cổ phần Đầu tư DEF"],
            ["Đơn vị tư vấn giám sát", "Công ty TNHH Tư vấn Xây dựng XYZ"],
            ["Nhà thầu thi công", "Công ty Cổ phần Xây dựng ABC"],
            ["Hợp đồng thi công số", "25/2024/HĐTC"],
            ["Nơi lập biên bản", "Hà Nội"],
        ],
        "Danh mục nghiệm thu": [
            ["DANH MỤC CÔNG VIỆC NGHIỆM THU – mỗi dòng là một biên bản"],
            [],
            [
                "STT",
                "Số biên bản",
                "Tên công việc",
                "Vị trí",
                "Bản vẽ số",
                "Tiêu chuẩn áp dụng",
                "Ngày nghiệm thu",
                "Giờ bắt đầu",
                "Giờ kết thúc",
                "Mã giám sát",
                "Mã kỹ thuật",
                "Kết quả",
            ],
            [
                1,
                "01/NTCV",
                "Lắp dựng cốt thép cột tầng 1",
                "Trục A–D / 1–5",
                "KC-01, KC-02",
                f"{std_bt}; TCVN 1651-2:2018",
                day(2),
                dt.time(8, 0),
                dt.time(9, 30),
                "GS01",
                "KT01",
                "Đạt",
            ],
            [
                2,
                "02/NTCV",
                "Lắp dựng ván khuôn cột tầng 1",
                "Trục A–D / 1–5",
                "KC-02",
                std_bt,
                day(3),
                dt.time(8, 0),
                dt.time(9, 0),
                "GS01",
                "KT01",
                "Đạt",
            ],
            [
                3,
                "03/NTCV",
                "Đổ bê tông cột tầng 1",
                "Trục A–D / 1–5",
                "KC-02",
                std_bt,
                day(4),
                dt.time(14, 0),
                dt.time(15, 30),
                "GS01",
                "KT02",
                "Đạt",
            ],
            [
                4,
                "04/NTCV",
                "Lắp dựng cốt thép dầm, sàn tầng 2",
                "Trục A–D / 1–5",
                "KC-05",
                f"{std_bt}; TCVN 1651-2:2018",
                day(10),
                dt.time(8, 0),
                dt.time(10, 0),
                "GS02",
                "KT01",
                "Đạt",
            ],
            [
                5,
                "05/NTCV",
                "Lắp dựng ván khuôn dầm, sàn tầng 2",
                "Trục A–D / 1–5",
                "KC-05",
                std_bt,
                day(11),
                dt.time(8, 30),
                dt.time(9, 30),
                "GS02",
                "KT02",
                "Đạt",
            ],
            [
                6,
                "06/NTCV",
                "Đổ bê tông dầm, sàn tầng 2",
                "Trục A–D / 1–5",
                "KC-05",
                std_bt,
                day(12),
                dt.time(7, 30),
                dt.time(11, 0),
                "GS02",
                "KT02",
                "Đạt",
            ],
        ],
        "Vật liệu đầu vào": [
            ["DANH MỤC VẬT LIỆU ĐẦU VÀO CẦN NGHIỆM THU – mỗi dòng là một biên bản"],
            [],
            [
                "STT",
                "Số biên bản",
                "Tên vật liệu",
                "Quy cách chủng loại",
                "Nhà sản xuất",
                "Khối lượng",
                "Đơn vị",
                "Hạng mục sử dụng",
                "Chứng chỉ chất lượng",
                "Kết quả thí nghiệm",
                "Ngày nghiệm thu",
                "Mã giám sát",
                "Mã kỹ thuật",
                "Kết quả",
            ],
            [
                1,
                "01/NTVL",
                "Thép cốt bê tông",
                "CB300-V, D10–D25",
                "Nhà máy thép (mẫu)",
                12.5,
                "tấn",
                "Cột, dầm, sàn tầng 1–2",
                "CQ số 1234/2024",
                "Phiếu KQTN số 45/2024",
                day(1),
                "GS01",
                "KT01",
                "Đạt",
            ],
            [
                2,
                "02/NTVL",
                "Xi măng",
                "PCB40, bao 50 kg",
                "Nhà máy xi măng (mẫu)",
                30,
                "tấn",
                "Bê tông, vữa phần thân",
                "CQ số 567/2024",
                "Phiếu KQTN số 46/2024",
                day(1),
                "GS01",
                "KT02",
                "Đạt",
            ],
            [
                3,
                "03/NTVL",
                "Cát vàng",
                "Cát hạt thô, Mđl ≥ 2,0",
                "Mỏ cát (mẫu)",
                45,
                "m³",
                "Bê tông phần thân",
                "—",
                "Phiếu KQTN số 47/2024",
                day(8),
                "GS02",
                "KT01",
                "Đạt",
            ],
            [
                4,
                "04/NTVL",
                "Đá dăm",
                "Đá 1x2 (10–20 mm)",
                "Mỏ đá (mẫu)",
                60,
                "m³",
                "Bê tông phần thân",
                "—",
                "Phiếu KQTN số 48/2024",
                day(8),
                "GS02",
                "KT02",
                "Đạt",
            ],
        ],
        "Nhân sự": [
            ["Mã", "Họ và tên", "Chức vụ", "Đơn vị"],
            ["GS01", "Nguyễn Văn Hùng", "Kỹ sư tư vấn giám sát", "Công ty TNHH Tư vấn Xây dựng XYZ"],
            ["GS02", "Lê Thị Mai", "Kỹ sư tư vấn giám sát", "Công ty TNHH Tư vấn Xây dựng XYZ"],
            ["KT01", "Trần Minh Tuấn", "Cán bộ kỹ thuật thi công", "Công ty Cổ phần Xây dựng ABC"],
            ["KT02", "Phạm Quốc Bảo", "Cán bộ kỹ thuật thi công", "Công ty Cổ phần Xây dựng ABC"],
        ],
    }


# ------------------------------------------------------------------ Word forms
def _document() -> DocxDocument:
    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)
    normal.paragraph_format.space_after = Pt(2)
    section = doc.sections[0]
    section.left_margin, section.right_margin = Pt(85), Pt(57)  # 3 cm / 2 cm, as for official papers
    section.top_margin = section.bottom_margin = Pt(50)
    return doc


def _line(
    doc: DocxDocument,
    text: str = "",
    *,
    bold: bool = False,
    italic: bool = False,
    align: WD_ALIGN_PARAGRAPH | None = None,
    size: int | None = None,
) -> Paragraph:
    p = doc.add_paragraph()
    if text:
        run = p.add_run(text)
        run.bold, run.italic = bold, italic
        if size:
            run.font.size = Pt(size)
    if align is not None:
        p.alignment = align
    return p


def _national_header(doc: DocxDocument, place_and_date: str) -> None:
    center = WD_ALIGN_PARAGRAPH.CENTER
    _line(doc, "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", bold=True, align=center)
    _line(doc, "Độc lập – Tự do – Hạnh phúc", bold=True, align=center)
    _line(doc, "———————", align=center)
    _line(doc, place_and_date, italic=True, align=WD_ALIGN_PARAGRAPH.RIGHT)


def _signatures(doc: DocxDocument, left: tuple[str, str], right: tuple[str, str]) -> None:
    table = doc.add_table(rows=3, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for col, (title, name) in enumerate((left, right)):
        cells = [table.cell(r, col) for r in range(3)]
        cells[0].paragraphs[0].add_run(title).bold = True
        cells[1].paragraphs[0].add_run("(Ký, ghi rõ họ tên)").italic = True
        for _ in range(3):
            cells[1].add_paragraph()
        cells[2].paragraphs[0].add_run(name).bold = True
        for cell in cells:
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER


def request_template(path: Path) -> None:
    doc = _document()
    _national_header(doc, "<<Nơi lập biên bản>>, <<Ngày yêu cầu bằng chữ>>")
    _line(doc, "PHIẾU YÊU CẦU NGHIỆM THU", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, size=15)
    _line(doc, "Số: PYC-<<Số biên bản>>", align=WD_ALIGN_PARAGRAPH.CENTER)
    _line(doc, "Kính gửi: <<Đơn vị tư vấn giám sát>>", bold=True)
    _line(doc, "Công trình: <<Tên công trình>>")
    _line(doc, "Hạng mục: <<Hạng mục>>")
    _line(doc, "Địa điểm xây dựng: <<Địa điểm xây dựng>>")
    _line(
        doc,
        "Căn cứ Hợp đồng thi công số <<Hợp đồng thi công số>>, <<Nhà thầu thi công>> đã hoàn thành "
        "công việc: <<Tên công việc>>, vị trí <<Vị trí>>, theo bản vẽ <<Bản vẽ số>>.",
    )
    _line(
        doc,
        "Đề nghị Quý đơn vị cử cán bộ tổ chức nghiệm thu vào lúc <<Giờ bắt đầu>> "
        "<<Ngày nghiệm thu bằng chữ>> tại hiện trường.",
    )
    _line(doc)
    _signatures(
        doc,
        ("ĐẠI DIỆN TƯ VẤN GIÁM SÁT", "<<Họ tên giám sát>>"),
        ("ĐẠI DIỆN NHÀ THẦU THI CÔNG", "<<Họ tên kỹ thuật>>"),
    )
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(path))


def record_template(path: Path) -> None:
    doc = _document()
    _national_header(doc, "<<Nơi lập biên bản>>, <<Ngày nghiệm thu bằng chữ>>")
    _line(doc, "BIÊN BẢN NGHIỆM THU CÔNG VIỆC XÂY DỰNG", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, size=15)
    _line(doc, "Số: <<Số biên bản>>", align=WD_ALIGN_PARAGRAPH.CENTER)
    _line(doc, "1. Công trình: <<Tên công trình>>", bold=True)
    _line(doc, "Hạng mục: <<Hạng mục>>")
    _line(doc, "Địa điểm xây dựng: <<Địa điểm xây dựng>>")
    _line(doc, "Chủ đầu tư: <<Chủ đầu tư>>")
    _line(doc, "2. Đối tượng nghiệm thu: <<Tên công việc>>", bold=True)
    _line(doc, "Vị trí: <<Vị trí>>")
    _line(doc, "3. Thành phần trực tiếp nghiệm thu:", bold=True)
    _line(doc, "a) Đại diện đơn vị tư vấn giám sát – <<Đơn vị tư vấn giám sát>>:")
    _line(doc, "Ông/Bà: <<Họ tên giám sát>> – Chức vụ: <<Chức vụ giám sát>>")
    _line(doc, "b) Đại diện nhà thầu thi công – <<Nhà thầu thi công>>:")
    _line(doc, "Ông/Bà: <<Họ tên kỹ thuật>> – Chức vụ: <<Chức vụ kỹ thuật>>")
    _line(doc, "4. Thời gian nghiệm thu:", bold=True)
    _line(doc, "Bắt đầu: <<Giờ bắt đầu>> <<Ngày nghiệm thu bằng chữ>>")
    _line(doc, "Kết thúc: <<Giờ kết thúc>> <<Ngày nghiệm thu bằng chữ>>")
    _line(doc, "Tại: <<Vị trí>> – <<Tên công trình>>")
    _line(doc, "5. Đánh giá công việc xây dựng đã thực hiện:", bold=True)
    _line(doc, "a) Tài liệu làm căn cứ nghiệm thu:")
    _line(doc, "– Phiếu yêu cầu nghiệm thu số PYC-<<Số biên bản>> của <<Nhà thầu thi công>>;")
    _line(doc, "– Hợp đồng thi công số <<Hợp đồng thi công số>>;")
    _line(doc, "– Bản vẽ thiết kế số: <<Bản vẽ số>>;")
    _line(doc, "– Tiêu chuẩn áp dụng: <<Tiêu chuẩn áp dụng>>.")
    _line(doc, "b) Về chất lượng công việc xây dựng: <<Kết quả>> yêu cầu của thiết kế và tiêu chuẩn áp dụng.")
    _line(doc, "c) Các ý kiến khác (nếu có): ..............................................................")
    _line(doc, "6. Kết luận:", bold=True)
    _line(doc, "Chấp nhận nghiệm thu công việc nêu trên, đồng ý cho triển khai các công việc tiếp theo.")
    _line(doc)
    _signatures(
        doc,
        ("ĐẠI DIỆN TƯ VẤN GIÁM SÁT", "<<Họ tên giám sát>>"),
        ("ĐẠI DIỆN NHÀ THẦU THI CÔNG", "<<Họ tên kỹ thuật>>"),
    )
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(path))


def materials_template(path: Path) -> None:
    doc = _document()
    _national_header(doc, "<<Nơi lập biên bản>>, <<Ngày nghiệm thu bằng chữ>>")
    _line(doc, "BIÊN BẢN NGHIỆM THU VẬT LIỆU ĐẦU VÀO", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, size=15)
    _line(doc, "Số: <<Số biên bản>>", align=WD_ALIGN_PARAGRAPH.CENTER)
    _line(doc, "1. Công trình: <<Tên công trình>>", bold=True)
    _line(doc, "Hạng mục: <<Hạng mục>>")
    _line(doc, "Địa điểm xây dựng: <<Địa điểm xây dựng>>")
    _line(doc, "2. Đối tượng nghiệm thu: <<Tên vật liệu>>", bold=True)
    _line(doc, "Quy cách, chủng loại: <<Quy cách chủng loại>>")
    _line(doc, "Nhà sản xuất, xuất xứ: <<Nhà sản xuất>>")
    _line(doc, "Khối lượng: <<Khối lượng>> <<Đơn vị>> – dùng cho: <<Hạng mục sử dụng>>")
    _line(doc, "3. Thành phần trực tiếp nghiệm thu:", bold=True)
    _line(doc, "a) Đại diện đơn vị tư vấn giám sát – <<Đơn vị tư vấn giám sát>>:")
    _line(doc, "Ông/Bà: <<Họ tên giám sát>> – Chức vụ: <<Chức vụ giám sát>>")
    _line(doc, "b) Đại diện nhà thầu thi công – <<Nhà thầu thi công>>:")
    _line(doc, "Ông/Bà: <<Họ tên kỹ thuật>> – Chức vụ: <<Chức vụ kỹ thuật>>")
    _line(doc, "4. Thời gian nghiệm thu: <<Ngày nghiệm thu bằng chữ>>, tại công trường.", bold=True)
    _line(doc, "5. Đánh giá vật liệu:", bold=True)
    _line(doc, "a) Tài liệu làm căn cứ nghiệm thu:")
    _line(doc, "– Chứng chỉ chất lượng, chứng nhận xuất xưởng: <<Chứng chỉ chất lượng>>;")
    _line(doc, "– Kết quả thí nghiệm: <<Kết quả thí nghiệm>>;")
    _line(doc, "– Hợp đồng thi công số <<Hợp đồng thi công số>>.")
    _line(doc, "b) Về chất lượng: <<Kết quả>> yêu cầu của thiết kế và tiêu chuẩn áp dụng.")
    _line(doc, "c) Các ý kiến khác (nếu có): ..............................................................")
    _line(doc, "6. Kết luận:", bold=True)
    _line(doc, "Chấp nhận nghiệm thu, đồng ý đưa vật liệu nêu trên vào sử dụng cho công trình.")
    _line(doc)
    _signatures(
        doc,
        ("ĐẠI DIỆN TƯ VẤN GIÁM SÁT", "<<Họ tên giám sát>>"),
        ("ĐẠI DIỆN NHÀ THẦU THI CÔNG", "<<Họ tên kỹ thuật>>"),
    )
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(path))


def write_construction_kit(directory: Path) -> tuple[Path, Path]:
    """Write both profiles, the Excel file and the Word forms into *directory*.

    Returns the work-items profile (``profile.yaml``) and the workbook; the materials profile is
    ``profile_vat_lieu.yaml`` next to it and uses the same workbook.
    """
    directory.mkdir(parents=True, exist_ok=True)
    profile = directory / "profile.yaml"
    profile.write_text(PROFILE, encoding="utf-8")
    (directory / MATERIALS_PROFILE_FILE).write_text(MATERIALS_PROFILE, encoding="utf-8")
    request_template(directory / "templates" / "PhieuYeuCauNghiemThu.docx")
    record_template(directory / "templates" / "BienBanNghiemThu.docx")
    materials_template(directory / "templates" / "BienBanNghiemThuVatLieu.docx")
    workbook = build_xlsx(directory / "data" / "HoSoNghiemThu.xlsx", sheets())
    return profile, workbook
