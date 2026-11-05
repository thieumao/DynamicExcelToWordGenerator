"""Form inputs in customer Word files: content controls, legacy form fields, label blanks."""

from __future__ import annotations

import datetime as dt
import io
import zipfile
from pathlib import Path

from docx import Document
from lxml import etree

from docgen.infrastructure.word.renderer import TemplateRenderer
from docgen.infrastructure.word.spots import KIND_BLANK, SpotOptions, process_spots
from docgen.infrastructure.word.verifier import verify_docx
from docgen.testing.docx_builder import (
    add_block_content_control,
    add_content_control,
    add_form_checkbox,
    add_form_text,
)

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def form_bytes() -> bytes:
    doc = Document()
    doc.sections[0].header.paragraphs[0].text = "Mã hồ sơ: ..........."
    doc.add_paragraph("Họ và tên: ................................")
    doc.add_paragraph("Sinh ngày ...... tháng ...... năm ........, tại: ____________")
    p = doc.add_paragraph("Phòng ban: ")
    add_content_control(p, "phong_ban", alias="Phòng ban", bold=True)
    p = doc.add_paragraph("Chức vụ: ")
    add_content_control(p, None, alias="Chức vụ", kind="dropdown", items=("NV", "TP"))
    p = doc.add_paragraph("Đảng viên ")
    add_content_control(p, "dang_vien", kind="checkbox")
    add_block_content_control(doc, "ghi_chu", alias="Ghi chú")
    p = doc.add_paragraph("Mã số thuế: ")
    add_form_text(p, "MST")
    p = doc.add_paragraph("Đã kết hôn ")
    add_form_checkbox(p, "KetHon")
    table = doc.add_table(rows=2, cols=4)
    table.cell(0, 0).text = "Số CCCD"
    table.cell(0, 2).text = "Điện thoại"
    table.cell(1, 0).text = "Email"
    doc.add_paragraph("Hà Nội, ngày ...... tháng ...... năm ........")
    doc.add_paragraph("Chữ ký: ............................")
    doc.add_paragraph("Ghi chú: xem thêm... tài liệu đính kèm")  # short ellipsis: not a blank
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def all_text(data: bytes) -> str:
    out = []
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        for name in sorted(zf.namelist()):
            if name.startswith("word/") and name.endswith(".xml"):
                root = etree.fromstring(zf.read(name))
                for p in root.iter(W + "p"):
                    out.append("".join(t.text or "" for t in p.iter(W + "t")))
    return "\n".join(out)


def test_detection_finds_every_kind_with_names_and_context():
    spots = process_spots(form_bytes(), SpotOptions()).spots
    by_name = {s.name: s for s in spots}
    assert by_name["phong_ban"].kind == "content_control" and by_name["phong_ban"].subtype == "text"
    assert by_name["Chức vụ"].subtype == "dropdown"  # no tag -> title (alias)
    assert by_name["dang_vien"].subtype == "checkbox"
    assert by_name["ghi_chu"].context.startswith("Ghi chú")
    assert by_name["MST"].kind == "form_field" and by_name["KetHon"].subtype == "checkbox"
    blanks = [s.name for s in spots if s.kind == KIND_BLANK]
    assert blanks[:2] == ["Mã hồ sơ", "Họ và tên"] or "Họ và tên" in blanks
    for name in (
        "Sinh ngày",
        "tháng",
        "năm",
        "tại",
        "ngày",
        "tháng (2)",
        "năm (2)",
        "Chữ ký",
        "Số CCCD",
        "Điện thoại",
        "Email",
    ):
        assert name in blanks, name
    assert "xem thêm" not in " ".join(blanks)
    assert "Sinh ngày [" in by_name["Sinh ngày"].context
    assert not any(s.filled for s in spots)


def test_detection_alone_changes_nothing():
    data = form_bytes()
    assert process_spots(data, SpotOptions()).data == data


def test_only_bound_spots_are_filled_and_values_land_in_place(tmp_path: Path):
    values = {
        "phong_ban": "Kế toán",
        "Chức vụ": "Trưởng phòng",
        "dang_vien": "Có",
        "ghi_chu": "A < B & C",
        "MST": "0101234567",
        "KetHon": "x",
        "Họ và tên": "Nguyễn Văn Ánh",
        "Sinh ngày": "05",
        "tháng": "09",
        "năm": "1990",
        "tại": "Hà Nội",
        "Số CCCD": "0010900",
        "Điện thoại": "0912",
        "Email": "a@b.vn",
        "Mã hồ sơ": "HS-01",
    }
    result = process_spots(form_bytes(), SpotOptions(), fill=lambda s: s.name in values)
    filled = {s.name for s in result.spots if s.filled}
    assert filled == set(values)
    out = tmp_path / "out.docx"
    TemplateRenderer().render(result.data, {"b": values}, out)
    text = all_text(out.read_bytes())
    for line in (
        "Họ và tên: Nguyễn Văn Ánh",
        "Sinh ngày 05 tháng 09 năm 1990, tại: Hà Nội",
        "Phòng ban: Kế toán",
        "Chức vụ: Trưởng phòng",
        "Đảng viên ☒",
        "A < B & C",
        "Mã số thuế: 0101234567",
        "Mã hồ sơ: HS-01",
    ):
        assert line in text, line
    # unbound inputs are untouched
    assert "Hà Nội, ngày ...... tháng ...... năm ........" in text
    assert "Chữ ký: ............................" in text
    doc = Document(str(out))
    assert [c.text for c in doc.tables[0].rows[0].cells] == ["Số CCCD", "0010900", "Điện thoại", "0912"]
    assert [c.text for c in doc.tables[0].rows[1].cells] == ["Email", "a@b.vn", "", ""]
    xml = zipfile.ZipFile(out).read("word/document.xml").decode()
    assert 'w14:checked w14:val="1"' in xml and '<w:checked w:val="1"/>' in xml
    assert xml.count("showingPlcHdr") == 0  # filled controls no longer show their placeholder
    assert "<w:sdt>" in xml  # controls are kept as controls
    bold_value = doc.paragraphs[2]._p.xpath(".//w:sdtContent//w:r[w:rPr/w:b]//w:t")
    assert bold_value and bold_value[0].text == "Kế toán"  # formatting of the control is kept
    assert verify_docx(out).ok


def test_unchecked_and_empty_values(tmp_path: Path):
    result = process_spots(form_bytes(), SpotOptions(), fill=lambda s: s.name in {"dang_vien", "KetHon"})
    out = tmp_path / "o.docx"
    TemplateRenderer().render(result.data, {"b": {"dang_vien": "Không", "KetHon": None}}, out)
    xml = zipfile.ZipFile(out).read("word/document.xml").decode()
    assert 'w14:checked w14:val="0"' in xml and '<w:checked w:val="0"/>' in xml
    assert "Đảng viên ☐" in all_text(out.read_bytes())


def test_options_switch_detectors_off():
    data = form_bytes()
    only_controls = process_spots(data, SpotOptions(form_fields=False, blanks=False)).spots
    assert {s.kind for s in only_controls} == {"content_control"}
    longer = process_spots(data, SpotOptions(content_controls=False, form_fields=False, blank_min=20)).spots
    assert [s.name for s in longer if s.subtype == "label"] == ["Họ và tên", "Chữ ký"]


def test_filled_table_cell_never_becomes_a_label():
    doc = Document()
    t = doc.add_table(rows=1, cols=3)
    t.cell(0, 0).text = "Email"
    buf = io.BytesIO()
    doc.save(buf)
    result = process_spots(buf.getvalue(), SpotOptions(), fill=lambda s: True)
    assert [s.name for s in result.spots] == ["Email"]


def test_date_value_into_date_control(tmp_path: Path):
    doc = Document()
    add_content_control(doc.add_paragraph("Ngày: "), "ngay", kind="date")
    buf = io.BytesIO()
    doc.save(buf)
    result = process_spots(buf.getvalue(), SpotOptions(), fill=lambda s: True)
    out = tmp_path / "d.docx"
    TemplateRenderer().render(result.data, {"b": {"ngay": dt.date(2024, 3, 9)}}, out)
    assert "Ngày: 09/03/2024" in all_text(out.read_bytes())


def test_filled_control_drops_custom_xml_binding_and_stale_date(tmp_path: Path):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    doc = Document()
    add_content_control(doc.add_paragraph("Ngày ký: "), "ngay_ky", kind="date")
    add_content_control(doc.add_paragraph("Ghi chú: "), "khong_gan")
    for sdt_pr in doc.element.body.iter(qn("w:sdtPr")):
        binding = OxmlElement("w:dataBinding")
        binding.set(qn("w:xpath"), "/root[1]/x[1]")
        binding.set(qn("w:storeItemID"), "{00000000-0000-0000-0000-000000000001}")
        sdt_pr.append(binding)
        date = sdt_pr.find(qn("w:date"))
        if date is not None:
            date.set(qn("w:fullDate"), "2020-01-01T00:00:00Z")
    buf = io.BytesIO()
    doc.save(buf)
    result = process_spots(buf.getvalue(), SpotOptions(), fill=lambda s: s.name == "ngay_ky")
    out = tmp_path / "b.docx"
    TemplateRenderer().render(result.data, {"b": {"ngay_ky": dt.date(2024, 3, 9)}}, out)
    xml = zipfile.ZipFile(out).read("word/document.xml").decode()
    assert xml.count("w:dataBinding") == 1  # only the unbound control keeps its XML binding
    assert "fullDate" not in xml
    assert "Ngày ký: 09/03/2024" in all_text(out.read_bytes())
