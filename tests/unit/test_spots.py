"""Form inputs in customer Word files: content controls, legacy form fields, label blanks."""

from __future__ import annotations

import datetime as dt
import io
import zipfile
from pathlib import Path

import pytest
from docx import Document
from lxml import etree

from docgen.domain.mapping.engine import ImageRef
from docgen.infrastructure.word.renderer import RenderError, TemplateRenderer
from docgen.infrastructure.word.spots import KIND_BLANK, SpotOptions, match_choice, process_spots
from docgen.infrastructure.word.verifier import verify_docx
from docgen.testing.docx_builder import (
    add_block_content_control,
    add_content_control,
    add_form_checkbox,
    add_form_dropdown,
    add_form_text,
    add_picture_content_control,
    tiny_png,
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


def picture_extents(path: Path) -> list[tuple[int, int]]:
    wp = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
    xml = etree.fromstring(zipfile.ZipFile(path).read("word/document.xml"))
    return [(int(e.get("cx")), int(e.get("cy"))) for e in xml.iter(wp + "extent")]


def dropdown_form() -> bytes:
    doc = Document()
    add_form_dropdown(doc.add_paragraph("Giới tính: "), "GioiTinh", ("Nam", "Nữ", "Khác"))
    add_form_dropdown(
        doc.add_paragraph("Trình độ: "), "TrinhDo", ("Trung cấp", "Cao đẳng", "Đại học"), selected=2
    )
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def selected_entries(path: Path) -> list[str]:
    xml = etree.fromstring(zipfile.ZipFile(path).read("word/document.xml"))
    return [dd.find(W + "result").get(W + "val") for dd in xml.iter(W + "ddList")]


def test_legacy_dropdown_is_detected_with_its_choices():
    spots = {s.name: s for s in process_spots(dropdown_form(), SpotOptions()).spots}
    assert spots["GioiTinh"].kind == "form_field" and spots["GioiTinh"].subtype == "dropdown"
    assert spots["GioiTinh"].context == "Giới tính: [Nam / Nữ / Khác]"


def test_legacy_dropdown_selects_the_matching_choice(tmp_path: Path):
    result = process_spots(dropdown_form(), SpotOptions(), fill=lambda s: True)
    renderer = TemplateRenderer()
    out = tmp_path / "a.docx"
    renderer.render(result.data, {"b": {"GioiTinh": "Nữ", "TrinhDo": "  đại HỌC "}}, out)
    assert selected_entries(out) == ["1", "2"]
    assert verify_docx(out).ok
    empty = tmp_path / "b.docx"
    renderer.render(result.data, {"b": {"GioiTinh": None, "TrinhDo": ""}}, empty)
    assert selected_entries(empty) == ["0", "2"]  # empty value keeps the form's own selection
    with pytest.raises(RenderError) as exc:
        renderer.render(result.data, {"b": {"GioiTinh": "F", "TrinhDo": "Đại học"}}, tmp_path / "c.docx")
    issue = exc.value.issues[-1]
    assert issue.code == "render.value_error"
    assert "'F'" in issue.message and "Nam / Nữ / Khác" in issue.message
    assert not (tmp_path / "c.docx").exists()


def test_choice_matching_never_guesses():
    assert match_choice("Nữ", ["Nam", "Nữ"]) == 1
    assert match_choice("nu", ["Nam", "Nữ"]) == 1  # accents ignored only as a last resort
    assert match_choice("nam", ["Nam", "nam"]) == 1  # exact match wins
    assert match_choice("NAM", ["Nam", "nam"]) is None  # ambiguous: no arbitrary pick
    assert match_choice(2, ["1", "2"]) == 1


def test_picture_control_gets_the_excel_image_fitted_into_its_frame(tmp_path: Path):
    placeholder = tmp_path / "frame.png"
    placeholder.write_bytes(tiny_png((220, 220, 220), (30, 40)))
    photo = tmp_path / "photo.png"
    photo.write_bytes(tiny_png((10, 120, 200), (200, 100)))  # wide picture, 2:1
    doc = Document()
    add_picture_content_control(doc.add_paragraph("Ảnh 3x4: "), "anh", placeholder, alias="Ảnh", width_mm=30)
    add_picture_content_control(doc.add_paragraph("Ảnh phụ: "), "anh_phu", placeholder)
    buf = io.BytesIO()
    doc.save(buf)
    spots = {s.name: s for s in process_spots(buf.getvalue(), SpotOptions()).spots}
    assert spots["anh"].subtype == "picture" and spots["anh"].context == "Ảnh 3x4: […]"
    result = process_spots(buf.getvalue(), SpotOptions(), fill=lambda s: s.name == "anh")
    out = tmp_path / "out.docx"
    TemplateRenderer().render(result.data, {"b": {"anh": ImageRef(str(photo))}}, out)
    shapes = picture_extents(out)
    assert len(shapes) == 2  # the photo, and the unbound frame keeps its own picture
    photo_w, photo_h = shapes[0]
    assert abs(photo_w - 30 * 36000) <= 1  # fills the 30 mm frame width
    assert abs(photo_w / photo_h - 2) < 0.01  # aspect ratio kept (15 mm high, inside the 40 mm frame)
    xml = zipfile.ZipFile(out).read("word/document.xml").decode()
    assert xml.count("<w:picture/>") == 2 and xml.count("showingPlcHdr") == 1
    assert verify_docx(out).ok
    # no image in Excel: the frame is emptied rather than showing the grey placeholder
    TemplateRenderer().render(result.data, {"b": {"anh": None}}, tmp_path / "empty.docx")
    assert len(picture_extents(tmp_path / "empty.docx")) == 1
