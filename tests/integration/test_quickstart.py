"""Quick start: Excel + Word files -> automatic matching -> a working profile, no configuration."""

from __future__ import annotations

import zipfile
from pathlib import Path

from lxml import etree

from docgen.application import quickstart as q
from docgen.application.generation.job import GenerationRequest
from docgen.application.generation.service import DocumentGenerationService
from docgen.domain.diagnostics import Severity
from docgen.domain.profile.io import parse_profile_text, profile_to_text
from docgen.domain.profile.validator import validate_profile
from docgen.testing.construction import write_construction_kit
from docgen.testing.samples import write_customer_c

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def text_of(path: Path) -> str:
    root = etree.fromstring(zipfile.ZipFile(path).read("word/document.xml"))
    return "\n".join("".join(t.text or "" for t in p.iter(W + "t")) for p in root.iter(W + "p"))


def test_construction_files_become_a_working_profile(tmp_path):
    _, workbook = write_construction_kit(tmp_path / "kit")
    analysis = q.analyze_workbook(workbook)
    assert {s.name: s.role for s in analysis.sheets} == {
        "Thông tin công trình": q.ROLE_INFO,
        "Danh mục nghiệm thu": q.ROLE_LIST,
        "Vật liệu đầu vào": q.ROLE_IGNORE,  # a second list: not used by this set
        "Nhân sự": q.ROLE_LOOKUP,
    }
    assert q.describe(analysis)[1] == "Sheet 'Danh mục nghiệm thu': 6 dòng – mỗi dòng sẽ tạo một bộ tài liệu."
    options = q.value_options(analysis)
    by_label = {o.label: o for o in options}
    assert by_label["Mã giám sát ▸ Họ và tên"].sample == "Nguyễn Văn Hùng"
    assert by_label["Mã kỹ thuật ▸ Họ và tên"].sample == "Trần Minh Tuấn"
    assert by_label["Ngày nghiệm thu (ngày … tháng … năm …)"].sample == "ngày 02 tháng 04 năm 2024"
    assert by_label["Tên công trình"].group == "Thông tin chung"
    files = ["PhieuYeuCauNghiemThu.docx", "BienBanNghiemThu.docx"]
    words = []
    for name in files:
        scan = q.scan_word(tmp_path / "kit" / "templates" / name)
        assert scan.style.value == "custom"
        words.append(q.WordChoice(f"templates/{name}", scan, q.suggest(scan.inputs, options)))
    record = words[1]
    assert [i.name for i in record.scan.inputs][:3] == [
        "Nơi lập biên bản",
        "Ngày nghiệm thu bằng chữ",
        "Số biên bản",
    ]
    assert record.bindings["Họ tên giám sát"] == "tra.ma_giam_sat.ho_va_ten"
    assert record.bindings["Chức vụ kỹ thuật"] == "tra.ma_ky_thuat.chuc_vu"
    assert "Ngày yêu cầu bằng chữ" not in words[0].bindings  # nothing like it in Excel: left empty
    assert "Các ý kiến khác" in [i.name for i in record.scan.inputs]  # blank kept for handwriting
    # the user prefers "08 giờ 00" for the start time
    q.apply_bindings(words, {(1, "Giờ bắt đầu"): "dong.gio_bat_dau_gio"})

    profile = q.build_profile(analysis, words, profile_id="cong_trinh", name=q.default_name(analysis))
    assert profile.profile.name == "Nhà làm việc 5 tầng – Công ty CP Đầu tư DEF"
    assert not [i for i in validate_profile(profile) if i.severity is Severity.ERROR]
    assert parse_profile_text(profile_to_text(profile)) == profile  # saves and loads back

    result = DocumentGenerationService().run(
        GenerationRequest(profile, tmp_path / "kit", workbook, tmp_path / "out")
    )
    assert result.status == "completed", [str(i) for i in result.all_issues()]
    first = Path(result.documents[0])
    assert first.name == "01 - Lắp dựng cốt thép cột tầng 1.docx"
    text = text_of(first)
    for expected in (
        "Kính gửi: Công ty TNHH Tư vấn Xây dựng XYZ",
        "Ông/Bà: Nguyễn Văn Hùng – Chức vụ: Kỹ sư tư vấn giám sát",
        "Ông/Bà: Trần Minh Tuấn – Chức vụ: Cán bộ kỹ thuật thi công",
        "Bắt đầu: 08 giờ 00 ngày 02 tháng 04 năm 2024",
        "Kết thúc: 09:30 ngày 02 tháng 04 năm 2024",
        "c) Các ý kiến khác (nếu có): ......",
    ):
        assert expected in text, expected
    assert "<<" not in text
    assert Path(result.combined_path).name == "Tổng hợp - Nhà làm việc 5 tầng – Công ty CP Đầu tư DEF.docx"


def test_existing_form_with_blanks_and_lookups(tmp_path):
    _, workbook = write_customer_c(tmp_path / "c")
    analysis = q.analyze_workbook(workbook)
    roles = {s.name: s.role for s in analysis.sheets}
    assert roles == {"DS nhân viên": q.ROLE_LIST, "Phòng ban": q.ROLE_LOOKUP, "Chức vụ": q.ROLE_LOOKUP}
    options = q.value_options(analysis)
    scan = q.scan_word(tmp_path / "c" / "templates" / "PhieuThongTin.docx")
    suggested = q.suggest(scan.inputs, options)
    assert suggested["Họ và tên"] == "dong.ho_va_ten"  # blank after the label
    assert suggested["Email"] == "dong.email"  # empty table cell right of the label
    assert suggested["Số CCCD"] == "dong.so_cccd"
    assert "tháng" not in suggested and "năm" not in suggested  # too generic to guess
    photo = next(o for o in options if o.path == "dong.anh")
    assert photo.label == "Ảnh (ảnh)" and photo.rule["type"] == "image"
    assert suggested["anh"] == "dong.anh"  # the 3x4 photo frame (picture content control)
    # choosing the month part of the birth date is one click in the wizard
    words = [q.WordChoice("templates/PhieuThongTin.docx", scan, suggested)]
    q.apply_bindings(words, {(0, "tháng"): "dong.ngay_sinh_thang", (0, "năm"): "dong.ngay_sinh_nam"})
    profile = q.build_profile(analysis, words, profile_id="phieu", name="Phiếu", combined=False)
    result = DocumentGenerationService().run(
        GenerationRequest(profile, tmp_path / "c", workbook, tmp_path / "o")
    )
    assert result.status == "completed", [str(i) for i in result.all_issues()]
    text = text_of(Path(result.documents[0]))
    assert "Họ và tên: nguyễn văn ánh" in text
    assert "tháng 09 năm 1990" in text
    assert result.combined_path == ""
    wp = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
    drawing = etree.fromstring(zipfile.ZipFile(result.documents[0]).read("word/document.xml"))
    assert len(list(drawing.iter(wp + "extent"))) == 1  # NV01's photo in the frame
    assert not [i for i in result.all_issues() if i.code == "template.picture_not_image"]


def test_suggestions_never_guess_between_equal_candidates():
    options = [
        q.ValueOption("dong.a", "Ngày ký", "Danh sách", "", {"source": "a"}),
        q.ValueOption("dong.b", "Ngay ky", "Danh sách", "", {"source": "b"}),
        q.ValueOption("dong.c", "Địa chỉ", "Danh sách", "", {"source": "c"}),
    ]
    inputs = [
        q.WordInput("Ngày ký", "marker", "", ""),
        q.WordInput("ngày", "blank", "label", ""),
        q.WordInput("Địa chỉ:", "blank", "label", ""),
    ]
    assert q.suggest(inputs, options) == {"Địa chỉ:": "dong.c"}
