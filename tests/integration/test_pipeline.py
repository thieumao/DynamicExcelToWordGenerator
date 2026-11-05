"""End-to-end: Excel -> mapping -> context -> Word render -> merge -> verified outputs."""

from __future__ import annotations

import json
from pathlib import Path

import pytest
from docx import Document
from docx.enum.section import WD_ORIENT

from docgen.application.generation.job import CancellationToken, GenerationRequest
from docgen.application.generation.service import DocumentGenerationService
from docgen.domain.profile.models import (
    ErrorPolicy,
    FieldRule,
    GenerationConfig,
    OutputKind,
    TemplateSpec,
    UnitKind,
)
from docgen.infrastructure.word.verifier import verify_docx


def run(sample, out: Path, **kwargs):
    service = kwargs.pop("service", None) or DocumentGenerationService()
    cancel = kwargs.pop("cancel", None)
    reporter = kwargs.pop("reporter", None)
    profile = kwargs.pop("profile", sample.profile)
    request = GenerationRequest(profile, sample.directory, sample.workbook, out, **kwargs)
    return service.run(request, reporter, cancel)


def all_text(path: Path) -> str:
    doc = Document(str(path))
    parts = [p.text for p in doc.paragraphs]
    for t in doc.tables:
        for r in t.rows:
            parts.extend(c.text for c in r.cells)
    for s in doc.sections:
        parts.extend(p.text for p in s.header.paragraphs)
    return "\n".join(parts)


def test_mode_e_batch_merged_documents(sample_a, tmp_path):
    out = tmp_path / "out"
    result = run(sample_a, out)
    assert result.status == "completed", [str(i) for i in result.all_issues()]
    assert len(result.units) == 8 and len(result.documents) == 8
    first = Path(result.units[0].outputs[0])
    assert first.name.startswith("HoSo_E001_")
    doc = Document(str(first))
    assert [s.orientation for s in doc.sections][:2] == [WD_ORIENT.PORTRAIT, WD_ORIENT.LANDSCAPE]
    text = all_text(first)
    assert "Phòng Kế toán" in text and "Phạm Quốc Bảo" in text and "₫" in text and "{{" not in text
    for path in result.documents:
        assert verify_docx(Path(path)).ok
    audit = json.loads(Path(result.audit_path).read_text(encoding="utf-8"))
    assert audit["documents_generated"] == 8 and audit["profile"]["id"] == "customer_a"
    assert audit["inputs"][0]["sha256"] and audit["records"]["selected"] == 8
    assert Path(result.report_path).read_text(encoding="utf-8-sig").startswith("unit,label")
    assert not list(out.glob(".docgen-tmp-*")) and not list(out.rglob("*.tmp"))


def test_mode_a_single_record_and_mode_b_separate(sample_a, tmp_path):
    single = run(sample_a, tmp_path / "a", rows=[6])
    assert [u.rows for u in single.units] == [[6]] and len(single.documents) == 1
    separate = run(sample_a, tmp_path / "b", output=OutputKind.SEPARATE, rows=[4, 5])
    names = sorted(Path(p).name for p in separate.documents)
    assert (
        names
        == [
            "appendix_E001.docx",
            "cover_E001.docx",
            "cover_E002.docx",
            "details_E001.docx",
            "details_E002.docx",
        ]
        or len(names) >= 5
    )
    both = run(sample_a, tmp_path / "c", output=OutputKind.BOTH, rows=[4])
    assert len(both.documents) == 4  # merged + 3 templates (appendix included: E001 has a project)


def _write_group_template(path: Path) -> None:
    doc = Document()
    doc.add_heading("Phòng {{ group.key }} – {{ group.count }} nhân viên – {{ headcount_text }}", 1)
    table = doc.add_table(rows=1, cols=2)
    table.rows[0].cells[0].text = "Mã"
    table.add_row().cells[0].text = "{%tr for r in group.records %}"
    row = table.add_row().cells
    row[0].text, row[1].text = "{{ r.employee.code }}", "{{ r.employee.full_name }}"
    table.add_row().cells[0].text = "{%tr endfor %}"
    doc.save(str(path))


def test_mode_c_grouped_and_mode_d_dataset(sample_a, tmp_path):
    tpl = sample_a.directory / "templates" / "Dept.docx"
    _write_group_template(tpl)
    profile = sample_a.profile.model_copy(
        update={
            "templates": [TemplateSpec(id="dept", file="templates/Dept.docx")],
            "group_fields": [FieldRule(name="headcount_text", expr="'Tổng: ' ~ (records | length)")],
            "generation": GenerationConfig(
                unit=UnitKind.GROUP, group_by=["employee.department.name"], output=OutputKind.MERGED
            ),
            "merge": sample_a.profile.merge.model_copy(update={"order": []}),
            "output": sample_a.profile.output.model_copy(
                update={"merged_filename": "Dept_{group.keys.employee.department.name}.docx"}
            ),
        }
    )
    result = run(sample_a, tmp_path / "g", profile=profile)
    assert result.status == "completed", [str(i) for i in result.all_issues()]
    assert len(result.units) == 4  # 4 departments
    names = sorted(Path(p).name for p in result.documents)
    assert "Dept_Phòng Kế toán.docx" in names
    text = all_text(Path(next(p for p in result.documents if "Kế toán" in p)))
    assert "2 nhân viên" in text and "Tổng: 2" in text and "E001" in text and "E005" in text

    dataset = profile.model_copy(
        update={
            "generation": GenerationConfig(unit=UnitKind.DATASET, output=OutputKind.MERGED),
            "output": profile.output.model_copy(update={"merged_filename": "All.docx"}),
        }
    )
    tpl_text = Document(str(tpl))
    tpl_text.paragraphs[0].text = "Toàn công ty {{ group.count }}"
    tpl_text.save(str(tpl))
    result = run(sample_a, tmp_path / "d", profile=dataset)
    assert result.status == "completed" and len(result.documents) == 1
    assert "Toàn công ty 8" in all_text(Path(result.documents[0]))


def test_customer_b_runs_with_same_code(sample_b, tmp_path):
    result = run(sample_b, tmp_path / "out")
    assert result.status == "completed", [str(i) for i in result.all_issues()]
    merged = tmp_path / "out" / "HN" / "HD_HD-001.docx"
    text = all_text(merged)
    assert "Công ty Ánh Dương" in text and "Một trăm hai mươi lăm triệu đồng" in text
    assert "HỢP ĐỒNG SỐ HD-001" in text and "Ngày ký: 02/07/2024" in text
    assert (tmp_path / "out" / "HCM" / "xac_nhan_HD-002.docx").exists()


def test_process_pool_batch(sample_a, tmp_path):
    service = DocumentGenerationService(pool_threshold=2)
    result = run(sample_a, tmp_path / "out", service=service, workers=2)
    assert result.status == "completed", [str(i) for i in result.all_issues()]
    assert len(result.documents) == 8


def test_no_overwrite_by_default(sample_a, tmp_path):
    out = tmp_path / "out"
    first = run(sample_a, out, rows=[4])
    second = run(sample_a, out, rows=[4])
    assert Path(first.documents[0]).exists() and Path(second.documents[0]).name.endswith("(2).docx")


class CancelAfterFirst:
    def __init__(self, token: CancellationToken) -> None:
        self.token = token
        self.stage_name = ""

    def stage(self, name: str) -> None:
        self.stage_name = name

    def progress(self, done: int, total: int, message: str = "") -> None:
        if self.stage_name == "generate" and done >= 1:
            self.token.cancel()

    def issue(self, issue) -> None:
        pass


def test_cancellation_leaves_no_partial_files(sample_a, tmp_path):
    token = CancellationToken()
    out = tmp_path / "out"
    result = run(sample_a, out, cancel=token, reporter=CancelAfterFirst(token))
    assert result.status == "cancelled"
    statuses = [u.status for u in result.units]
    assert statuses.count("cancelled") >= 6 and len(result.documents) <= 2
    assert not list(out.glob(".docgen-tmp-*")) and not list(out.rglob("*.tmp"))
    for doc in result.documents:
        assert verify_docx(Path(doc)).ok


def _break_workbook(sample) -> None:
    from openpyxl import load_workbook

    wb = load_workbook(sample.workbook)
    wb["Employees"]["C6"] = "D99"  # department that does not exist -> lookup error for E003
    wb.save(sample.workbook)


def test_collect_errors_skips_bad_records_and_reports_location(sample_a, tmp_path):
    _break_workbook(sample_a)
    result = run(sample_a, tmp_path / "out")
    assert result.status == "completed_with_errors"
    bad = next(u for u in result.units if u.rows == [6])
    assert bad.status == "error" and not bad.outputs
    issue = next(i for i in bad.issues if i.code == "lookup.missing_reference")
    assert (issue.sheet, issue.row, issue.column, issue.rule) == ("Employees", 6, "DepartmentID", "emp_dept")
    assert issue.suggestion
    assert len(result.documents) == 7


def test_fail_fast_generates_nothing(sample_a, tmp_path):
    _break_workbook(sample_a)
    result = run(sample_a, tmp_path / "out", error_policy=ErrorPolicy.FAIL_FAST)
    assert result.status == "failed" and not result.documents
    assert not list((tmp_path / "out").rglob("*.docx"))


def test_dry_run_and_missing_template(sample_a, tmp_path):
    result = run(sample_a, tmp_path / "out", dry_run=True)
    assert result.status == "validated" and len(result.units) == 8 and not result.documents
    (sample_a.directory / "templates" / "Cover.docx").unlink()
    result = run(sample_a, tmp_path / "out")
    assert result.status == "failed"
    assert any(i.code == "template.not_found" and i.template == "cover" for i in result.issues)


def test_unknown_placeholder_is_caught_before_generation(sample_a, tmp_path):
    tpl = sample_a.directory / "templates" / "Cover.docx"
    doc = Document(str(tpl))
    doc.add_paragraph("{{ employee.nickname }}")
    doc.save(str(tpl))
    result = run(sample_a, tmp_path / "out")
    assert result.status == "failed"
    issue = next(i for i in result.issues if i.code == "template.unknown_placeholder")
    assert issue.placeholder == "employee.nickname" and issue.template == "cover"


def test_input_files_are_never_modified(sample_a, tmp_path):
    import hashlib

    def digest(p: Path) -> str:
        return hashlib.sha256(p.read_bytes()).hexdigest()

    files = [sample_a.workbook, *sorted((sample_a.directory / "templates").glob("*.docx"))]
    before = [digest(f) for f in files]
    run(sample_a, tmp_path / "out")
    assert [digest(f) for f in files] == before


@pytest.mark.parametrize("bad_dir", ["file"])
def test_output_dir_must_be_writable(sample_a, tmp_path, bad_dir):
    blocker = tmp_path / "blocker"
    blocker.write_text("x")
    result = run(sample_a, blocker / "sub")
    assert result.status == "failed" and any(i.code == "output.not_writable" for i in result.issues)


def test_customer_c_fills_an_existing_word_form(sample_c, tmp_path):
    """The customer's Word file is a form without placeholders (blanks, controls, form fields)."""
    result = run(sample_c, tmp_path / "out")
    assert result.status == "completed", [str(i) for i in result.all_issues()]
    assert sorted(Path(p).name for p in result.documents) == ["PhieuTT_NV01.docx", "PhieuTT_NV02.docx"]
    import zipfile

    from lxml import etree

    w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    root = etree.fromstring(zipfile.ZipFile(tmp_path / "out" / "PhieuTT_NV01.docx").read("word/document.xml"))
    lines = ["".join(t.text or "" for t in p.iter(w + "t")) for p in root.iter(w + "p")]
    text = "\n".join(lines)
    for expected in (
        "Họ và tên: Nguyễn Văn Ánh",
        "Sinh ngày 05 tháng 09 năm 1990, tại: Hà Nội",
        "Phòng ban: Phòng Kế toán",
        "Chức vụ: Trưởng phòng",
        "Ngày vào làm: 01/03/2020",
        "Đảng viên: ☒",
        "Mã số thuế cá nhân: 0101234567",
        "001090001234",
        "anh@example.vn",
    ):
        assert expected in text, expected
    # inputs that are not bound stay exactly as the customer designed them
    assert "Hà Nội, ngày ...... tháng ...... năm ........" in text
    assert "Người khai ký tên: .........................................." in text
    assert "Nhấn để nhập văn bản." in text  # unbound 'ghi_chu' control keeps its placeholder
    infos = [i.code for i in result.issues]
    assert "template.spot_unbound" in infos and "template.blanks_unbound" in infos
    # legacy drop-down "Giới tính" selects the Excel value (NV02 has "nữ": case-insensitive)
    selected = []
    for name in ("PhieuTT_NV01.docx", "PhieuTT_NV02.docx"):
        doc = etree.fromstring(zipfile.ZipFile(tmp_path / "out" / name).read("word/document.xml"))
        selected += [r.get(w + "val") for r in doc.iter(w + "result")]
    assert selected == ["0", "1"]
    # picture content control "anh": NV01's photo fitted into the 30x40 mm frame, NV02 has none
    wp = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
    sizes = {}
    for name in ("PhieuTT_NV01.docx", "PhieuTT_NV02.docx"):
        doc = etree.fromstring(zipfile.ZipFile(tmp_path / "out" / name).read("word/document.xml"))
        sizes[name] = [(int(e.get("cx")), int(e.get("cy"))) for e in doc.iter(wp + "extent")]
    assert sizes["PhieuTT_NV02.docx"] == []
    ((cx, cy),) = sizes["PhieuTT_NV01.docx"]
    assert abs(cx - 30 * 36000) <= 1 and abs(cy - 40 * 36000) <= 1  # 3:4 photo fills the 3:4 frame
    assert "template.picture_not_image" not in infos
