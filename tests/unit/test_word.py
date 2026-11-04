"""Marker adapters, inspector, renderer and verifier."""

from __future__ import annotations

import datetime as dt
import io
import zipfile
from decimal import Decimal
from pathlib import Path

import pytest
from docx import Document
from docx.shared import Pt

from docgen.domain.mapping.engine import ImageRef
from docgen.domain.profile.models import MarkerConfig, MarkerStyle, MissingPolicy
from docgen.infrastructure.word.inspector import RefKind, inspect_template
from docgen.infrastructure.word.markers import convert_custom_markers, convert_mergefields
from docgen.infrastructure.word.renderer import RenderError, TemplateRenderer
from docgen.infrastructure.word.verifier import verify_docx
from docgen.testing.docx_builder import add_merge_field, tiny_png


def docx_bytes(build) -> bytes:
    doc = Document()
    build(doc)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def texts(data: bytes | Path) -> list[str]:
    doc = Document(io.BytesIO(data) if isinstance(data, bytes) else str(data))
    return [p.text for p in doc.paragraphs]


# ------------------------------------------------------------------ markers
def test_custom_markers_split_across_runs_keep_first_run_format():
    def build(doc):
        p = doc.add_paragraph("Kính gửi ")
        r = p.add_run("<<Họ")
        r.bold = True
        p.add_run(" tên>> và <<Mã>>.")
        doc.sections[0].header.paragraphs[0].text = "Số: [[SoHD]]"

    data = docx_bytes(build)
    result = convert_custom_markers(data, "<<", ">>")
    assert result.names == ["Họ tên", "Mã"]
    assert texts(result.data)[0] == 'Kính gửi {{ b["Họ tên"] }} và {{ b["Mã"] }}.'
    doc = Document(io.BytesIO(result.data))
    bold_runs = [r.text for r in doc.paragraphs[0].runs if r.bold]
    assert bold_runs == ['{{ b["Họ tên"] }}']
    header = convert_custom_markers(data, "[[", "]]")
    assert header.names == ["SoHD"] and header.hits[0].part == "word/header1.xml"


def test_custom_markers_reject_unsafe_names():
    data = docx_bytes(lambda d: d.add_paragraph('<<a"b>> <<ok>>'))
    result = convert_custom_markers(data, "<<", ">>")
    assert result.names == ["ok"] and result.problems


def test_mergefields_simple_and_complex():
    def build(doc):
        p = doc.add_paragraph("A: ")
        add_merge_field(p, "Họ_tên", bold=True)
        p = doc.add_paragraph("B: ")
        add_merge_field(p, "Code", simple=True)

    result = convert_mergefields(docx_bytes(build))
    assert result.names == ["Code", "Họ_tên"] or set(result.names) == {"Code", "Họ_tên"}
    assert texts(result.data)[:2] == ['A: {{ b["Họ_tên"] }}', 'B: {{ b["Code"] }}']
    doc = Document(io.BytesIO(result.data))
    assert doc.paragraphs[0].runs[-1].bold
    assert "fldChar" not in doc.element.body.xml


def test_mergefield_region_markers_are_reported():
    data = docx_bytes(lambda d: add_merge_field(d.add_paragraph(), "TableStart:Items"))
    result = convert_mergefields(data)
    assert not result.names and "unsupported" in result.problems[0]


# ---------------------------------------------------------------- inspector
def test_inspector_resolves_loops_conditions_and_headers(tmp_path):
    def build(doc):
        doc.sections[0].header.paragraphs[0].text = "{{ report.profile.name }}"
        doc.add_paragraph("{{ employee.name | upper }}")
        doc.add_paragraph("{%p if employee.projects %}")
        doc.add_paragraph("{%p for p in employee.projects %}")
        doc.add_paragraph("{{ p.name }} {{ loop.index }} {{ p.owner.name }}")
        doc.add_paragraph("{%p endfor %}")
        doc.add_paragraph("{%p endif %}")

    path = tmp_path / "t.docx"
    path.write_bytes(docx_bytes(build))
    result = inspect_template(str(path))
    refs = {(r.path, r.kind) for r in result.refs}
    assert ("employee.name", RefKind.VALUE) in refs
    assert ("employee.projects", RefKind.CONDITION) in refs
    assert ("employee.projects", RefKind.LOOP) in refs
    assert ("employee.projects[].name", RefKind.VALUE) in refs
    assert ("employee.projects[].owner.name", RefKind.VALUE) in refs
    assert ("report.profile.name", RefKind.VALUE) in refs
    assert not any("loop" in r.path for r in result.refs)


def test_inspector_reports_syntax_errors_and_foreign_markers(tmp_path):
    path = tmp_path / "bad.docx"
    path.write_bytes(docx_bytes(lambda d: (d.add_paragraph("{% if x %} no end"), d.add_paragraph("«Name»"))))
    result = inspect_template(str(path))
    assert result.errors and "syntax" in result.errors[0]
    assert any("mergefield" in w for w in result.warnings)


def test_inspector_custom_markers_give_binding_names(tmp_path):
    path = tmp_path / "c.docx"
    path.write_bytes(docx_bytes(lambda d: d.add_paragraph("<<Số HĐ>> – {{ hop_dong.ngay }}")))
    result = inspect_template(str(path), MarkerConfig(style=MarkerStyle.CUSTOM))
    assert result.binding_names == ["Số HĐ"] and result.paths == ["hop_dong.ngay"]


# ----------------------------------------------------------------- renderer
@pytest.fixture()
def renderer(tmp_path):
    return TemplateRenderer(image_root=tmp_path)


def test_render_values_tables_conditions_and_escaping(tmp_path, renderer):
    def build(doc):
        doc.sections[0].footer.paragraphs[0].text = "Footer {{ c.code }}"
        p = doc.add_paragraph("Tên: ")
        run = p.add_run("{{ c.name }}")
        run.bold = True
        run.font.size = Pt(15)
        doc.add_paragraph("Ngày {{ c.date }} – số {{ c.amount }} – {{ c.flag }} – {{ c.note }}")
        doc.add_paragraph("{%p if c.vip %}")
        doc.add_paragraph("Khách VIP")
        doc.add_paragraph("{%p endif %}")
        table = doc.add_table(rows=1, cols=2)
        table.rows[0].cells[0].text = "Item"
        table.add_row().cells[0].text = "{%tr for it in c.items %}"
        row = table.add_row().cells
        row[0].text, row[1].text = "{{ it.name }}", "{{ it.qty }}"
        table.add_row().cells[0].text = "{%tr endfor %}"

    ctx = {
        "c": {
            "name": "Công ty <A&B>",
            "date": dt.date(2024, 2, 9),
            "amount": Decimal("1234.50"),
            "flag": True,
            "note": "dòng 1\ndòng 2",
            "vip": False,
            "code": "X1",
            "items": [{"name": "Bút", "qty": 2}, {"name": "Vở", "qty": 3}],
        }
    }
    out = tmp_path / "out.docx"
    result = renderer.render(docx_bytes(build), ctx, out)
    assert not result.issues
    doc = Document(str(out))
    assert doc.paragraphs[0].text == "Tên: Công ty <A&B>"
    styled = next(r for r in doc.paragraphs[0].runs if r.text.startswith("Công ty"))
    assert styled.bold and styled.font.size == Pt(15)
    assert doc.paragraphs[1].text == "Ngày 09/02/2024 – số 1234,5 – Có – dòng 1\ndòng 2"
    assert "Khách VIP" not in [p.text for p in doc.paragraphs]
    assert [[c.text for c in r.cells] for r in doc.tables[0].rows] == [
        ["Item", ""],
        ["Bút", "2"],
        ["Vở", "3"],
    ]
    assert doc.sections[0].footer.paragraphs[0].text == "Footer X1"
    assert verify_docx(out).ok


def test_missing_placeholder_policies(tmp_path, renderer):
    data = docx_bytes(lambda d: d.add_paragraph("A {{ c.missing }} B {{ c.ok }}"))
    with pytest.raises(RenderError) as exc:
        renderer.render(data, {"c": {"ok": 1}}, tmp_path / "x.docx", template_id="t")
    assert exc.value.issues[0].placeholder == "missing" and exc.value.issues[0].template == "t"
    assert not (tmp_path / "x.docx").exists()
    res = renderer.render(data, {"c": {"ok": 1}}, tmp_path / "w.docx", missing_policy=MissingPolicy.WARN)
    assert res.missing == ["missing"] and texts(tmp_path / "w.docx")[0] == "A  B 1"
    res = renderer.render(data, {"c": {"ok": 1}}, tmp_path / "e.docx", missing_policy=MissingPolicy.EMPTY)
    assert not res.issues


def test_template_syntax_error_is_reported(tmp_path, renderer):
    data = docx_bytes(lambda d: d.add_paragraph("{% if x %} never closed"))
    with pytest.raises(RenderError) as exc:
        renderer.render(data, {}, tmp_path / "x.docx")
    assert exc.value.issues[-1].code == "render.template_error"


def test_images_are_resolved_safely(tmp_path, renderer):
    (tmp_path / "img").mkdir()
    (tmp_path / "img" / "a.png").write_bytes(tiny_png())
    outside = tmp_path.parent / "outside.png"
    outside.write_bytes(tiny_png())
    data = docx_bytes(lambda d: d.add_paragraph("{{ a }}|{{ b }}|{{ c }}"))
    ctx = {
        "a": ImageRef("img/a.png", width_mm=10),
        "b": ImageRef("../outside.png"),
        "c": ImageRef("img/x.png"),
    }
    res = renderer.render(data, ctx, tmp_path / "img.docx")
    doc = Document(str(tmp_path / "img.docx"))
    assert len(doc.inline_shapes) == 1
    problems = " ".join(i.message for i in res.issues)
    assert "outside the allowed folder" in problems and "not found" in problems


# ----------------------------------------------------------------- verifier
def test_verifier_detects_problems(tmp_path):
    bad = tmp_path / "bad.docx"
    bad.write_bytes(b"garbage")
    assert not verify_docx(bad).ok
    good = tmp_path / "good.docx"
    good.write_bytes(docx_bytes(lambda d: d.add_paragraph("left {{ over }}")))
    result = verify_docx(good)
    assert not result.ok and result.leftover_markers
    broken = tmp_path / "broken.docx"
    with zipfile.ZipFile(good) as zin, zipfile.ZipFile(broken, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/_rels/document.xml.rels":
                data = data.replace(b"styles.xml", b"missing.xml")
            zout.writestr(item, data)
    assert any("Broken relationship" in p for p in verify_docx(broken).problems)
