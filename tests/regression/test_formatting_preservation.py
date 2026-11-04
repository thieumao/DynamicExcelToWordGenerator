"""Rendering must not lose template formatting (fonts, styles, tables, images, sections...).

The template is built with real Word features via python-docx/OOXML. After
rendering, the XML-level properties are compared with the source template.
This proves preservation for the features covered here only; real customer
templates must be added to tests/fixtures/real (see docs/TEST_STRATEGY.md).
"""

from __future__ import annotations

import io
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_BREAK
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor

from docgen.infrastructure.word.merger import MergePart, WordMergeEngine
from docgen.infrastructure.word.renderer import TemplateRenderer
from docgen.testing.docx_builder import add_hyperlink, add_page_number_field, tiny_png


def build_rich_template(path: Path) -> Path:
    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(11)
    sec = doc.sections[0]
    sec.left_margin = Mm(25)
    sec.header.paragraphs[0].text = "Header {{ x.code }}"
    sec.header.paragraphs[0].add_run().add_picture(io.BytesIO(tiny_png()), width=Mm(5))
    fpar = sec.footer.paragraphs[0]
    fpar.text = "Trang "
    add_page_number_field(fpar)
    doc.add_heading("Tiêu đề {{ x.title }}", level=1)
    p = doc.add_paragraph()
    r = p.add_run("{{ x.name }}")
    r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
    r.italic = True
    r.font.name = "Courier New"
    doc.add_paragraph("Mục {{ x.a }}", style="List Number")
    doc.add_paragraph("Mục {{ x.b }}", style="List Number")
    doc.add_paragraph("Ý {{ x.a }}", style="List Bullet")
    table = doc.add_table(rows=2, cols=2)
    table.style = "Light Grid Accent 1"
    table.cell(0, 0).merge(table.cell(0, 1)).text = "Gộp {{ x.code }}"
    table.cell(1, 0).text, table.cell(1, 1).text = "{{ x.a }}", "{{ x.b }}"
    doc.add_picture(io.BytesIO(tiny_png((0, 128, 0))), width=Mm(30))
    link = doc.add_paragraph("Xem ")
    add_hyperlink(link, "https://example.com/{{ x.code }}", "{{ x.name }}")
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    land = doc.add_section()
    land.orientation = WD_ORIENT.LANDSCAPE
    land.page_width, land.page_height = sec.page_height, sec.page_width
    doc.add_paragraph("Trang ngang {{ x.code }}")
    doc.save(str(path))
    return path


def snapshot(path: Path) -> dict:
    doc = Document(str(path))
    body = doc.element.body
    return {
        "styles": sorted(s.name for s in doc.styles),
        "normal_font": (doc.styles["Normal"].font.name, doc.styles["Normal"].font.size),
        "paragraph_styles": [p.style.name for p in doc.paragraphs],
        "run_formats": [
            (r.italic, r.font.name, str(r.font.color.rgb) if r.font.color.type else None)
            for p in doc.paragraphs
            for r in p.runs
            if r.italic
        ],
        "tables": [(t.style.name, len(t.rows), len(t.columns)) for t in doc.tables],
        "merged_cells": len(body.xpath(".//w:gridSpan")),
        "numbering": len(body.xpath(".//w:numPr"))
        + sum(1 for p in doc.paragraphs if p.style.name in ("List Number", "List Bullet")),
        "images": len(doc.inline_shapes),
        "header_images": len(doc.sections[0].header.part.element.xpath(".//a:blip")),
        "page_breaks": len(body.xpath(".//w:br[@w:type='page']")),
        "sections": [(s.orientation, s.page_width, s.page_height, s.left_margin) for s in doc.sections],
        "footer_fields": len(doc.sections[0].footer.part.element.xpath(".//w:instrText")),
        "hyperlinks": len(body.xpath(".//w:hyperlink")),
    }


CTX = {"x": {"code": "K-01", "title": "Báo cáo", "name": "Nguyễn Ánh", "a": "một", "b": "hai"}}


def test_render_preserves_formatting(tmp_path):
    template = build_rich_template(tmp_path / "rich.docx")
    out = tmp_path / "out.docx"
    TemplateRenderer().render(template.read_bytes(), CTX, out)
    before, after = snapshot(template), snapshot(out)
    assert before == after
    doc = Document(str(out))
    assert doc.sections[0].header.paragraphs[0].text == "Header K-01"
    links = [r for r in doc.part.rels.values() if r.reltype == RT.HYPERLINK]
    assert links and doc.element.body.xpath(".//w:hyperlink//w:t")[0].text == "Nguyễn Ánh"
    assert doc.element.body.xpath(".//w:sectPr/w:pgSz[@w:orient='landscape']")


def test_merge_keeps_both_documents_formatting(tmp_path):
    template = build_rich_template(tmp_path / "rich.docx")
    a, b = tmp_path / "a.docx", tmp_path / "b.docx"
    renderer = TemplateRenderer()
    renderer.render(template.read_bytes(), CTX, a)
    renderer.render(template.read_bytes(), {"x": {**CTX["x"], "code": "K-02"}}, b)
    merged = tmp_path / "m.docx"
    WordMergeEngine().merge([MergePart(a), MergePart(b, section_mode="preserve")], merged)
    snap = snapshot(merged)
    single = snapshot(a)
    assert snap["images"] == 2 * single["images"]
    assert snap["tables"] == 2 * single["tables"]
    assert snap["hyperlinks"] == 2 * single["hyperlinks"]
    assert len(snap["sections"]) == 2 * len(single["sections"])
    assert snap["sections"] == single["sections"] * 2
    doc = Document(str(merged))
    headers = [s.header.paragraphs[0].text for s in doc.sections]
    assert headers[0] == "Header K-01" and headers[2] == "Header K-02"
    assert doc.sections[2].header.part.element.xpath(".//a:blip")
    assert doc.element.body.xpath(".//w:sectPr/w:pgSz[@w:orient='landscape']")
    assert qn("w:sectPr")
