"""Synthetic DOCX fixture builders.

Used by tests, benchmarks and the sample generator. The builders only use
python-docx plus a little raw OOXML for features python-docx has no API for
(hyperlinks, page-number fields), so fixtures are reproducible on any OS.
"""

from __future__ import annotations

import struct
import zlib
from collections.abc import Sequence
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor
from docx.text.paragraph import Paragraph


def tiny_png(rgb: tuple[int, int, int] = (200, 30, 30), size: tuple[int, int] = (8, 8)) -> bytes:
    """Return bytes of a small solid-colour PNG (no external imaging library needed)."""
    width, height = size
    raw = b"".join(b"\x00" + bytes(rgb) * width for _ in range(height))

    def chunk(tag: bytes, data: bytes) -> bytes:
        body = tag + data
        return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body) & 0xFFFFFFFF)

    ihdr = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", ihdr)
        + chunk(b"IDAT", zlib.compress(raw, 9))
        + chunk(b"IEND", b"")
    )


def add_hyperlink(paragraph: Paragraph, url: str, text: str) -> None:
    """Append an external hyperlink run to *paragraph*."""
    part = paragraph.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    rpr.append(color)
    rpr.append(underline)
    run.append(rpr)
    text_el = OxmlElement("w:t")
    text_el.text = text
    run.append(text_el)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def add_page_number_field(paragraph: Paragraph) -> None:
    """Append a PAGE field (complex field) to *paragraph*."""
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    num = OxmlElement("w:t")
    num.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for el in (begin, instr, sep, num, end):
        run._r.append(el)


def add_merge_field(paragraph: Paragraph, name: str, *, simple: bool = False, bold: bool = False) -> None:
    """Append a Word MERGEFIELD (simple or complex form) to *paragraph*."""
    if simple:
        fld = OxmlElement("w:fldSimple")
        fld.set(qn("w:instr"), f" MERGEFIELD {name} \\* MERGEFORMAT ")
        run = OxmlElement("w:r")
        if bold:
            rpr = OxmlElement("w:rPr")
            rpr.append(OxmlElement("w:b"))
            run.append(rpr)
        t = OxmlElement("w:t")
        t.text = f"«{name}»"
        run.append(t)
        fld.append(run)
        paragraph._p.append(fld)
        return

    def new_run() -> Any:
        r = OxmlElement("w:r")
        if bold:
            rpr = OxmlElement("w:rPr")
            rpr.append(OxmlElement("w:b"))
            r.append(rpr)
        return r

    r1 = new_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    r1.append(begin)
    r2 = new_run()
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" MERGEFIELD  {name} \\* MERGEFORMAT "
    r2.append(instr)
    r3 = new_run()
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    r3.append(sep)
    r4 = new_run()
    t = OxmlElement("w:t")
    t.text = f"«{name}»"
    r4.append(t)
    r5 = new_run()
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    r5.append(end)
    for r in (r1, r2, r3, r4, r5):
        paragraph._p.append(r)


@dataclass
class DocSpec:
    """Declarative description of a synthetic DOCX fixture."""

    title: str = "Document"
    paragraphs: Sequence[str] = ()
    header_text: str | None = None
    footer_page_number: bool = False
    footer_text: str | None = None
    landscape: bool = False
    margins_mm: float | None = None
    table_header: Sequence[str] | None = None
    table_rows: Sequence[Sequence[str]] = ()
    numbered_items: Sequence[str] = ()
    bullet_items: Sequence[str] = ()
    image: bool = False
    header_image: bool = False
    hyperlink: tuple[str, str] | None = None
    font_name: str | None = None
    font_size_pt: float | None = None
    custom_style: str | None = None
    split_runs: Sequence[str] = field(default_factory=tuple)
    """Paragraph texts written as one run *per character*, mimicking Word's run splitting."""


def build_docx(path: Path, spec: DocSpec) -> Path:
    """Create a DOCX file described by *spec* and return its path."""
    doc = Document()
    if spec.font_name or spec.font_size_pt:
        normal = doc.styles["Normal"]
        if spec.font_name:
            normal.font.name = spec.font_name
            rpr = normal.element.get_or_add_rPr()
            rfonts = rpr.find(qn("w:rFonts"))
            if rfonts is None:
                rfonts = OxmlElement("w:rFonts")
                rpr.append(rfonts)
            for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
                rfonts.set(qn(attr), spec.font_name)
        if spec.font_size_pt:
            normal.font.size = Pt(spec.font_size_pt)

    section = doc.sections[0]
    if spec.landscape:
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width, section.page_height = section.page_height, section.page_width
    if spec.margins_mm is not None:
        for attr in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
            setattr(section, attr, Mm(spec.margins_mm))

    if spec.header_text is not None or spec.header_image:
        header_par = section.header.paragraphs[0]
        if spec.header_text is not None:
            header_par.text = spec.header_text
        if spec.header_image:
            import io

            header_par.add_run().add_picture(io.BytesIO(tiny_png((30, 30, 200))), width=Mm(5))
    if spec.footer_text is not None or spec.footer_page_number:
        footer_par = section.footer.paragraphs[0]
        if spec.footer_text is not None:
            footer_par.text = spec.footer_text
        if spec.footer_page_number:
            footer_par.add_run(" Page ")
            add_page_number_field(footer_par)

    doc.add_heading(spec.title, level=1)
    if spec.custom_style:
        from docx.enum.style import WD_STYLE_TYPE

        style = doc.styles.add_style(spec.custom_style, WD_STYLE_TYPE.PARAGRAPH)
        style.font.color.rgb = RGBColor(0x11, 0x88, 0x11)
        style.font.bold = True
        doc.add_paragraph(f"Styled with {spec.custom_style}", style=spec.custom_style)

    for text in spec.paragraphs:
        doc.add_paragraph(text)
    for text in spec.split_runs:
        par = doc.add_paragraph()
        for ch in text:
            par.add_run(ch)
    for item in spec.numbered_items:
        doc.add_paragraph(item, style="List Number")
    for item in spec.bullet_items:
        doc.add_paragraph(item, style="List Bullet")

    if spec.table_header:
        table = doc.add_table(rows=1, cols=len(spec.table_header))
        table.style = "Table Grid"
        for cell, text in zip(table.rows[0].cells, spec.table_header, strict=True):
            cell.text = text
        for row in spec.table_rows:
            cells = table.add_row().cells
            for cell, text in zip(cells, row, strict=True):
                cell.text = text

    if spec.image:
        import io

        doc.add_picture(io.BytesIO(tiny_png()), width=Mm(20))
    if spec.hyperlink:
        par = doc.add_paragraph("Link: ")
        add_hyperlink(par, spec.hyperlink[0], spec.hyperlink[1])

    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(path))
    return path
