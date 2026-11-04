"""Proof-of-concept: does merging keep sections, headers/footers and content?

These tests verify the OOXML structure of merged output. When LibreOffice is
available they also render to PDF and check page count / per-page header text,
which is an independent layout engine check (not a substitute for MS Word).
"""

from __future__ import annotations

import re
import shutil
import subprocess
from pathlib import Path

import pytest
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn

from docgen.infrastructure.word.merger import MergePart, WordMergeEngine
from docgen.testing.docx_builder import DocSpec, build_docx


@pytest.fixture()
def parts(tmp_path: Path) -> list[Path]:
    cover = build_docx(
        tmp_path / "Cover.docx",
        DocSpec(
            title="Báo cáo nhân sự",
            paragraphs=["Trang bìa – Tiếng Việt có dấu: Nguyễn Văn Ánh"],
            header_text="COVER HEADER",
            footer_page_number=True,
            font_name="Times New Roman",
        ),
    )
    details = build_docx(
        tmp_path / "Details.docx",
        DocSpec(
            title="Chi tiết",
            header_text="DETAILS HEADER",
            header_image=True,
            footer_text="DETAILS FOOTER",
            landscape=True,
            margins_mm=12,
            table_header=["Mã", "Tên", "Phòng ban"],
            table_rows=[["E01", "Lê Thị Hoa", "Kế toán"], ["E02", "Trần Bình", "IT"]],
            numbered_items=["Bước một", "Bước hai"],
            image=True,
            hyperlink=("https://example.com/policy", "Chính sách"),
        ),
    )
    appendix = build_docx(
        tmp_path / "Appendix.docx",
        DocSpec(
            title="Phụ lục",
            header_text="APPENDIX HEADER",
            numbered_items=["Mục A", "Mục B"],
            custom_style="AppendixNote",
        ),
    )
    return [cover, details, appendix]


def _section_header_texts(doc_path: Path) -> list[str]:
    doc = Document(str(doc_path))
    texts = []
    for section in doc.sections:
        texts.append(" ".join(p.text for p in section.header.paragraphs).strip())
    return texts


def test_preserve_mode_keeps_sections_and_headers(tmp_path: Path, parts: list[Path]) -> None:
    out = tmp_path / "out" / "Merged.docx"
    report = WordMergeEngine().merge(
        [
            MergePart(parts[0], "cover"),
            MergePart(parts[1], "details", section_mode="preserve"),
            MergePart(parts[2], "appendix", section_mode="preserve"),
        ],
        out,
    )
    assert report.sections == 3
    doc = Document(str(out))
    orientations = [s.orientation for s in doc.sections]
    assert orientations == [WD_ORIENT.PORTRAIT, WD_ORIENT.LANDSCAPE, WD_ORIENT.PORTRAIT]
    assert doc.sections[1].left_margin == Document(str(parts[1])).sections[0].left_margin
    assert _section_header_texts(out) == ["COVER HEADER", "DETAILS HEADER", "APPENDIX HEADER"]
    assert "DETAILS FOOTER" in " ".join(p.text for p in doc.sections[1].footer.paragraphs)
    # parts without their own footer get an explicit empty one (no inheritance)
    assert doc.sections[2].footer.is_linked_to_previous is False
    assert " ".join(p.text for p in doc.sections[2].footer.paragraphs).strip() == ""

    # header image of section 2 must resolve to an image part
    hdr_part = doc.sections[1].header.part
    image_rels = [r for r in hdr_part.rels.values() if r.reltype == RT.IMAGE]
    assert image_rels and image_rels[0].target_part.blob.startswith(b"\x89PNG")
    blips = hdr_part.element.xpath(".//a:blip")
    assert blips[0].get(qn("r:embed")) in hdr_part.rels

    # content survived: table, numbered list, hyperlink, body image
    assert len(doc.tables) == 1
    assert doc.tables[0].cell(1, 1).text == "Lê Thị Hoa"
    hyperlinks = [r for r in doc.part.rels.values() if r.reltype == RT.HYPERLINK]
    assert any(r.target_ref == "https://example.com/policy" for r in hyperlinks)
    assert len(doc.inline_shapes) == 1
    numbered = [p for p in doc.paragraphs if p.style.name == "List Number"]
    assert [p.text for p in numbered] == ["Bước một", "Bước hai", "Mục A", "Mục B"]
    assert "AppendixNote" in [s.name for s in doc.styles]


def test_inherit_mode_uses_master_header_and_warns(tmp_path: Path, parts: list[Path]) -> None:
    out = tmp_path / "Merged.docx"
    report = WordMergeEngine().merge(
        [MergePart(p, p.stem, section_mode="inherit", break_before="page") for p in parts], out
    )
    doc = Document(str(out))
    assert len(doc.sections) == 1
    assert _section_header_texts(out) == ["COVER HEADER"]
    assert any("orientation" in w for w in report.warnings)
    body_xml = doc.element.body.xml
    assert body_xml.count('w:type="page"') == 2


def test_section_break_mode_creates_sections_with_master_layout(tmp_path: Path, parts: list[Path]) -> None:
    out = tmp_path / "Merged.docx"
    WordMergeEngine().merge(
        [
            MergePart(parts[0], "cover"),
            MergePart(parts[2], "appendix", break_before="section", restart_page_numbering=True),
        ],
        out,
    )
    doc = Document(str(out))
    assert len(doc.sections) == 2
    assert doc.sections[-1]._sectPr.find(qn("w:pgNumType")).get(qn("w:start")) == "1"


def test_merged_output_is_structurally_valid(tmp_path: Path, parts: list[Path]) -> None:
    from docgen.infrastructure.word.verifier import verify_docx

    out = tmp_path / "Merged.docx"
    WordMergeEngine().merge(
        [MergePart(parts[0]), *[MergePart(p, section_mode="preserve") for p in parts[1:]]], out
    )
    result = verify_docx(out)
    assert result.ok, result.problems


@pytest.mark.libreoffice
@pytest.mark.skipif(shutil.which("soffice") is None, reason="LibreOffice not installed")
@pytest.mark.skipif(shutil.which("pdftotext") is None, reason="pdftotext not installed")
def test_libreoffice_renders_each_section_with_its_header(tmp_path: Path, parts: list[Path]) -> None:
    out = tmp_path / "Merged.docx"
    WordMergeEngine().merge(
        [MergePart(parts[0]), *[MergePart(p, section_mode="preserve") for p in parts[1:]]], out
    )
    profile_dir = tmp_path / "lo-profile"
    subprocess.run(
        [
            "soffice",
            f"-env:UserInstallation=file://{profile_dir}",
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(tmp_path),
            str(out),
        ],
        check=True,
        capture_output=True,
        timeout=180,
    )
    pdf = tmp_path / "Merged.pdf"
    assert pdf.exists()
    pages = []
    for page in (1, 2, 3):
        text = subprocess.run(
            ["pdftotext", "-f", str(page), "-l", str(page), "-layout", str(pdf), "-"],
            check=True,
            capture_output=True,
            text=True,
        ).stdout
        pages.append(text)
    assert "COVER HEADER" in pages[0]
    assert "DETAILS HEADER" in pages[1] and "Lê Thị Hoa" in pages[1]
    assert "APPENDIX HEADER" in pages[2]
    # numbering restarts per appended document
    assert re.search(r"1\.\s+Bước một", pages[1]), pages[1]
    assert re.search(r"1\.\s+Mục A", pages[2]), pages[2]
    info = subprocess.run(["pdfinfo", str(pdf)], capture_output=True, text=True, check=True).stdout
    assert re.search(r"^Pages:\s+3$", info, re.MULTILINE), info
    # appendix defines no footer -> must not inherit the details footer in preserve mode
    assert "DETAILS FOOTER" not in pages[2]


def test_preserve_mode_can_inherit_missing_headers(tmp_path: Path, parts: list[Path]) -> None:
    out = tmp_path / "Merged.docx"
    WordMergeEngine().merge(
        [
            MergePart(parts[0]),
            MergePart(parts[1], section_mode="preserve"),
            MergePart(parts[2], section_mode="preserve", inherit_missing_headers=True),
        ],
        out,
    )
    doc = Document(str(out))
    assert doc.sections[2].footer.is_linked_to_previous is True
