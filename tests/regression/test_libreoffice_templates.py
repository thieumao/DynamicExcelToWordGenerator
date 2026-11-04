"""Templates authored/re-saved by another editor (LibreOffice) must work like python-docx ones.

LibreOffice writes different OOXML (run splitting, table/style markup) and flattens unbound
MERGEFIELDs into plain «Name» text. Customer B's templates are round-tripped docx → odt → docx and
the full pipeline must still fill every placeholder.
"""

from __future__ import annotations

import shutil
import subprocess
from pathlib import Path

import pytest
from docx import Document

from docgen.application.generation.job import GenerationRequest
from docgen.application.generation.service import DocumentGenerationService
from docgen.domain.profile.io import load_profile
from docgen.testing.samples import write_customer_a, write_customer_b

pytestmark = [
    pytest.mark.libreoffice,
    pytest.mark.skipif(shutil.which("soffice") is None, reason="LibreOffice not installed"),
]


def roundtrip(folder: Path, work: Path) -> None:
    profile_dir = work / "lo-profile"
    for target, fmt in ((work / "odt", "odt"), (work / "docx", "docx:MS Word 2007 XML")):
        source = sorted(folder.glob("*.docx")) if fmt == "odt" else sorted((work / "odt").glob("*.odt"))
        subprocess.run(
            [
                "soffice",
                f"-env:UserInstallation=file://{profile_dir}",
                "--headless",
                "--convert-to",
                fmt,
                "--outdir",
                str(target),
                *map(str, source),
            ],
            check=True,
            capture_output=True,
            timeout=300,
        )
    for docx in (work / "docx").glob("*.docx"):
        shutil.copyfile(docx, folder / docx.name)


@pytest.mark.parametrize("writer", [write_customer_a, write_customer_b], ids=["customer_a", "customer_b"])
def test_libreoffice_saved_templates(writer, tmp_path):
    profile_path, workbook = writer(tmp_path / "src")
    roundtrip(profile_path.parent / "templates", tmp_path / "lo")
    result = DocumentGenerationService().run(
        GenerationRequest(load_profile(profile_path), profile_path.parent, workbook, tmp_path / "out")
    )
    assert result.status == "completed", [str(i) for i in result.all_issues()]
    text = "\n".join(p.text for doc in result.documents for p in Document(doc).paragraphs)
    assert "«" not in text and "{{" not in text and "<<" not in text
    if writer is write_customer_b:
        assert "Hợp đồng: HD-001" in text and "Ngày ký: 02/07/2024" in text
