"""Validate generated .docx files with Microsoft's Open XML SDK (OpenXmlValidator).

This is the strictest automated Word-compatibility check available without Microsoft Word:
schema + semantic constraints for the chosen Office version. Requires the .NET SDK (preinstalled
on GitHub runners). Skipped when `dotnet` is missing, unless DOCGEN_REQUIRE_OOXML=1 (set in CI).
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
from pathlib import Path

import pytest

from docgen.application.generation.job import GenerationRequest
from docgen.application.generation.service import DocumentGenerationService
from docgen.domain.profile.io import load_profile
from docgen.domain.profile.models import OutputKind
from docgen.infrastructure.word.merger import MergePart, WordMergeEngine
from docgen.infrastructure.word.renderer import TemplateRenderer
from docgen.testing.samples import write_customer_a, write_customer_b, write_customer_c

from .test_formatting_preservation import CTX, build_rich_template

ROOT = Path(__file__).resolve().parents[2]
PROJECT = ROOT / "tools" / "ooxml-validator"
REQUIRED = os.environ.get("DOCGEN_REQUIRE_OOXML") == "1"

pytestmark = pytest.mark.ooxml


@pytest.fixture(scope="module")
def validator(tmp_path_factory: pytest.TempPathFactory) -> Path:
    if shutil.which("dotnet") is None:
        if REQUIRED:
            pytest.fail("dotnet is required (DOCGEN_REQUIRE_OOXML=1) but not on PATH")
        pytest.skip("dotnet SDK not installed")
    out = tmp_path_factory.mktemp("ooxmlval")
    env = dict(os.environ, DOTNET_CLI_TELEMETRY_OPTOUT="1", DOTNET_NOLOGO="1")
    build = subprocess.run(
        ["dotnet", "build", str(PROJECT), "-c", "Release", "-o", str(out)],
        capture_output=True,
        text=True,
        timeout=600,
        env=env,
    )
    if build.returncode != 0:
        if REQUIRED:
            pytest.fail(f"validator build failed:\n{build.stdout[-3000:]}\n{build.stderr[-3000:]}")
        pytest.skip("validator could not be built (offline NuGet?)")
    dll = out / "OoxmlValidator.dll"
    assert dll.exists()
    return dll


def run_validator(dll: Path, target: Path, fmt: str = "Office2019") -> tuple[int, list[dict], str]:
    proc = subprocess.run(
        ["dotnet", str(dll), "--format", fmt, str(target)],
        capture_output=True,
        text=True,
        encoding="utf-8",
        timeout=600,
    )
    errors = [json.loads(line) for line in proc.stdout.splitlines() if line.startswith("{")]
    return proc.returncode, errors, proc.stderr


def test_generated_documents_conform_to_office_schemas(validator: Path, tmp_path: Path) -> None:
    corpus = tmp_path / "corpus"
    service = DocumentGenerationService()
    for writer, name in ((write_customer_a, "a"), (write_customer_b, "b"), (write_customer_c, "c")):
        profile_path, workbook = writer(tmp_path / f"src_{name}")
        result = service.run(
            GenerationRequest(
                load_profile(profile_path),
                profile_path.parent,
                workbook,
                corpus / name,
                output=OutputKind.BOTH,
            )
        )
        assert result.status == "completed", [str(i) for i in result.all_issues()]
    rich = build_rich_template(tmp_path / "rich.docx")
    renderer = TemplateRenderer()
    a, b = corpus / "rich" / "a.docx", corpus / "rich" / "b.docx"
    renderer.render(rich.read_bytes(), CTX, a)
    renderer.render(rich.read_bytes(), CTX, b)
    engine = WordMergeEngine()
    engine.merge([MergePart(a), MergePart(b, section_mode="preserve")], corpus / "rich" / "preserve.docx")
    engine.merge(
        [MergePart(a), MergePart(b, break_before="section", restart_page_numbering=True)],
        corpus / "rich" / "section.docx",
    )
    engine.merge([MergePart(a), MergePart(b, section_mode="inherit")], corpus / "rich" / "inherit.docx")
    files = list(corpus.rglob("*.docx"))
    assert len(files) > 30
    for fmt in ("Office2010", "Office2019", "Microsoft365"):
        code, errors, summary = run_validator(validator, corpus, fmt)
        assert code == 0 and not errors, f"{fmt}: {summary}\n" + "\n".join(
            f"{e['file']} {e.get('part')} {e.get('path')}: {e['description']}" for e in errors[:20]
        )
