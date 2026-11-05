"""The user guide is shipped with the application and identical to docs/HUONG_DAN_SU_DUNG.md."""

from __future__ import annotations

from pathlib import Path

from docgen.help import GUIDE, guide_text

ROOT = Path(__file__).resolve().parents[2]


def test_packaged_guide_is_the_docs_guide():
    docs = (ROOT / "docs" / GUIDE).read_text(encoding="utf-8")
    packaged = (ROOT / "src" / "docgen" / "help" / GUIDE).read_text(encoding="utf-8")
    assert packaged == docs, "copy docs/HUONG_DAN_SU_DUNG.md to src/docgen/help/ after editing it"
    assert guide_text() == docs
    assert "## A1. Cài đặt" in docs and "## A5. Tạo bộ hồ sơ mới (4 bước)" in docs
