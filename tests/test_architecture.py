"""Layering rules: business logic must not depend on the UI toolkit."""

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1] / "src" / "docgen"


def test_core_layers_do_not_import_qt():
    for layer in ("domain", "application", "infrastructure", "cli"):
        for path in (ROOT / layer).rglob("*.py"):
            text = path.read_text(encoding="utf-8")
            assert "PySide6" not in text and "from docgen.ui" not in text, path


def test_domain_does_not_import_infrastructure():
    for path in (ROOT / "domain").rglob("*.py"):
        text = path.read_text(encoding="utf-8")
        assert "docgen.infrastructure" not in text and "docgen.application" not in text, path


def test_no_eval_or_exec_in_sources():
    import re

    # Python's eval()/exec() builtins; Qt's QApplication.exec()/QDialog.exec() are method calls
    pattern = re.compile(r"(?<![.\w])(eval|exec)\(")
    for path in ROOT.rglob("*.py"):
        assert not pattern.search(path.read_text(encoding="utf-8")), path
