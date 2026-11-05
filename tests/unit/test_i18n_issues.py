"""Every diagnostic the application can emit is shown in Vietnamese, with its details.

The app is used mostly by Vietnamese speakers: an error that appears in English (or with an
empty placeholder) is a defect. The checks are static, so new issue codes cannot slip through:

* every issue code emitted in ``src/`` has an ``issue:<code>`` entry in ``vi.json``;
* every ``{param}`` used by that entry (and by its ``hint:``) is passed at each call site;
* every catalog template formats without errors (e.g. literal braces must be doubled).
"""

from __future__ import annotations

import ast
import json
import string
from pathlib import Path

import pytest

from docgen.domain.diagnostics import Category, Severity, make_issue
from docgen.i18n.catalog import issue_text, set_language

SRC = Path(__file__).resolve().parents[2] / "src" / "docgen"
CATALOG = json.loads((SRC / "i18n" / "vi.json").read_text(encoding="utf-8"))
ISSUE_FUNCS = {"make_issue", "_err", "_warn"}
LOCATION_KEYS = {"sheet", "row", "column", "rule", "field", "template", "placeholder", "unit"}
# codes built at runtime: (code, parameters always passed)
VALIDATION_PARAMS = {"field", "rule", "row", "value", "min", "max", "user_message"}
DYNAMIC = {
    "validation.required": VALIDATION_PARAMS,
    "validation.regex": VALIDATION_PARAMS | {"pattern"},
    "validation.range": VALIDATION_PARAMS,
    "validation.length": VALIDATION_PARAMS | {"length"},
    "validation.one_of": VALIDATION_PARAMS | {"values"},
    "validation.unique": VALIDATION_PARAMS | {"first_row"},
    "validation.expr": VALIDATION_PARAMS | {"expr", "detail"},
}
# parameters that arrive through **kwargs at specific call sites
STAR_PARAMS = {"field", "sheet", "row", "column", "rule", "template"}


def call_sites() -> list[tuple[str, set[str], bool, str]]:
    sites = []
    for path in sorted(SRC.rglob("*.py")):
        tree = ast.parse(path.read_text(encoding="utf-8"))
        for node in ast.walk(tree):
            if not isinstance(node, ast.Call):
                continue
            name = getattr(node.func, "id", getattr(node.func, "attr", ""))
            if name not in ISSUE_FUNCS or not node.args:
                continue
            first = node.args[0]
            if not (isinstance(first, ast.Constant) and isinstance(first.value, str)):
                continue
            kwargs = {k.arg for k in node.keywords if k.arg}
            star = any(k.arg is None for k in node.keywords)
            sites.append((first.value, kwargs, star, f"{path.relative_to(SRC)}:{node.lineno}"))
    # _fatal(result, code, message) always passes detail
    for path in sorted(SRC.rglob("*.py")):
        for node in ast.walk(ast.parse(path.read_text(encoding="utf-8"))):
            if isinstance(node, ast.Call) and getattr(node.func, "attr", "") == "_fatal":
                code = node.args[1]
                assert isinstance(code, ast.Constant)
                sites.append((code.value, {"detail"}, False, f"{path.relative_to(SRC)}:{node.lineno}"))
    return sites


def fields_of(template: str) -> set[str]:
    return {name for _, name, _, _ in string.Formatter().parse(template) if name}


SITES = call_sites()


def test_every_emitted_code_is_translated():
    codes = {code for code, *_ in SITES} | set(DYNAMIC)
    missing = sorted(c for c in codes if f"issue:{c}" not in CATALOG)
    assert not missing, f"no Vietnamese text for: {missing}"
    assert len(codes) > 80  # the scan really sees the code base


@pytest.mark.parametrize(("code", "kwargs", "star", "where"), SITES, ids=[s[3] for s in SITES])
def test_translation_parameters_are_passed(code, kwargs, star, where):
    available = kwargs | (STAR_PARAMS if star else set())
    for key in (f"issue:{code}", f"hint:{code}"):
        if key in CATALOG:
            unknown = fields_of(CATALOG[key]) - available
            assert not unknown, f"{key} uses {unknown}, not passed at {where}"


@pytest.mark.parametrize("code", sorted(DYNAMIC))
def test_validation_rule_parameters_are_passed(code):
    assert fields_of(CATALOG[f"issue:{code}"]) <= DYNAMIC[code]


def test_all_catalog_templates_format():
    for key, text in CATALOG.items():
        if key.startswith(("issue:", "hint:")):
            text.format(**dict.fromkeys(fields_of(text), "x"))  # raises on stray braces


def test_messages_are_vietnamese_and_tolerate_missing_values():
    set_language("vi")
    issue = make_issue(
        "output.renamed",
        Severity.INFO,
        Category.OUTPUT,
        "a.docx existed",
        name="a.docx",
        new_name="a (2).docx",
    )
    assert issue_text(issue)[0] == "File a.docx đã có sẵn; tài liệu mới được lưu thành a (2).docx"
    partial = make_issue(
        "cell.error_value", Severity.WARNING, Category.FORMULA, "x", value="#REF!", sheet="S"
    )
    message = issue_text(partial)[0]
    assert message.startswith("Ô có lỗi Excel #REF!") and "—" in message  # unknown row shows as —
    custom = make_issue(
        "validation.required",
        Severity.ERROR,
        Category.VALIDATION,
        "Thiếu mã số thuế",
        user_message="Thiếu mã số thuế",
    )
    assert issue_text(custom)[0] == "Thiếu mã số thuế"  # a message written in the profile wins
    hint = issue_text(
        make_issue("template.syntax", Severity.ERROR, Category.TEMPLATE, "x", template="t", detail="d")
    )[1]
    assert "{% if %}" in hint
