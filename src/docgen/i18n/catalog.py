"""Tiny message catalog (Vietnamese / English) shared by the UI and core messages.

UI strings use the English text as the key: ``tr("Generate")``. Diagnostics
are translated by issue code (``issue:<code>`` keys) with ``{param}``
placeholders; unknown codes fall back to the English message from the core.
"""

from __future__ import annotations

import json
import string
from collections.abc import Mapping
from functools import cache
from importlib import resources
from typing import Any

from docgen.domain.diagnostics import Issue

SUPPORTED = {"vi": "Tiếng Việt", "en": "English"}
_language = "vi"


@cache
def _catalog(lang: str) -> dict[str, str]:
    if lang == "en":
        return {}
    try:
        text = resources.files("docgen.i18n").joinpath(f"{lang}.json").read_text(encoding="utf-8")
    except (FileNotFoundError, OSError):
        return {}
    data: dict[str, str] = json.loads(text)
    return data


def set_language(lang: str) -> None:
    global _language
    _language = lang if lang in SUPPORTED else "en"


def language() -> str:
    return _language


def tr(text: str, **params: Any) -> str:
    translated = _catalog(_language).get(text, text)
    if params:
        try:
            return translated.format(**params)
        except (KeyError, IndexError, ValueError):
            return text.format(**params)
    return translated


class _Params(dict[str, Any]):
    """Message parameters; a value that is not known for this issue shows as '—'."""

    def __missing__(self, key: str) -> str:
        return "—"


def _fill(template: str, params: Mapping[str, Any]) -> str | None:
    try:
        return string.Formatter().vformat(template, (), _Params(params))
    except (IndexError, ValueError, AttributeError):
        return None


def issue_text(issue: Issue) -> tuple[str, str]:
    """Localised (message, suggestion) for an issue."""
    catalog = _catalog(_language)
    message, suggestion = issue.message, issue.suggestion or ""
    template = catalog.get(f"issue:{issue.code}")
    if issue.params.get("user_message"):
        template = None  # a message the customer wrote in the profile is shown as written
    if template:
        message = _fill(template, issue.params) or issue.message
    hint = catalog.get(f"hint:{issue.code}")
    if hint:
        suggestion = _fill(hint, issue.params) or suggestion
    return message, suggestion
