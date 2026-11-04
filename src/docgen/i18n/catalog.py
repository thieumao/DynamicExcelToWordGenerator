"""Tiny message catalog (Vietnamese / English) shared by the UI and core messages.

UI strings use the English text as the key: ``tr("Generate")``. Diagnostics
are translated by issue code (``issue:<code>`` keys) with ``{param}``
placeholders; unknown codes fall back to the English message from the core.
"""

from __future__ import annotations

import json
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


def issue_text(issue: Issue) -> tuple[str, str]:
    """Localised (message, suggestion) for an issue."""
    catalog = _catalog(_language)
    message, suggestion = issue.message, issue.suggestion or ""
    template = catalog.get(f"issue:{issue.code}")
    if template:
        try:
            message = template.format(**{k: v for k, v in issue.params.items()})
        except (KeyError, IndexError, ValueError):
            message = issue.message
    hint = catalog.get(f"hint:{issue.code}")
    if hint:
        try:
            suggestion = hint.format(**dict(issue.params))
        except (KeyError, IndexError, ValueError):
            pass
    return message, suggestion
