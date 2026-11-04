"""Output filename patterns and Windows-safe sanitisation.

Pattern syntax: ``Report_{employee.code}_{report.generated_at:%Y%m%d}.docx``.
``{path}`` takes a value from the unit context, ``{path:fmt}`` applies a
strftime format for dates or a Python format spec for numbers. ``{{``/``}}``
produce literal braces.
"""

from __future__ import annotations

import datetime as dt
import hashlib
import re
import unicodedata
from collections.abc import Mapping
from decimal import Decimal
from typing import Any

from docgen.domain.data.values import to_text
from docgen.domain.mapping.context import get_path

WINDOWS_RESERVED = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}
_INVALID = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_TOKEN = re.compile(r"\{\{|\}\}|\{([^{}:]+)(?::([^{}]*))?\}")


class PatternError(ValueError):
    pass


def pattern_fields(pattern: str) -> list[str]:
    return [m.group(1).strip() for m in _TOKEN.finditer(pattern) if m.group(1)]


def render_pattern(
    pattern: str, values: Mapping[str, Any], *, escape_separators: bool = True
) -> tuple[str, list[str]]:
    """Return (rendered text, list of missing paths).

    With *escape_separators* a value such as 'IT/Kế toán' cannot create sub-folders:
    path separators inside values become '-'. Separators written in the pattern stay.
    """
    missing: list[str] = []

    def repl(m: re.Match[str]) -> str:
        if m.group(0) == "{{":
            return "{"
        if m.group(0) == "}}":
            return "}"
        path, fmt = m.group(1).strip(), m.group(2)
        value = get_path(values, path)
        if value is None or value == "":
            missing.append(path)
            return ""
        if fmt:
            try:
                if isinstance(value, (dt.date, dt.datetime)):
                    return _esc(value.strftime(fmt), escape_separators)
                if isinstance(value, (int, float, Decimal)):
                    return _esc(format(value, fmt), escape_separators)
            except (ValueError, TypeError) as exc:
                raise PatternError(f"bad format '{fmt}' for {path}: {exc}") from exc
        if isinstance(value, dt.datetime):
            return _esc(value.strftime("%Y%m%d-%H%M%S"), escape_separators)
        if isinstance(value, dt.date):
            return value.strftime("%Y%m%d")
        return _esc(to_text(value), escape_separators)

    return _TOKEN.sub(repl, pattern), missing


def _esc(text: str, escape: bool) -> str:
    return re.sub(r"[\\/]", "-", text) if escape else text


def sanitize_component(name: str, *, max_length: int = 150, replacement: str = "_") -> str:
    """Make one path component safe on Windows, macOS and Linux."""
    name = unicodedata.normalize("NFC", name)
    name = _INVALID.sub(replacement, name)
    name = re.sub(r"\s+", " ", name).strip()
    name = name.rstrip(". ")
    if not name:
        name = "untitled"
    stem, dot, ext = name.rpartition(".")
    if not dot:
        stem, ext = name, ""
    if stem.upper() in WINDOWS_RESERVED or stem.split(".")[0].upper() in WINDOWS_RESERVED:
        stem = f"_{stem}"
    full = f"{stem}.{ext}" if ext else stem
    if len(full) > max_length:
        digest = hashlib.sha1(full.encode("utf-8")).hexdigest()[:8]
        keep = max_length - len(digest) - 1 - (len(ext) + 1 if ext else 0)
        full = f"{stem[: max(1, keep)]}~{digest}" + (f".{ext}" if ext else "")
    return full


def sanitize_relative_path(text: str, *, max_component: int = 150) -> str:
    """Sanitise a relative path ('a/b/c.docx'); '..' and absolute paths are neutralised."""
    parts = [p for p in re.split(r"[\\/]+", text) if p not in ("", ".", "..")]
    return "/".join(sanitize_component(p, max_length=max_component) for p in parts)
