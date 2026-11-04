"""Workbook inspection for the Excel configuration page and the CLI."""

from __future__ import annotations

from dataclasses import dataclass, field
from itertools import islice
from pathlib import Path
from typing import Any

from docgen.domain.data.values import is_empty, normalize_label
from docgen.domain.profile.models import ColumnSpec, ColumnType, HeaderConfig, SheetConfig, SheetMatch
from docgen.infrastructure.excel.header_detect import detect_header_row
from docgen.infrastructure.excel.sources import column_letter, open_workbook


@dataclass(slots=True)
class SheetSummary:
    name: str
    index: int
    hidden: bool
    max_row: int | None
    header_row: int | None
    header_confidence: float
    headers: list[str] = field(default_factory=list)
    preview: list[list[Any]] = field(default_factory=list)
    """First rows of the sheet (raw values) for display."""


@dataclass(slots=True)
class WorkbookSummary:
    path: Path
    format: str
    sheets: list[SheetSummary]


def inspect_workbook(path: Path, *, preview_rows: int = 50, scan_rows: int = 30) -> WorkbookSummary:
    source = open_workbook(path)
    try:
        sheets = []
        for info in source.sheets():
            rows = [list(r) for r in islice(source.iter_rows(info.name, 1, preview_rows), preview_rows)]
            detection = detect_header_row(rows[:scan_rows])
            headers: list[str] = []
            if detection.row is not None:
                raw = rows[detection.row - 1]
                headers = ["" if is_empty(v) else str(v).strip() for v in raw]
                while headers and not headers[-1]:
                    headers.pop()
            sheets.append(
                SheetSummary(
                    info.name,
                    info.index,
                    info.hidden,
                    info.max_row,
                    detection.row,
                    detection.confidence,
                    headers,
                    rows,
                )
            )
        return WorkbookSummary(path, source.format, sheets)
    finally:
        source.close()


def suggest_identifier(text: str, taken: set[str]) -> str:
    """'Mã nhân viên' -> 'ma_nhan_vien' (unique within *taken*)."""
    import re

    base = normalize_label(text, accents=True)
    base = re.sub(r"[^a-z0-9]+", "_", base).strip("_") or "col"
    if base[0].isdigit():
        base = f"c_{base}"
    base = base[:60]
    candidate, n = base, 2
    while candidate in taken:
        candidate = f"{base}_{n}"
        n += 1
    taken.add(candidate)
    return candidate


def sheet_config_from_summary(summary: SheetSummary, sheet_id: str | None = None) -> SheetConfig:
    """Draft a SheetConfig (all detected columns, optional) that the user can refine."""
    taken: set[str] = set()
    columns = []
    for i, header in enumerate(summary.headers):
        if not header:
            continue
        columns.append(
            ColumnSpec(
                id=suggest_identifier(header, taken),
                header=header,
                type=ColumnType.AUTO,
                required=False,
                description=f"column {column_letter(i)}",
            )
        )
    if not columns:
        columns.append(ColumnSpec(id="col_a", position="A", required=False))
    sid = sheet_id or suggest_identifier(summary.name, set())
    header_cfg = (
        HeaderConfig(mode="fixed", row=summary.header_row)
        if summary.header_row and summary.header_confidence < 0.5
        else HeaderConfig()
    )
    return SheetConfig(id=sid, match=SheetMatch(name=summary.name), header=header_cfg, columns=columns)
