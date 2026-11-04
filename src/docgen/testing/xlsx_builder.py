"""Synthetic workbook builders for tests, benchmarks and samples."""

from __future__ import annotations

import re
import shutil
import tempfile
import zipfile
from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import Any

from openpyxl import Workbook


def build_xlsx(
    path: Path,
    sheets: Mapping[str, Sequence[Sequence[Any]]],
    *,
    hidden: Sequence[str] = (),
    merged: Mapping[str, Sequence[str]] | None = None,
) -> Path:
    """Write *sheets* ({name: rows}) to an .xlsx file. Values starting with '=' become formulas."""
    wb = Workbook(write_only=False)
    default = wb.active
    for name, rows in sheets.items():
        ws = wb.create_sheet(title=name)
        for row in rows:
            ws.append(list(row))
        if name in hidden:
            ws.sheet_state = "hidden"
        for rng in (merged or {}).get(name, ()):
            ws.merge_cells(rng)
    if default is not None:
        wb.remove(default)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(path))
    return path


def build_xlsx_streaming(path: Path, sheets: Mapping[str, Any]) -> Path:
    """Write large sheets with openpyxl's write-only mode. Values may be generators of rows."""
    wb = Workbook(write_only=True)
    for name, rows in sheets.items():
        ws = wb.create_sheet(title=name)
        for row in rows:
            ws.append(list(row))
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(path))
    return path


def inject_cached_values(path: Path, sheet_index: int, values: Mapping[str, Any]) -> None:
    """Add cached results (<v>) to formula cells, as Excel would when saving.

    openpyxl never writes cached values, so tests use this to emulate a workbook
    that was saved by Excel. *values* maps cell refs ('C2') to cached results.
    """
    member = f"xl/worksheets/sheet{sheet_index}.xml"
    tmp_dir = Path(tempfile.mkdtemp())
    try:
        tmp = tmp_dir / "out.xlsx"
        with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == member:
                    xml = data.decode("utf-8")
                    for ref, value in values.items():
                        pattern = re.compile(
                            rf'(<c r="{ref}"[^>]*>)(<f>.*?</f>)(<v>.*?</v>)?(</c>)', re.DOTALL
                        )
                        xml, n = pattern.subn(rf"\1\2<v>{value}</v>\4", xml)
                        if n != 1:
                            raise ValueError(f"formula cell {ref} not found in {member}")
                    data = xml.encode("utf-8")
                zout.writestr(item, data)
        shutil.move(str(tmp), str(path))
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def build_xls(path: Path, sheets: Mapping[str, Sequence[Sequence[Any]]]) -> Path:
    """Write an Excel 97-2003 .xls file (requires the dev dependency xlwt)."""
    import datetime as dt

    import xlwt  # type: ignore[import-untyped]

    wb = xlwt.Workbook(encoding="utf-8")
    date_style = xlwt.easyxf(num_format_str="DD/MM/YYYY")
    for name, rows in sheets.items():
        ws = wb.add_sheet(name)
        for r, row in enumerate(rows):
            for c, value in enumerate(row):
                if value is None:
                    continue
                if isinstance(value, (dt.date, dt.datetime)):
                    ws.write(r, c, value, date_style)
                else:
                    ws.write(r, c, value)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(path))
    return path
