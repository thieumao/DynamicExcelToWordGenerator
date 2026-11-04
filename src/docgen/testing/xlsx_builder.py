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


def build_xlsx_streaming(path: Path, sheets: Mapping[str, Any], *, add_dimension: bool = True) -> Path:
    """Write large sheets with openpyxl's write-only mode. Values may be generators of rows.

    openpyxl's write-only mode omits ``<dimension>``, which Excel always writes; without it
    openpyxl's read-only reader scans every sheet when opening the file. ``add_dimension``
    restores it so benchmarks reflect files saved by Excel.
    """
    wb = Workbook(write_only=True)
    sizes: dict[str, tuple[int, int]] = {}
    for name, rows in sheets.items():
        ws = wb.create_sheet(title=name)
        count = width = 0
        for row in rows:
            values = list(row)
            ws.append(values)
            count += 1
            width = max(width, len(values))
        sizes[name] = (count, width)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(path))
    if add_dimension:
        _add_dimensions(path, [sizes[n] for n in sheets])
    return path


def _add_dimensions(path: Path, sizes: Sequence[tuple[int, int]]) -> None:
    from docgen.infrastructure.excel.sources import column_letter

    tmp_dir = Path(tempfile.mkdtemp())
    try:
        tmp = tmp_dir / "out.xlsx"
        with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                m = re.fullmatch(r"xl/worksheets/sheet(\d+)\.xml", item.filename)
                if m and int(m.group(1)) <= len(sizes):
                    rows, cols = sizes[int(m.group(1)) - 1]
                    ref = f"A1:{column_letter(max(cols, 1) - 1)}{max(rows, 1)}"
                    xml = data.decode("utf-8")
                    if "<dimension" not in xml:
                        end_pr = xml.find("</sheetPr>")
                        if end_pr >= 0:
                            pos = end_pr + len("</sheetPr>")
                        else:
                            pos = xml.index(">", xml.index("<worksheet")) + 1
                        xml = xml[:pos] + f'<dimension ref="{ref}"/>' + xml[pos:]
                        data = xml.encode("utf-8")
                zout.writestr(item, data)
        shutil.move(str(tmp), str(path))
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


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
