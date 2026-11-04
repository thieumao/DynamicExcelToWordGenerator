"""Workbook sources: .xlsx/.xlsm (openpyxl), .xls (xlrd) and .csv (stdlib).

All sources expose raw cell values row by row and never write to the input.
openpyxl is used in read-only streaming mode; it does NOT evaluate formulas,
so formula cells are checked separately (see :meth:`XlsxSource.formula_cells`).
"""

from __future__ import annotations

import csv
import datetime as dt
import io
import re
import zipfile
from collections.abc import Iterator, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Protocol

from docgen.infrastructure.fs.safety import UnsafeFileError, ZipLimits, check_zip


class WorkbookOpenError(OSError):
    """The workbook could not be opened (missing, locked, unsupported, corrupt)."""


@dataclass(frozen=True, slots=True)
class SheetInfo:
    name: str
    index: int
    """1-based position in the workbook."""
    hidden: bool = False
    max_row: int | None = None
    max_column: int | None = None


class WorkbookSource(Protocol):
    path: Path
    format: str
    supports_formulas: bool

    def sheets(self) -> list[SheetInfo]: ...

    def iter_rows(
        self, sheet: str, min_row: int = 1, max_row: int | None = None
    ) -> Iterator[tuple[Any, ...]]: ...

    def formula_cells(
        self, sheet: str, columns: set[int], min_row: int, max_row: int | None
    ) -> dict[tuple[int, int], str]: ...

    def full_calc_on_load(self) -> bool: ...

    def close(self) -> None: ...


# ----------------------------------------------------------------------- xlsx
class XlsxSource:
    format = "xlsx"
    supports_formulas = True

    def __init__(self, path: Path, limits: ZipLimits | None = None) -> None:
        from openpyxl import load_workbook

        self.path = path
        check_zip(path, limits or ZipLimits())
        try:
            self._wb = load_workbook(filename=str(path), read_only=True, data_only=True, keep_links=False)
        except PermissionError as exc:
            raise WorkbookOpenError(
                f"{path.name} is locked by another program (close it in Excel and retry)"
            ) from exc
        except (zipfile.BadZipFile, KeyError, ValueError, OSError) as exc:
            raise WorkbookOpenError(f"Cannot open {path.name}: {exc}") from exc
        self._members: dict[str, str] | None = None

    def sheets(self) -> list[SheetInfo]:
        result = []
        for i, ws in enumerate(self._wb.worksheets, start=1):
            max_row = max_col = None
            try:
                max_row, max_col = ws.max_row, ws.max_column
            except (AttributeError, TypeError):
                pass
            result.append(SheetInfo(ws.title, i, ws.sheet_state != "visible", max_row, max_col))
        return result

    def iter_rows(
        self, sheet: str, min_row: int = 1, max_row: int | None = None
    ) -> Iterator[tuple[Any, ...]]:
        ws = self._wb[sheet]
        yield from ws.iter_rows(min_row=min_row, max_row=max_row, values_only=True)

    def formula_cells(
        self, sheet: str, columns: set[int], min_row: int, max_row: int | None, *, need_text: bool = False
    ) -> dict[tuple[int, int], str]:
        """Return {(excel_row, col_index0): formula} for formula cells in *columns*.

        openpyxl exposes either cached values or formulas, never both, and re-opening the
        workbook per sheet is slow, so formula *presence* is detected with one streaming
        lxml pass over the sheet XML. Shared-formula children carry no text; their full
        formula is only needed for the raw_formula policy (``need_text``), which then uses
        openpyxl's formula translation.
        """
        if need_text:
            return self._formula_text_openpyxl(sheet, columns, min_row, max_row)
        member = self._sheet_members().get(sheet)
        if member is None:
            return {}
        return scan_formula_cells(self.path, member, columns, min_row, max_row)

    def _sheet_members(self) -> dict[str, str]:
        if self._members is None:
            self._members = sheet_members(self.path)
        return self._members

    def _formula_text_openpyxl(
        self, sheet: str, columns: set[int], min_row: int, max_row: int | None
    ) -> dict[tuple[int, int], str]:
        from openpyxl import load_workbook

        found: dict[tuple[int, int], str] = {}
        wb = load_workbook(filename=str(self.path), read_only=True, data_only=False, keep_links=False)
        try:
            ws = wb[sheet]
            for offset, row in enumerate(ws.iter_rows(min_row=min_row, max_row=max_row, values_only=True)):
                for col in columns:
                    if col < len(row):
                        value = row[col]
                        if isinstance(value, str) and value.startswith("="):
                            found[(min_row + offset, col)] = value
                        elif value is not None and type(value).__name__ in (
                            "ArrayFormula",
                            "DataTableFormula",
                        ):
                            found[(min_row + offset, col)] = str(getattr(value, "text", value))
        finally:
            wb.close()
        return found

    def full_calc_on_load(self) -> bool:
        try:
            with zipfile.ZipFile(self.path) as zf:
                xml = zf.read("xl/workbook.xml")[:200_000].decode("utf-8", "replace")
        except (KeyError, zipfile.BadZipFile):
            return False
        return bool(re.search(r"<(?:\w+:)?calcPr[^>]*fullCalcOnLoad=\"(?:1|true)\"", xml))

    def close(self) -> None:
        self._wb.close()


# ------------------------------------------------------------------------ xls
class XlsSource:
    format = "xls"
    supports_formulas = False

    def __init__(self, path: Path, max_bytes: int = 200 * 1024 * 1024) -> None:
        import xlrd

        self.path = path
        if path.stat().st_size > max_bytes:
            raise UnsafeFileError(f"{path.name}: file larger than {max_bytes / 1e6:.0f} MB")
        try:
            self._book = xlrd.open_workbook(str(path), on_demand=True)
        except PermissionError as exc:
            raise WorkbookOpenError(f"{path.name} is locked by another program") from exc
        except xlrd.XLRDError as exc:
            raise WorkbookOpenError(f"Cannot open {path.name}: {exc}") from exc

    def sheets(self) -> list[SheetInfo]:
        result = []
        for i, name in enumerate(self._book.sheet_names(), start=1):
            sh = self._book.sheet_by_index(i - 1)
            result.append(SheetInfo(name, i, sh.visibility != 0, sh.nrows, sh.ncols))
        return result

    def iter_rows(
        self, sheet: str, min_row: int = 1, max_row: int | None = None
    ) -> Iterator[tuple[Any, ...]]:
        import xlrd

        sh = self._book.sheet_by_name(sheet)
        last = sh.nrows if max_row is None else min(max_row, sh.nrows)
        for r in range(min_row - 1, last):
            values: list[Any] = []
            for c in range(sh.ncols):
                cell = sh.cell(r, c)
                ctype, value = cell.ctype, cell.value
                if ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                    values.append(None)
                elif ctype == xlrd.XL_CELL_DATE:
                    try:
                        values.append(xlrd.xldate.xldate_as_datetime(value, self._book.datemode))
                    except (xlrd.xldate.XLDateError, OverflowError):
                        values.append(value)
                elif ctype == xlrd.XL_CELL_NUMBER:
                    values.append(int(value) if float(value).is_integer() else value)
                elif ctype == xlrd.XL_CELL_BOOLEAN:
                    values.append(bool(value))
                elif ctype == xlrd.XL_CELL_ERROR:
                    values.append(xlrd.error_text_from_code.get(value, "#ERROR"))
                else:
                    values.append(value)
            yield tuple(values)

    def formula_cells(
        self, sheet: str, columns: set[int], min_row: int, max_row: int | None
    ) -> dict[tuple[int, int], str]:
        return {}

    def full_calc_on_load(self) -> bool:
        return False

    def close(self) -> None:
        self._book.release_resources()


# ------------------------------------------------------------------------ csv
class CsvSource:
    """One CSV file = one sheet; a directory of CSV files = a workbook."""

    format = "csv"
    supports_formulas = False

    def __init__(
        self,
        path: Path,
        *,
        encodings: Sequence[str] = ("utf-8-sig", "utf-16", "cp1258"),
        delimiter: str | None = None,
        max_bytes: int = 500 * 1024 * 1024,
    ) -> None:
        self.path = path
        self.encodings = list(encodings)
        self.delimiter = delimiter
        files = sorted(path.glob("*.csv")) if path.is_dir() else [path]
        if not files:
            raise WorkbookOpenError(f"No .csv files in {path}")
        self._files = {f.stem: f for f in files}
        for f in files:
            if f.stat().st_size > max_bytes:
                raise UnsafeFileError(f"{f.name}: file larger than {max_bytes / 1e6:.0f} MB")
        self._cache: dict[str, list[tuple[Any, ...]]] = {}
        self.detected_encoding: dict[str, str] = {}

    def sheets(self) -> list[SheetInfo]:
        return [SheetInfo(name, i) for i, name in enumerate(self._files, start=1)]

    def _decode(self, name: str) -> str:
        raw = self._files[name].read_bytes()
        candidates = list(self.encodings)
        if raw.startswith(b"\xef\xbb\xbf") and "utf-8-sig" not in candidates:
            candidates.insert(0, "utf-8-sig")
        if raw[:2] in (b"\xff\xfe", b"\xfe\xff") and "utf-16" not in candidates:
            candidates.insert(0, "utf-16")
        for enc in candidates:
            if enc.startswith("utf-16") and raw[:2] not in (b"\xff\xfe", b"\xfe\xff"):
                continue
            try:
                text = raw.decode(enc)
            except (UnicodeDecodeError, LookupError):
                continue
            self.detected_encoding[name] = enc
            return text
        self.detected_encoding[name] = "cp1252(replace)"
        return raw.decode("cp1252", errors="replace")

    def _rows(self, name: str) -> list[tuple[Any, ...]]:
        if name not in self._cache:
            text = self._decode(name)
            delimiter = self.delimiter
            if delimiter is None:
                try:
                    delimiter = csv.Sniffer().sniff(text[:20_000], delimiters=",;\t|").delimiter
                except csv.Error:
                    delimiter = ","
            reader = csv.reader(io.StringIO(text), delimiter=delimiter)
            self._cache[name] = [tuple(cell if cell != "" else None for cell in row) for row in reader]
        return self._cache[name]

    def iter_rows(
        self, sheet: str, min_row: int = 1, max_row: int | None = None
    ) -> Iterator[tuple[Any, ...]]:
        rows = self._rows(sheet)
        yield from rows[min_row - 1 : max_row]

    def formula_cells(
        self, sheet: str, columns: set[int], min_row: int, max_row: int | None
    ) -> dict[tuple[int, int], str]:
        return {}

    def full_calc_on_load(self) -> bool:
        return False

    def close(self) -> None:
        self._cache.clear()


_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def sheet_members(path: Path) -> dict[str, str]:
    """Map sheet names to their ZIP member ('xl/worksheets/sheet1.xml')."""
    import posixpath

    from lxml import etree

    parser = etree.XMLParser(resolve_entities=False, no_network=True)
    with zipfile.ZipFile(path) as zf:
        workbook = etree.fromstring(zf.read("xl/workbook.xml"), parser)
        rels = etree.fromstring(zf.read("xl/_rels/workbook.xml.rels"), parser)
    targets = {r.get("Id"): r.get("Target", "") for r in rels if _local(r.tag) == "Relationship"}
    result: dict[str, str] = {}
    for el in workbook.iter():
        if _local(el.tag) != "sheet":
            continue
        rid = next((v for k, v in el.attrib.items() if _local(str(k)) == "id"), None)
        target = targets.get(rid, "")
        if not target:
            continue
        member = (
            target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
        )
        result[el.get("name", "")] = member
    return result


def _split_ref(ref: str) -> tuple[int, int]:
    """'AB12' -> (12, 27) as (row, 0-based column)."""
    i = 0
    while i < len(ref) and ref[i].isalpha():
        i += 1
    return int(ref[i:]), column_index(ref[:i])


_F_TAG = re.compile(rb"<(?:\w+:)?f[ >/]")


def _has_formula_tags(zf: zipfile.ZipFile, member: str) -> bool:
    """Fast byte-level pre-check (C speed) for any <f> element in a worksheet part."""
    tail = b""
    with zf.open(member) as fh:
        while chunk := fh.read(1 << 20):
            if _F_TAG.search(tail + chunk):
                return True
            tail = chunk[-16:]
    return False


def scan_formula_cells(
    path: Path, member: str, columns: set[int], min_row: int, max_row: int | None
) -> dict[tuple[int, int], str]:
    """Stream one worksheet part and return formula cells in *columns* (single pass, O(1) memory)."""
    from lxml import etree

    found: dict[tuple[int, int], str] = {}
    with zipfile.ZipFile(path) as zf:
        if not _has_formula_tags(zf, member):
            return found  # common case: plain data sheets skip the XML parse entirely
    with zipfile.ZipFile(path) as zf, zf.open(member) as fh:
        for _event, el in etree.iterparse(
            fh, events=("end",), resolve_entities=False, no_network=True, huge_tree=False
        ):
            name = _local(el.tag)
            if name == "row":
                row_attr = el.get("r")
                el.clear()
                while el.getprevious() is not None:
                    del el.getparent()[0]
                if row_attr and max_row is not None and int(row_attr) >= max_row:
                    break
                continue
            if name != "c":
                continue
            ref = el.get("r")
            if not ref:
                continue
            row, col = _split_ref(ref)
            if row < min_row or (max_row is not None and row > max_row) or col not in columns:
                continue
            formula = next((child for child in el if _local(child.tag) == "f"), None)
            if formula is not None:
                text = formula.text or ""
                found[(row, col)] = f"={text}" if text else "=(shared formula)"
    return found


def detect_format(path: Path) -> str:
    if path.is_dir():
        return "csv"
    suffix = path.suffix.lower()
    try:
        with path.open("rb") as fh:
            magic = fh.read(8)
    except PermissionError as exc:
        raise WorkbookOpenError(f"{path.name} is locked by another program") from exc
    except OSError as exc:
        raise WorkbookOpenError(f"Cannot read {path}: {exc}") from exc
    if magic.startswith(b"PK"):
        return "xlsx"
    if magic.startswith(b"\xd0\xcf\x11\xe0"):
        return "xls"
    if suffix in (".csv", ".txt", ".tsv"):
        return "csv"
    raise WorkbookOpenError(f"{path.name}: unsupported file type (expected .xlsx, .xlsm, .xls or .csv)")


def open_workbook(
    path: Path,
    *,
    csv_encodings: Sequence[str] = ("utf-8-sig", "utf-16", "cp1258"),
    csv_delimiter: str | None = None,
) -> WorkbookSource:
    if not path.exists():
        raise WorkbookOpenError(f"File not found: {path}")
    fmt = detect_format(path)
    if fmt == "xlsx":
        return XlsxSource(path)
    if fmt == "xls":
        return XlsSource(path)
    return CsvSource(path, encodings=csv_encodings, delimiter=csv_delimiter)


def column_letter(index0: int) -> str:
    """0 -> 'A', 25 -> 'Z', 26 -> 'AA'."""
    n = index0 + 1
    letters = ""
    while n:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def column_index(letter: str) -> int:
    """'A' -> 0, 'AA' -> 26."""
    n = 0
    for ch in letter.upper():
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def excel_time_to_text(value: Any) -> Any:
    if isinstance(value, dt.time):
        return value.isoformat(timespec="minutes")
    return value
