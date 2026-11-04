"""In-memory representation of extracted sheet data.

Only the columns declared in the profile are kept (projection), as tuples, with
the original Excel row numbers so diagnostics and traces can point users to
the exact cell.
"""

from __future__ import annotations

from collections.abc import Iterator, Mapping
from dataclasses import dataclass, field
from typing import Any


@dataclass(frozen=True, slots=True)
class CellRef:
    sheet: str
    row: int
    column: str
    """Header text as found in the workbook (or the column id when headerless)."""
    letter: str = ""

    def __str__(self) -> str:
        cell = f"{self.letter}{self.row}" if self.letter else f"row {self.row}"
        return f"{self.sheet}!{cell} [{self.column}]"


@dataclass(slots=True)
class ColumnBinding:
    """Where a profile column was found in the sheet."""

    column_id: str
    header: str
    index: int | None
    """0-based column index in the sheet, or None if the (optional) column is absent."""
    letter: str = ""


@dataclass(slots=True)
class SheetTable:
    sheet_id: str
    sheet_name: str
    columns: list[ColumnBinding]
    rows: list[tuple[Any, ...]] = field(default_factory=list)
    row_numbers: list[int] = field(default_factory=list)
    header_row: int | None = None
    invalid: dict[tuple[int, int], tuple[str, str]] = field(default_factory=dict)
    """(row index, column index) -> (reason code, severity) for cells whose value must not be
    trusted: formula without a cached result, or a value that failed type conversion."""
    _index: dict[str, int] = field(init=False, repr=False, default_factory=dict)

    def __post_init__(self) -> None:
        self._index = {c.column_id: i for i, c in enumerate(self.columns)}

    def __len__(self) -> int:
        return len(self.rows)

    def column_position(self, column_id: str) -> int:
        return self._index[column_id]

    def has_column(self, column_id: str) -> bool:
        return column_id in self._index

    def value(self, row_index: int, column_id: str) -> Any:
        return self.rows[row_index][self._index[column_id]]

    def cell_ref(self, row_index: int, column_id: str) -> CellRef:
        binding = self.columns[self._index[column_id]]
        return CellRef(self.sheet_name, self.row_numbers[row_index], binding.header, binding.letter)

    def row_view(self, row_index: int) -> RowView:
        return RowView(self, row_index)

    def find_row_index(self, excel_row: int) -> int | None:
        """Map an Excel row number back to the table row index (binary search)."""
        import bisect

        pos = bisect.bisect_left(self.row_numbers, excel_row)
        if pos < len(self.row_numbers) and self.row_numbers[pos] == excel_row:
            return pos
        return None


class RowView(Mapping[str, Any]):
    """Read-only mapping view of one row: ``view['column_id']``."""

    __slots__ = ("index", "table")

    def __init__(self, table: SheetTable, index: int) -> None:
        self.table = table
        self.index = index

    def __getitem__(self, key: str) -> Any:
        try:
            pos = self.table._index[key]
        except KeyError:
            raise KeyError(key) from None
        return self.table.rows[self.index][pos]

    def __iter__(self) -> Iterator[str]:
        return iter(self.table._index)

    def __len__(self) -> int:
        return len(self.table._index)

    @property
    def excel_row(self) -> int:
        return self.table.row_numbers[self.index]

    def __repr__(self) -> str:
        return f"RowView({self.table.sheet_name}!{self.excel_row})"


@dataclass(slots=True)
class DataSet:
    """All extracted sheets for one job, keyed by profile sheet id."""

    tables: dict[str, SheetTable] = field(default_factory=dict)
    source_path: str = ""

    def table(self, sheet_id: str) -> SheetTable:
        return self.tables[sheet_id]
