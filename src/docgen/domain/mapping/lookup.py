"""Lookup indexes and relationship resolution.

Each relationship gets one hash index per job (built lazily, reused for every
record), so lookups are O(1) instead of scanning sheets. Duplicate keys are
detected at build time and reported with all offending row numbers.
"""

from __future__ import annotations

from collections.abc import Callable, Sequence
from dataclasses import dataclass, field
from typing import Any

from docgen.domain.data.table import RowView, SheetTable
from docgen.domain.data.values import canonical_key, is_empty
from docgen.domain.diagnostics import Category, Issue, Severity, make_issue
from docgen.domain.profile.models import (
    KeyNormalize,
    MatchMode,
    OnDuplicate,
    OnMissing,
    Relationship,
    SortKey,
)

Key = tuple[str | None, ...]


def make_key(values: Sequence[Any], norm: KeyNormalize) -> Key:
    return tuple(
        canonical_key(
            v,
            trim=norm.trim,
            casefold=norm.casefold,
            numeric_text=norm.numeric_text,
            strip_leading_zeros=norm.strip_leading_zeros,
            accents=norm.strip_accents,
        )
        for v in values
    )


def sort_rows(table: SheetTable, indices: list[int], sort: Sequence[SortKey]) -> list[int]:
    """Stable multi-key sort of row indices by target columns; ties keep sheet order."""
    result = list(indices)
    for key in reversed(sort):
        if not table.has_column(key.by):
            continue
        pos = table.column_position(key.by)
        present = [i for i in result if not is_empty(table.rows[i][pos])]
        missing = [i for i in result if is_empty(table.rows[i][pos])]
        present.sort(key=lambda i: _sortable(table.rows[i][pos]), reverse=key.order == "desc")
        result = missing + present if key.nulls == "first" else present + missing
    return result


def _sortable(value: Any) -> tuple[int, Any]:
    from datetime import date, datetime
    from decimal import Decimal

    if isinstance(value, bool):
        return (0, int(value))
    if isinstance(value, (int, float, Decimal)):
        return (0, value)
    if isinstance(value, datetime):
        return (1, value)
    if isinstance(value, date):
        return (1, datetime(value.year, value.month, value.day))
    return (2, str(value).casefold())


@dataclass(slots=True)
class LookupIndex:
    relationship: Relationship
    table: SheetTable
    entries: dict[Key, list[int]] = field(default_factory=dict)
    duplicates: dict[Key, list[int]] = field(default_factory=dict)
    issues: list[Issue] = field(default_factory=list)

    @classmethod
    def build(
        cls,
        relationship: Relationship,
        table: SheetTable,
        *,
        row_filter: Callable[[RowView], bool] | None = None,
    ) -> LookupIndex:
        index = cls(relationship, table)
        positions = [table.column_position(p.to) for p in relationship.keys]
        norm = relationship.key_normalize
        for i, row in enumerate(table.rows):
            if row_filter is not None and not row_filter(RowView(table, i)):
                continue
            key = make_key([row[p] for p in positions], norm)
            if all(part is None for part in key):
                continue
            index.entries.setdefault(key, []).append(i)
        if relationship.sort:
            for key, rows in index.entries.items():
                index.entries[key] = sort_rows(table, rows, relationship.sort)
        if relationship.effective_match is not MatchMode.ALL:
            for key, rows in index.entries.items():
                if len(rows) > 1:
                    index.duplicates[key] = rows
            if index.duplicates and relationship.effective_match is MatchMode.UNIQUE:
                severity = (
                    Severity.ERROR if relationship.on_duplicate is OnDuplicate.ERROR else Severity.WARNING
                )
                for key, rows in list(index.duplicates.items())[:200]:
                    excel_rows = [table.row_numbers[r] for r in rows]
                    index.issues.append(
                        make_issue(
                            "lookup.duplicate_key",
                            severity,
                            Category.DUPLICATE,
                            f"Key {format_key(key)} appears {len(rows)} times in sheet "
                            f"'{table.sheet_name}' (rows {excel_rows[:10]})",
                            suggestion="Remove the duplicate rows, extend the key with more "
                            "columns (composite key), or set match: first if any match is acceptable.",
                            sheet=table.sheet_name,
                            row=excel_rows[0],
                            column=", ".join(p.to for p in relationship.keys),
                            rule=relationship.id,
                            key=format_key(key),
                            rows=excel_rows,
                        )
                    )
        return index

    def lookup(self, key: Key) -> list[int]:
        return self.entries.get(key, [])


def format_key(key: Key) -> str:
    return (
        "(" + ", ".join("∅" if k is None else repr(k) for k in key) + ")"
        if len(key) > 1
        else ("∅" if key[0] is None else repr(key[0]))
    )


@dataclass(slots=True)
class Resolution:
    """Result of following one relationship from one source row."""

    rows: list[int]
    issues: list[Issue]
    missing: bool = False
    empty_key: bool = False


def resolve(
    index: LookupIndex,
    source: SheetTable,
    source_row: int,
) -> Resolution:
    rel = index.relationship
    key_values = [source.value(source_row, p.from_) for p in rel.keys]
    key = make_key(key_values, rel.key_normalize)
    excel_row = source.row_numbers[source_row]
    if all(part is None for part in key) or any(part is None for part in key):
        if not rel.empty_key_is_missing:
            return Resolution([], [], empty_key=True)
    matches = index.lookup(key)
    issues: list[Issue] = []
    if not matches:
        if rel.on_missing in (OnMissing.ERROR, OnMissing.WARN):
            severity = Severity.ERROR if rel.on_missing is OnMissing.ERROR else Severity.WARNING
            issues.append(
                make_issue(
                    "lookup.missing_reference",
                    severity,
                    Category.MISSING_REFERENCE,
                    f"No row in sheet '{index.table.sheet_name}' matches key {format_key(key)} "
                    f"from sheet '{source.sheet_name}' row {excel_row}",
                    suggestion=f"Add the key to sheet '{index.table.sheet_name}', correct the "
                    "value in the source row, or set on_missing: default/warn/ignore.",
                    sheet=source.sheet_name,
                    row=excel_row,
                    column=", ".join(source.cell_ref(source_row, p.from_).column for p in rel.keys),
                    rule=rel.id,
                    key=format_key(key),
                )
            )
        return Resolution([], issues, missing=True)
    match_mode = rel.effective_match
    if match_mode is MatchMode.ALL:
        return Resolution(list(matches), issues)
    if len(matches) > 1 and match_mode is MatchMode.UNIQUE:
        excel_rows = [index.table.row_numbers[m] for m in matches]
        if rel.on_duplicate is OnDuplicate.ERROR:
            issues.append(
                make_issue(
                    "lookup.ambiguous",
                    Severity.ERROR,
                    Category.DUPLICATE,
                    f"Key {format_key(key)} from sheet '{source.sheet_name}' row {excel_row} "
                    f"matches {len(matches)} rows in '{index.table.sheet_name}' {excel_rows[:10]};"
                    " no value was chosen",
                    suggestion="Make the key unique or use match: first / on_duplicate: warn.",
                    sheet=source.sheet_name,
                    row=excel_row,
                    rule=rel.id,
                    key=format_key(key),
                )
            )
            return Resolution([], issues)
        issues.append(
            make_issue(
                "lookup.ambiguous_first_used",
                Severity.WARNING,
                Category.DUPLICATE,
                f"Key {format_key(key)} matches {len(matches)} rows in "
                f"'{index.table.sheet_name}'; using row {excel_rows[0]}",
                sheet=source.sheet_name,
                row=excel_row,
                rule=rel.id,
                key=format_key(key),
                target=index.table.sheet_name,
                count=len(matches),
                used_row=excel_rows[0],
            )
        )
    return Resolution([matches[0]], issues)
