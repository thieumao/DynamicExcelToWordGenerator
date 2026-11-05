"""Pass 1: map + validate every selected record and form generation units.

Contexts are discarded after this pass (only row indices, sort and group keys
are kept) so very large workbooks do not hold every context in memory.
"""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass, field
from typing import Any

from docgen.application.generation.job import CancellationToken, ProgressReporter
from docgen.domain.data import values as V
from docgen.domain.diagnostics import Category, Issue, Severity, make_issue
from docgen.domain.mapping.context import get_path
from docgen.domain.mapping.engine import MappingEngine
from docgen.domain.mapping.expressions import ExpressionError
from docgen.domain.mapping.lookup import _sortable
from docgen.domain.profile.models import SortKey, UnitKind
from docgen.domain.validation.engine import ValidationEngine


@dataclass(slots=True)
class PlannedUnit:
    index: int
    key: tuple[Any, ...]
    label: str
    row_indices: list[int]
    excel_rows: list[int]
    issues: list[Issue] = field(default_factory=list)

    @property
    def has_errors(self) -> bool:
        return any(i.severity is Severity.ERROR for i in self.issues)


@dataclass(slots=True)
class Plan:
    unit_kind: UnitKind
    units: list[PlannedUnit]
    records_total: int
    records_selected: int
    filtered_out: int
    issues: list[Issue] = field(default_factory=list)


def _sort_rows(
    records: list[tuple[int, dict[str, Any]]], sort: Sequence[SortKey]
) -> list[tuple[int, dict[str, Any]]]:
    """Stable multi-key sort; records carry flat {sort path: value} dicts."""
    result = list(records)
    for key in reversed(sort):
        present = [r for r in result if not V.is_empty(r[1]["sort"].get(key.by))]
        missing = [r for r in result if V.is_empty(r[1]["sort"].get(key.by))]
        present.sort(key=lambda r: _sortable(r[1]["sort"][key.by]), reverse=key.order == "desc")
        result = missing + present if key.nulls == "first" else present + missing
    return result


def plan_units(
    engine: MappingEngine,
    *,
    unit_kind: UnitKind,
    rows: Sequence[int] | None,
    validation: ValidationEngine | None,
    reporter: ProgressReporter,
    cancel: CancellationToken,
) -> Plan:
    profile = engine.profile
    table = engine.primary
    gen = profile.generation
    issues: list[Issue] = []
    if rows is None:
        indices = list(range(len(table)))
    else:
        indices = []
        for excel_row in rows:
            ri = table.find_row_index(excel_row)
            if ri is None:
                issues.append(
                    make_issue(
                        "plan.row_not_found",
                        Severity.WARNING,
                        Category.SHEET,
                        f"Row {excel_row} is not a data row of sheet '{table.sheet_name}'",
                        sheet=table.sheet_name,
                        row=excel_row,
                    )
                )
            else:
                indices.append(ri)
    filter_fn = engine.expressions.compile(gen.filter) if gen.filter else None
    kept: list[tuple[int, dict[str, Any]]] = []
    record_issues: dict[int, list[Issue]] = {}
    filtered_out = 0
    total = len(indices)
    for n, ri in enumerate(indices, start=1):
        if n % 200 == 0:
            cancel.raise_if_cancelled()
            reporter.progress(n, total, "Mapping records")
        result = engine.map_record(ri)
        rec_issues = list(result.issues)
        if validation is not None:
            rec_issues.extend(validation.check(result))
        if filter_fn is not None:
            try:
                ns = {**result.context, engine.primary_id: table.row_view(ri), "report": engine.report}
                if not filter_fn(ns):
                    filtered_out += 1
                    continue
            except ExpressionError as exc:
                rec_issues.append(
                    make_issue(
                        "plan.filter_failed",
                        Severity.ERROR,
                        Category.EXPRESSION,
                        str(exc),
                        row=result.excel_row,
                        sheet=table.sheet_name,
                        rule="generation.filter",
                        detail=str(exc),
                    )
                )
        record_issues[ri] = rec_issues
        # keep only what is needed for sorting / grouping
        slim: dict[str, Any] = {"sort": {k.by: get_path(result.context, k.by) for k in gen.sort}}
        if unit_kind is UnitKind.GROUP:
            slim["group"] = tuple(V.canonical_key(get_path(result.context, g)) or "" for g in gen.group_by)
            slim["group_raw"] = tuple(get_path(result.context, g) for g in gen.group_by)
        kept.append((ri, slim))
    reporter.progress(total, total, "Mapping records")

    if gen.sort:
        kept = _sort_rows(kept, gen.sort)

    units: list[PlannedUnit] = []
    if unit_kind is UnitKind.RECORD:
        for ri, _ in kept:
            excel_row = table.row_numbers[ri]
            units.append(
                PlannedUnit(
                    len(units) + 1,
                    (excel_row,),
                    f"{table.sheet_name} row {excel_row}",
                    [ri],
                    [excel_row],
                    record_issues.get(ri, []),
                )
            )
    elif unit_kind is UnitKind.DATASET:
        all_rows = [ri for ri, _ in kept]
        all_issues = [i for ri in all_rows for i in record_issues.get(ri, [])]
        units.append(
            PlannedUnit(
                1,
                (),
                f"All records ({len(all_rows)})",
                all_rows,
                [table.row_numbers[r] for r in all_rows],
                all_issues,
            )
        )
    else:
        groups: dict[tuple[Any, ...], list[int]] = {}
        raw_keys: dict[tuple[Any, ...], tuple[Any, ...]] = {}
        for ri, slim in kept:
            key = slim["group"]
            groups.setdefault(key, []).append(ri)
            raw_keys.setdefault(key, slim["group_raw"])
        for key, members in groups.items():
            raw = raw_keys[key]
            label = " / ".join(V.to_text(v) for v in raw) or "(empty)"
            units.append(
                PlannedUnit(
                    len(units) + 1,
                    raw,
                    label,
                    members,
                    [table.row_numbers[r] for r in members],
                    [i for r in members for i in record_issues.get(r, [])],
                )
            )
    return Plan(unit_kind, units, len(table), len(indices), filtered_out, issues)
