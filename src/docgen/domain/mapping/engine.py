"""The mapping engine: Excel rows -> normalised, traceable document context.

Independent of the UI and of file formats: it works on a :class:`DataSet` and a
validated :class:`Profile`. Rules are compiled once per job; lookup indexes are
built once per relationship and reused for every record.
"""

from __future__ import annotations

import datetime as dt
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass, field
from decimal import Decimal
from typing import Any

from docgen.domain.data import values as V
from docgen.domain.data.table import CellRef, DataSet, RowView, SheetTable
from docgen.domain.diagnostics import Category, Issue, Severity, make_issue
from docgen.domain.mapping.context import set_path
from docgen.domain.mapping.expressions import (
    ExpressionEngine,
    ExpressionError,
    expression_references,
)
from docgen.domain.mapping.graph import Node, order_nodes
from docgen.domain.mapping.lookup import LookupIndex, resolve
from docgen.domain.mapping.transforms import (
    DEFAULT_REGISTRY,
    TransformContext,
    TransformError,
    TransformRegistry,
)
from docgen.domain.profile.models import (
    AggregationSpec,
    CollectionSpec,
    FieldRule,
    FieldType,
    FormattingConfig,
    OnMissing,
    Profile,
    SortKey,
    TransformCall,
)


@dataclass(frozen=True, slots=True)
class ImageRef:
    """Image placeholder value; the Word renderer turns it into an inline picture."""

    path: str
    width_mm: float | None = None
    height_mm: float | None = None
    max_bytes: int = 20 * 1024 * 1024
    base_dir: str | None = None

    def __str__(self) -> str:
        return self.path


@dataclass(slots=True)
class FieldTrace:
    path: str
    kind: str
    sources: list[CellRef] = field(default_factory=list)
    steps: list[str] = field(default_factory=list)
    value: Any = None
    untrusted: bool = False

    def describe(self) -> str:
        src = ", ".join(str(s) for s in self.sources) or "-"
        return f"{self.path} <- {self.kind} [{src}] " + " -> ".join(self.steps)


@dataclass(slots=True)
class RecordResult:
    row_index: int
    excel_row: int
    context: dict[str, Any]
    traces: dict[str, FieldTrace]
    issues: list[Issue]

    @property
    def has_errors(self) -> bool:
        return any(i.severity is Severity.ERROR for i in self.issues)


@dataclass(slots=True)
class GroupResult:
    key: tuple[Any, ...]
    records: list[RecordResult]
    context: dict[str, Any]
    issues: list[Issue]

    @property
    def has_errors(self) -> bool:
        return any(i.severity is Severity.ERROR for i in self.issues) or any(
            r.has_errors for r in self.records
        )


def transform_context(fmt: FormattingConfig) -> TransformContext:
    return TransformContext(
        date_format=fmt.date_format,
        datetime_format=fmt.datetime_format,
        decimal_separator=fmt.decimal_separator,
        thousands_separator=fmt.thousands_separator,
        currency_symbol=fmt.currency_symbol,
        currency_position=fmt.currency_position,
        currency_decimals=fmt.currency_decimals,
        true_text=fmt.true_text,
        false_text=fmt.false_text,
        input_date_formats=tuple(fmt.input_date_formats),
    )


# A rule node in a scope: one of FieldRule / CollectionSpec / AggregationSpec
RuleNode = FieldRule | CollectionSpec | AggregationSpec


def rule_expressions(rule: RuleNode) -> list[str]:
    exprs: list[str] = []
    if isinstance(rule, FieldRule):
        if rule.expr:
            exprs.append(rule.expr)
        for clause in rule.when or ():
            if clause.if_:
                exprs.append(clause.if_)
            if clause.expr:
                exprs.append(clause.expr)
    return exprs


def order_rules(rules: Sequence[RuleNode]) -> tuple[list[RuleNode], list[list[str]]]:
    nodes = []
    for rule in rules:
        refs: set[tuple[str, ...]] = set()
        for source in rule_expressions(rule):
            try:
                refs |= expression_references(source)
            except ExpressionError:
                continue
        nodes.append(Node(rule.name, frozenset(refs)))
    ordering = order_nodes(nodes)
    return [rules[i] for i in ordering.order], ordering.cycles


@dataclass(slots=True)
class _Scope:
    """Evaluation scope: a base sheet row plus the context built so far."""

    table: SheetTable
    row: int
    ctx: dict[str, Any]
    extra: dict[str, Any]
    base_name: str

    def namespace(self) -> dict[str, Any]:
        ns: dict[str, Any] = dict(self.extra)
        ns.update(self.ctx)
        ns[self.base_name] = RowView(self.table, self.row)
        return ns


class MappingEngine:
    def __init__(
        self,
        profile: Profile,
        dataset: DataSet,
        *,
        registry: TransformRegistry = DEFAULT_REGISTRY,
        trace: bool = False,
        report: Mapping[str, Any] | None = None,
    ) -> None:
        self.profile = profile
        self.dataset = dataset
        self.registry = registry
        self.trace_enabled = trace
        self.tctx = transform_context(profile.formatting)
        self.expressions = ExpressionEngine(registry=registry, ctx=self.tctx)
        self.report = dict(report or {})
        self.primary_id = profile.excel.primary
        self.primary = dataset.table(self.primary_id)
        self._indexes: dict[str, LookupIndex] = {}
        self.index_issues: list[Issue] = []
        # cells whose values must not be trusted, keyed like CellRef for O(1) checks
        self._invalid: dict[tuple[str, int, str], tuple[str, str]] = {}
        for table in dataset.tables.values():
            for (ri, ci), reason in table.invalid.items():
                key = (table.sheet_name, table.row_numbers[ri], table.columns[ci].header)
                self._invalid[key] = reason
        self._ordered_cache: dict[int, list[RuleNode]] = {}
        self._expr_columns: dict[tuple[str, str], list[str]] = {}
        top: list[RuleNode] = [*profile.fields, *profile.collections, *profile.aggregations]
        self.top_rules, self.cycles = order_rules(top)
        self.group_rules, group_cycles = order_rules(list(profile.group_fields))
        self.cycles += group_cycles

    # ----------------------------------------------------------- indexes
    def index_for(self, rel_id: str) -> LookupIndex:
        index = self._indexes.get(rel_id)
        if index is None:
            rel = self.profile.relationship(rel_id)
            if rel is None:
                raise KeyError(f"unknown relationship '{rel_id}'")
            target = self.dataset.table(rel.to)
            row_filter: Callable[[RowView], bool] | None = None
            if rel.filter:
                compiled = self.expressions.compile(rel.filter)

                def row_filter(view: RowView, _c: Any = compiled, _sid: str = rel.to) -> bool:
                    ns = {_sid: view, "row": view, **{k: view[k] for k in view}}
                    try:
                        return bool(_c(ns))
                    except ExpressionError:
                        return False

            index = LookupIndex.build(rel, target, row_filter=row_filter)
            self.index_issues.extend(index.issues)
            self._indexes[rel_id] = index
        return index

    def build_all_indexes(self) -> list[Issue]:
        for rel in self.profile.relationships:
            if rel.from_ in self.dataset.tables and rel.to in self.dataset.tables:
                self.index_for(rel.id)
        return list(self.index_issues)

    # ----------------------------------------------------------- records
    def map_record(self, row_index: int) -> RecordResult:
        table = self.primary
        ctx: dict[str, Any] = {}
        traces: dict[str, FieldTrace] = {}
        issues: list[Issue] = []
        scope = _Scope(table, row_index, ctx, {"report": self.report}, self.primary_id)
        self._eval_rules(self.top_rules, scope, traces, issues, prefix="")
        excel_row = table.row_numbers[row_index]
        return RecordResult(row_index, excel_row, ctx, traces, _dedupe(issues, table.sheet_name, excel_row))

    def map_group(
        self, key: tuple[Any, ...], records: Sequence[RecordResult], key_names: Sequence[str]
    ) -> GroupResult:
        contexts = [r.context for r in records]
        ctx: dict[str, Any] = {
            "group": {
                "key": " / ".join(V.to_text(k) for k in key) if key else "",
                "keys": {},
                "count": len(records),
                "records": contexts,
            },
            "records": contexts,
            "first": contexts[0] if contexts else {},
        }
        for name, value in zip(key_names, key, strict=False):
            set_path(ctx["group"]["keys"], name, value)
        issues: list[Issue] = []
        traces: dict[str, FieldTrace] = {}
        if self.group_rules:
            row = records[0].row_index if records else 0
            scope = _Scope(self.primary, row, ctx, {"report": self.report}, self.primary_id)
            self._eval_rules(self.group_rules, scope, traces, issues, prefix="")
        return GroupResult(key, list(records), ctx, issues)

    # ------------------------------------------------------------- rules
    def _eval_rules(
        self,
        rules: Sequence[RuleNode],
        scope: _Scope,
        traces: dict[str, FieldTrace],
        issues: list[Issue],
        prefix: str,
    ) -> None:
        for rule in rules:
            if isinstance(rule, FieldRule):
                value, trace = self._eval_field(rule, scope, issues, prefix)
            elif isinstance(rule, CollectionSpec):
                value, trace = self._eval_collection(rule, scope, issues, prefix)
            else:
                value, trace = self._eval_aggregation(rule, scope, issues, prefix)
            try:
                set_path(scope.ctx, rule.name, value)
            except ValueError as exc:
                issues.append(
                    make_issue(
                        "mapping.path_conflict",
                        Severity.ERROR,
                        Category.CONFIG,
                        str(exc),
                        field=prefix + rule.name,
                    )
                )
            if trace is not None:
                traces[prefix + rule.name] = trace

    def _new_trace(self, name: str, kind: str) -> FieldTrace | None:
        return FieldTrace(name, kind) if self.trace_enabled else None

    def _eval_field(
        self, rule: FieldRule, scope: _Scope, issues: list[Issue], prefix: str
    ) -> tuple[Any, FieldTrace | None]:
        name = prefix + rule.name
        trace = self._new_trace(name, rule.kind)
        sources: list[CellRef] = []
        try:
            value = self._raw_value(rule, scope, issues, name, sources, trace)
        except ExpressionError as exc:
            issues.append(
                make_issue(
                    "mapping.expression_failed",
                    Severity.ERROR,
                    Category.EXPRESSION,
                    str(exc),
                    suggestion="Check the expression syntax and referenced field names.",
                    field=name,
                    rule=name,
                )
            )
            value = None
        untrusted = self._check_untrusted(sources, issues, name) if self._invalid else False
        value = self._finish_value(value, rule, issues, name, trace, sources, scope)
        if trace is not None:
            trace.sources = sources
            trace.value = value
            trace.untrusted = untrusted
        return value, trace

    def _check_untrusted(self, sources: Sequence[CellRef], issues: list[Issue], name: str) -> bool:
        found = False
        for ref in sources:
            reason = self._invalid.get((ref.sheet, ref.row, ref.column))
            if reason is None:
                continue
            found = True
            code, severity = reason
            text = (
                "a formula without a calculated result"
                if code == "formula.no_cached_value"
                else "a cell whose value could not be converted"
            )
            issues.append(
                make_issue(
                    "mapping.untrusted_source",
                    Severity(severity),
                    Category.FORMULA if code.startswith("formula") else Category.TYPE,
                    f"Field '{name}' uses {text} ({ref})",
                    suggestion="Fix the cell in Excel (recalculate and save for formulas).",
                    field=name,
                    sheet=ref.sheet,
                    row=ref.row,
                    column=ref.column,
                )
            )
        return found

    def _raw_value(
        self,
        rule: FieldRule,
        scope: _Scope,
        issues: list[Issue],
        name: str,
        sources: list[CellRef],
        trace: FieldTrace | None,
    ) -> Any:
        if rule.source is not None:
            return self._column_value(rule.source, scope, issues, name, sources, trace)
        if rule.path is not None:
            return self._path_value(
                rule.path, scope, issues, name, sources, trace, rule.aggregate, rule.separator
            )
        if rule.coalesce is not None:
            for ref in rule.coalesce:
                if "." in ref and not ref.startswith(scope.base_name + "."):
                    value = self._path_value(
                        ref, scope, issues, name, sources, trace, rule.aggregate, rule.separator
                    )
                else:
                    value = self._column_value(ref, scope, issues, name, sources, trace)
                if not V.is_empty(value):
                    if trace is not None:
                        trace.steps.append(f"coalesce picked {ref}")
                    return value
            return None
        if rule.expr is not None:
            self._expression_sources(rule.expr, scope, sources)
            if trace is not None:
                trace.steps.append(f"expr {rule.expr}")
            return self.expressions.compile(rule.expr)(scope.namespace())
        if rule.when is not None:
            for clause in rule.when:
                for text in (clause.if_, clause.expr):
                    if text:
                        self._expression_sources(text, scope, sources)
            ns = scope.namespace()
            for clause in rule.when:
                if clause.is_else:
                    if trace is not None:
                        trace.steps.append("when: else")
                    return clause.else_
                assert clause.if_ is not None
                if self.expressions.compile(clause.if_)(ns):
                    if trace is not None:
                        trace.steps.append(f"when: {clause.if_}")
                    if clause.expr is not None:
                        return self.expressions.compile(clause.expr)(ns)
                    return clause.value
            if trace is not None:
                trace.steps.append("when: no clause matched")
            return None
        return rule.const

    def _expression_sources(self, source: str, scope: _Scope, sources: list[CellRef]) -> None:
        """Record base-sheet cells read by an expression (traceability + untrusted checks)."""
        columns = self._expr_columns.get((source, scope.base_name))
        if columns is None:
            columns = []
            try:
                for chain in expression_references(source):
                    if len(chain) >= 2 and chain[0] in (scope.base_name, "item"):
                        if scope.table.has_column(chain[1]) and chain[1] not in columns:
                            columns.append(chain[1])
            except ExpressionError:
                pass
            self._expr_columns[(source, scope.base_name)] = columns
        for column in columns:
            sources.append(scope.table.cell_ref(scope.row, column))

    def _column_value(
        self,
        ref: str,
        scope: _Scope,
        issues: list[Issue],
        name: str,
        sources: list[CellRef],
        trace: FieldTrace | None,
    ) -> Any:
        column = ref.split(".", 1)[1] if ref.startswith(scope.base_name + ".") else ref
        table = scope.table
        if not table.has_column(column):
            issues.append(
                make_issue(
                    "mapping.unknown_column",
                    Severity.ERROR,
                    Category.CONFIG,
                    f"Column '{column}' is not declared for sheet '{table.sheet_id}'",
                    field=name,
                )
            )
            return None
        sources.append(table.cell_ref(scope.row, column))
        return table.value(scope.row, column)

    def _follow(
        self,
        rel_ids: Sequence[str],
        table: SheetTable,
        rows: list[int],
        issues: list[Issue],
        name: str,
        trace: FieldTrace | None,
    ) -> tuple[SheetTable, list[int], bool, Any]:
        """Follow a relationship chain. Returns (table, rows, to_many, default_if_missing)."""
        to_many = False
        default: Any = None
        for rel_id in rel_ids:
            rel = self.profile.relationship(rel_id)
            if rel is None or rel.from_ != table.sheet_id:
                issues.append(
                    make_issue(
                        "mapping.bad_relationship_chain",
                        Severity.ERROR,
                        Category.CONFIG,
                        f"Relationship '{rel_id}' cannot be followed from sheet '{table.sheet_id}'",
                        field=name,
                        rule=rel_id,
                    )
                )
                return table, [], to_many, None
            index = self.index_for(rel_id)
            next_rows: list[int] = []
            for r in rows:
                res = resolve(index, table, r)
                for issue in res.issues:
                    issues.append(issue.with_context(field=name))
                if res.missing and rel.on_missing is OnMissing.DEFAULT:
                    default = rel.default
                next_rows.extend(res.rows)
            if trace is not None:
                target = index.table
                trace.steps.append(
                    f"{rel_id}: {table.sheet_name} -> {target.sheet_name}"
                    f" rows {[target.row_numbers[x] for x in next_rows[:10]]}"
                )
            if rel.cardinality.to_many or len(rows) > 1:
                to_many = True
            table = index.table
            rows = next_rows
            if not rows:
                break
        return table, rows, to_many, default

    def _path_value(
        self,
        path: str,
        scope: _Scope,
        issues: list[Issue],
        name: str,
        sources: list[CellRef],
        trace: FieldTrace | None,
        aggregate: str | None,
        separator: str,
    ) -> Any:
        parts = path.split(".")
        rel_ids, column = parts[:-1], parts[-1]
        table, rows, to_many, default = self._follow(rel_ids, scope.table, [scope.row], issues, name, trace)
        if not rows:
            return default
        target_table = table
        if not target_table.has_column(column):
            issues.append(
                make_issue(
                    "mapping.unknown_column",
                    Severity.ERROR,
                    Category.CONFIG,
                    f"Column '{column}' is not declared for sheet '{target_table.sheet_id}'",
                    field=name,
                )
            )
            return None
        values = []
        for r in rows:
            sources.append(target_table.cell_ref(r, column))
            values.append(target_table.value(r, column))
        if not to_many and len(values) == 1:
            return values[0]
        if aggregate:
            return aggregate_values(aggregate, values, separator)
        return values

    def _finish_value(
        self,
        value: Any,
        rule: FieldRule,
        issues: list[Issue],
        name: str,
        trace: FieldTrace | None,
        sources: list[CellRef],
        scope: _Scope,
    ) -> Any:
        loc: dict[str, Any] = {"field": name}
        if sources:
            loc.update(sheet=sources[0].sheet, row=sources[0].row, column=sources[0].column)
        if rule.type not in (FieldType.AUTO, FieldType.IMAGE) and not V.is_empty(value):
            try:
                value = convert_type(value, rule.type, self.tctx)
                if trace is not None:
                    trace.steps.append(f"type {rule.type.value}")
            except V.ConversionError as exc:
                issues.append(
                    make_issue(
                        "mapping.type_conversion",
                        Severity.ERROR,
                        Category.TYPE,
                        f"Cannot convert {value!r} to {rule.type.value}: {exc}",
                        suggestion="Fix the source value or change the field type.",
                        **loc,
                    )
                )
                value = None
        value = self._apply_transforms(value, rule.transforms, issues, loc, trace)
        if V.is_empty(value) and rule.default is not None:
            value = rule.default
            if trace is not None:
                trace.steps.append("default")
        if rule.required and V.is_empty(value):
            issues.append(
                make_issue(
                    "mapping.required_empty",
                    Severity.ERROR,
                    Category.VALIDATION,
                    f"Required field '{name}' is empty",
                    suggestion="Fill the source cell or give the field a default value.",
                    **loc,
                )
            )
        if rule.type is FieldType.IMAGE and not V.is_empty(value):
            assert rule.image is not None
            value = ImageRef(
                V.to_text(value).strip(),
                rule.image.width_mm,
                rule.image.height_mm,
                rule.image.max_bytes,
                rule.image.base_dir,
            )
        return value

    def _apply_transforms(
        self,
        value: Any,
        transforms: Sequence[TransformCall],
        issues: list[Issue],
        loc: Mapping[str, Any],
        trace: FieldTrace | None,
    ) -> Any:
        for call in transforms:
            try:
                value = self.registry.apply(call.name, value, self.tctx, call.params)
                if trace is not None:
                    trace.steps.append(call.name)
            except TransformError as exc:
                issues.append(
                    make_issue(
                        "mapping.transform_failed",
                        Severity.ERROR,
                        Category.TRANSFORM,
                        f"Transform '{call.name}' failed on {value!r}: {exc}",
                        suggestion="Check the transform parameters or the source value.",
                        rule=call.name,
                        **loc,
                    )
                )
                return None
        return value

    # -------------------------------------------------------- collections
    def _collection_rows(
        self, relationship: str, scope: _Scope, issues: list[Issue], name: str, trace: FieldTrace | None
    ) -> tuple[SheetTable | None, list[int]]:
        rel_ids = relationship.split(".")
        _, rows, _, _ = self._follow(rel_ids, scope.table, [scope.row], issues, name, trace)
        last = self.profile.relationship(rel_ids[-1])
        if last is None:
            return None, []
        return self.dataset.table(last.to), rows

    def _eval_collection(
        self, spec: CollectionSpec, scope: _Scope, issues: list[Issue], prefix: str
    ) -> tuple[list[dict[str, Any]], FieldTrace | None]:
        name = prefix + spec.name
        trace = self._new_trace(name, "collection")
        table, rows = self._collection_rows(spec.relationship, scope, issues, name, trace)
        if table is None:
            return [], trace
        item_rules = self._ordered_cache.get(id(spec))
        if item_rules is None:
            item_rules, _ = order_rules([*spec.fields, *spec.collections, *spec.aggregations])
            self._ordered_cache[id(spec)] = item_rules
        filter_fn = self.expressions.compile(spec.filter) if spec.filter else None
        items: list[tuple[int, dict[str, Any]]] = []
        for r in rows:
            item_ctx: dict[str, Any] = {}
            view = RowView(table, r)
            extra = {"report": self.report, "parent": scope.ctx, "item": view, "this": item_ctx}
            if filter_fn is not None:
                try:
                    if not filter_fn({**extra, table.sheet_id: view}):
                        continue
                except ExpressionError as exc:
                    issues.append(
                        make_issue(
                            "mapping.expression_failed",
                            Severity.ERROR,
                            Category.EXPRESSION,
                            str(exc),
                            field=name,
                        )
                    )
                    continue
            item_scope = _Scope(table, r, item_ctx, extra, table.sheet_id)
            item_traces: dict[str, FieldTrace] = {}
            self._eval_rules(item_rules, item_scope, item_traces, issues, prefix=f"{name}[].")
            if trace is not None:
                trace.sources.append(CellRef(table.sheet_name, table.row_numbers[r], "*"))
            items.append((r, item_ctx))
        if spec.distinct_by:
            seen: set[tuple[str, ...]] = set()
            unique_items = []
            for r, item in items:
                key = tuple(V.to_text(_lookup_item(item, table, r, k)) for k in spec.distinct_by)
                if key not in seen:
                    seen.add(key)
                    unique_items.append((r, item))
            items = unique_items
        if spec.sort:
            items = sort_items(items, table, spec.sort)
        if spec.limit is not None:
            items = items[: spec.limit]
        if trace is not None:
            trace.steps.append(f"{len(items)} item(s)")
            trace.value = f"[{len(items)} items]"
        return [item for _, item in items], trace

    def _eval_aggregation(
        self, spec: AggregationSpec, scope: _Scope, issues: list[Issue], prefix: str
    ) -> tuple[Any, FieldTrace | None]:
        name = prefix + spec.name
        trace = self._new_trace(name, f"aggregate {spec.op}")
        table, rows = self._collection_rows(spec.relationship, scope, issues, name, trace)
        if table is None:
            return spec.default, trace
        if spec.filter:
            compiled = self.expressions.compile(spec.filter)
            kept = []
            for r in rows:
                view = RowView(table, r)
                try:
                    if compiled({"item": view, table.sheet_id: view, "report": self.report}):
                        kept.append(r)
                except ExpressionError as exc:
                    issues.append(
                        make_issue(
                            "mapping.expression_failed",
                            Severity.ERROR,
                            Category.EXPRESSION,
                            str(exc),
                            field=name,
                        )
                    )
            rows = kept
        if spec.op == "count" and spec.of is None:
            value: Any = len(rows)
        else:
            assert spec.of is not None
            if not table.has_column(spec.of):
                issues.append(
                    make_issue(
                        "mapping.unknown_column",
                        Severity.ERROR,
                        Category.CONFIG,
                        f"Column '{spec.of}' is not declared for sheet '{table.sheet_id}'",
                        field=name,
                    )
                )
                return spec.default, trace
            values = [table.value(r, spec.of) for r in rows]
            if spec.distinct:
                seen: set[str] = set()
                deduped = []
                for v in values:
                    key = V.to_text(v)
                    if key not in seen:
                        seen.add(key)
                        deduped.append(v)
                values = deduped
            try:
                value = aggregate_values(spec.op, values, spec.separator)
            except (V.ConversionError, ArithmeticError) as exc:
                issues.append(
                    make_issue(
                        "mapping.aggregate_failed",
                        Severity.ERROR,
                        Category.TRANSFORM,
                        f"Aggregation '{spec.op}' failed: {exc}",
                        field=name,
                    )
                )
                value = None
        value = self._apply_transforms(value, spec.transforms, issues, {"field": name}, trace)
        if V.is_empty(value) and spec.default is not None:
            value = spec.default
        if trace is not None:
            trace.value = value
            trace.steps.append(f"{len(rows)} row(s)")
        return value, trace


def _dedupe(issues: list[Issue], sheet: str, row: int) -> list[Issue]:
    """One lookup problem used by several fields is reported once per record."""
    seen: set[tuple[Any, ...]] = set()
    result = []
    for issue in issues:
        if issue.sheet is None:
            issue = issue.with_context(sheet=sheet, row=row)
        key = (issue.code, issue.rule, issue.sheet, issue.row, issue.message)
        if issue.code.startswith("lookup.") and key in seen:
            continue
        seen.add(key)
        result.append(issue)
    return result


def _lookup_item(item: Mapping[str, Any], table: SheetTable, row: int, key: str) -> Any:
    from docgen.domain.mapping.context import get_path

    value = get_path(item, key, default=None)
    if value is None and table.has_column(key):
        value = table.value(row, key)
    return value


def sort_items(
    items: list[tuple[int, dict[str, Any]]], table: SheetTable, sort: Sequence[SortKey]
) -> list[tuple[int, dict[str, Any]]]:
    from docgen.domain.mapping.lookup import _sortable

    result = list(items)
    for key in reversed(sort):
        present = [it for it in result if not V.is_empty(_lookup_item(it[1], table, it[0], key.by))]
        missing = [it for it in result if V.is_empty(_lookup_item(it[1], table, it[0], key.by))]
        present.sort(
            key=lambda it: _sortable(_lookup_item(it[1], table, it[0], key.by)),
            reverse=key.order == "desc",
        )
        result = missing + present if key.nulls == "first" else present + missing
    return result


def aggregate_values(op: str, values: Sequence[Any], separator: str = ", ") -> Any:
    non_empty = [v for v in values if not V.is_empty(v)]
    if op == "count":
        return len(non_empty)
    if op == "count_distinct":
        return len({V.to_text(v) for v in non_empty})
    if op == "list":
        return list(values)
    if op == "join":
        return separator.join(V.to_text(v) for v in non_empty)
    if op == "first":
        return non_empty[0] if non_empty else None
    if op == "last":
        return non_empty[-1] if non_empty else None
    if op in ("sum", "avg"):
        numbers = [V.to_decimal(v) for v in non_empty]
        if op == "sum":
            return sum(numbers, Decimal(0))
        return (sum(numbers, Decimal(0)) / len(numbers)) if numbers else None
    if op in ("min", "max"):
        if not non_empty:
            return None
        from docgen.domain.mapping.lookup import _sortable

        chooser = min if op == "min" else max
        return chooser(non_empty, key=_sortable)
    raise ValueError(f"unknown aggregate '{op}'")


def convert_type(value: Any, field_type: FieldType, tctx: TransformContext) -> Any:
    if field_type is FieldType.STRING:
        return V.to_text(value)
    if field_type is FieldType.INTEGER:
        if isinstance(value, str):
            return V.to_int(V.parse_localized_number(value, tctx.decimal_separator, tctx.thousands_separator))
        return V.to_int(value)
    if field_type is FieldType.DECIMAL:
        if isinstance(value, str):
            return V.parse_localized_number(value, tctx.decimal_separator, tctx.thousands_separator)
        return V.to_decimal(value)
    if field_type is FieldType.DATE:
        return V.to_date(value, tctx.input_date_formats)
    if field_type is FieldType.DATETIME:
        return V.to_datetime(value)
    if field_type is FieldType.BOOLEAN:
        return V.to_bool(value)
    return value


def context_to_jsonable(value: Any) -> Any:
    """Convert a context to JSON-friendly data for previews and audit files."""
    if isinstance(value, Mapping):
        return {str(k): context_to_jsonable(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [context_to_jsonable(v) for v in value]
    if isinstance(value, Decimal):
        return format(value, "f")
    if isinstance(value, (dt.date, dt.datetime, dt.time)):
        return value.isoformat()
    if isinstance(value, ImageRef):
        return {"image": value.path}
    return value
