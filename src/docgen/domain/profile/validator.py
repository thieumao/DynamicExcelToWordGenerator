"""Static (data-independent) validation of a profile.

Checks references between sheets, columns, relationships, fields and
templates; expression syntax; transform names/parameters; dependency cycles;
path conflicts; and template/merge configuration. Produces :class:`Issue`s so
the UI can show precise, actionable messages.
"""

from __future__ import annotations

from collections.abc import Sequence

from docgen.domain.diagnostics import Category, Issue, Severity, make_issue
from docgen.domain.mapping.engine import RuleNode, order_rules
from docgen.domain.mapping.expressions import ExpressionError, check_syntax
from docgen.domain.mapping.transforms import DEFAULT_REGISTRY, TransformRegistry
from docgen.domain.profile.models import (
    RESERVED_ROOTS,
    AggregationSpec,
    CollectionSpec,
    FieldRule,
    Profile,
    TransformCall,
    UnitKind,
)


def _err(code: str, message: str, *, suggestion: str | None = None, **ctx: object) -> Issue:
    return make_issue(code, Severity.ERROR, Category.CONFIG, message, suggestion=suggestion, **ctx)


def _warn(code: str, message: str, *, suggestion: str | None = None, **ctx: object) -> Issue:
    return make_issue(code, Severity.WARNING, Category.CONFIG, message, suggestion=suggestion, **ctx)


class ProfileValidator:
    def __init__(self, profile: Profile, registry: TransformRegistry = DEFAULT_REGISTRY) -> None:
        self.profile = profile
        self.registry = registry
        self.issues: list[Issue] = []
        self.sheet_ids = {s.id for s in profile.excel.sheets}
        self.columns = {s.id: {c.id for c in s.columns} for s in profile.excel.sheets}
        self.rels = {r.id: r for r in profile.relationships}

    def validate(self) -> list[Issue]:
        self._relationships()
        primary = self.profile.excel.primary
        self._rules(
            [*self.profile.fields, *self.profile.collections, *self.profile.aggregations],
            base=primary,
            scope="fields",
            top_level=True,
        )
        if self.profile.group_fields:
            if self.profile.generation.unit is UnitKind.RECORD:
                self.issues.append(
                    _warn(
                        "profile.group_fields_unused",
                        "group_fields are only evaluated when generation.unit is 'group' or 'dataset'",
                    )
                )
            self._rules(list(self.profile.group_fields), base=primary, scope="group_fields", top_level=True)
        self._generation()
        self._templates()
        self._validation_rules()
        return self.issues

    # ----------------------------------------------------------------
    def _relationships(self) -> None:
        for rel in self.profile.relationships:
            for side, sheet in (("from", rel.from_), ("to", rel.to)):
                if sheet not in self.sheet_ids:
                    self.issues.append(
                        _err(
                            "relationship.unknown_sheet",
                            f"Relationship '{rel.id}' {side} sheet '{sheet}' is not declared",
                            rule=rel.id,
                            target=sheet,
                        )
                    )
            for pair in rel.keys:
                if rel.from_ in self.columns and pair.from_ not in self.columns[rel.from_]:
                    self.issues.append(
                        _err(
                            "relationship.unknown_column",
                            f"Relationship '{rel.id}': column '{pair.from_}' is not declared in "
                            f"sheet '{rel.from_}'",
                            rule=rel.id,
                            col=pair.from_,
                            sheet_id=rel.from_,
                        )
                    )
                if rel.to in self.columns and pair.to not in self.columns[rel.to]:
                    self.issues.append(
                        _err(
                            "relationship.unknown_column",
                            f"Relationship '{rel.id}': column '{pair.to}' is not declared in "
                            f"sheet '{rel.to}'",
                            rule=rel.id,
                            col=pair.to,
                            sheet_id=rel.to,
                        )
                    )
            if rel.filter:
                self._expr(rel.filter, rule=rel.id)

    def _chain(self, chain: str, base: str, name: str) -> tuple[str | None, bool]:
        """Validate 'rel1.rel2' from *base*; return (final sheet, passes through to-many)."""
        sheet: str | None = base
        to_many = False
        for rel_id in chain.split("."):
            rel = self.rels.get(rel_id)
            if rel is None:
                self.issues.append(
                    _err(
                        "field.unknown_relationship",
                        f"'{name}' uses unknown relationship '{rel_id}'",
                        suggestion=f"Declared relationships: {sorted(self.rels)}",
                        field=name,
                        relationship=rel_id,
                        available=", ".join(sorted(self.rels)),
                    )
                )
                return None, to_many
            if rel.from_ != sheet:
                self.issues.append(
                    _err(
                        "field.bad_chain",
                        f"'{name}': relationship '{rel_id}' starts at sheet '{rel.from_}' but the "
                        f"path is at sheet '{sheet}'",
                        field=name,
                        relationship=rel_id,
                        start=rel.from_,
                        sheet_id=str(sheet),
                    )
                )
                return None, to_many
            to_many = to_many or rel.cardinality.to_many
            sheet = rel.to
        return sheet, to_many

    def _rules(self, rules: Sequence[RuleNode], *, base: str, scope: str, top_level: bool) -> None:
        names = [r.name for r in rules]
        dupes = sorted({n for n in names if names.count(n) > 1})
        for d in dupes:
            self.issues.append(_err("field.duplicate", f"Output '{d}' is defined more than once", field=d))
        for a in names:
            for b in names:
                if a != b and b.startswith(a + "."):
                    self.issues.append(
                        _err(
                            "field.path_conflict",
                            f"'{a}' is a value but '{b}' needs it to be an object",
                            suggestion=f"Rename '{a}' (e.g. '{a}.value') or '{b}'.",
                            field=a,
                            other=b,
                        )
                    )
        if top_level:
            sheet_like = self.sheet_ids | set(self.rels)
            for n in names:
                root = n.split(".")[0]
                if root in RESERVED_ROOTS:
                    self.issues.append(
                        _err(
                            "field.reserved_name",
                            f"'{n}' uses the reserved root name '{root}'",
                            field=n,
                            root=root,
                        )
                    )
                if root in sheet_like:
                    self.issues.append(
                        _err(
                            "field.shadows_sheet",
                            f"'{n}' starts with '{root}', which is also a sheet/relationship id; "
                            "expressions would be ambiguous",
                            field=n,
                            root=root,
                        )
                    )
        _, cycles = order_rules(list(rules))
        for cycle in cycles:
            self.issues.append(
                _err(
                    "field.cycle",
                    "Circular dependency: " + " -> ".join(cycle),
                    suggestion="Break the cycle so no field (indirectly) depends on itself.",
                    field=cycle[0],
                    cycle=" → ".join(cycle),
                )
            )
        for rule in rules:
            if isinstance(rule, FieldRule):
                self._field(rule, base, scope)
            elif isinstance(rule, CollectionSpec):
                self._collection(rule, base)
            elif isinstance(rule, AggregationSpec):
                self._aggregation(rule, base)

    def _column(self, sheet: str | None, column: str, name: str) -> None:
        if sheet is None or sheet not in self.columns:
            return
        if column not in self.columns[sheet]:
            self.issues.append(
                _err(
                    "field.unknown_column",
                    f"'{name}': column '{column}' is not declared in sheet '{sheet}'",
                    suggestion=f"Declared columns: {sorted(self.columns[sheet])}",
                    field=name,
                    col=column,
                    sheet_id=sheet,
                    available=", ".join(sorted(self.columns[sheet])),
                )
            )

    def _source_ref(self, ref: str, base: str, name: str, aggregate: str | None) -> None:
        if ref.startswith(base + "."):
            self._column(base, ref.split(".", 1)[1], name)
        elif "." in ref:
            chain, column = ref.rsplit(".", 1)
            sheet, to_many = self._chain(chain, base, name)
            self._column(sheet, column, name)
            if to_many and aggregate is None:
                self.issues.append(
                    _err(
                        "field.to_many_without_aggregate",
                        f"'{name}' follows a one-to-many relationship and would produce a list",
                        suggestion="Add 'aggregate' (join/first/sum/...) or use a collection.",
                        field=name,
                    )
                )
        else:
            self._column(base, ref, name)

    def _field(self, rule: FieldRule, base: str, scope: str) -> None:
        name = rule.name
        if rule.source is not None:
            self._source_ref(rule.source, base, name, rule.aggregate)
        if rule.path is not None:
            if "." not in rule.path:
                self.issues.append(
                    _err("field.bad_path", f"'{name}': path must be 'relationship.column'", field=name)
                )
            else:
                self._source_ref(rule.path, base, name, rule.aggregate)
        for ref in rule.coalesce or ():
            self._source_ref(ref, base, name, rule.aggregate)
        if rule.expr:
            self._expr(rule.expr, field=name)
        for clause in rule.when or ():
            if clause.if_:
                self._expr(clause.if_, field=name)
            if clause.expr:
                self._expr(clause.expr, field=name)
        self._transforms(rule.transforms, name)

    def _collection(self, spec: CollectionSpec, base: str) -> None:
        sheet, to_many = self._chain(spec.relationship, base, spec.name)
        if sheet is None:
            return
        if not to_many:
            self.issues.append(
                _warn(
                    "collection.not_to_many",
                    f"Collection '{spec.name}' follows only to-one relationships (at most one item)",
                    field=spec.name,
                )
            )
        if spec.filter:
            self._expr(spec.filter, field=spec.name)
        self._rules(
            [*spec.fields, *spec.collections, *spec.aggregations],
            base=sheet,
            scope=spec.name,
            top_level=False,
        )

    def _aggregation(self, spec: AggregationSpec, base: str) -> None:
        sheet, _ = self._chain(spec.relationship, base, spec.name)
        if spec.of is not None:
            self._column(sheet, spec.of, spec.name)
        if spec.filter:
            self._expr(spec.filter, field=spec.name)
        self._transforms(spec.transforms, spec.name)

    def _transforms(self, transforms: Sequence[TransformCall], name: str) -> None:
        for call in transforms:
            for problem in self.registry.validate_params(call.name, call.params):
                self.issues.append(
                    _err(
                        "transform.invalid",
                        f"'{name}': {problem}",
                        suggestion=f"Available transforms: {', '.join(self.registry.names())}",
                        field=name,
                        rule=call.name,
                        detail=problem,
                        available=", ".join(self.registry.names()),
                    )
                )

    def _expr(self, source: str, **ctx: str) -> None:
        try:
            check_syntax(source)
        except ExpressionError as exc:
            self.issues.append(
                make_issue(
                    "expression.syntax",
                    Severity.ERROR,
                    Category.EXPRESSION,
                    str(exc),
                    suggestion="Expressions use Jinja syntax, e.g. emp.grade >= 5 and emp.city == 'Hà Nội'",
                    source=source,
                    detail=str(exc),
                    **ctx,
                )
            )

    def _generation(self) -> None:
        gen = self.profile.generation
        known = {f.name for f in self.profile.fields} | {a.name for a in self.profile.aggregations}
        for path in gen.group_by:
            if path not in known:
                self.issues.append(
                    _err(
                        "generation.unknown_group_by",
                        f"group_by '{path}' is not a declared field",
                        suggestion=f"Use one of: {sorted(known)[:10]}",
                        field=path,
                        available=", ".join(sorted(known)[:10]),
                    )
                )
        if gen.filter:
            self._expr(gen.filter, rule="generation.filter")

    def _templates(self) -> None:
        if not self.profile.templates:
            self.issues.append(
                _warn(
                    "template.none",
                    "The profile has no Word templates yet",
                    suggestion="Import templates on the Templates page.",
                )
            )
        ids = {t.id for t in self.profile.templates}
        for tid in self.profile.merge.order:
            if tid not in ids:
                self.issues.append(
                    _err(
                        "merge.unknown_template",
                        f"merge.order references unknown template '{tid}'",
                        template=tid,
                    )
                )
        for t in self.profile.templates:
            if t.include_if:
                self._expr(t.include_if, template=t.id)
            if t.marker.style.value == "custom" and t.marker.open == t.marker.close:
                self.issues.append(
                    _warn(
                        "template.marker_same_delimiters",
                        f"Template '{t.id}': identical open/close markers make parsing ambiguous",
                        template=t.id,
                    )
                )

    def _validation_rules(self) -> None:
        for rule in self.profile.validation:
            if rule.expr:
                self._expr(rule.expr, field=rule.field)
            if rule.pattern:
                import re

                try:
                    re.compile(rule.pattern)
                except re.error as exc:
                    self.issues.append(
                        _err(
                            "validation.bad_regex",
                            f"Invalid regex for '{rule.field}': {exc}",
                            field=rule.field,
                            detail=str(exc),
                        )
                    )


def validate_profile(profile: Profile, registry: TransformRegistry = DEFAULT_REGISTRY) -> list[Issue]:
    return ProfileValidator(profile, registry).validate()
