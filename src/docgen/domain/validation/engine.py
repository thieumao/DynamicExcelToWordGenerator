"""Data validation rules applied to mapped record contexts."""

from __future__ import annotations

import re
from collections.abc import Sequence
from typing import Any

from docgen.domain.data import values as V
from docgen.domain.diagnostics import Category, Issue, Severity, make_issue
from docgen.domain.mapping.context import get_path
from docgen.domain.mapping.engine import RecordResult
from docgen.domain.mapping.expressions import ExpressionEngine, ExpressionError
from docgen.domain.profile.models import ValidationRuleSpec


def _severity(rule: ValidationRuleSpec) -> Severity:
    return Severity(rule.severity.value)


def _compare(value: Any, bound: Any) -> int:
    try:
        a, b = V.to_decimal(value), V.to_decimal(bound)
    except V.ConversionError:
        try:
            a2, b2 = V.to_date(value), V.to_date(bound)
        except V.ConversionError:
            a3, b3 = V.to_text(value), V.to_text(bound)
            return (a3 > b3) - (a3 < b3)
        return (a2 > b2) - (a2 < b2)
    return (a > b) - (a < b)


class ValidationEngine:
    def __init__(self, rules: Sequence[ValidationRuleSpec], expressions: ExpressionEngine) -> None:
        self.rules = list(rules)
        self.expressions = expressions
        self._patterns = {r.pattern: re.compile(r.pattern) for r in self.rules if r.pattern}
        self._seen: dict[str, dict[str, int]] = {}

    def check(self, record: RecordResult) -> list[Issue]:
        issues: list[Issue] = []
        for rule in self.rules:
            value = get_path(record.context, rule.field)
            found = self._check_one(rule, value, record)
            if found:
                problem, details = found
                issues.append(
                    make_issue(
                        f"validation.{rule.rule}",
                        _severity(rule),
                        Category.VALIDATION,
                        rule.message or problem,
                        field=rule.field,
                        rule=rule.rule,
                        row=record.excel_row,
                        value=V.to_text(value)[:100],
                        min="" if rule.min is None else V.to_text(rule.min),
                        max="" if rule.max is None else V.to_text(rule.max),
                        user_message=rule.message or "",  # a message written in the profile wins
                        **details,
                    )
                )
        return issues

    def _check_one(
        self, rule: ValidationRuleSpec, value: Any, record: RecordResult
    ) -> tuple[str, dict[str, Any]] | None:
        """(English message, extra message parameters) when the value breaks the rule."""
        empty = V.is_empty(value)
        if rule.rule == "required":
            return (f"'{rule.field}' is required but empty", {}) if empty else None
        if empty:
            return None
        text = V.to_text(value)
        if rule.rule == "regex":
            assert rule.pattern is not None
            if not self._patterns[rule.pattern].fullmatch(text):
                return f"'{rule.field}' value {text!r} does not match {rule.pattern}", {
                    "pattern": rule.pattern
                }
        elif rule.rule == "range":
            if rule.min is not None and _compare(value, rule.min) < 0:
                return f"'{rule.field}' value {text} is below {rule.min}", {}
            if rule.max is not None and _compare(value, rule.max) > 0:
                return f"'{rule.field}' value {text} is above {rule.max}", {}
        elif rule.rule == "length":
            n = len(text)
            if rule.min is not None and n < int(rule.min):
                return f"'{rule.field}' is shorter than {rule.min} characters", {"length": n}
            if rule.max is not None and n > int(rule.max):
                return f"'{rule.field}' is longer than {rule.max} characters", {"length": n}
        elif rule.rule == "one_of":
            allowed = {V.canonical_key(v) for v in rule.values or ()}
            if V.canonical_key(value) not in allowed:
                values = ", ".join(V.to_text(v) for v in rule.values or ())
                return f"'{rule.field}' value {text!r} is not one of {rule.values}", {"values": values}
        elif rule.rule == "unique":
            seen = self._seen.setdefault(rule.field, {})
            key = V.canonical_key(value) or ""
            if key in seen:
                return (
                    f"'{rule.field}' value {text!r} is duplicated (first seen at row {seen[key]})",
                    {"first_row": seen[key]},
                )
            seen[key] = record.excel_row
        elif rule.rule == "expr":
            assert rule.expr is not None
            try:
                ok = self.expressions.compile(rule.expr)({**record.context, "value": value})
            except ExpressionError as exc:
                return str(exc), {"expr": rule.expr, "detail": str(exc)}
            if not ok:
                return f"'{rule.field}' fails check: {rule.expr}", {"expr": rule.expr, "detail": ""}
        return None
