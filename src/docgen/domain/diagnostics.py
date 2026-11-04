"""Structured diagnostics shared by every layer.

An :class:`Issue` carries *where* (sheet/row/column/field/template/placeholder),
*what* (a stable ``code`` plus parameters) and *how to fix it* (suggestion).
Messages are rendered from the code by :mod:`docgen.i18n` so the UI can show
Vietnamese or English text; ``message`` holds the English fallback.
"""

from __future__ import annotations

from collections import Counter
from collections.abc import Iterable, Iterator, Mapping
from dataclasses import dataclass, field, replace
from enum import StrEnum
from typing import Any


class Severity(StrEnum):
    INFO = "info"
    WARNING = "warning"
    ERROR = "error"

    @property
    def rank(self) -> int:
        return {"info": 0, "warning": 1, "error": 2}[self.value]


class Category(StrEnum):
    CONFIG = "config"
    WORKBOOK = "workbook"
    SHEET = "sheet"
    COLUMN = "column"
    HEADER = "header"
    LOOKUP = "lookup"
    DUPLICATE = "duplicate"
    MISSING_REFERENCE = "missing_reference"
    FORMULA = "formula"
    TYPE = "type"
    TRANSFORM = "transform"
    EXPRESSION = "expression"
    VALIDATION = "validation"
    TEMPLATE = "template"
    PLACEHOLDER = "placeholder"
    RENDER = "render"
    MERGE = "merge"
    OUTPUT = "output"
    IO = "io"
    SECURITY = "security"
    CANCELLED = "cancelled"
    INTERNAL = "internal"


@dataclass(frozen=True, slots=True)
class Issue:
    """One diagnostic produced anywhere in the pipeline."""

    code: str
    severity: Severity
    category: Category
    message: str
    params: Mapping[str, Any] = field(default_factory=dict)
    sheet: str | None = None
    row: int | None = None
    column: str | None = None
    rule: str | None = None
    field: str | None = None
    template: str | None = None
    placeholder: str | None = None
    unit: str | None = None
    suggestion: str | None = None

    def with_context(self, **kwargs: Any) -> Issue:
        """Return a copy with location fields filled in (existing values win)."""
        updates = {k: v for k, v in kwargs.items() if getattr(self, k) is None and v is not None}
        return replace(self, **updates) if updates else self

    def location(self) -> str:
        parts: list[str] = []
        if self.template:
            parts.append(f"template={self.template}")
        if self.sheet:
            parts.append(f"sheet={self.sheet}")
        if self.row is not None:
            parts.append(f"row={self.row}")
        if self.column:
            parts.append(f"column={self.column}")
        if self.field:
            parts.append(f"field={self.field}")
        if self.placeholder:
            parts.append(f"placeholder={self.placeholder}")
        if self.rule:
            parts.append(f"rule={self.rule}")
        if self.unit:
            parts.append(f"unit={self.unit}")
        return ", ".join(parts)

    def to_dict(self) -> dict[str, Any]:
        return {
            "code": self.code,
            "severity": self.severity.value,
            "category": self.category.value,
            "message": self.message,
            "sheet": self.sheet,
            "row": self.row,
            "column": self.column,
            "rule": self.rule,
            "field": self.field,
            "template": self.template,
            "placeholder": self.placeholder,
            "unit": self.unit,
            "suggestion": self.suggestion,
        }

    def __str__(self) -> str:
        loc = self.location()
        text = f"[{self.severity.value.upper()}] {self.code}: {self.message}"
        if loc:
            text += f" ({loc})"
        if self.suggestion:
            text += f" -> {self.suggestion}"
        return text


def make_issue(
    code: str,
    severity: Severity,
    category: Category,
    message: str,
    *,
    suggestion: str | None = None,
    **context: Any,
) -> Issue:
    """Helper that splits location kwargs from free-form message params."""
    location_keys = {"sheet", "row", "column", "rule", "field", "template", "placeholder", "unit"}
    location = {k: v for k, v in context.items() if k in location_keys}
    params = {k: v for k, v in context.items() if k not in location_keys}
    params.update({k: v for k, v in location.items() if v is not None})
    return Issue(
        code=code,
        severity=severity,
        category=category,
        message=message,
        params=params,
        suggestion=suggestion,
        **location,
    )


class FailFastError(Exception):
    """Raised by :class:`IssueCollector` when the policy is fail-fast and an error arrives."""

    def __init__(self, issue: Issue) -> None:
        super().__init__(str(issue))
        self.issue = issue


class TooManyIssuesError(Exception):
    """Raised when the configured error budget is exhausted."""


class IssueCollector:
    """Accumulates issues with an optional fail-fast policy and memory cap.

    Only ``max_stored`` issues are kept in memory; counts remain exact so the
    audit report stays truthful even for very noisy inputs.
    """

    def __init__(
        self,
        *,
        fail_fast: bool = False,
        max_errors: int | None = None,
        max_stored: int = 10_000,
    ) -> None:
        self.fail_fast = fail_fast
        self.max_errors = max_errors
        self.max_stored = max_stored
        self._issues: list[Issue] = []
        self._counts: Counter[Severity] = Counter()
        self.dropped = 0

    def add(self, issue: Issue) -> None:
        self._counts[issue.severity] += 1
        if len(self._issues) < self.max_stored:
            self._issues.append(issue)
        else:
            self.dropped += 1
        if issue.severity is Severity.ERROR:
            if self.fail_fast:
                raise FailFastError(issue)
            if self.max_errors is not None and self._counts[Severity.ERROR] > self.max_errors:
                raise TooManyIssuesError(f"More than {self.max_errors} errors; stopping (last: {issue})")

    def extend(self, issues: Iterable[Issue]) -> None:
        for issue in issues:
            self.add(issue)

    def __iter__(self) -> Iterator[Issue]:
        return iter(self._issues)

    def __len__(self) -> int:
        return len(self._issues)

    @property
    def issues(self) -> list[Issue]:
        return list(self._issues)

    @property
    def error_count(self) -> int:
        return self._counts[Severity.ERROR]

    @property
    def warning_count(self) -> int:
        return self._counts[Severity.WARNING]

    @property
    def has_errors(self) -> bool:
        return self._counts[Severity.ERROR] > 0

    def errors(self) -> list[Issue]:
        return [i for i in self._issues if i.severity is Severity.ERROR]

    def warnings(self) -> list[Issue]:
        return [i for i in self._issues if i.severity is Severity.WARNING]
