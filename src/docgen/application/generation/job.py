"""Job primitives shared by the CLI and the desktop UI (no Qt dependency)."""

from __future__ import annotations

import threading
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Protocol

from docgen.domain.diagnostics import Issue, Severity
from docgen.domain.profile.models import ErrorPolicy, OutputKind, Profile, UnitKind


class JobCancelledError(Exception):
    """Raised at a safe checkpoint after cancellation was requested."""


class CancellationToken:
    def __init__(self) -> None:
        self._event = threading.Event()

    def cancel(self) -> None:
        self._event.set()

    @property
    def cancelled(self) -> bool:
        return self._event.is_set()

    def raise_if_cancelled(self) -> None:
        if self._event.is_set():
            raise JobCancelledError()


class ProgressReporter(Protocol):
    def stage(self, name: str) -> None: ...

    def progress(self, done: int, total: int, message: str = "") -> None: ...

    def issue(self, issue: Issue) -> None: ...


class NullReporter:
    def stage(self, name: str) -> None:
        pass

    def progress(self, done: int, total: int, message: str = "") -> None:
        pass

    def issue(self, issue: Issue) -> None:
        pass


@dataclass(slots=True)
class GenerationRequest:
    profile: Profile
    profile_dir: Path
    excel_path: Path
    output_dir: Path
    unit: UnitKind | None = None
    output: OutputKind | None = None
    rows: list[int] | None = None
    """Excel row numbers of the primary sheet to include (None = all)."""
    unit_indices: list[int] | None = None
    """1-based unit numbers to generate after planning (None = all)."""
    template_ids: list[str] | None = None
    workers: int | None = None
    error_policy: ErrorPolicy | None = None
    dry_run: bool = False
    header_overrides: dict[str, int] = field(default_factory=dict)
    combined: bool | None = None
    """One extra file with every generated document, to print (None = output.combined_filename)."""


STATUS_OK = "ok"
STATUS_WARNING = "warning"
STATUS_ERROR = "error"
STATUS_SKIPPED = "skipped"
STATUS_CANCELLED = "cancelled"


@dataclass(slots=True)
class UnitOutcome:
    index: int
    label: str
    rows: list[int]
    status: str = STATUS_OK
    outputs: list[str] = field(default_factory=list)
    templates: list[str] = field(default_factory=list)
    issues: list[Issue] = field(default_factory=list)
    seconds: float = 0.0

    def to_dict(self) -> dict[str, Any]:
        return {
            "index": self.index,
            "label": self.label,
            "excel_rows": self.rows,
            "status": self.status,
            "outputs": self.outputs,
            "templates": self.templates,
            "errors": sum(1 for i in self.issues if i.severity is Severity.ERROR),
            "warnings": sum(1 for i in self.issues if i.severity is Severity.WARNING),
            "seconds": round(self.seconds, 3),
            "issues": [i.to_dict() for i in self.issues[:50]],
        }


@dataclass(slots=True)
class JobResult:
    job_id: str
    status: str
    started: str
    finished: str = ""
    profile_id: str = ""
    profile_name: str = ""
    profile_version: int = 0
    inputs: list[dict[str, Any]] = field(default_factory=list)
    output_dir: str = ""
    unit_kind: str = ""
    output_kind: str = ""
    records_total: int = 0
    records_selected: int = 0
    units: list[UnitOutcome] = field(default_factory=list)
    issues: list[Issue] = field(default_factory=list)
    audit_path: str = ""
    report_path: str = ""
    combined_path: str = ""
    """The single file with every document of the job, when requested."""
    stage_seconds: dict[str, float] = field(default_factory=dict)

    @property
    def documents(self) -> list[str]:
        return [o for u in self.units for o in u.outputs]

    @property
    def error_count(self) -> int:
        return sum(1 for i in self.issues if i.severity is Severity.ERROR) + sum(
            1 for u in self.units for i in u.issues if i.severity is Severity.ERROR
        )

    @property
    def warning_count(self) -> int:
        return sum(1 for i in self.issues if i.severity is Severity.WARNING) + sum(
            1 for u in self.units for i in u.issues if i.severity is Severity.WARNING
        )

    def all_issues(self) -> list[Issue]:
        return [*self.issues, *(i for u in self.units for i in u.issues)]
