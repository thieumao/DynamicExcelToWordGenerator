"""Preview: resolved values with traces for selected records, plus the generation plan."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from docgen.application.generation.job import GenerationRequest, JobResult
from docgen.application.generation.service import DocumentGenerationService, JobSession
from docgen.domain.diagnostics import Issue
from docgen.domain.mapping.engine import FieldTrace, context_to_jsonable
from docgen.domain.mapping.lookup import format_key
from docgen.domain.profile.models import Profile


@dataclass(slots=True)
class RecordPreview:
    excel_row: int
    context: dict[str, Any]
    traces: list[FieldTrace]
    issues: list[Issue]


@dataclass(slots=True)
class RelationshipStats:
    relationship: str
    target_sheet: str
    keys: int
    duplicate_keys: int
    sample_duplicates: list[str] = field(default_factory=list)


@dataclass(slots=True)
class PreviewResult:
    ok: bool
    issues: list[Issue]
    records: list[RecordPreview] = field(default_factory=list)
    relationships: list[RelationshipStats] = field(default_factory=list)
    plan: list[tuple[int, str, list[int], list[str]]] = field(default_factory=list)
    """(unit index, label, excel rows, output file names)"""
    row_numbers: list[int] = field(default_factory=list)
    session: JobSession | None = None


def preview(
    profile: Profile,
    profile_dir: Path,
    excel: Path,
    *,
    rows: list[int] | None = None,
    limit: int = 5,
    with_plan: bool = True,
) -> PreviewResult:
    service = DocumentGenerationService()
    request = GenerationRequest(profile, profile_dir, excel, profile_dir, dry_run=True)
    job = JobResult("preview", "running", "")
    session = service.prepare(request, job_id="preview", result=job, plan=False)
    if session is None:
        return PreviewResult(False, job.issues)
    engine = session.engine
    engine.trace_enabled = True
    table = engine.primary
    result = PreviewResult(True, list(job.issues), session=session, row_numbers=list(table.row_numbers))
    for rel in profile.relationships:
        if rel.to not in session.load.dataset.tables or rel.from_ not in session.load.dataset.tables:
            continue
        index = engine.index_for(rel.id)
        result.relationships.append(
            RelationshipStats(
                rel.id,
                index.table.sheet_name,
                len(index.entries),
                len(index.duplicates),
                [format_key(k) for k in list(index.duplicates)[:5]],
            )
        )
    wanted = rows or table.row_numbers[:limit]
    for excel_row in wanted:
        ri = table.find_row_index(excel_row)
        if ri is None:
            continue
        rec = engine.map_record(ri)
        result.records.append(
            RecordPreview(
                rec.excel_row,
                context_to_jsonable(rec.context),
                list(rec.traces.values()),
                rec.issues,
            )
        )
    if with_plan:
        from docgen.application.generation.job import CancellationToken, NullReporter
        from docgen.application.generation.planner import plan_units
        from docgen.domain.validation.engine import ValidationEngine

        engine.trace_enabled = False
        plan = plan_units(
            engine,
            unit_kind=session.unit_kind,
            rows=None,
            validation=ValidationEngine(profile.validation, engine.expressions),
            reporter=NullReporter(),
            cancel=CancellationToken(),
        )
        session.plan = plan
        for unit in plan.units[:200]:
            ctx = service.unit_context(session, unit, len(plan.units))
            issues: list[Issue] = []
            _, filename = service._names(session, ctx, unit, issues, "preview")
            names = []
            if session.output_kind.value in ("merged", "both"):
                names.append(filename(profile.output.merged_filename))
            if session.output_kind.value in ("separate", "both"):
                for t in session.templates:
                    names.append(filename(profile.output.individual_filename, t.spec.id))
            result.plan.append((unit.index, unit.label, unit.excel_rows[:20], names))
    return result
