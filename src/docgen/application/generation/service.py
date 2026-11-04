"""DocumentGenerationService: the end-to-end pipeline used by the CLI and the UI.

Excel input -> snapshot -> workbook loading -> profile/template validation ->
lookup indexes -> pass 1 (map + validate + plan) -> pass 2 (per unit: context ->
render -> merge -> verify -> atomic finalise) -> audit report -> cleanup.
"""

from __future__ import annotations

import datetime as dt
import hashlib
import logging
import multiprocessing
import os
import shutil
import tempfile
import time
import uuid
from collections.abc import Callable, Iterator, Mapping
from concurrent.futures import FIRST_COMPLETED, Future, ProcessPoolExecutor, wait
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from docgen.application.generation.audit import write_audit
from docgen.application.generation.job import (
    STATUS_CANCELLED,
    STATUS_ERROR,
    STATUS_OK,
    STATUS_SKIPPED,
    STATUS_WARNING,
    CancellationToken,
    GenerationRequest,
    JobCancelledError,
    JobResult,
    NullReporter,
    ProgressReporter,
    UnitOutcome,
)
from docgen.application.generation.planner import Plan, PlannedUnit, plan_units
from docgen.application.generation.worker import TemplateJob, UnitJob, UnitJobResult, render_unit
from docgen.application.templates import PreparedTemplate, prepare_template
from docgen.domain.data.table import DataSet
from docgen.domain.diagnostics import (
    Category,
    FailFastError,
    Issue,
    IssueCollector,
    Severity,
    TooManyIssuesError,
    make_issue,
)
from docgen.domain.mapping.context import build_schema, get_path
from docgen.domain.mapping.engine import MappingEngine, transform_context
from docgen.domain.mapping.expressions import ExpressionError
from docgen.domain.naming import PatternError, render_pattern, sanitize_component, sanitize_relative_path
from docgen.domain.profile.models import ErrorPolicy, OutputKind, Profile, UnitKind
from docgen.domain.profile.validator import validate_profile
from docgen.domain.validation.engine import ValidationEngine
from docgen.infrastructure.excel.loader import LoadResult, WorkbookLoader
from docgen.infrastructure.excel.sources import WorkbookOpenError
from docgen.infrastructure.fs.atomic import (
    OutputConflictError,
    OutputLockedError,
    check_writable_dir,
    finalize,
)
from docgen.infrastructure.fs.safety import UnsafeFileError

FATAL_CATEGORIES = {
    Category.CONFIG,
    Category.WORKBOOK,
    Category.SHEET,
    Category.COLUMN,
    Category.HEADER,
    Category.SECURITY,
    Category.IO,
    Category.TEMPLATE,
    Category.PLACEHOLDER,
    Category.EXPRESSION,
}
POOL_THRESHOLD = 20
LOG = logging.getLogger("docgen.job")
WORKER_RECYCLE_UNITS = 50


def _now() -> str:
    return dt.datetime.now().astimezone().isoformat(timespec="seconds")


def sha256_of(path: Path) -> str:
    digest = hashlib.sha256()
    files = sorted(path.glob("*.csv")) if path.is_dir() else [path]
    for f in files:
        with f.open("rb") as fh:
            while chunk := fh.read(1024 * 1024):
                digest.update(chunk)
    return digest.hexdigest()


def snapshot_input(path: Path, directory: Path) -> Path:
    """Copy the input so the job reads a stable file even if Excel keeps it open/changes it."""
    if path.is_dir():
        target = directory / path.name
        target.mkdir(parents=True, exist_ok=True)
        for f in sorted(path.glob("*.csv")):
            shutil.copy2(f, target / f.name)
        return target
    target = directory / path.name
    try:
        shutil.copy2(path, target)
    except PermissionError as exc:
        raise WorkbookOpenError(
            f"{path.name} cannot be read because another program locks it; close it and retry"
        ) from exc
    return target


def default_workers() -> int:
    return max(1, min(4, (os.cpu_count() or 2) - 1))


@dataclass(slots=True)
class JobSession:
    """Everything prepared for a job; also used by the Preview page."""

    profile: Profile
    request: GenerationRequest
    load: LoadResult
    engine: MappingEngine
    templates: list[PreparedTemplate]
    plan: Plan | None
    unit_kind: UnitKind
    output_kind: OutputKind
    report: dict[str, Any]
    snapshot_dir: Path
    issues: list[Issue] = field(default_factory=list)


class _StageTimer:
    def __init__(self, result: JobResult, reporter: ProgressReporter) -> None:
        self.result = result
        self.reporter = reporter
        self._name: str | None = None
        self._start = 0.0

    def __call__(self, name: str) -> None:
        self.stop()
        self._name, self._start = name, time.perf_counter()
        self.reporter.stage(name)

    def stop(self) -> None:
        if self._name:
            self.result.stage_seconds[self._name] = (
                self.result.stage_seconds.get(self._name, 0.0) + time.perf_counter() - self._start
            )
        self._name = None


class DocumentGenerationService:
    def __init__(self, *, pool_threshold: int = POOL_THRESHOLD) -> None:
        self.pool_threshold = pool_threshold

    # ================================================================ run
    def run(
        self,
        request: GenerationRequest,
        reporter: ProgressReporter | None = None,
        cancel: CancellationToken | None = None,
    ) -> JobResult:
        reporter = reporter or NullReporter()
        cancel = cancel or CancellationToken()
        profile = request.profile
        job_id = dt.datetime.now().strftime("%Y%m%d-%H%M%S-") + uuid.uuid4().hex[:6]
        result = JobResult(
            job_id=job_id,
            status="running",
            started=_now(),
            profile_id=profile.profile.id,
            profile_name=profile.profile.name,
            profile_version=profile.profile.version,
            output_dir=str(request.output_dir),
        )
        timer = _StageTimer(result, reporter)
        policy = request.error_policy or profile.generation.error_policy
        collector = IssueCollector(
            fail_fast=policy is ErrorPolicy.FAIL_FAST, max_errors=profile.generation.max_errors
        )
        session: JobSession | None = None
        temp_dir = request.output_dir / f".docgen-tmp-{job_id}"
        # log metadata only (never cell values): ids, counts, durations, file names
        LOG.info(
            "job %s started: profile=%s v%s input=%s dry_run=%s",
            job_id,
            profile.profile.id,
            profile.profile.version,
            request.excel_path.name,
            request.dry_run,
        )
        try:
            session = self.prepare(
                request, reporter, cancel, job_id=job_id, result=result, timer=timer, collector=collector
            )
            if session is None:
                result.status = "failed"
                return result
            if request.dry_run:
                result.status = "validated"
                self._plan_outcomes(session, result)
                return result
            timer("generate")
            self._execute(session, result, reporter, cancel, temp_dir, policy)
            if cancel.cancelled:
                result.status = "cancelled"
            elif any(u.status == STATUS_ERROR for u in result.units):
                result.status = "completed_with_errors"
            else:
                result.status = "completed"
        except JobCancelledError:
            result.status = "cancelled"
        except (FailFastError, TooManyIssuesError) as exc:
            result.status = "failed"
            issue = (
                exc.issue
                if isinstance(exc, FailFastError)
                else make_issue("job.too_many_errors", Severity.ERROR, Category.VALIDATION, str(exc))
            )
            if issue not in result.issues:
                result.issues.append(issue)
        finally:
            timer.stop()
            if session is not None:
                shutil.rmtree(session.snapshot_dir, ignore_errors=True)
            shutil.rmtree(temp_dir, ignore_errors=True)
            result.finished = _now()
            LOG.info(
                "job %s %s: units=%d documents=%d errors=%d warnings=%d stages=%s",
                job_id,
                result.status,
                len(result.units),
                len(result.documents),
                result.error_count,
                result.warning_count,
                {k: round(v, 2) for k, v in result.stage_seconds.items()},
            )
            if profile.output.write_audit and not request.dry_run and result.status != "running":
                try:
                    audit, report = write_audit(result, request.output_dir)
                    result.audit_path, result.report_path = str(audit), str(report)
                except OSError as exc:
                    result.issues.append(
                        make_issue(
                            "output.audit_failed",
                            Severity.WARNING,
                            Category.OUTPUT,
                            f"Audit report could not be written: {exc}",
                        )
                    )
        return result

    # ============================================================ prepare
    def prepare(
        self,
        request: GenerationRequest,
        reporter: ProgressReporter | None = None,
        cancel: CancellationToken | None = None,
        *,
        job_id: str = "preview",
        result: JobResult | None = None,
        timer: Callable[[str], None] | None = None,
        collector: IssueCollector | None = None,
        plan: bool = True,
    ) -> JobSession | None:
        reporter = reporter or NullReporter()
        cancel = cancel or CancellationToken()
        if result is None:
            result = JobResult(job_id, "running", _now())
        if collector is None:
            collector = IssueCollector()
        stage = timer or reporter.stage
        profile = request.profile
        unit_kind = request.unit or profile.generation.unit
        output_kind = request.output or profile.generation.output
        if unit_kind is UnitKind.GROUP and not profile.generation.group_by:
            self._fatal(result, "generation.group_by_missing", "Grouped generation needs group_by fields")
            return None

        def fatal_issues(issues: list[Issue]) -> bool:
            fatal = False
            for issue in issues:
                result.issues.append(issue)
                reporter.issue(issue)
                if issue.severity is Severity.ERROR and issue.category in FATAL_CATEGORIES:
                    fatal = True
            return fatal

        stage("validate")
        if fatal_issues(validate_profile(profile)):
            return None
        if not request.dry_run:
            problem = check_writable_dir(request.output_dir)
            if problem:
                self._fatal(result, "output.not_writable", problem, Category.IO)
                return None

        schema = build_schema(profile, unit_kind)
        prepared: list[PreparedTemplate] = []
        selected = [
            t
            for t in profile.ordered_templates()
            if t.enabled and (request.template_ids is None or t.id in request.template_ids)
        ]
        if not selected and not request.dry_run and plan:
            self._fatal(result, "template.none_selected", "No enabled template selected", Category.TEMPLATE)
            return None
        template_issues: list[Issue] = []
        for spec in selected:
            tpl, issues = prepare_template(profile, request.profile_dir, spec, unit=unit_kind, schema=schema)
            template_issues.extend(issues)
            if tpl is not None:
                prepared.append(tpl)
        if fatal_issues(template_issues):
            return None
        cancel.raise_if_cancelled()

        stage("read_excel")
        snapshot_dir = Path(tempfile.mkdtemp(prefix="docgen-input-"))
        try:
            snap = snapshot_input(request.excel_path, snapshot_dir)
            result.inputs = [
                {
                    "file": str(request.excel_path),
                    "name": request.excel_path.name,
                    "sha256": sha256_of(snap),
                }
            ]
            for t in prepared:
                result.inputs.append(
                    {
                        "template": t.spec.id,
                        "file": str(t.source),
                        "sha256": hashlib.sha256(t.source.read_bytes()).hexdigest(),
                    }
                )
            loader = WorkbookLoader(profile.excel, profile.formatting)
            load = loader.load(
                snap,
                progress=lambda m: reporter.progress(0, 0, m),
                header_overrides=request.header_overrides or None,
            )
        except (WorkbookOpenError, UnsafeFileError, OSError) as exc:
            shutil.rmtree(snapshot_dir, ignore_errors=True)
            self._fatal(result, "workbook.open_failed", str(exc), Category.WORKBOOK)
            return None
        # data is in memory now; the input snapshot is no longer needed
        shutil.rmtree(snapshot_dir, ignore_errors=True)
        if fatal_issues(load.issues) or profile.excel.primary not in load.dataset.tables:
            shutil.rmtree(snapshot_dir, ignore_errors=True)
            if not any(i.severity is Severity.ERROR for i in result.issues):
                self._fatal(
                    result, "sheet.primary_missing", "Primary sheet could not be loaded", Category.SHEET
                )
            return None
        cancel.raise_if_cancelled()

        stage("index")
        report = {
            "generated_at": dt.datetime.now().replace(microsecond=0),
            "generated_date": dt.date.today(),
            "job_id": job_id,
            "source_file": request.excel_path.name,
            "profile": {
                "id": profile.profile.id,
                "name": profile.profile.name,
                "version": profile.profile.version,
            },
        }
        engine = MappingEngine(profile, load.dataset, report=report)
        try:
            index_issues = engine.build_all_indexes()
        except ExpressionError as exc:
            shutil.rmtree(snapshot_dir, ignore_errors=True)
            self._fatal(result, "relationship.filter_failed", str(exc), Category.EXPRESSION)
            return None
        for issue in index_issues:
            result.issues.append(issue)
            reporter.issue(issue)
        session = JobSession(
            profile, request, load, engine, prepared, None, unit_kind, output_kind, report, snapshot_dir
        )
        if not plan:
            return session

        stage("map")
        validation = ValidationEngine(profile.validation, engine.expressions)
        session.plan = plan_units(
            engine,
            unit_kind=unit_kind,
            rows=request.rows,
            validation=validation,
            reporter=reporter,
            cancel=cancel,
        )
        result.records_total = session.plan.records_total
        result.records_selected = session.plan.records_selected
        result.unit_kind, result.output_kind = unit_kind.value, output_kind.value
        result.issues.extend(session.plan.issues)
        for unit in session.plan.units:
            for issue in unit.issues:
                collector.add(issue)  # raises on fail-fast
        return session

    def _fatal(
        self, result: JobResult, code: str, message: str, category: Category = Category.CONFIG
    ) -> None:
        result.issues.append(make_issue(code, Severity.ERROR, category, message))

    def _plan_outcomes(self, session: JobSession, result: JobResult) -> None:
        assert session.plan is not None
        for unit in session.plan.units:
            status = STATUS_ERROR if unit.has_errors else (STATUS_WARNING if unit.issues else STATUS_OK)
            result.units.append(
                UnitOutcome(unit.index, unit.label, unit.excel_rows, status, issues=list(unit.issues))
            )

    # ============================================================ execute
    def unit_context(self, session: JobSession, unit: PlannedUnit, total: int) -> dict[str, Any]:
        engine = session.engine
        if session.unit_kind is UnitKind.RECORD:
            ctx = engine.map_record(unit.row_indices[0]).context
        else:
            records = [engine.map_record(ri) for ri in unit.row_indices]
            names = list(session.profile.generation.group_by) if session.unit_kind is UnitKind.GROUP else []
            ctx = engine.map_group(unit.key if names else (), records, names).context
        ctx = dict(ctx)
        ctx["report"] = session.report
        ctx["unit"] = {"index": unit.index, "count": total, "key": unit.label, "label": unit.label}
        return ctx

    def template_context(self, tpl: PreparedTemplate, ctx: Mapping[str, Any]) -> dict[str, Any]:
        base: dict[str, Any]
        if tpl.spec.context_root:
            root = get_path(ctx, tpl.spec.context_root)
            base = dict(root) if isinstance(root, Mapping) else {"value": root}
            base.setdefault("report", ctx.get("report"))
            base.setdefault("unit", ctx.get("unit"))
        else:
            base = dict(ctx)
        base["b"] = {name: get_path(ctx, path) for name, path in tpl.bindings.items()}
        return base

    def _included(
        self, session: JobSession, tpl: PreparedTemplate, ctx: dict[str, Any], issues: list[Issue]
    ) -> bool:
        if not tpl.spec.include_if:
            return True
        try:
            return bool(session.engine.expressions.compile(tpl.spec.include_if)(ctx))
        except ExpressionError as exc:
            issues.append(
                make_issue(
                    "template.include_if_failed",
                    Severity.ERROR,
                    Category.EXPRESSION,
                    str(exc),
                    template=tpl.spec.id,
                )
            )
            return False

    def _names(
        self, session: JobSession, ctx: dict[str, Any], unit: PlannedUnit, issues: list[Issue], job_id: str
    ) -> tuple[str, Callable[..., str]]:
        profile = session.profile
        values = dict(ctx)
        values["profile"] = {
            "id": profile.profile.id,
            "name": profile.profile.name,
            "version": profile.profile.version,
        }
        values["job"] = {"id": job_id}

        def render(pattern: str, template_id: str | None = None) -> str:
            v = values if template_id is None else {**values, "template": {"id": template_id}}
            try:
                text, missing = render_pattern(pattern, v)
            except PatternError as exc:
                issues.append(make_issue("output.pattern", Severity.WARNING, Category.OUTPUT, str(exc)))
                text, missing = "", []
            if missing:
                issues.append(
                    make_issue(
                        "output.pattern_missing",
                        Severity.WARNING,
                        Category.OUTPUT,
                        f"Filename pattern values are empty: {missing}",
                        suggestion="Use fields that are always filled in output filename patterns.",
                    )
                )
            return text

        subfolder = ""
        if profile.output.subfolder:
            subfolder = sanitize_relative_path(render(profile.output.subfolder))

        def filename(pattern: str, template_id: str | None = None) -> str:
            """Relative output path; '/' in the *pattern* creates sub-folders."""
            rel = sanitize_relative_path(render(pattern, template_id).strip())
            head, _, last = rel.rpartition("/")
            if not last or last.lower() in (".docx", "untitled"):
                last = f"unit-{unit.index}" + (f"-{template_id}" if template_id else "")
            if not last.lower().endswith(".docx"):
                last += ".docx"
            return f"{head}/{last}" if head else last

        return subfolder, filename

    def _target(self, session: JobSession, subfolder: str, name: str, issues: list[Issue]) -> Path:
        base = session.request.output_dir / subfolder if subfolder else session.request.output_dir
        target = base / name
        limit = session.profile.output.max_path_length
        if len(str(target)) > limit:
            base = target.parent
            room = max(20, limit - len(str(base)) - 1)
            short = sanitize_component(target.name, max_length=room)
            issues.append(
                make_issue(
                    "output.path_shortened",
                    Severity.WARNING,
                    Category.OUTPUT,
                    f"Output path longer than {limit} characters; file name shortened to {short}",
                )
            )
            target = base / short
        return target

    def _build_jobs(
        self,
        session: JobSession,
        units: list[PlannedUnit],
        temp_dir: Path,
        outcomes: dict[int, UnitOutcome],
        job_id: str,
    ) -> Iterator[tuple[UnitJob, dict[str, Any]]]:
        profile = session.profile
        tctx = transform_context(profile.formatting)
        want_individual = (
            session.output_kind in (OutputKind.SEPARATE, OutputKind.BOTH) or not profile.merge.enabled
        )
        want_merged = session.output_kind in (OutputKind.MERGED, OutputKind.BOTH) and profile.merge.enabled
        tpl_dir = temp_dir / "templates"
        tpl_dir.mkdir(parents=True, exist_ok=True)
        compiled_paths = {}
        for tpl in session.templates:
            path = tpl_dir / f"{tpl.spec.id}.docx"
            path.write_bytes(tpl.compiled)
            compiled_paths[tpl.spec.id] = str(path)
        image_root = str(session.request.excel_path.parent)
        for unit in units:
            outcome = outcomes[unit.index]
            ctx = self.unit_context(session, unit, len(units))
            issues = outcome.issues
            templates, contexts = [], []
            for tpl in session.templates:
                if not self._included(session, tpl, ctx, issues):
                    continue
                s = tpl.spec.section
                templates.append(
                    TemplateJob(
                        tpl.spec.id,
                        compiled_paths[tpl.spec.id],
                        tpl.spec.missing_placeholder,
                        s.mode,
                        s.break_before,
                        s.start,
                        s.restart_page_numbering,
                        s.inherit_missing_headers,
                    )
                )
                contexts.append(self.template_context(tpl, ctx))
            outcome.templates = [t.template_id for t in templates]
            if not templates:
                outcome.status = STATUS_SKIPPED
                issues.append(
                    make_issue(
                        "unit.no_templates",
                        Severity.WARNING,
                        Category.TEMPLATE,
                        "No template applies to this unit (include_if)",
                    )
                )
                continue
            yield (
                UnitJob(
                    unit.index,
                    str(temp_dir),
                    templates,
                    contexts,
                    want_individual,
                    want_merged,
                    profile.merge.preserve_styles,
                    profile.merge.update_fields_on_open,
                    profile.generation.verify_output,
                    tctx,
                    image_root,
                ),
                ctx,
            )

    def _execute(
        self,
        session: JobSession,
        result: JobResult,
        reporter: ProgressReporter,
        cancel: CancellationToken,
        temp_dir: Path,
        policy: ErrorPolicy,
    ) -> None:
        assert session.plan is not None
        units = session.plan.units
        if session.request.unit_indices:
            wanted = set(session.request.unit_indices)
            units = [u for u in units if u.index in wanted]
        outcomes: dict[int, UnitOutcome] = {}
        runnable: list[PlannedUnit] = []
        for unit in units:
            outcome = UnitOutcome(unit.index, unit.label, unit.excel_rows, issues=list(unit.issues))
            outcomes[unit.index] = outcome
            if unit.has_errors:
                outcome.status = STATUS_ERROR
            else:
                runnable.append(unit)
        result.units = [outcomes[u.index] for u in units]
        temp_dir.mkdir(parents=True, exist_ok=True)
        reserved: set[Path] = set()
        contexts: dict[int, dict[str, Any]] = {}
        total = len(units)
        done = total - len(runnable)
        reporter.progress(done, total, "Generating documents")
        workers = session.request.workers or (
            default_workers()
            if session.profile.generation.workers == "auto"
            else int(session.profile.generation.workers)
        )
        stop = False
        processed: set[int] = set()

        def handle(res: UnitJobResult) -> None:
            nonlocal done, stop
            processed.add(res.index)
            outcome = outcomes[res.index]
            outcome.issues.extend(res.issues)
            outcome.seconds = res.seconds
            if res.failed:
                outcome.status = STATUS_ERROR
            else:
                self._finalize_outputs(
                    session, res, outcome, contexts.pop(res.index, {}), reserved, result.job_id
                )
            for issue in res.issues:
                reporter.issue(issue.with_context(unit=outcome.label))
            done += 1
            reporter.progress(done, total, outcome.label)
            if outcome.status == STATUS_ERROR and policy is ErrorPolicy.FAIL_FAST:
                stop = True

        jobs = self._build_jobs(session, runnable, temp_dir, outcomes, result.job_id)
        # Large jobs always render in worker processes (even with one worker): native memory
        # used by lxml/docxcompose while merging is not fully returned to the OS (~1 MB per merge
        # measured), so workers are recycled every WORKER_RECYCLE_UNITS units and the UI/CLI
        # process stays small. Small jobs render in-process to avoid process start-up cost.
        use_pool = len(runnable) >= self.pool_threshold
        if not use_pool:
            for job, ctx in jobs:
                if cancel.cancelled or stop:
                    break
                contexts[job.index] = ctx
                handle(render_unit(job))
        else:
            mp = multiprocessing.get_context("spawn")
            in_flight: dict[Future[UnitJobResult], int] = {}
            with ProcessPoolExecutor(
                max_workers=workers, mp_context=mp, max_tasks_per_child=WORKER_RECYCLE_UNITS
            ) as pool:
                exhausted = False
                while True:
                    while (
                        not exhausted and not cancel.cancelled and not stop and len(in_flight) < workers * 2
                    ):
                        try:
                            job, ctx = next(jobs)
                        except StopIteration:
                            exhausted = True
                            break
                        contexts[job.index] = ctx
                        in_flight[pool.submit(render_unit, job)] = job.index
                    if not in_flight:
                        break
                    finished, _ = wait(list(in_flight), timeout=0.5, return_when=FIRST_COMPLETED)
                    for fut in finished:
                        index = in_flight.pop(fut)
                        try:
                            handle(fut.result())
                        except Exception as exc:  # worker crashed (e.g. memory)
                            outcomes[index].status = STATUS_ERROR
                            outcomes[index].issues.append(
                                make_issue(
                                    "render.worker_failed",
                                    Severity.ERROR,
                                    Category.INTERNAL,
                                    f"Worker process failed: {type(exc).__name__}: {exc}",
                                )
                            )
                    if (cancel.cancelled or stop) and exhausted is False:
                        exhausted = True
                pool.shutdown(wait=True, cancel_futures=True)
        if cancel.cancelled or stop:
            for unit in runnable:
                outcome = outcomes[unit.index]
                if unit.index not in processed and outcome.status == STATUS_OK:
                    outcome.status = STATUS_CANCELLED if cancel.cancelled else STATUS_SKIPPED

    def _finalize_outputs(
        self,
        session: JobSession,
        res: UnitJobResult,
        outcome: UnitOutcome,
        ctx: dict[str, Any],
        reserved: set[Path],
        job_id: str,
    ) -> None:
        profile = session.profile
        issues = outcome.issues
        unit = next(u for u in session.plan.units if u.index == res.index) if session.plan else None
        assert unit is not None
        subfolder, filename = self._names(session, ctx, unit, issues, job_id)
        pairs: list[tuple[Path, Path]] = []
        if res.merged:
            pairs.append(
                (
                    Path(res.merged),
                    self._target(session, subfolder, filename(profile.output.merged_filename), issues),
                )
            )
        for template_id, path in res.individual:
            name = filename(profile.output.individual_filename, template_id)
            pairs.append((Path(path), self._target(session, subfolder, name, issues)))
        for temp, target in pairs:
            try:
                # temp files live in <output>/.docgen-tmp-<job>, i.e. on the same volume as the
                # target, so finalize() can use an atomic os.replace
                fin = finalize(temp, target, profile.output.on_conflict, reserved=reserved)
            except (OutputConflictError, OutputLockedError, OSError) as exc:
                hint = (
                    "Close the file in Word or choose another folder."
                    if isinstance(exc, (OutputLockedError, PermissionError))
                    else "Change on_conflict (suffix/overwrite) or the output folder."
                    if isinstance(exc, OutputConflictError)
                    else "Check the output folder and the file name pattern."
                )
                issues.append(
                    make_issue(
                        "output.write_failed", Severity.ERROR, Category.OUTPUT, str(exc), suggestion=hint
                    )
                )
                outcome.status = STATUS_ERROR
                continue
            if fin.skipped:
                issues.append(
                    make_issue(
                        "output.skipped_existing",
                        Severity.WARNING,
                        Category.OUTPUT,
                        f"{target.name} exists; skipped (on_conflict: skip)",
                    )
                )
                continue
            assert fin.path is not None
            if fin.renamed_from is not None:
                issues.append(
                    make_issue(
                        "output.renamed",
                        Severity.INFO,
                        Category.OUTPUT,
                        f"{target.name} existed; saved as {fin.path.name}",
                    )
                )
            outcome.outputs.append(str(fin.path))
        if outcome.status != STATUS_ERROR:
            outcome.status = (
                STATUS_WARNING if any(i.severity is Severity.WARNING for i in issues) else STATUS_OK
            )


def load_dataset_for_preview(profile: Profile, excel: Path) -> tuple[DataSet, LoadResult]:
    load = WorkbookLoader(profile.excel, profile.formatting).load(excel)
    return load.dataset, load
