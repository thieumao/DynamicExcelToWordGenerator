"""Render + merge + verify one generation unit.

This module is import-safe for ``multiprocessing`` (spawn on Windows): the job
description is plain picklable data and :func:`render_unit` is top-level.
Outputs are written to the job's temporary folder; the parent process
finalises (moves) them into place.
"""

from __future__ import annotations

import shutil
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from docgen.domain.diagnostics import Category, Issue, Severity, make_issue
from docgen.domain.mapping.transforms import TransformContext
from docgen.domain.profile.models import MissingPolicy
from docgen.infrastructure.word.merger import MergePart, WordMergeEngine
from docgen.infrastructure.word.renderer import RenderError, TemplateRenderer
from docgen.infrastructure.word.verifier import verify_docx


@dataclass(frozen=True, slots=True)
class TemplateJob:
    template_id: str
    compiled_path: str
    missing_policy: MissingPolicy
    section_mode: str
    break_before: str
    section_start: str
    restart_page_numbering: bool
    inherit_missing_headers: bool
    leftover_patterns: tuple[str, ...] = ()
    forbid_mergefields: bool = False


@dataclass(slots=True)
class UnitJob:
    index: int
    temp_dir: str
    templates: list[TemplateJob]
    contexts: list[dict[str, Any]]
    """One context per template (bindings / context_root already applied)."""
    want_individual: bool
    want_merged: bool
    preserve_styles: bool
    update_fields_on_open: bool
    verify: bool
    tctx: TransformContext
    image_root: str | None


@dataclass(slots=True)
class UnitJobResult:
    index: int
    individual: list[tuple[str, str]] = field(default_factory=list)
    """(template id, temp path)"""
    merged: str | None = None
    issues: list[Issue] = field(default_factory=list)
    seconds: float = 0.0
    failed: bool = False


_RENDERERS: dict[tuple[Any, ...], TemplateRenderer] = {}
_TEMPLATE_BYTES: dict[str, bytes] = {}


def _renderer(tctx: TransformContext, image_root: str | None) -> TemplateRenderer:
    key = (tctx, image_root)
    renderer = _RENDERERS.get(key)
    if renderer is None:
        renderer = TemplateRenderer(tctx, image_root=Path(image_root) if image_root else None)
        _RENDERERS[key] = renderer
    return renderer


def _template_bytes(path: str) -> bytes:
    data = _TEMPLATE_BYTES.get(path)
    if data is None:
        data = Path(path).read_bytes()
        _TEMPLATE_BYTES[path] = data
    return data


def render_unit(job: UnitJob) -> UnitJobResult:
    start = time.perf_counter()
    result = UnitJobResult(job.index)
    renderer = _renderer(job.tctx, job.image_root)
    unit_dir = Path(job.temp_dir) / f"unit-{job.index:06d}"
    unit_dir.mkdir(parents=True, exist_ok=True)
    rendered: list[tuple[TemplateJob, Path]] = []
    for tjob, ctx in zip(job.templates, job.contexts, strict=True):
        out = unit_dir / f"{tjob.template_id}.docx"
        try:
            res = renderer.render(
                _template_bytes(tjob.compiled_path),
                ctx,
                out,
                template_id=tjob.template_id,
                missing_policy=tjob.missing_policy,
            )
            result.issues.extend(res.issues)
        except RenderError as exc:
            result.issues.extend(exc.issues)
            result.failed = True
            continue
        except Exception as exc:  # unexpected library failure: report, never crash the job
            result.issues.append(
                make_issue(
                    "render.unexpected",
                    Severity.ERROR,
                    Category.INTERNAL,
                    f"Unexpected error while rendering: {type(exc).__name__}: {exc}",
                    template=tjob.template_id,
                )
            )
            result.failed = True
            continue
        if job.verify and not job.want_merged:
            _verify(out, result, tjob.template_id, (tjob,))
        rendered.append((tjob, out))
    if result.failed:
        result.seconds = time.perf_counter() - start
        return result
    if job.want_individual:
        result.individual = [(t.template_id, str(p)) for t, p in rendered]
    if job.want_merged and rendered:
        merged = unit_dir / "__merged__.docx"
        if len(rendered) == 1:
            shutil.copyfile(rendered[0][1], merged)
        else:
            parts = [
                MergePart(
                    p,
                    t.template_id,
                    section_mode=t.section_mode,  # type: ignore[arg-type]
                    break_before=t.break_before,  # type: ignore[arg-type]
                    section_start=t.section_start,  # type: ignore[arg-type]
                    restart_page_numbering=t.restart_page_numbering,
                    inherit_missing_headers=t.inherit_missing_headers,
                )
                for t, p in rendered
            ]
            try:
                report = WordMergeEngine(
                    preserve_styles=job.preserve_styles,
                    update_fields_on_open=job.update_fields_on_open,
                ).merge(parts, merged)
                for warning in report.warnings:
                    result.issues.append(
                        make_issue("merge.warning", Severity.WARNING, Category.MERGE, warning)
                    )
            except Exception as exc:
                result.issues.append(
                    make_issue(
                        "merge.failed",
                        Severity.ERROR,
                        Category.MERGE,
                        f"Merging failed: {type(exc).__name__}: {exc}",
                        suggestion="Generate separate documents, or report the template combination.",
                    )
                )
                result.failed = True
                result.seconds = time.perf_counter() - start
                return result
        if job.verify:
            _verify(merged, result, None, tuple(t for t, _ in rendered))
        result.merged = str(merged)
        if job.verify and job.want_individual:
            by_id = {t.template_id: t for t, _ in rendered}
            for tid, path in result.individual:
                _verify(Path(path), result, tid, (by_id[tid],))
    result.seconds = time.perf_counter() - start
    return result


def _verify(
    path: Path, result: UnitJobResult, template_id: str | None, templates: tuple[TemplateJob, ...]
) -> None:
    extra = tuple(dict.fromkeys(p for t in templates for p in t.leftover_patterns))
    check = verify_docx(
        path, extra_patterns=extra, forbid_mergefields=any(t.forbid_mergefields for t in templates)
    )
    for problem in check.problems:
        result.issues.append(
            make_issue(
                "output.verification_failed",
                Severity.ERROR,
                Category.OUTPUT,
                problem,
                template=template_id,
                suggestion="The document was not saved. Check the template for unsupported content.",
            )
        )
        result.failed = True
    for warning in check.warnings:
        result.issues.append(
            make_issue(
                "output.verification_warning",
                Severity.WARNING,
                Category.OUTPUT,
                warning,
                template=template_id,
            )
        )
