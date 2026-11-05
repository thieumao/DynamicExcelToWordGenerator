"""Audit report (JSON) and per-unit processing report (CSV, Excel-friendly UTF-8 BOM)."""

from __future__ import annotations

import csv
import json
from pathlib import Path
from typing import Any

from docgen import __version__
from docgen.application.generation.job import JobResult
from docgen.domain.diagnostics import Severity
from docgen.i18n.catalog import issue_text


def audit_dict(result: JobResult) -> dict[str, Any]:
    issues = result.all_issues()
    return {
        "application": {"name": "DocGen", "version": __version__},
        "job_id": result.job_id,
        "status": result.status,
        "started": result.started,
        "finished": result.finished,
        "profile": {
            "id": result.profile_id,
            "name": result.profile_name,
            "version": result.profile_version,
        },
        "inputs": result.inputs,
        "mode": {"unit": result.unit_kind, "output": result.output_kind},
        "output_dir": result.output_dir,
        "records": {"total": result.records_total, "selected": result.records_selected},
        "units": {
            "total": len(result.units),
            "ok": sum(1 for u in result.units if u.status in ("ok", "warning")),
            "failed": sum(1 for u in result.units if u.status == "error"),
            "skipped": sum(1 for u in result.units if u.status == "skipped"),
            "cancelled": sum(1 for u in result.units if u.status == "cancelled"),
        },
        "documents_generated": len(result.documents),
        "outputs": result.documents,
        "combined_output": result.combined_path,
        "errors": sum(1 for i in issues if i.severity is Severity.ERROR),
        "warnings": sum(1 for i in issues if i.severity is Severity.WARNING),
        "stage_seconds": {k: round(v, 3) for k, v in result.stage_seconds.items()},
        "job_issues": [i.to_dict() for i in result.issues[:500]],
        "unit_details": [u.to_dict() for u in result.units],
    }


def write_audit(result: JobResult, directory: Path) -> tuple[Path, Path]:
    directory.mkdir(parents=True, exist_ok=True)
    audit_path = directory / f"docgen-audit-{result.job_id}.json"
    report_path = directory / f"docgen-report-{result.job_id}.csv"
    tmp = audit_path.with_suffix(".json.tmp")
    tmp.write_text(
        json.dumps(audit_dict(result), ensure_ascii=False, indent=2, default=str), encoding="utf-8"
    )
    tmp.replace(audit_path)
    tmp_csv = report_path.with_suffix(".csv.tmp")
    with tmp_csv.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(
            [
                "unit",
                "label",
                "excel_rows",
                "status",
                "errors",
                "warnings",
                "outputs",
                "first_error",
                "suggestion",
            ]
        )
        for u in result.units:
            errors = [i for i in u.issues if i.severity is Severity.ERROR]
            writer.writerow(
                [
                    u.index,
                    u.label,
                    " ".join(str(r) for r in u.rows[:50]),
                    u.status,
                    len(errors),
                    sum(1 for i in u.issues if i.severity is Severity.WARNING),
                    " | ".join(u.outputs),
                    issue_text(errors[0])[0] if errors else "",
                    issue_text(errors[0])[1] if errors else "",
                ]
            )
    tmp_csv.replace(report_path)
    return audit_path, report_path
