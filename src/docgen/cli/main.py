"""Command-line interface (automation, tests, CI smoke tests on Windows).

Exit codes: 0 = success, 1 = completed with errors / validation errors, 2 = usage error.
"""

from __future__ import annotations

import argparse
import json
import sys
from collections.abc import Sequence
from pathlib import Path

from docgen import __version__
from docgen.domain.diagnostics import Issue, Severity


def _configure_stdout() -> None:
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            try:
                reconfigure(encoding="utf-8", errors="replace")
            except (ValueError, OSError):
                pass


def parse_rows(text: str | None) -> list[int] | None:
    """'5,7-9' -> [5, 7, 8, 9]"""
    if not text:
        return None
    rows: list[int] = []
    for part in text.split(","):
        part = part.strip()
        if "-" in part:
            a, b = part.split("-", 1)
            rows.extend(range(int(a), int(b) + 1))
        elif part:
            rows.append(int(part))
    return rows


def _print_issues(
    issues: Sequence[Issue], *, min_severity: Severity = Severity.INFO, limit: int = 200
) -> None:
    shown = [i for i in issues if i.severity.rank >= min_severity.rank]
    for issue in shown[:limit]:
        print(f"  {issue}")
    if len(shown) > limit:
        print(f"  ... {len(shown) - limit} more")


# ------------------------------------------------------------------ commands
def cmd_inspect(args: argparse.Namespace) -> int:
    from docgen.application.workbook import inspect_workbook

    summary = inspect_workbook(Path(args.excel), preview_rows=args.rows)
    print(f"{summary.path.name} ({summary.format}) – {len(summary.sheets)} sheet(s)")
    for sheet in summary.sheets:
        hidden = " [hidden]" if sheet.hidden else ""
        print(f"\n# {sheet.index}. {sheet.name}{hidden} rows≈{sheet.max_row}")
        print(f"  header row: {sheet.header_row} (confidence {sheet.header_confidence:.0%})")
        print(f"  headers: {sheet.headers}")
    return 0


def cmd_templates_scan(args: argparse.Namespace) -> int:
    from docgen.domain.profile.models import MarkerConfig, MarkerStyle
    from docgen.infrastructure.word.inspector import inspect_template

    marker = MarkerConfig(style=MarkerStyle(args.marker), open=args.open, close=args.close)
    result = inspect_template(str(args.docx), marker)
    print(f"{args.docx}: parts={result.parts}")
    for ref in result.refs:
        print(f"  {ref.kind.value:9} {ref.path}  ({ref.part})")
    for name in result.marker_names:
        print(f"  marker    {name}")
    for err in result.errors:
        print(f"  ERROR {err}")
    for w in result.warnings:
        print(f"  WARNING {w}")
    return 1 if result.errors else 0


def cmd_profile_validate(args: argparse.Namespace) -> int:
    from docgen.domain.profile.io import ProfileLoadError, load_profile
    from docgen.domain.profile.validator import validate_profile

    path = Path(args.profile)
    try:
        profile = load_profile(path)
    except ProfileLoadError as exc:
        print(f"INVALID: {exc}")
        _print_issues(exc.issues)
        return 1
    issues = validate_profile(profile)
    if args.excel:
        from docgen.application.preview import preview

        result = preview(profile, path.parent, Path(args.excel), limit=0, with_plan=False)
        issues = result.issues
        for stats in result.relationships:
            print(
                f"  relationship {stats.relationship} -> {stats.target_sheet}: "
                f"{stats.keys} keys, {stats.duplicate_keys} duplicated {stats.sample_duplicates}"
            )
    errors = [i for i in issues if i.severity is Severity.ERROR]
    print(
        f"{profile.profile.id} v{profile.profile.version}: {len(errors)} error(s), "
        f"{sum(1 for i in issues if i.severity is Severity.WARNING)} warning(s)"
    )
    _print_issues(issues)
    return 1 if errors else 0


def cmd_profile_schema(args: argparse.Namespace) -> int:
    from docgen.domain.profile.io import profile_json_schema

    text = json.dumps(profile_json_schema(), ensure_ascii=False, indent=2)
    if args.output:
        Path(args.output).write_text(text, encoding="utf-8")
    else:
        print(text)
    return 0


def cmd_preview(args: argparse.Namespace) -> int:
    from docgen.application.preview import preview
    from docgen.domain.profile.io import load_profile

    path = Path(args.profile)
    profile = load_profile(path)
    result = preview(
        profile,
        path.parent,
        Path(args.excel),
        rows=parse_rows(args.rows),
        limit=args.limit,
        with_plan=args.plan,
    )
    if not result.ok:
        print("Preview failed:")
        _print_issues(result.issues)
        return 1
    if args.json:
        payload = [{"row": r.excel_row, "context": r.context} for r in result.records]
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return 0
    for rec in result.records:
        print(f"\n## Row {rec.excel_row}")
        print(json.dumps(rec.context, ensure_ascii=False, indent=2))
        if args.trace:
            for t in rec.traces:
                print(f"   trace: {t.describe()}")
        _print_issues(rec.issues)
    if result.plan:
        print(f"\n## Plan: {len(result.plan)} unit(s)")
        for index, label, rows, names in result.plan[:50]:
            print(f"  {index:>4}. {label} rows={rows} -> {names}")
    _print_issues(result.issues, min_severity=Severity.WARNING)
    return 0


class _ConsoleReporter:
    def __init__(self, quiet: bool) -> None:
        self.quiet = quiet
        self._last = -1

    def stage(self, name: str) -> None:
        if not self.quiet:
            print(f"[{name}]", flush=True)

    def progress(self, done: int, total: int, message: str = "") -> None:
        if self.quiet or total <= 0:
            return
        pct = int(done * 100 / total)
        if pct // 10 != self._last // 10 or done == total:
            self._last = pct
            print(f"  {done}/{total} ({pct}%) {message}", flush=True)

    def issue(self, issue: Issue) -> None:
        pass


def cmd_generate(args: argparse.Namespace) -> int:
    from docgen.application.generation.job import CancellationToken, GenerationRequest
    from docgen.application.generation.service import DocumentGenerationService
    from docgen.domain.profile.io import load_profile
    from docgen.domain.profile.models import ErrorPolicy, OutputKind, UnitKind

    path = Path(args.profile)
    profile = load_profile(path)
    request = GenerationRequest(
        profile=profile,
        profile_dir=path.parent,
        excel_path=Path(args.excel),
        output_dir=Path(args.out),
        unit=UnitKind(args.unit) if args.unit else None,
        output=OutputKind(args.output) if args.output else None,
        rows=parse_rows(args.rows),
        template_ids=args.templates.split(",") if args.templates else None,
        workers=args.workers,
        error_policy=ErrorPolicy.FAIL_FAST if args.fail_fast else None,
        dry_run=args.dry_run,
    )
    cancel = CancellationToken()
    try:
        result = DocumentGenerationService().run(request, _ConsoleReporter(args.quiet), cancel)
    except KeyboardInterrupt:
        cancel.cancel()
        print("Cancelled")
        return 1
    print(
        f"\nStatus: {result.status}  units={len(result.units)}  documents={len(result.documents)}  "
        f"errors={result.error_count}  warnings={result.warning_count}"
    )
    _print_issues(result.issues, min_severity=Severity.WARNING)
    for unit in result.units:
        if unit.status in ("error", "cancelled", "skipped") or args.verbose:
            print(f"  unit {unit.index} {unit.label}: {unit.status}")
            _print_issues(unit.issues, min_severity=Severity.WARNING, limit=10)
    if result.audit_path:
        print(f"Audit: {result.audit_path}")
    return 0 if result.status in ("completed", "validated") else 1


def cmd_samples(args: argparse.Namespace) -> int:
    from docgen.testing.samples import write_customer_a, write_customer_b

    out = Path(args.directory)
    a = write_customer_a(out / "customer_a", n_employees=args.employees)
    b = write_customer_b(out / "customer_b")
    print(f"customer_a: {a[0]}  {a[1]}")
    print(f"customer_b: {b[0]}  {b[1]}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="docgen", description="Dynamic Excel-to-Word document generator")
    parser.add_argument("--version", action="version", version=f"docgen {__version__}")
    sub = parser.add_subparsers(dest="command", required=True)

    p = sub.add_parser("inspect", help="list sheets, detected header rows and headers")
    p.add_argument("excel")
    p.add_argument("--rows", type=int, default=30)
    p.set_defaults(func=cmd_inspect)

    tp = sub.add_parser("templates", help="template utilities").add_subparsers(dest="sub", required=True)
    p = tp.add_parser("scan", help="list placeholders in a Word template")
    p.add_argument("docx")
    p.add_argument("--marker", choices=["jinja", "custom", "mergefield"], default="jinja")
    p.add_argument("--open", default="<<")
    p.add_argument("--close", default=">>")
    p.set_defaults(func=cmd_templates_scan)

    pp = sub.add_parser("profile", help="profile utilities").add_subparsers(dest="sub", required=True)
    p = pp.add_parser("validate", help="validate a profile (optionally against a workbook)")
    p.add_argument("profile")
    p.add_argument("--excel")
    p.set_defaults(func=cmd_profile_validate)
    p = pp.add_parser("schema", help="print the profile JSON schema")
    p.add_argument("-o", "--output")
    p.set_defaults(func=cmd_profile_schema)

    p = sub.add_parser("preview", help="show resolved data and traces for some records")
    p.add_argument("--profile", required=True)
    p.add_argument("--excel", required=True)
    p.add_argument("--rows", help="Excel row numbers, e.g. 5,7-9")
    p.add_argument("--limit", type=int, default=3)
    p.add_argument("--trace", action="store_true")
    p.add_argument("--json", action="store_true")
    p.add_argument("--no-plan", dest="plan", action="store_false")
    p.set_defaults(func=cmd_preview)

    p = sub.add_parser("generate", help="generate Word documents")
    p.add_argument("--profile", required=True)
    p.add_argument("--excel", required=True)
    p.add_argument("--out", required=True)
    p.add_argument("--unit", choices=["record", "group", "dataset"])
    p.add_argument("--output", choices=["separate", "merged", "both"])
    p.add_argument("--rows", help="Excel row numbers of the primary sheet, e.g. 5,7-9")
    p.add_argument("--templates", help="comma separated template ids")
    p.add_argument("--workers", type=int)
    p.add_argument("--fail-fast", action="store_true")
    p.add_argument("--dry-run", action="store_true", help="validate and plan only")
    p.add_argument("--quiet", action="store_true")
    p.add_argument("--verbose", action="store_true")
    p.set_defaults(func=cmd_generate)

    p = sub.add_parser("samples", help="write the sample customers (profiles, workbooks, templates)")
    p.add_argument("directory")
    p.add_argument("--employees", type=int, default=8)
    p.set_defaults(func=cmd_samples)
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    import multiprocessing

    multiprocessing.freeze_support()
    _configure_stdout()
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        code: int = args.func(args)
    except Exception as exc:  # report cleanly instead of a traceback for users
        from docgen.domain.profile.io import ProfileLoadError

        if isinstance(exc, (ProfileLoadError, FileNotFoundError, OSError, ValueError)):
            print(f"Error: {exc}", file=sys.stderr)
            return 1
        raise
    return code


if __name__ == "__main__":
    sys.exit(main())
