"""Template preparation: compile markers, inspect placeholders, resolve bindings."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from docgen.domain.data.values import normalize_label
from docgen.domain.diagnostics import Category, Issue, Severity, make_issue
from docgen.domain.mapping.context import ContextSchema, build_schema
from docgen.domain.profile.models import MissingPolicy, Profile, TemplateSpec, UnitKind
from docgen.infrastructure.fs.safety import UnsafeFileError, ZipLimits, check_zip
from docgen.infrastructure.word.inspector import (
    TemplateInspection,
    compile_template,
    inspect_compiled,
)


@dataclass(slots=True)
class PreparedTemplate:
    spec: TemplateSpec
    source: Path
    compiled: bytes
    inspection: TemplateInspection
    bindings: dict[str, str] = field(default_factory=dict)
    """placeholder name -> context path (explicit or auto-resolved)."""
    unbound: list[str] = field(default_factory=list)
    unknown_paths: list[str] = field(default_factory=list)


def resolve_template_path(profile_dir: Path, spec: TemplateSpec) -> Path:
    path = Path(spec.file)
    return path if path.is_absolute() else (profile_dir / path)


def auto_bind(name: str, schema: ContextSchema) -> str | None:
    """Map a mail-merge style name ('Họ tên', 'employee_code') to a context path."""
    candidates = sorted(p for p in schema.paths if "[]" not in p)
    if name in schema.paths:
        return name

    def norm(text: str) -> str:
        return normalize_label(text, accents=True).replace("_", "").replace(" ", "").replace(".", "")

    target = norm(name)
    full = [p for p in candidates if norm(p) == target]
    if len(full) == 1:
        return full[0]
    last = [p for p in candidates if norm(p.rsplit(".", 1)[-1]) == target]
    if len(last) == 1:
        return last[0]
    return None


def prepare_template(
    profile: Profile,
    profile_dir: Path,
    spec: TemplateSpec,
    *,
    unit: UnitKind | None = None,
    schema: ContextSchema | None = None,
) -> tuple[PreparedTemplate | None, list[Issue]]:
    issues: list[Issue] = []
    path = resolve_template_path(profile_dir, spec)
    if not path.is_file():
        issues.append(
            make_issue(
                "template.not_found",
                Severity.ERROR,
                Category.TEMPLATE,
                f"Template file not found: {path}",
                suggestion="Re-import the template on the Templates page or fix 'file'.",
                template=spec.id,
            )
        )
        return None, issues
    try:
        check_zip(path, ZipLimits(max_file_bytes=200 * 1024 * 1024))
        data = path.read_bytes()
        converted = compile_template(data, spec.marker)
        inspection = inspect_compiled(converted.data)
    except (UnsafeFileError, OSError, ValueError) as exc:
        issues.append(
            make_issue(
                "template.unreadable",
                Severity.ERROR,
                Category.TEMPLATE,
                f"Cannot read template {path.name}: {exc}",
                template=spec.id,
            )
        )
        return None, issues
    inspection.marker_names = converted.names
    inspection.warnings.extend(converted.problems)
    for err in inspection.errors:
        issues.append(
            make_issue(
                "template.syntax",
                Severity.ERROR,
                Category.TEMPLATE,
                err,
                template=spec.id,
                suggestion="Fix the tag in Word; every {% if %}/{% for %} needs its end tag.",
            )
        )
    for warning in inspection.warnings:
        issues.append(
            make_issue("template.marker", Severity.WARNING, Category.TEMPLATE, warning, template=spec.id)
        )

    if not inspection.refs and not inspection.errors:
        issues.append(
            make_issue(
                "template.no_placeholders",
                Severity.WARNING,
                Category.TEMPLATE,
                f"Template '{spec.id}' contains no placeholders for marker style "
                f"'{spec.marker.style.value}'; it will be copied without data",
                suggestion="Check the marker style (jinja / custom / mergefield) and delimiters.",
                template=spec.id,
            )
        )
    schema = schema or build_schema(profile, unit)
    prepared = PreparedTemplate(spec, path, converted.data, inspection)
    severity = {
        MissingPolicy.ERROR: Severity.ERROR,
        MissingPolicy.WARN: Severity.WARNING,
        MissingPolicy.EMPTY: Severity.INFO,
    }[spec.missing_placeholder]

    # placeholder bindings (custom markers / MERGEFIELD / explicit b["..."])
    for name in inspection.binding_names:
        if name in spec.bindings:
            prepared.bindings[name] = spec.bindings[name]
            if not schema.knows(spec.bindings[name]):
                issues.append(
                    make_issue(
                        "template.binding_unknown_path",
                        severity,
                        Category.PLACEHOLDER,
                        f"Placeholder '{name}' is bound to '{spec.bindings[name]}', which the "
                        "profile does not produce",
                        template=spec.id,
                        placeholder=name,
                        suggestion=f"Did you mean {schema.suggestions(spec.bindings[name])}?",
                    )
                )
            continue
        guess = auto_bind(name, schema)
        if guess is not None:
            prepared.bindings[name] = guess
            issues.append(
                make_issue(
                    "template.auto_bound",
                    Severity.INFO,
                    Category.PLACEHOLDER,
                    f"Placeholder '{name}' was matched to field '{guess}'",
                    template=spec.id,
                    placeholder=name,
                )
            )
            continue
        prepared.unbound.append(name)
        issues.append(
            make_issue(
                "template.unbound_placeholder",
                severity,
                Category.PLACEHOLDER,
                f"Placeholder '{name}' is not bound to any field",
                suggestion="Bind it on the Templates page (bindings) or rename it to a field path.",
                template=spec.id,
                placeholder=name,
            )
        )
    # direct Jinja paths
    root = spec.context_root
    for path_ref in inspection.paths:
        full = f"{root}.{path_ref}" if root and path_ref.split(".")[0] not in ("report", "unit") else path_ref
        if not schema.knows(full):
            prepared.unknown_paths.append(path_ref)
            issues.append(
                make_issue(
                    "template.unknown_placeholder",
                    severity,
                    Category.PLACEHOLDER,
                    f"Placeholder '{path_ref}' is not produced by the profile",
                    suggestion=(
                        f"Did you mean {schema.suggestions(full)}?"
                        if schema.suggestions(full)
                        else "Add a field with this name or fix the placeholder."
                    ),
                    template=spec.id,
                    placeholder=path_ref,
                )
            )
    return prepared, issues
