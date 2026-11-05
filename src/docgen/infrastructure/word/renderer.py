"""Render one compiled Word template with one data context (docxtpl).

* sandboxed Jinja environment with autoescape (values containing ``<`` or
  ``&`` cannot break the XML),
* default output formatting for dates/decimals/booleans (profile locale),
* image references resolved safely to inline pictures,
* undefined placeholders tracked and handled by the template's policy.
"""

from __future__ import annotations

import datetime as dt
import io
from collections.abc import Mapping
from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path
from typing import Any

from docx.image.exceptions import InvalidImageStreamError, UnexpectedEndOfFileError, UnrecognizedImageError
from docx.image.image import Image as DocxImage
from docx.shared import Emu, Mm
from docxtpl import DocxTemplate, InlineImage
from jinja2.exceptions import TemplateError

from docgen.domain.data import values as V
from docgen.domain.diagnostics import Category, Issue, Severity, make_issue
from docgen.domain.mapping.engine import ImageRef
from docgen.domain.mapping.expressions import ExpressionError, make_environment
from docgen.domain.mapping.transforms import DEFAULT_REGISTRY, TransformContext, format_decimal
from docgen.domain.profile.models import MissingPolicy
from docgen.infrastructure.fs.safety import is_within
from docgen.infrastructure.word.spots import FORM_FILTERS

IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".gif", ".bmp"}


class RenderError(RuntimeError):
    def __init__(self, message: str, issues: list[Issue]) -> None:
        super().__init__(message)
        self.issues = issues


@dataclass(slots=True)
class RenderResult:
    output: Path
    issues: list[Issue] = field(default_factory=list)
    missing: list[str] = field(default_factory=list)


def _finalizer(tctx: TransformContext) -> Any:
    def finalize(value: Any) -> Any:
        if value is None:
            return ""
        if isinstance(value, bool):
            return tctx.true_text if value else tctx.false_text
        if isinstance(value, dt.datetime):
            if value.time() == dt.time():
                return value.strftime(tctx.date_format)
            return value.strftime(tctx.datetime_format)
        if isinstance(value, dt.date):
            return value.strftime(tctx.date_format)
        if isinstance(value, Decimal):
            decimals = max(0, -value.normalize().as_tuple().exponent)  # type: ignore[operator]
            return format_decimal(value, min(decimals, 6), tctx.decimal_separator, "")
        if isinstance(value, float):
            return V.number_to_text(value).replace(".", tctx.decimal_separator)
        return value

    return finalize


def picture_fit(value: Any, cx: int, cy: int) -> Any:
    """Fit an image placed in a picture content control into the control's frame (EMU).

    The aspect ratio is kept. An explicit width/height on the image field wins; values that
    are not images (empty or text) are returned unchanged.
    """
    if not isinstance(value, InlineImage) or value.width or value.height or cx <= 0 or cy <= 0:
        return value
    try:
        image = DocxImage.from_file(value.image_descriptor)
    except (OSError, UnrecognizedImageError, InvalidImageStreamError, UnexpectedEndOfFileError):
        return value
    if not image.width or not image.height:
        return value
    scale = min(cx / image.width, cy / image.height)
    return InlineImage(
        value.tpl,
        image_descriptor=value.image_descriptor,
        width=Emu(round(image.width * scale)),
        height=Emu(round(image.height * scale)),
    )


class TemplateRenderer:
    """Reusable renderer; create one per worker thread/process."""

    def __init__(self, tctx: TransformContext | None = None, *, image_root: Path | None = None) -> None:
        self.tctx = tctx or TransformContext()
        self.image_root = image_root
        self._missing: list[str] = []
        self.env = make_environment(
            registry=DEFAULT_REGISTRY, ctx=self.tctx, sink=self._missing, autoescape=True
        )
        self.env.finalize = _finalizer(self.tctx)
        self.env.filters.update(FORM_FILTERS)
        self.env.filters["picture_fit"] = picture_fit

    def render(
        self,
        template_bytes: bytes,
        context: Mapping[str, Any],
        output: Path,
        *,
        template_id: str = "",
        missing_policy: MissingPolicy = MissingPolicy.ERROR,
    ) -> RenderResult:
        issues: list[Issue] = []
        tpl = DocxTemplate(io.BytesIO(template_bytes))
        prepared = self._prepare(context, tpl, issues, template_id)
        self._missing.clear()
        try:
            tpl.render(prepared, jinja_env=self.env, autoescape=True)
        except ExpressionError as exc:  # a value that cannot be placed (e.g. not a drop-down choice)
            issue = make_issue(
                "render.value_error",
                Severity.ERROR,
                Category.RENDER,
                f"Cannot fill the document: {exc}",
                suggestion="Check the value in Excel, or convert it on the field with a "
                "'map_values' transform.",
                template=template_id,
                detail=str(exc),
            )
            raise RenderError(str(issue), [*issues, issue]) from exc
        except TemplateError as exc:
            lines = getattr(exc, "docx_context", None)
            near = " | ".join(x.strip() for x in lines if x.strip())[:200] if lines else ""
            issue = make_issue(
                "render.template_error",
                Severity.ERROR,
                Category.RENDER,
                f"Template error: {exc}" + (f" near '{near}'" if near else ""),
                suggestion="Check the Jinja tags in the Word template (matching endif/endfor, "
                "{%tr %} in table rows, {%p %} alone in a paragraph).",
                template=template_id,
                detail=str(exc) + (f" – '{near}'" if near else ""),
            )
            raise RenderError(str(issue), [*issues, issue]) from exc
        missing = sorted(set(self._missing))
        if missing and missing_policy is not MissingPolicy.EMPTY:
            severity = Severity.ERROR if missing_policy is MissingPolicy.ERROR else Severity.WARNING
            for name in missing:
                issues.append(
                    make_issue(
                        "render.missing_placeholder",
                        severity,
                        Category.PLACEHOLDER,
                        f"Placeholder '{name}' has no value in the data context",
                        suggestion="Bind the placeholder to a field on the Templates page, or set "
                        "missing_placeholder: warn/empty for optional content.",
                        template=template_id,
                        placeholder=name,
                    )
                )
            if severity is Severity.ERROR:
                raise RenderError(f"Missing placeholders: {missing}", issues)
        output.parent.mkdir(parents=True, exist_ok=True)
        tpl.save(str(output))
        return RenderResult(output, issues, missing)

    # ------------------------------------------------------------------
    def _prepare(self, value: Any, tpl: DocxTemplate, issues: list[Issue], template_id: str) -> Any:
        if isinstance(value, Mapping):
            return {k: self._prepare(v, tpl, issues, template_id) for k, v in value.items()}
        if isinstance(value, list):
            return [self._prepare(v, tpl, issues, template_id) for v in value]
        if isinstance(value, str):
            return V.sanitize_text(value)
        if isinstance(value, ImageRef):
            return self._image(value, tpl, issues, template_id)
        return value

    def _image(self, ref: ImageRef, tpl: DocxTemplate, issues: list[Issue], template_id: str) -> Any:
        base = Path(ref.base_dir) if ref.base_dir else self.image_root
        path = Path(ref.path)
        if not path.is_absolute() and base is not None:
            path = base / path
        problem = None
        if base is not None and not is_within(path, base):
            problem = f"image path {ref.path!r} points outside the allowed folder {base}"
        elif path.suffix.lower() not in IMAGE_SUFFIXES:
            problem = f"unsupported image type {path.suffix!r} for {ref.path!r}"
        elif not path.is_file():
            problem = f"image file not found: {path}"
        elif path.stat().st_size > ref.max_bytes:
            problem = f"image {path.name} is larger than {ref.max_bytes // 1024} KB"
        if problem:
            issues.append(
                make_issue(
                    "render.image",
                    Severity.WARNING,
                    Category.RENDER,
                    problem,
                    suggestion="Check the image path in Excel and the field's image.base_dir.",
                    template=template_id,
                    detail=problem,
                )
            )
            return ""
        return InlineImage(
            tpl,
            image_descriptor=str(path),
            width=Mm(ref.width_mm) if ref.width_mm else None,
            height=Mm(ref.height_mm) if ref.height_mm else None,
        )
