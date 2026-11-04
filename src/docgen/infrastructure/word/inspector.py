"""Word template inspection: placeholders, loops and conditions with full paths.

The docx parts are pre-processed with docxtpl's own XML patching (so split
runs and ``{%tr %}``/``{%p %}`` tags behave exactly as during rendering), then
parsed by Jinja; the AST is walked to collect dotted references, resolving
loop variables (``{% for p in employee.projects %}{{ p.name }}`` ->
``employee.projects[].name``).
"""

from __future__ import annotations

import io
import re
import zipfile
from dataclasses import dataclass, field
from enum import StrEnum

from docxtpl import DocxTemplate
from jinja2 import nodes
from jinja2.exceptions import TemplateSyntaxError
from lxml import etree

from docgen.domain.mapping.expressions import SAFE_FUNCTIONS, dotted_chain, make_environment
from docgen.domain.profile.models import MarkerConfig, MarkerStyle
from docgen.infrastructure.fs.safety import ZipLimits, check_zip
from docgen.infrastructure.word.markers import (
    ConversionResult,
    convert_custom_markers,
    convert_mergefields,
)

_PARTS = re.compile(r"^word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$")
_PARSER = etree.XMLParser(resolve_entities=False, no_network=True)
_JINJA_GLOBALS = {
    "loop",
    "range",
    "true",
    "false",
    "none",
    "True",
    "False",
    "None",
    "lipsum",
    "dict",
    "cycler",
    "joiner",
    "namespace",
    "caller",
    "varargs",
    "kwargs",
}


class RefKind(StrEnum):
    VALUE = "value"
    LOOP = "loop"
    CONDITION = "condition"


@dataclass(frozen=True, slots=True)
class PlaceholderRef:
    path: str
    kind: RefKind
    part: str

    @property
    def binding_name(self) -> str | None:
        """For ``b.<name>`` references (custom markers / MERGEFIELDs) return the name."""
        if self.path.startswith("b."):
            return self.path[2:]
        return None


@dataclass(slots=True)
class TemplateInspection:
    refs: list[PlaceholderRef] = field(default_factory=list)
    marker_names: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    parts: list[str] = field(default_factory=list)

    @property
    def paths(self) -> list[str]:
        seen: dict[str, None] = {}
        for r in self.refs:
            if r.binding_name is None:
                seen.setdefault(r.path, None)
        return list(seen)

    @property
    def binding_names(self) -> list[str]:
        seen: dict[str, None] = {}
        for r in self.refs:
            if r.binding_name is not None:
                seen.setdefault(r.binding_name, None)
        return list(seen)


def compile_template(data: bytes, marker: MarkerConfig) -> ConversionResult:
    """Return docx bytes where all placeholders are Jinja tags (original untouched)."""
    if marker.style is MarkerStyle.CUSTOM:
        return convert_custom_markers(data, marker.open, marker.close)
    if marker.style is MarkerStyle.MERGEFIELD:
        return convert_mergefields(data)
    return ConversionResult(data)


def _patched_parts(data: bytes) -> list[tuple[str, str]]:
    patcher = DocxTemplate.__new__(DocxTemplate)
    parts = []
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        for name in sorted(zf.namelist()):
            if not _PARTS.match(name):
                continue
            root = etree.fromstring(zf.read(name), _PARSER)
            xml = etree.tostring(root, encoding="unicode")
            parts.append((name, DocxTemplate.patch_xml(patcher, xml)))
    return parts


def _line_snippet(xml: str, lineno: int | None) -> str:
    if not lineno:
        return ""
    lines = re.sub(r"<w:p([ >])", r"\n<w:p\1", xml).splitlines()
    idx = max(0, min(len(lines) - 1, lineno - 1))
    text = re.sub(r"<[^>]+>", "", lines[idx])
    return text.strip()[:120]


class _RefCollector:
    def __init__(self, part: str) -> None:
        self.part = part
        self.refs: list[PlaceholderRef] = []

    def visit(self, node: nodes.Node, scope: dict[str, str], kind: RefKind = RefKind.VALUE) -> None:
        if isinstance(node, nodes.For):
            iter_path = self._chain_path(node.iter, scope)
            if iter_path:
                self._add(iter_path, RefKind.LOOP)
            else:
                self.visit(node.iter, scope, RefKind.LOOP)
            inner = dict(scope)
            targets = (
                [node.target]
                if isinstance(node.target, nodes.Name)
                else (list(node.target.items) if isinstance(node.target, nodes.Tuple) else [])
            )
            for t in targets:
                if isinstance(t, nodes.Name):
                    inner[t.name] = (iter_path + "[]") if iter_path else ""
            for child in node.body:
                self.visit(child, inner)
            for child in node.else_:
                self.visit(child, scope)
            if node.test is not None:
                self.visit(node.test, inner, RefKind.CONDITION)
            return
        if isinstance(node, nodes.If):
            self.visit(node.test, scope, RefKind.CONDITION)
            for child in node.body:
                self.visit(child, scope)
            for elif_ in node.elif_:
                self.visit(elif_, scope)
            for child in node.else_:
                self.visit(child, scope)
            return
        if isinstance(node, nodes.Assign) and isinstance(node.target, nodes.Name):
            path = self._chain_path(node.node, scope)
            self.visit(node.node, scope)
            scope[node.target.name] = path or ""
            return
        if isinstance(node, (nodes.Getattr, nodes.Getitem, nodes.Name)):
            path = self._chain_path(node, scope)
            if path:
                self._add(path, kind)
                return
            if path == "":
                return
        if isinstance(node, nodes.Filter) or isinstance(node, nodes.Test):
            if node.node is not None:
                self.visit(node.node, scope, kind)
            for arg in node.args:
                self.visit(arg, scope, kind)
            return
        for child in node.iter_child_nodes():
            self.visit(child, scope, kind)

    def _chain_path(self, node: nodes.Node, scope: dict[str, str]) -> str | None:
        chain = dotted_chain(node)
        if chain is None:
            return None
        root = chain[0]
        if root in _JINJA_GLOBALS or root in SAFE_FUNCTIONS:
            return ""
        if root in scope:
            base = scope[root]
            if not base:
                return ""
            return ".".join((base, *chain[1:])) if len(chain) > 1 else base
        return ".".join(chain)

    def _add(self, path: str, kind: RefKind) -> None:
        self.refs.append(PlaceholderRef(path, kind, self.part))


def inspect_compiled(data: bytes) -> TemplateInspection:
    """Inspect docx bytes that already use Jinja syntax."""
    result = TemplateInspection()
    env = make_environment()
    for part, xml in _patched_parts(data):
        result.parts.append(part)
        try:
            ast = env.parse(xml)
        except TemplateSyntaxError as exc:
            snippet = _line_snippet(xml, exc.lineno)
            result.errors.append(f"{part}: template syntax error: {exc.message} near '{snippet}'")
            continue
        collector = _RefCollector(part)
        collector.visit(ast, {})
        result.refs.extend(collector.refs)
    return result


def inspect_template(path_or_bytes: bytes | str, marker: MarkerConfig | None = None) -> TemplateInspection:
    """Inspect a template file (or bytes) with the given marker style."""
    if isinstance(path_or_bytes, bytes):
        data = path_or_bytes
    else:
        from pathlib import Path

        p = Path(path_or_bytes)
        check_zip(p, ZipLimits(max_file_bytes=200 * 1024 * 1024))
        data = p.read_bytes()
    marker = marker or MarkerConfig()
    converted = compile_template(data, marker)
    result = inspect_compiled(converted.data)
    result.marker_names = converted.names
    result.warnings.extend(converted.problems)
    if marker.style is MarkerStyle.JINJA:
        # remind authors about common foreign marker styles left in a Jinja template
        raw = inspect_raw_text(data)
        foreign = re.findall(r"«[^»]{1,60}»|<<[^>]{1,60}>>", raw)
        if foreign:
            result.warnings.append(
                f"Found non-Jinja markers {sorted(set(foreign))[:5]}; set the template marker "
                "style to 'custom' or 'mergefield' if they are placeholders."
            )
    return result


def inspect_raw_text(data: bytes) -> str:
    texts = []
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        for name in zf.namelist():
            if _PARTS.match(name):
                root = etree.fromstring(zf.read(name), _PARSER)
                texts.append("".join(str(t) for t in root.itertext()))
    return "\n".join(texts)
