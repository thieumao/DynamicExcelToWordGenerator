"""Output verification for generated DOCX files.

Checks are structural (ZIP integrity, content types, relationship targets,
python-docx can reopen the file) plus a scan for template markers that were
not rendered. They do not prove that Microsoft Word renders identically.
"""

from __future__ import annotations

import posixpath
import re
import zipfile
from collections.abc import Sequence
from dataclasses import dataclass, field
from pathlib import Path

from lxml import etree

PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"

_TEXT_PARTS = re.compile(r"^word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$")
_SAFE_PARSER = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=False)


@dataclass(slots=True)
class VerificationResult:
    path: Path
    problems: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    leftover_markers: list[tuple[str, str]] = field(default_factory=list)
    """(part name, marker text) pairs that still look like unrendered placeholders."""

    @property
    def ok(self) -> bool:
        return not self.problems


def _paragraph_texts(xml: bytes) -> list[str]:
    root = etree.fromstring(xml, _SAFE_PARSER)
    texts = []
    for p in root.iter(f"{{{W_NS}}}p"):
        texts.append("".join(t.text or "" for t in p.iter(f"{{{W_NS}}}t")))
    return texts


def verify_docx(
    path: Path,
    *,
    marker_patterns: Sequence[re.Pattern[str]] = (
        re.compile(r"\{\{.*?\}\}"),
        re.compile(r"\{%.*?%\}"),
    ),
    check_reopen: bool = True,
) -> VerificationResult:
    result = VerificationResult(path=path)
    try:
        zf = zipfile.ZipFile(path)
    except (zipfile.BadZipFile, OSError) as exc:
        result.problems.append(f"Not a valid ZIP/DOCX file: {exc}")
        return result
    with zf:
        bad = zf.testzip()
        if bad is not None:
            result.problems.append(f"Corrupted ZIP member: {bad}")
        names = set(zf.namelist())
        for required in ("[Content_Types].xml", "word/document.xml", "_rels/.rels"):
            if required not in names:
                result.problems.append(f"Missing required part: {required}")
        for name in sorted(n for n in names if n.endswith(".rels")):
            base_dir = posixpath.dirname(posixpath.dirname(name))
            try:
                rels = etree.fromstring(zf.read(name), _SAFE_PARSER)
            except etree.XMLSyntaxError as exc:
                result.problems.append(f"Invalid XML in {name}: {exc}")
                continue
            for rel in rels.iter(f"{{{PKG_REL_NS}}}Relationship"):
                if rel.get("TargetMode") == "External":
                    continue
                target = rel.get("Target", "")
                resolved = (
                    target.lstrip("/")
                    if target.startswith("/")
                    else posixpath.normpath(posixpath.join(base_dir, target))
                )
                if resolved not in names:
                    result.problems.append(
                        f"Broken relationship {rel.get('Id')} in {name}: target {target} missing"
                    )
        doc_pr_ids: list[str] = []
        for name in sorted(names):
            if not _TEXT_PARTS.match(name):
                continue
            data = zf.read(name)
            try:
                root = etree.fromstring(data, _SAFE_PARSER)
            except etree.XMLSyntaxError as exc:
                result.problems.append(f"Invalid XML in {name}: {exc}")
                continue
            doc_pr_ids.extend(el.get("id", "") for el in root.iter(f"{{{WP_NS}}}docPr"))
            for text in _paragraph_texts(data):
                for pattern in marker_patterns:
                    for match in pattern.finditer(text):
                        result.leftover_markers.append((name, match.group(0)))
        dupes = {i for i in doc_pr_ids if doc_pr_ids.count(i) > 1}
        if dupes:
            result.warnings.append(f"Duplicate drawing ids: {sorted(dupes)[:5]}")
    if result.leftover_markers:
        sample = ", ".join(m for _, m in result.leftover_markers[:5])
        result.problems.append(f"Unrendered template markers remain: {sample}")
    if check_reopen and not result.problems:
        try:
            from docx import Document

            Document(str(path))
        except Exception as exc:  # python-docx raises various exception types
            result.problems.append(f"python-docx cannot reopen the document: {exc}")
    return result
