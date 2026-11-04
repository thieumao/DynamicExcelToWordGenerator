"""Marker adapters: convert customer placeholder styles into docxtpl/Jinja tags.

The original template is never modified: adapters read the .docx bytes and
return new bytes for an internal "compiled" copy.

* ``custom``  – ``<<Họ tên>>``, ``[MaNV]``, ``${code}`` ... with configurable
  delimiters, even when Word split the marker across several runs.
* ``mergefield`` – Word mail-merge fields, both ``w:fldSimple`` and complex
  ``w:fldChar``/``w:instrText`` fields.

Each placeholder name becomes ``{{ b["name"] }}``; ``b`` (bindings) is filled
per template from ``TemplateSpec.bindings`` or by exact field-path match.
"""

from __future__ import annotations

import io
import re
import zipfile
from collections.abc import Callable, Iterator
from copy import deepcopy
from dataclasses import dataclass, field

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
_TEXT_PART = re.compile(r"^word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$")
_PARSER = etree.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False)
_MERGEFIELD = re.compile(r"^\s*MERGEFIELD\s+(?:\"([^\"]+)\"|(\S+))", re.IGNORECASE)


def w(tag: str) -> str:
    return f"{{{W}}}{tag}"


@dataclass(frozen=True, slots=True)
class MarkerHit:
    name: str
    part: str
    style: str


@dataclass(slots=True)
class ConversionResult:
    data: bytes
    hits: list[MarkerHit] = field(default_factory=list)
    problems: list[str] = field(default_factory=list)

    @property
    def names(self) -> list[str]:
        seen: dict[str, None] = {}
        for h in self.hits:
            seen.setdefault(h.name, None)
        return list(seen)


def binding_tag(name: str) -> str:
    return '{{ b["' + name + '"] }}'


def _valid_name(name: str) -> bool:
    return bool(name) and '"' not in name and "\\" not in name and "{" not in name and "}" not in name


def _rewrite_parts(
    data: bytes, fn: Callable[[etree._Element, str, ConversionResult], bool]
) -> ConversionResult:
    result = ConversionResult(b"")
    src = zipfile.ZipFile(io.BytesIO(data))
    out_buf = io.BytesIO()
    with src, zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as out:
        for item in src.infolist():
            payload = src.read(item.filename)
            if _TEXT_PART.match(item.filename):
                root = etree.fromstring(payload, _PARSER)
                if fn(root, item.filename, result):
                    payload = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            out.writestr(item, payload)
    result.data = out_buf.getvalue()
    return result


def _own_text_nodes(p: etree._Element) -> list[etree._Element]:
    """w:t nodes whose nearest paragraph ancestor is *p* (skips nested text boxes)."""
    nodes = []
    for t in p.iter(w("t")):
        parent = t.getparent()
        while parent is not None and parent.tag != w("p"):
            parent = parent.getparent()
        if parent is p:
            nodes.append(t)
    return nodes


def _set_text(node: etree._Element, text: str) -> None:
    node.text = text
    node.set(f"{{{XML_NS}}}space", "preserve")


# ----------------------------------------------------------------- custom
def convert_custom_markers(
    data: bytes, open_mark: str, close_mark: str, *, skip: Callable[[str], bool] | None = None
) -> ConversionResult:
    pattern = re.compile(re.escape(open_mark) + r"\s*(.+?)\s*" + re.escape(close_mark), re.DOTALL)

    def handle(root: etree._Element, part: str, result: ConversionResult) -> bool:
        changed = False
        for p in root.iter(w("p")):
            nodes = _own_text_nodes(p)
            if not nodes:
                continue
            texts = [n.text or "" for n in nodes]
            full = "".join(texts)
            matches = list(pattern.finditer(full))
            if not matches:
                continue
            original = list(texts)
            starts = []
            pos = 0
            for t in original:
                starts.append(pos)
                pos += len(t)

            def node_of(char_index: int, starts: list[int] = starts, original: list[str] = original) -> int:
                for i in range(len(original)):
                    if starts[i] <= char_index < starts[i] + len(original[i]):
                        return i
                return len(original) - 1

            for m in reversed(matches):
                name = m.group(1).strip()
                if skip is not None and skip(name):
                    continue
                if not _valid_name(name):
                    result.problems.append(f"{part}: unsupported characters in marker {m.group(0)!r}")
                    continue
                si = node_of(m.start())
                so = m.start() - starts[si]
                ei = node_of(m.end() - 1)
                eo = m.end() - starts[ei]
                replacement = binding_tag(name)
                if si == ei:
                    texts[si] = texts[si][:so] + replacement + texts[si][eo:]
                else:
                    texts[si] = texts[si][:so] + replacement
                    for k in range(si + 1, ei):
                        texts[k] = ""
                    texts[ei] = texts[ei][eo:]
                result.hits.append(MarkerHit(name, part, "custom"))
                changed = True
            for node, text in zip(nodes, texts, strict=True):
                if (node.text or "") != text:
                    _set_text(node, text)
        return changed

    result = _rewrite_parts(data, handle)
    result.hits.reverse()
    return result


# ------------------------------------------------------------- mergefield
def _run_with_text(template_run: etree._Element | None, text: str) -> etree._Element:
    run = etree.Element(w("r"))
    if template_run is not None:
        rpr = template_run.find(w("rPr"))
        if rpr is not None:
            run.append(deepcopy(rpr))
    t = etree.SubElement(run, w("t"))
    _set_text(t, text)
    return run


def _field_runs(p: etree._Element) -> Iterator[list[etree._Element]]:
    """Yield lists of runs forming one complete top-level complex field in paragraph *p*."""
    depth = 0
    current: list[etree._Element] = []
    for run in list(p.iter(w("r"))):
        fld = run.find(w("fldChar"))
        kind = fld.get(w("fldCharType")) if fld is not None else None
        if kind == "begin":
            depth += 1
            current.append(run)
            continue
        if depth > 0:
            current.append(run)
            if kind == "end":
                depth -= 1
                if depth == 0:
                    yield current
                    current = []


def convert_mergefields(data: bytes) -> ConversionResult:
    def handle(root: etree._Element, part: str, result: ConversionResult) -> bool:
        changed = False
        # simple fields
        for fld in list(root.iter(w("fldSimple"))):
            instr = fld.get(w("instr"), "")
            m = _MERGEFIELD.match(instr)
            if not m:
                continue
            name = (m.group(1) or m.group(2) or "").strip()
            if name.lower().startswith(("tablestart:", "tableend:")) or not _valid_name(name):
                result.problems.append(f"{part}: unsupported MERGEFIELD {name!r}")
                continue
            runs = fld.findall(w("r"))
            new_run = _run_with_text(runs[0] if runs else None, binding_tag(name))
            parent = fld.getparent()
            assert parent is not None
            parent.replace(fld, new_run)
            result.hits.append(MarkerHit(name, part, "mergefield"))
            changed = True
        # complex fields
        for p in list(root.iter(w("p"))):
            for runs in list(_field_runs(p)):
                nested = sum(
                    1
                    for r in runs
                    if (f := r.find(w("fldChar"))) is not None and f.get(w("fldCharType")) == "begin"
                )
                instr = "".join((t.text or "") for r in runs for t in r.findall(w("instrText")))
                m = _MERGEFIELD.match(instr)
                if not m:
                    continue
                name = (m.group(1) or m.group(2) or "").strip()
                if nested > 1:
                    result.problems.append(
                        f"{part}: MERGEFIELD {name!r} is nested inside another field (IF/NEXT)"
                        " – not supported"
                    )
                    continue
                if name.lower().startswith(("tablestart:", "tableend:")) or not _valid_name(name):
                    result.problems.append(f"{part}: unsupported MERGEFIELD {name!r}")
                    continue
                # formatting from the first result run (after 'separate'), else the begin run
                template_run = runs[0]
                seen_sep = False
                for r in runs:
                    f = r.find(w("fldChar"))
                    if f is not None and f.get(w("fldCharType")) == "separate":
                        seen_sep = True
                        continue
                    if seen_sep and r.find(w("t")) is not None:
                        template_run = r
                        break
                new_run = _run_with_text(template_run, binding_tag(name))
                first = runs[0]
                parent = first.getparent()
                assert parent is not None
                parent.insert(parent.index(first), new_run)
                for r in runs:
                    rp = r.getparent()
                    if rp is not None:
                        rp.remove(r)
                result.hits.append(MarkerHit(name, part, "mergefield"))
                changed = True
        return changed

    fields = _rewrite_parts(data, handle)
    # Editors such as LibreOffice flatten unbound mail-merge fields into plain text that only shows
    # the field result («Name»). Treat that visible form as a placeholder too, otherwise it would be
    # copied to every output unchanged.
    flattened = convert_custom_markers(
        fields.data, "«", "»", skip=lambda n: n.lower().startswith(("tablestart:", "tableend:"))
    )
    fields.data = flattened.data
    fields.hits.extend(MarkerHit(h.name, h.part, "mergefield-text") for h in flattened.hits)
    fields.problems.extend(flattened.problems)
    return fields


DEFAULT_LEFTOVER_PATTERNS = (r"\{\{.*?\}\}", r"\{%.*?%\}")


def leftover_patterns(style: str, open_mark: str = "<<", close_mark: str = ">>") -> tuple[str, ...]:
    """Regexes (text) that must not survive rendering for a template using *style* markers."""
    if style == "custom":
        return (re.escape(open_mark) + r".+?" + re.escape(close_mark),)
    if style == "mergefield":
        return (r"«[^»]{1,100}»",)
    return ()


def find_custom_markers(data: bytes, open_mark: str, close_mark: str) -> list[MarkerHit]:
    return convert_custom_markers(data, open_mark, close_mark).hits
