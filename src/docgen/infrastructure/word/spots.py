"""Input spots in customer Word documents that are *not* explicit placeholders.

Customers often send forms instead of prepared templates. Three kinds of spots are
detected and, when bound to data, turned into ``{{ b["name"] }}`` placeholders on the
internal compiled copy (the customer's file is never modified):

* ``content_control`` – Word content controls (``w:sdt``): plain/rich text, date,
  drop-down/combo box, check box and picture (the frame's picture is replaced by the
  image from Excel, fitted into the frame). Named by their *tag*, else *title (alias)*.
  The control is kept (still a form in the output); only its content is filled and its
  "showing placeholder" state is cleared.
* ``form_field`` – legacy form fields: ``FORMTEXT`` (result text replaced),
  ``FORMCHECKBOX`` (checked state set) and ``FORMDROPDOWN`` (the entry equal to the value
  is selected; a value that is not one of the entries is an error, never a guess).
  Named by the field's bookmark name.
* ``blank`` – a run of dots/underscores/ellipsis (or underlined spaces) after a label,
  e.g. ``Họ và tên: ..........`` or ``ngày ..... tháng ..... năm .....``, and an empty
  table cell to the right of a label cell. Named by the label text; repeated labels get
  `` (2)``, `` (3)``…

Detection never changes anything; a spot is only converted when the caller's ``fill``
callback accepts it (i.e. the user bound it to a field), so blanks that must stay empty
(signature lines, "for office use") are left untouched.
"""

from __future__ import annotations

import base64
import io
import json
import re
import zipfile
from collections.abc import Callable
from copy import deepcopy
from dataclasses import dataclass, field
from typing import Any

from jinja2 import Undefined
from lxml import etree

from docgen.domain.data import values as V
from docgen.domain.mapping.expressions import ExpressionError
from docgen.infrastructure.word.markers import XML_NS, binding_tag, w

W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
_TEXT_PART = re.compile(r"^word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$")
_PARSER = etree.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False)
_FORM = re.compile(r"^\s*(FORMTEXT|FORMCHECKBOX|FORMDROPDOWN)\b", re.IGNORECASE)
_LABEL_CUT = re.compile(r"[.;,!?()\[\]]")
_NAME_FORBIDDEN = re.compile(r"[\"'\\{}]")
MAX_TABLE_SPOTS = 300

KIND_CONTENT_CONTROL = "content_control"
KIND_FORM_FIELD = "form_field"
KIND_BLANK = "blank"


@dataclass(frozen=True, slots=True)
class SpotOptions:
    content_controls: bool = True
    form_fields: bool = True
    blanks: bool = True
    blank_min: int = 5
    """Minimum filler length (dots/underscores; '…' counts 3) for a paragraph blank."""


@dataclass(frozen=True, slots=True)
class Spot:
    name: str
    kind: str
    subtype: str
    """text | date | dropdown | checkbox | label | table"""
    part: str
    context: str
    filled: bool = False


@dataclass(slots=True)
class SpotResult:
    data: bytes
    spots: list[Spot] = field(default_factory=list)
    problems: list[str] = field(default_factory=list)


FillDecision = Callable[[Spot], bool]


# ------------------------------------------- render-time filters (drop-downs)
def encode_choices(choices: list[str]) -> str:
    """Choices embedded in a template expression; URL-safe base64 survives XML and docxtpl."""
    return base64.urlsafe_b64encode(json.dumps(choices, ensure_ascii=False).encode("utf-8")).decode("ascii")


def _decode_choices(encoded: str) -> list[str]:
    return [str(c) for c in json.loads(base64.urlsafe_b64decode(encoded.encode("ascii")).decode("utf-8"))]


def match_choice(value: Any, choices: list[str]) -> int | None:
    """Index of the choice equal to *value*: exact, then ignoring case/spaces, then accents."""
    text = V.to_text(value).strip()
    for strip in (None, False, True):
        if strip is None:
            hits = [i for i, c in enumerate(choices) if c.strip() == text]
        else:
            key = V.normalize_label(text, accents=strip)
            hits = [i for i, c in enumerate(choices) if V.normalize_label(c, accents=strip) == key]
        if len(hits) == 1 or (hits and strip is None):
            return hits[0]
    return None


def _choice(value: Any, name: str, encoded: str, default: int) -> int:
    choices = _decode_choices(encoded)
    if isinstance(value, Undefined) or V.is_empty(value):
        return default if 0 <= default < len(choices) else 0
    index = match_choice(value, choices)
    if index is None:
        raise ExpressionError(
            f"value '{V.to_text(value)}' is not one of the choices of drop-down '{name}' "
            f"({' / '.join(choices)})"
        )
    return index


def dropdown_index(value: Any, name: str, encoded: str, default: int = 0) -> str:
    return str(_choice(value, name, encoded, default))


def dropdown_text(value: Any, name: str, encoded: str, default: int = 0) -> str:
    choices = _decode_choices(encoded)
    return choices[_choice(value, name, encoded, default)] if choices else ""


FORM_FILTERS: dict[str, Callable[..., str]] = {
    "dropdown_index": dropdown_index,
    "dropdown_text": dropdown_text,
}
"""Jinja filters the renderer must provide for filled drop-down form fields."""


def _clean_name(text: str) -> str:
    name = _NAME_FORBIDDEN.sub("", re.sub(r"\s+", " ", text)).strip(" :\t ")
    return name[:60].strip()


def _text_of(el: etree._Element) -> str:
    return "".join(t.text or "" for t in el.iter(w("t")))


def _ancestor(el: etree._Element, tag: str) -> etree._Element | None:
    parent = el.getparent()
    while parent is not None:
        if parent.tag == tag:
            return parent
        parent = parent.getparent()
    return None


def _short(text: str, limit: int = 100) -> str:
    text = re.sub(r"\s+", " ", text).strip()
    return text if len(text) <= limit else text[: limit - 1] + "…"


class _Namer:
    """Unique names per document for blanks; controls/fields keep their (shared) names."""

    def __init__(self) -> None:
        self.seen: dict[str, int] = {}
        self.anonymous = 0

    def unique(self, label: str, prefix: str) -> str:
        if not label:
            self.anonymous += 1
            return f"{prefix}_{self.anonymous}"
        count = self.seen.get(label, 0) + 1
        self.seen[label] = count
        return label if count == 1 else f"{label} ({count})"


# --------------------------------------------------------- content controls
def _sdt_subtype(pr: etree._Element) -> str:
    if pr.find(f"{{{W14}}}checkbox") is not None:
        return "checkbox"
    if pr.find(w("date")) is not None:
        return "date"
    if pr.find(w("dropDownList")) is not None or pr.find(w("comboBox")) is not None:
        return "dropdown"
    if pr.find(w("picture")) is not None:
        return "picture"
    return "text"


def _sdt_texts(content: etree._Element) -> list[etree._Element]:
    return [t for t in content.iter(w("t")) if _ancestor(t, w("sdt")) is _ancestor(content, w("sdt"))]


def _strip_placeholder_style(content: etree._Element) -> None:
    for style in content.iter(w("rStyle")):
        if (style.get(w("val")) or "").lower().replace(" ", "") == "placeholdertext":
            rpr = style.getparent()
            if rpr is not None:
                rpr.remove(style)


def _set_text(node: etree._Element, text: str) -> None:
    node.text = text
    node.set(f"{{{XML_NS}}}space", "preserve")


def _fill_text_nodes(nodes: list[etree._Element], text: str) -> bool:
    if not nodes:
        return False
    _set_text(nodes[0], text)
    for node in nodes[1:]:
        _set_text(node, "")
    return True


def _new_run(text: str, rpr: etree._Element | None = None) -> etree._Element:
    run = etree.Element(w("r"))
    if rpr is not None:
        copied = deepcopy(rpr)
        copied.tag = w("rPr")
        run.append(copied)
    t = etree.SubElement(run, w("t"))
    _set_text(t, text)
    return run


def _fill_sdt(sdt: etree._Element, name: str, subtype: str) -> None:
    pr = sdt.find(w("sdtPr"))
    content = sdt.find(w("sdtContent"))
    if pr is None or content is None:
        return
    for flag in pr.findall(w("showingPlcHdr")):
        pr.remove(flag)
    # Word refreshes a control bound to custom XML from that XML on open, which would
    # overwrite the Excel value; the user bound it to Excel, so that binding wins.
    for binding in pr.findall(w("dataBinding")):
        pr.remove(binding)
    date = pr.find(w("date"))
    if date is not None:  # stale picker value would disagree with the filled text
        if w("fullDate") in date.attrib:
            del date.attrib[w("fullDate")]
    _strip_placeholder_style(content)
    if subtype == "picture":
        _fill_picture(content, name)
        return
    if subtype == "checkbox":
        box = pr.find(f"{{{W14}}}checkbox")
        assert box is not None
        checked = box.find(f"{{{W14}}}checked")
        if checked is None:
            checked = etree.Element(f"{{{W14}}}checked")
            box.insert(0, checked)
        checked.set(f"{{{W14}}}val", "{{ b['" + name + "'] | checkbox_value }}")
        text = '{{ b["' + name + '"] | checkbox_symbol }}'
    else:
        text = binding_tag(name)
    if _fill_text_nodes(_sdt_texts(content), text):
        return
    paragraph = content.find(w("p"))
    if paragraph is not None:  # block-level control without text
        paragraph.append(_new_run(text))
        return
    cell = content.find(w("tc"))
    if cell is not None:
        p = cell.find(w("p"))
        if p is not None:
            p.append(_new_run(text))
        return
    content.append(_new_run(text))  # run-level control


def _fill_picture(content: etree._Element, name: str) -> None:
    """Replace the frame's placeholder picture by the bound image, fitted into the frame."""
    drawing = content.find(f".//{w('drawing')}")
    extent = drawing.find(f".//{{{WP}}}extent") if drawing is not None else None
    cx = int(extent.get("cx", "0")) if extent is not None else 0
    cy = int(extent.get("cy", "0")) if extent is not None else 0
    text = "{{ b['" + name + "'] | picture_fit(" + f"{cx}, {cy}" + ") }}"
    run = _ancestor(drawing, w("r")) if drawing is not None else None
    if run is not None:
        run.addprevious(_new_run(text, run.find(w("rPr"))))
        parent = run.getparent()
        assert parent is not None
        parent.remove(run)
        return
    paragraph = content.find(f".//{w('p')}")
    (paragraph if paragraph is not None else content).append(_new_run(text))


def _paragraph_context(el: etree._Element, mark: etree._Element | None = None) -> str:
    p = el if el.tag == w("p") else _ancestor(el, w("p"))
    if p is None:
        return ""
    parts: list[str] = []
    for node in p.iter(w("t"), mark.tag if mark is not None else w("t")):
        if node is mark:
            parts.append("[…]")  # the input itself, even when it holds no text (pictures)
        elif node.tag == w("t") and (mark is None or _ancestor(node, mark.tag) is not mark):
            parts.append(node.text or "")
    return _short("".join(parts))


def _content_controls(
    root: etree._Element, part: str, fill: FillDecision, result: SpotResult, counter: list[int]
) -> bool:
    changed = False
    for sdt in list(root.iter(w("sdt"))):
        content = sdt.find(w("sdtContent"))
        if content is None or content.find(f".//{w('sdt')}") is not None:
            continue  # container controls (repeating sections, groups) are not filled themselves
        pr = sdt.find(w("sdtPr"))
        if pr is None:
            continue
        subtype = _sdt_subtype(pr)
        tag = pr.find(w("tag"))
        alias = pr.find(w("alias"))
        raw = (tag.get(w("val")) if tag is not None else None) or (
            alias.get(w("val")) if alias is not None else None
        )
        name = _clean_name(raw or "")
        if not name:
            counter[0] += 1
            name = f"control_{counter[0]}"
        label = alias.get(w("val")) if alias is not None else ""
        parent = sdt.getparent()
        if parent is not None and parent.tag == w("p"):
            context = _paragraph_context(parent, sdt)  # inline control: show the sentence
        else:
            context = f"[{_short(_text_of(content), 60)}]"
        if label and label != name and label not in context:
            context = f"{label}: {context}"
        spot = Spot(name, KIND_CONTENT_CONTROL, subtype, part, context)
        if fill(spot):
            _fill_sdt(sdt, name, subtype)
            spot = Spot(name, KIND_CONTENT_CONTROL, subtype, part, context, filled=True)
            changed = True
        result.spots.append(spot)
    return changed


# -------------------------------------------------------- legacy form fields
def _field_runs(p: etree._Element) -> list[list[etree._Element]]:
    fields: list[list[etree._Element]] = []
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
                    fields.append(current)
                    current = []
    return fields


def _form_fields(
    root: etree._Element, part: str, fill: FillDecision, result: SpotResult, counter: list[int]
) -> bool:
    changed = False
    for p in list(root.iter(w("p"))):
        for runs in _field_runs(p):
            instr = "".join(t.text or "" for r in runs for t in r.findall(w("instrText")))
            m = _FORM.match(instr)
            if not m:
                continue
            kind = m.group(1).upper()
            begin = runs[0].find(w("fldChar"))
            ff = begin.find(w("ffData")) if begin is not None else None
            name_el = ff.find(w("name")) if ff is not None else None
            name = _clean_name(name_el.get(w("val"), "") if name_el is not None else "")
            if not name:
                counter[0] += 1
                name = f"field_{counter[0]}"
            context = _short(_paragraph_context(p))
            dd = ff.find(w("ddList")) if ff is not None else None
            choices = [e.get(w("val"), "") for e in dd.findall(w("listEntry"))] if dd is not None else []
            if kind == "FORMDROPDOWN":
                if dd is None or not choices:
                    result.problems.append(
                        f"{part}: drop-down form field '{name}' has no choices (left unchanged)"
                    )
                    continue
                context = _short(f"{context} [{' / '.join(choices)}]")
                subtype = "dropdown"
            else:
                subtype = "checkbox" if kind == "FORMCHECKBOX" else "text"
            spot = Spot(name, KIND_FORM_FIELD, subtype, part, context)
            if fill(spot):
                if subtype == "dropdown":
                    assert dd is not None
                    _fill_form_dropdown(runs, dd, name, choices)
                elif subtype == "checkbox":
                    box = ff.find(w("checkBox")) if ff is not None else None
                    if box is None:
                        result.problems.append(f"{part}: check box form field '{name}' has no checkBox data")
                        result.spots.append(spot)
                        continue
                    checked = box.find(w("checked"))
                    if checked is None:
                        checked = etree.SubElement(box, w("checked"))
                    checked.set(w("val"), "{{ b['" + name + "'] | checkbox_value }}")
                else:
                    _fill_form_text(runs, name)
                spot = Spot(name, KIND_FORM_FIELD, subtype, part, context, filled=True)
                changed = True
            result.spots.append(spot)
    return changed


def _fill_form_dropdown(
    runs: list[etree._Element], dd: etree._Element, name: str, choices: list[str]
) -> None:
    selected = dd.find(w("result"))
    if selected is None:
        default_el = dd.find(w("default"))
        current = default_el.get(w("val"), "0") if default_el is not None else "0"
        selected = etree.Element(w("result"))
        dd.insert(0, selected)  # CT_DDList: result, default, listEntry*
    else:
        current = selected.get(w("val"), "0")
    default = int(current) if current.isdigit() else 0
    args = f"'{name}', '{encode_choices(choices)}', {default}"
    selected.set(w("val"), "{{ b['" + name + "'] | dropdown_index(" + args + ") }}")
    # Word shows the selected entry; some editors also write it as result text - keep that in step
    for i, run in enumerate(runs):
        fld = run.find(w("fldChar"))
        if fld is not None and fld.get(w("fldCharType")) == "separate":
            nodes = [t for r in runs[i + 1 : -1] for t in r.findall(w("t"))]
            _fill_text_nodes(nodes, "{{ b['" + name + "'] | dropdown_text(" + args + ") }}")
            break


def _fill_form_text(runs: list[etree._Element], name: str) -> None:
    separate_index = None
    for i, run in enumerate(runs):
        fld = run.find(w("fldChar"))
        if fld is not None and fld.get(w("fldCharType")) == "separate":
            separate_index = i
            break
    tag = binding_tag(name)
    if separate_index is None:
        # no result section yet: add "separate" + result run before the end run
        end = runs[-1]
        sep = etree.Element(w("r"))
        etree.SubElement(sep, w("fldChar")).set(w("fldCharType"), "separate")
        end.addprevious(sep)
        end.addprevious(_new_run(tag))
        return
    result_nodes = [t for run in runs[separate_index + 1 : -1] for t in run.findall(w("t"))]
    if not _fill_text_nodes(result_nodes, tag):
        runs[-1].addprevious(_new_run(tag, runs[separate_index].find(w("rPr"))))


# ------------------------------------------------------------------- blanks
def _runs_in_fields(p: etree._Element) -> set[etree._Element]:
    inside: set[etree._Element] = set()
    for runs in _field_runs(p):
        inside.update(runs)
    return inside


def _underlined(run: etree._Element | None) -> bool:
    if run is None:
        return False
    rpr = run.find(w("rPr"))
    u = rpr.find(w("u")) if rpr is not None else None
    return u is not None and (u.get(w("val")) or "single") != "none"


def _left_cell_label(p: etree._Element) -> str:
    cell = _ancestor(p, w("tc"))
    if cell is None:
        return ""
    previous = cell.getprevious()
    while previous is not None and previous.tag != w("tc"):
        previous = previous.getprevious()
    return _clean_name(_text_of(previous)) if previous is not None else ""


def _label_before(text: str) -> str:
    segment = text.rstrip().rstrip(":：").rstrip()
    # 'Các ý kiến khác (nếu có):' -> 'Các ý kiến khác': a trailing note in brackets is not the label
    segment = re.sub(r"\s*\([^()]{0,40}\)$", "", segment).rstrip().rstrip(":：").rstrip()
    cut = max((m.end() for m in _LABEL_CUT.finditer(segment)), default=0)
    words = segment[cut:].split()
    return _clean_name(" ".join(words[-8:]))


def _paragraph_blanks(
    root: etree._Element,
    part: str,
    fill: FillDecision,
    result: SpotResult,
    options: SpotOptions,
    namer: _Namer,
) -> bool:
    changed = False
    filler = re.compile(r"[.…_](?:[ \t ]*[.…_])+")
    for p in list(root.iter(w("p"))):
        in_fields = _runs_in_fields(p)
        nodes = []
        for t in p.iter(w("t")):
            if _ancestor(t, w("p")) is not p or _ancestor(t, w("sdt")) is not None:
                continue
            run = t.getparent()
            if run in in_fields:
                continue
            nodes.append(t)
        if not nodes:
            continue
        texts = [n.text or "" for n in nodes]
        analysis = [
            ("_" * len(txt)) if txt and not txt.strip() and _underlined(n.getparent()) else txt
            for n, txt in zip(nodes, texts, strict=True)
        ]
        full = "".join(analysis)
        matches = [
            m
            for m in filler.finditer(full)
            if sum(3 if ch == "…" else 1 for ch in m.group(0) if ch in ".…_") >= options.blank_min
        ]
        if not matches:
            continue
        starts, pos = [], 0
        for txt in texts:
            starts.append(pos)
            pos += len(txt)

        def node_of(index: int, starts: list[int] = starts, texts: list[str] = texts) -> int:
            for i in range(len(texts)):
                if starts[i] <= index < starts[i] + len(texts[i]):
                    return i
            return len(texts) - 1

        decided: list[tuple[re.Match[str], Spot]] = []
        previous_end = 0
        for m in matches:
            label = _label_before(full[previous_end : m.start()].replace("_", " "))
            if not label and previous_end == 0:
                label = _left_cell_label(p)
            name = namer.unique(label, "blank")
            context = _short(full[: m.start()].replace("_", " ") + "[…]" + full[m.end() :].replace("_", " "))
            subtype = "table" if _ancestor(p, w("tc")) is not None else "label"
            spot = Spot(name, KIND_BLANK, subtype, part, context)
            decided.append((m, spot))
            previous_end = m.end()
        new_texts = list(texts)
        final: list[Spot] = []
        for m, spot in reversed(decided):
            if not fill(spot):
                final.append(spot)
                continue
            si, ei = node_of(m.start()), node_of(m.end() - 1)
            so, eo = m.start() - starts[si], m.end() - starts[ei]
            tag = binding_tag(spot.name)
            if si == ei:
                new_texts[si] = new_texts[si][:so] + tag + new_texts[si][eo:]
            else:
                new_texts[si] = new_texts[si][:so] + tag
                for k in range(si + 1, ei):
                    new_texts[k] = ""
                new_texts[ei] = new_texts[ei][eo:]
            final.append(Spot(spot.name, spot.kind, spot.subtype, spot.part, spot.context, True))
            changed = True
        result.spots.extend(reversed(final))  # document order
        for node, old, new in zip(nodes, texts, new_texts, strict=True):
            if old != new:
                _set_text(node, new)
    return changed


def _table_blanks(
    root: etree._Element, part: str, fill: FillDecision, result: SpotResult, namer: _Namer
) -> bool:
    changed = False
    count = 0
    for row in list(root.iter(w("tr"))):
        cells = [c for c in row if c.tag == w("tc")]
        # decide on the original row content first: a cell filled below must never become
        # the "label" of the next empty cell
        texts = [_text_of(c) for c in cells]
        candidates = []
        for i in range(1, len(cells)):
            cell = cells[i]
            if texts[i].strip() or cell.find(f".//{w('sdt')}") is not None:
                continue
            if cell.find(f".//{w('drawing')}") is not None or cell.find(f".//{w('tbl')}") is not None:
                continue
            label = _clean_name(texts[i - 1])
            if not label or not re.search(r"[^\W\d_]", label) or len(label) > 60:
                continue
            if "{{" in texts[i - 1] or "{%" in texts[i - 1]:
                continue
            candidates.append((cell, label))
        for cell, label in candidates:
            count += 1
            if count > MAX_TABLE_SPOTS:
                result.problems.append(f"{part}: more than {MAX_TABLE_SPOTS} empty table cells; rest ignored")
                return changed
            spot = Spot(namer.unique(label, "cell"), KIND_BLANK, "table", part, _short(f"{label} | […]"))
            if fill(spot):
                p = cell.find(w("p"))
                if p is None:
                    p = etree.SubElement(cell, w("p"))
                ppr = p.find(w("pPr"))
                mark_rpr = ppr.find(w("rPr")) if ppr is not None else None
                p.append(_new_run(binding_tag(spot.name), mark_rpr))
                spot = Spot(spot.name, spot.kind, spot.subtype, spot.part, spot.context, True)
                changed = True
            result.spots.append(spot)
    return changed


# -------------------------------------------------------------------- entry
def process_spots(data: bytes, options: SpotOptions, fill: FillDecision | None = None) -> SpotResult:
    """Detect spots (and convert those *fill* accepts). Returns new bytes + all spots in order."""
    decide: FillDecision = fill or (lambda _spot: False)
    result = SpotResult(data)
    if not (options.content_controls or options.form_fields or options.blanks):
        return result
    namer = _Namer()
    counters = {"control": [0], "field": [0]}
    src = zipfile.ZipFile(io.BytesIO(data))
    out_buf = io.BytesIO()
    any_change = False
    with src, zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as out:
        for item in src.infolist():
            payload = src.read(item.filename)
            if _TEXT_PART.match(item.filename):
                root = etree.fromstring(payload, _PARSER)
                changed = False
                if options.content_controls:
                    changed |= _content_controls(root, item.filename, decide, result, counters["control"])
                if options.form_fields:
                    changed |= _form_fields(root, item.filename, decide, result, counters["field"])
                if options.blanks:
                    changed |= _paragraph_blanks(root, item.filename, decide, result, options, namer)
                    changed |= _table_blanks(root, item.filename, decide, result, namer)
                if changed:
                    any_change = True
                    payload = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            out.writestr(item, payload)
    if any_change:
        result.data = out_buf.getvalue()
    return result
