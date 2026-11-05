"""Quick start for non-IT users: an Excel file and Word files become a ready profile.

The user never sees sheets ids, field paths or profile options. This module

1. reads the workbook and guesses each sheet's role: the *list* (one row = one set of documents),
   *general information* entered once (label | value, e.g. the project details), *lookup* tables
   whose codes appear in the list (e.g. 'Mã giám sát' -> sheet 'Nhân sự'), or unused;
2. offers every usable value in plain words ("Tên công việc", "Ngày nghiệm thu (ngày … tháng …
   năm …)", "Mã giám sát ▸ Họ và tên"), with an example from the first row;
3. finds what to fill in each Word file (``<<…>>`` markers, mail-merge fields, blanks after
   labels, content controls, form fields) and suggests a value for each by name;
4. builds the profile from the user's choices.

Suggestions are only suggestions: the user confirms them in the wizard, and anything left
unmatched stays as it is in the Word file.
"""

from __future__ import annotations

import datetime as dt
import difflib
import io
import re
import zipfile
from collections.abc import Iterable, Mapping, Sequence
from dataclasses import dataclass, field
from decimal import Decimal
from itertools import islice
from pathlib import Path
from typing import Any

from lxml import etree

from docgen.application.workbook import suggest_identifier
from docgen.domain.data import values as V
from docgen.domain.mapping.transforms import DEFAULT_REGISTRY, TransformContext, TransformError
from docgen.domain.profile.models import MarkerConfig, MarkerStyle, Profile
from docgen.infrastructure.excel.header_detect import detect_header_row
from docgen.infrastructure.excel.sources import open_workbook
from docgen.infrastructure.word.inspector import compile_template
from docgen.infrastructure.word.spots import SpotOptions, process_spots

ROLE_LIST = "list"
ROLE_INFO = "info"
ROLE_LOOKUP = "lookup"
ROLE_IGNORE = "ignore"

PRIMARY_ID = "ds"
MAX_ROWS = 5000
_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_TEXT_PART = re.compile(r"^word/(document|header\d*|footer\d*)\.xml$")
_MARKER = re.compile(r"<<\s*([^<>]{1,80}?)\s*>>")
_STOP = {"va", "cua", "la", "cac", "nhung"}
_GENERIC = {"ngay", "thang", "nam", "gio", "phut", "so", "tai", "ten", "ong", "ba", "ongba", "ky", "noi"}
_CTX = TransformContext()


# --------------------------------------------------------------- workbook
@dataclass(slots=True)
class SheetPlan:
    """A sheet of the workbook and the role it plays."""

    name: str
    role: str
    header_row: int | None
    headers: list[str] = field(default_factory=list)
    rows: list[tuple[Any, ...]] = field(default_factory=list)
    """Data rows below the header (table sheets)."""
    info: dict[str, Any] = field(default_factory=dict)
    """label -> value (general information sheets)."""
    links: list[tuple[int, int]] = field(default_factory=list)
    """Lookup sheets: (column of the list, key column of this sheet)."""
    hidden: bool = False

    @property
    def row_count(self) -> int:
        return len(self.rows)


@dataclass(slots=True)
class QuickAnalysis:
    path: Path
    sheets: list[SheetPlan]

    @property
    def primary(self) -> SheetPlan | None:
        return next((s for s in self.sheets if s.role == ROLE_LIST), None)

    def set_role(self, sheet_name: str, role: str) -> None:
        """Change a sheet's role (there is only one list sheet); lookups are re-detected."""
        for sheet in self.sheets:
            if sheet.name == sheet_name:
                sheet.role = role
            elif role == ROLE_LIST and sheet.role == ROLE_LIST:
                sheet.role = ROLE_IGNORE
        _detect_links(self)


def _filled(row: Sequence[Any]) -> list[tuple[int, Any]]:
    return [(i, v) for i, v in enumerate(row) if not V.is_empty(v)]


def _key_value_pairs(rows: Sequence[Sequence[Any]]) -> dict[str, Any] | None:
    """label -> value when the sheet looks like 'label | value' lines, else None."""
    lines = [f for f in (_filled(r) for r in rows) if f]
    pairs: dict[str, Any] = {}
    shaped = 0
    for filled in lines:
        if isinstance(filled[0][1], str) and 2 <= len(filled) <= 3:
            shaped += 1
            label = filled[0][1].strip().rstrip(":：").strip()
            if label and label not in pairs:
                pairs[label] = filled[1][1]
    if shaped >= 2 and shaped >= 0.6 * len(lines) and max(len(f) for f in lines) <= 4:
        return pairs
    return None


def _is_sequence(values: list[Any]) -> bool:
    """1, 2, 3 … (an 'STT' column) – never a code to look up."""
    numbers = [v for v in values if isinstance(v, (int, float)) and not isinstance(v, bool)]
    return len(numbers) == len(values) and [int(n) for n in numbers] == list(range(1, len(numbers) + 1))


def _column(rows: Sequence[Sequence[Any]], index: int) -> list[Any]:
    return [r[index] for r in rows if index < len(r) and not V.is_empty(r[index])]


def _detect_links(analysis: QuickAnalysis) -> None:
    primary = analysis.primary
    for sheet in analysis.sheets:
        sheet.links = []
    if primary is None:
        return
    taken: set[int] = set()
    candidates = [
        s for s in analysis.sheets if s is not primary and s.role in (ROLE_LOOKUP, ROLE_IGNORE, ROLE_INFO)
    ]
    for ci, header in enumerate(primary.headers):
        values = _column(primary.rows, ci)
        if not header or len(values) < min(2, len(primary.rows)) or _is_sequence(values):
            continue
        wanted = {V.canonical_key(v, casefold=True) for v in values}
        best: tuple[float, SheetPlan, int] | None = None
        for sheet in candidates:
            if not sheet.headers or not sheet.rows:
                continue
            for ki, key_header in enumerate(sheet.headers):
                keys = _column(sheet.rows, ki)
                if not key_header or not keys or _is_sequence(keys):
                    continue
                key_set = {V.canonical_key(k, casefold=True) for k in keys}
                if len(key_set) < 0.9 * len(keys):
                    continue  # not unique: not a code column
                coverage = len(wanted & key_set) / len(wanted)
                if coverage >= 0.8 and (best is None or coverage > best[0]):
                    best = (coverage, sheet, ki)
        if best is not None and ci not in taken:
            taken.add(ci)
            best[1].links.append((ci, best[2]))
            if best[1].role != ROLE_LOOKUP:
                best[1].role = ROLE_LOOKUP
    for sheet in analysis.sheets:
        if sheet.role == ROLE_LOOKUP and not sheet.links:
            sheet.role = ROLE_INFO if sheet.info else ROLE_IGNORE


def analyze_workbook(path: Path, *, max_rows: int = MAX_ROWS) -> QuickAnalysis:
    source = open_workbook(path)
    sheets: list[SheetPlan] = []
    try:
        for info in source.sheets():
            raw = [tuple(r) for r in islice(source.iter_rows(info.name, 1, max_rows), max_rows)]
            detection = detect_header_row([list(r) for r in raw[:30]])
            headers: list[str] = []
            rows: list[tuple[Any, ...]] = []
            if detection.row is not None:
                headers = ["" if V.is_empty(v) else str(v).strip() for v in raw[detection.row - 1]]
                while headers and not headers[-1]:
                    headers.pop()
                rows = [r for r in raw[detection.row :] if _filled(r)]
            pairs = _key_value_pairs(raw)
            sheet = SheetPlan(
                info.name, ROLE_IGNORE, detection.row, headers, rows, pairs or {}, hidden=info.hidden
            )
            sheets.append(sheet)
    finally:
        source.close()
    analysis = QuickAnalysis(path, sheets)
    tables = [s for s in sheets if not s.hidden and len([h for h in s.headers if h]) >= 2 and s.rows]
    # the list: the biggest table that is not a 'label | value' sheet (fall back to any table)
    pool = [s for s in tables if not s.info] or tables
    if pool:
        max(pool, key=lambda s: (s.row_count, len(s.headers))).role = ROLE_LIST
    for sheet in sheets:
        if sheet.role != ROLE_LIST and sheet.info and not sheet.hidden:
            sheet.role = ROLE_INFO
    _detect_links(analysis)
    return analysis


# ----------------------------------------------------------------- values
@dataclass(slots=True)
class ValueOption:
    """One value the user can put into a Word input."""

    path: str
    label: str
    group: str
    sample: str
    rule: dict[str, Any]
    names: tuple[str, ...] = ()
    """Other ways the value may be called in a Word file (for suggestions)."""
    base: bool = True
    """False for derived variants (date parts, numbers in words): created only when used."""


@dataclass(slots=True)
class ProfileParts:
    """Ids chosen for the analysis (kept stable between option listing and profile building)."""

    sheet_ids: dict[str, str]
    column_ids: dict[tuple[str, int], str]
    """(sheet name, column index or info label index) -> column id."""


def _kind(values: Iterable[Any]) -> str:
    seen = [v for v in values if not V.is_empty(v)][:200]
    if not seen:
        return "text"
    if all(isinstance(v, dt.time) for v in seen):
        return "time"
    if all(isinstance(v, (dt.date, dt.datetime)) for v in seen):
        return "date"
    if all(isinstance(v, (int, float, Decimal)) and not isinstance(v, bool) for v in seen):
        return "integer" if all(float(v).is_integer() for v in seen) else "number"
    return "text"


def _sample(value: Any, transforms: Sequence[Any]) -> str:
    if V.is_empty(value):
        return ""
    try:
        for call in transforms:
            name, params = (call, {}) if isinstance(call, str) else next(iter(call.items()))
            value = DEFAULT_REGISTRY.apply(name, value, _CTX, params)
    except TransformError:
        return V.to_text(value)
    if isinstance(value, dt.datetime):
        return value.strftime("%d/%m/%Y") if value.time() == dt.time() else value.strftime("%d/%m/%Y %H:%M")
    if isinstance(value, dt.date):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, dt.time):
        return value.strftime("%H:%M")
    return V.to_text(value)[:80]


def _variants(
    path: str, label: str, group: str, rule: dict[str, Any], values: list[Any], names: tuple[str, ...]
) -> list[ValueOption]:
    """The value itself and, for dates/times/numbers, the usual ways to write it."""
    first = values[0] if values else None
    kind = _kind(values)
    words = "ngày %d tháng %m năm %Y"
    out: list[ValueOption] = []
    if kind == "date":
        out.append(
            ValueOption(
                path,
                label,
                group,
                _sample(first, ["format_date"]),
                {**rule, "transforms": ["format_date"]},
                names,
            )
        )
        for suffix, text, fmt, extra in (
            ("bang_chu", "(ngày … tháng … năm …)", words, ("bằng chữ", "đầy đủ")),
            ("ngay", "– ngày", "%d", ("ngày",)),
            ("thang", "– tháng", "%m", ("tháng",)),
            ("nam", "– năm", "%Y", ("năm",)),
        ):
            transforms = [{"format_date": {"format": fmt}}]
            out.append(
                ValueOption(
                    f"{path}_{suffix}",
                    f"{label} {text}",
                    group,
                    _sample(first, transforms),
                    {**rule, "transforms": transforms},
                    tuple(f"{n} {e}" for n in (label, *names) for e in extra),
                    base=False,
                )
            )
    elif kind == "time":
        out.append(
            ValueOption(
                path,
                label,
                group,
                _sample(first, ["format_time"]),
                {**rule, "transforms": ["format_time"]},
                names,
            )
        )
        transforms = [{"format_time": {"format": "%H giờ %M"}}]
        out.append(
            ValueOption(
                f"{path}_gio",
                f"{label} (08 giờ 30)",
                group,
                _sample(first, transforms),
                {**rule, "transforms": transforms},
                tuple(f"{n} giờ" for n in (label, *names)),
                base=False,
            )
        )
    else:
        out.append(ValueOption(path, label, group, _sample(first, []), dict(rule), names))
        if kind in ("integer", "number") and any(
            abs(float(v)) >= 1000 for v in values[:200] if not V.is_empty(v)
        ):
            words_rule = [{"number_to_words_vi": {"capitalize": True}}]
            out.append(
                ValueOption(
                    f"{path}_bang_chu",
                    f"{label} (bằng chữ)",
                    group,
                    _sample(first, words_rule),
                    {**rule, "transforms": words_rule},
                    tuple(f"{n} bằng chữ" for n in (label, *names)),
                    base=False,
                )
            )
            if kind == "integer":
                grouped = [{"format_number": {"decimals": 0}}]
                out.append(
                    ValueOption(
                        f"{path}_phan_cach",
                        f"{label} (1.234.567)",
                        group,
                        _sample(first, grouped),
                        {**rule, "transforms": grouped},
                        (),
                        base=False,
                    )
                )
    return out


def _strip_code_word(header: str) -> str:
    """'Mã giám sát' -> 'giám sát' (what a code column refers to)."""
    return (
        re.sub(r"^\s*(mã|ma|id|code|số hiệu|so hieu)\s+", "", header, flags=re.IGNORECASE).strip() or header
    )


def plan_ids(analysis: QuickAnalysis) -> ProfileParts:
    sheet_ids: dict[str, str] = {}
    column_ids: dict[tuple[str, int], str] = {}
    n_info = n_lookup = 0
    for sheet in analysis.sheets:
        taken: set[str] = set()
        if sheet.role == ROLE_LIST:
            sheet_ids[sheet.name] = PRIMARY_ID
            headers = sheet.headers
        elif sheet.role == ROLE_INFO:
            n_info += 1
            sheet_ids[sheet.name] = f"tt{n_info}"
            headers = list(sheet.info)
        elif sheet.role == ROLE_LOOKUP:
            n_lookup += 1
            sheet_ids[sheet.name] = f"tc{n_lookup}"
            headers = sheet.headers
        else:
            continue
        for i, header in enumerate(headers):
            if header:
                column_ids[(sheet.name, i)] = suggest_identifier(header, taken)
    return ProfileParts(sheet_ids, column_ids)


def value_options(analysis: QuickAnalysis) -> list[ValueOption]:
    """Everything the user can put into the Word files, grouped and with examples."""
    parts = plan_ids(analysis)
    primary = analysis.primary
    options: list[ValueOption] = []
    if primary is not None:
        for i, header in enumerate(primary.headers):
            cid = parts.column_ids.get((primary.name, i))
            if cid:
                values = _column(primary.rows, i)
                options += _variants(f"dong.{cid}", header, "Danh sách", {"source": cid}, values, ())
    used_info: set[str] = set()
    for sheet in analysis.sheets:
        if sheet.role != ROLE_INFO:
            continue
        sid = parts.sheet_ids[sheet.name]
        for i, (label, value) in enumerate(sheet.info.items()):
            cid = parts.column_ids[(sheet.name, i)]
            path = f"chung.{cid}" if cid not in used_info else f"chung.{sid}_{cid}"
            used_info.add(cid)
            options += _variants(path, label, "Thông tin chung", {"path": f"r_{sid}.{cid}"}, [value], ())
    if primary is not None:
        for sheet in analysis.sheets:
            if sheet.role != ROLE_LOOKUP:
                continue
            for ci, ki in sheet.links:
                pcid = parts.column_ids[(primary.name, ci)]
                code_header = primary.headers[ci]
                about = _strip_code_word(code_header)
                index = {
                    V.canonical_key(r[ki], casefold=True): r
                    for r in sheet.rows
                    if ki < len(r) and not V.is_empty(r[ki])
                }
                firsts = _column(primary.rows, ci)
                match = index.get(V.canonical_key(firsts[0], casefold=True)) if firsts else None
                for xi, header in enumerate(sheet.headers):
                    xid = parts.column_ids.get((sheet.name, xi))
                    if xi == ki or not xid:
                        continue
                    sample_values = [match[xi]] if match is not None and xi < len(match) else []
                    sample_values += _column(sheet.rows, xi)
                    options += _variants(
                        f"tra.{pcid}.{xid}",
                        f"{code_header} ▸ {header}",
                        f"Tra cứu theo {code_header}",
                        {"path": f"r_{pcid}.{xid}"},
                        sample_values,
                        (f"{header} {about}", f"{header} của {about}"),
                    )
    return options


# ------------------------------------------------------------------ Word
@dataclass(slots=True)
class WordInput:
    name: str
    kind: str
    """marker | blank | content_control | form_field"""
    subtype: str
    context: str


@dataclass(slots=True)
class WordScan:
    path: Path
    style: MarkerStyle
    open: str = "<<"
    close: str = ">>"
    inputs: list[WordInput] = field(default_factory=list)
    uses_jinja: bool = False


def _paragraph_texts(data: bytes) -> list[str]:
    texts: list[str] = []
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        for name in zf.namelist():
            if _TEXT_PART.match(name):
                root = etree.fromstring(
                    zf.read(name), etree.XMLParser(resolve_entities=False, no_network=True)
                )
                texts += ["".join(t.text or "" for t in p.iter(_W + "t")) for p in root.iter(_W + "p")]
    return texts


def _has_mergefields(data: bytes) -> bool:
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        return any(b"MERGEFIELD" in zf.read(name) for name in zf.namelist() if _TEXT_PART.match(name))


def scan_word(path: Path) -> WordScan:
    """What can be filled in this Word file, in document order."""
    data = path.read_bytes()
    texts = _paragraph_texts(data)
    joined = "\n".join(texts)
    if _has_mergefields(data):
        style = MarkerStyle.MERGEFIELD
    elif _MARKER.search(joined):
        style = MarkerStyle.CUSTOM
    elif "{{" in joined or "{%" in joined:
        style = MarkerStyle.JINJA
    else:
        style = MarkerStyle.CUSTOM
    scan = WordScan(path, style, uses_jinja=style is MarkerStyle.JINJA)
    converted = compile_template(data, MarkerConfig(style=style))
    contexts: dict[str, str] = {}
    order: dict[str, int] = {}
    for text in texts:
        for m in _MARKER.finditer(text):
            contexts.setdefault(m.group(1).strip(), _short(text))
            order.setdefault(m.group(1).strip(), len(order))
    seen: set[str] = set()
    for name in sorted(converted.names, key=lambda n: order.get(n, len(order))):  # document order
        if name not in seen:
            seen.add(name)
            scan.inputs.append(WordInput(name, "marker", "", contexts.get(name, "")))
    for spot in process_spots(converted.data, SpotOptions()).spots:
        if spot.name not in seen:
            seen.add(spot.name)
            scan.inputs.append(WordInput(spot.name, spot.kind, spot.subtype, spot.context))
    return scan


def _short(text: str, limit: int = 100) -> str:
    text = re.sub(r"\s+", " ", text).strip()
    return text if len(text) <= limit else text[: limit - 1] + "…"


# ----------------------------------------------------------- suggestions
def _compact(text: str) -> str:
    words = re.findall(r"[a-z0-9]+", V.normalize_label(text, accents=True))
    return "".join(w for w in words if w not in _STOP)


def _tokens(text: str) -> frozenset[str]:
    return frozenset(
        w for w in re.findall(r"[a-z0-9]+", V.normalize_label(text, accents=True)) if w not in _STOP
    )


def similarity(a: str, b: str) -> float:
    ca, cb = _compact(a), _compact(b)
    if not ca or not cb:
        return 0.0
    if ca == cb:
        return 1.0
    if _tokens(a) == _tokens(b):
        return 0.95
    return difflib.SequenceMatcher(None, ca, cb).ratio()


def suggest(
    inputs: Sequence[WordInput], options: Sequence[ValueOption], *, threshold: float = 0.86
) -> dict[str, str]:
    """input name -> option path, only for clear matches (never for 'ngày', 'tháng', …)."""
    result: dict[str, str] = {}
    for item in inputs:
        if _compact(item.name) in _GENERIC or len(_compact(item.name)) < 3:
            continue
        scored = sorted(
            ((max(similarity(item.name, n) for n in (o.label, *o.names)), o.path) for o in options),
            reverse=True,
        )
        if not scored or scored[0][0] < threshold:
            continue
        if len(scored) > 1 and scored[1][0] > scored[0][0] - 0.02 and scored[1][1] != scored[0][1]:
            continue  # two values fit equally well: let the user choose
        result[item.name] = scored[0][1]
    return result


# --------------------------------------------------------------- profile
@dataclass(slots=True)
class WordChoice:
    """A Word file of the profile (already copied into the profile folder) and its bindings."""

    file: str
    scan: WordScan
    bindings: dict[str, str] = field(default_factory=dict)


def _title_column(primary: SheetPlan, parts: ProfileParts) -> str | None:
    """A column that names each row ('Tên công việc'), for file names."""
    best: str | None = None
    for i, header in enumerate(primary.headers):
        cid = parts.column_ids.get((primary.name, i))
        values = _column(primary.rows, i)
        if not cid or _kind(values) != "text" or len(set(map(str, values))) < len(values):
            continue
        if "ten" in _tokens(header):
            return cid
        best = best or cid
    return best


def build_profile(
    analysis: QuickAnalysis,
    words: Sequence[WordChoice],
    *,
    profile_id: str,
    name: str,
    one_file_per_row: bool = True,
    combined: bool = True,
) -> Profile:
    primary = analysis.primary
    if primary is None:
        raise ValueError("Chưa chọn sheet danh sách (mỗi dòng là một bộ tài liệu)")
    parts = plan_ids(analysis)
    options = {o.path: o for o in value_options(analysis)}
    used = {path for w in words for path in w.bindings.values()}
    unknown = sorted(used - set(options))
    if unknown:
        raise ValueError(f"unknown values: {unknown}")

    sheets: list[dict[str, Any]] = []
    relationships: list[dict[str, Any]] = []
    for sheet in analysis.sheets:
        sid = parts.sheet_ids.get(sheet.name)
        if sid is None:
            continue
        labels = list(sheet.info) if sheet.role == ROLE_INFO else sheet.headers
        columns = [
            {"id": parts.column_ids[(sheet.name, i)], "header": label, "required": False}
            for i, label in enumerate(labels)
            if (sheet.name, i) in parts.column_ids
        ]
        config: dict[str, Any] = {"id": sid, "match": {"name": sheet.name}, "columns": columns}
        if sheet.role == ROLE_INFO:
            config["layout"] = "key_value"
            relationships.append({"id": f"r_{sid}", "from": PRIMARY_ID, "to": sid, "keys": []})
        elif sheet.header_row and sheet.header_row > 1:
            config["header"] = {"mode": "fixed", "row": sheet.header_row}
        sheets.append(config)
        if sheet.role == ROLE_LOOKUP:
            for ci, ki in sheet.links:
                pcid = parts.column_ids[(primary.name, ci)]
                relationships.append(
                    {
                        "id": f"r_{pcid}",
                        "from": PRIMARY_ID,
                        "to": sid,
                        "keys": [{"from": pcid, "to": parts.column_ids[(sheet.name, ki)]}],
                        "key_normalize": {"casefold": True},
                    }
                )
    fields = [{"name": o.path, **o.rule} for o in options.values() if o.base or o.path in used]
    templates: list[dict[str, Any]] = []
    taken: set[str] = set()
    for choice in words:
        tid = suggest_identifier(Path(choice.file).stem, taken)
        templates.append(
            {
                "id": tid,
                "file": choice.file,
                "marker": {
                    "style": choice.scan.style.value,
                    "open": choice.scan.open,
                    "close": choice.scan.close,
                },
                "bindings": dict(choice.bindings),
                "missing_placeholder": "empty",  # unmatched inputs were left empty on purpose
            }
        )
    title = _title_column(primary, parts)
    merged_name = "{unit.index:02d}" + (f" - {{dong.{title}}}" if title else "") + ".docx"
    data: dict[str, Any] = {
        "profile": {
            "id": profile_id,
            "name": name,
            "description": f"Tạo nhanh từ {analysis.path.name}: mỗi dòng của sheet '{primary.name}' "
            "tạo một bộ tài liệu.",
        },
        "excel": {"primary": PRIMARY_ID, "sheets": sheets, "formula_policy": "cached_warn"},
        "relationships": relationships,
        "fields": fields,
        "templates": templates,
        "merge": {"order": [t["id"] for t in templates]},
        "generation": {"unit": "record", "output": "merged" if one_file_per_row else "separate"},
        "output": {
            "merged_filename": merged_name,
            "individual_filename": "{unit.index:02d} - {template.id}.docx",
            "combined_filename": "Tổng hợp - {profile.name}.docx" if combined else None,
        },
    }
    return Profile.model_validate(data)


def default_name(analysis: QuickAnalysis) -> str:
    """The project name when the workbook has one ('Tên công trình'), else the file name."""
    for sheet in analysis.sheets:
        if sheet.role == ROLE_INFO:
            for label, value in sheet.info.items():
                if {"ten", "cong", "trinh"} <= _tokens(label) or {"ten", "du", "an"} <= _tokens(label):
                    if isinstance(value, str) and value.strip():
                        return value.strip()[:80]
    return analysis.path.stem


def describe(analysis: QuickAnalysis) -> list[str]:
    """Plain sentences about what was found (shown in the wizard)."""
    lines: list[str] = []
    for sheet in analysis.sheets:
        if sheet.role == ROLE_LIST:
            lines.append(f"Sheet '{sheet.name}': {sheet.row_count} dòng – mỗi dòng sẽ tạo một bộ tài liệu.")
        elif sheet.role == ROLE_INFO:
            lines.append(
                f"Sheet '{sheet.name}': thông tin chung ({len(sheet.info)} mục), dùng cho mọi tài liệu."
            )
        elif sheet.role == ROLE_LOOKUP and analysis.primary is not None:
            cols = ", ".join(f"'{analysis.primary.headers[ci]}'" for ci, _ in sheet.links)
            lines.append(f"Sheet '{sheet.name}': bảng tra cứu theo cột {cols}.")
    return lines


def apply_bindings(words: Sequence[WordChoice], chosen: Mapping[tuple[int, str], str | None]) -> None:
    """Store the user's choices ((file index, input name) -> option path or None)."""
    for (index, name), path in chosen.items():
        if path:
            words[index].bindings[name] = path
        else:
            words[index].bindings.pop(name, None)
