"""Transform registry.

Every transform is a pure function ``f(value, ctx, **params)``. The same
registry backs YAML transform pipelines, Jinja filters in expressions and
Jinja filters inside Word templates, so users learn one vocabulary.
"""

from __future__ import annotations

import datetime as dt
import re
from collections.abc import Callable, Iterable, Mapping
from dataclasses import dataclass, field
from decimal import ROUND_HALF_UP, Decimal
from typing import Any

from docgen.domain.data import values as V
from docgen.domain.mapping.vietnamese import number_to_words_vi


class TransformError(ValueError):
    """A transform could not be applied to a value."""


@dataclass(frozen=True, slots=True)
class TransformContext:
    """Locale/formatting settings visible to transforms."""

    date_format: str = "%d/%m/%Y"
    datetime_format: str = "%d/%m/%Y %H:%M"
    decimal_separator: str = ","
    thousands_separator: str = "."
    currency_symbol: str = "₫"
    currency_position: str = "suffix"
    currency_decimals: int = 0
    true_text: str = "Có"
    false_text: str = "Không"
    input_date_formats: tuple[str, ...] = ("%d/%m/%Y", "%Y-%m-%d")


@dataclass(frozen=True, slots=True)
class ParamSpec:
    type: str  # "str" | "int" | "number" | "bool" | "mapping" | "any"
    required: bool = False
    default: Any = None
    description: str = ""
    choices: tuple[str, ...] = ()


TransformFunc = Callable[..., Any]


@dataclass(frozen=True, slots=True)
class TransformDef:
    name: str
    func: TransformFunc
    description: str
    params: Mapping[str, ParamSpec] = field(default_factory=dict)
    category: str = "text"
    accepts_empty: bool = False
    """If False, empty input values pass through unchanged without calling ``func``."""


class TransformRegistry:
    def __init__(self) -> None:
        self._defs: dict[str, TransformDef] = {}

    def register(self, definition: TransformDef) -> None:
        if definition.name in self._defs:
            raise ValueError(f"transform '{definition.name}' already registered")
        self._defs[definition.name] = definition

    def get(self, name: str) -> TransformDef | None:
        return self._defs.get(name)

    def names(self) -> list[str]:
        return sorted(self._defs)

    def definitions(self) -> list[TransformDef]:
        return [self._defs[n] for n in self.names()]

    def validate_params(self, name: str, params: Mapping[str, Any]) -> list[str]:
        """Return human-readable problems with a transform call (empty = OK)."""
        definition = self._defs.get(name)
        if definition is None:
            return [f"unknown transform '{name}'"]
        problems = []
        for key in params:
            if key not in definition.params:
                problems.append(
                    f"transform '{name}' has no parameter '{key}' "
                    f"(allowed: {sorted(definition.params) or 'none'})"
                )
        for key, spec in definition.params.items():
            if spec.required and key not in params:
                problems.append(f"transform '{name}' requires parameter '{key}'")
            if key in params and spec.choices and params[key] not in spec.choices:
                problems.append(f"transform '{name}': '{key}' must be one of {list(spec.choices)}")
        return problems

    def apply(self, name: str, value: Any, ctx: TransformContext, params: Mapping[str, Any]) -> Any:
        definition = self._defs.get(name)
        if definition is None:
            raise TransformError(f"unknown transform '{name}'")
        if not definition.accepts_empty and V.is_empty(value):
            return value
        kwargs = {k: spec.default for k, spec in definition.params.items() if not spec.required}
        kwargs.update(params)
        try:
            return definition.func(value, ctx, **kwargs)
        except TransformError:
            raise
        except (V.ConversionError, ValueError, TypeError, ArithmeticError) as exc:
            raise TransformError(f"{name}: {exc}") from exc


# ------------------------------------------------------------------ helpers
def _text(value: Any) -> str:
    return V.to_text(value)


def _decimal(value: Any, ctx: TransformContext) -> Decimal:
    if isinstance(value, str):
        try:
            return V.to_decimal(value)
        except V.ConversionError:
            return V.to_decimal(
                value,
                decimal_separator=ctx.decimal_separator,
                thousands_separator=ctx.thousands_separator,
            )
    return V.to_decimal(value)


def format_decimal(value: Decimal, decimals: int, decimal_separator: str, thousands_separator: str) -> str:
    quantum = Decimal(1).scaleb(-decimals)
    rounded = value.quantize(quantum, rounding=ROUND_HALF_UP)
    text = f"{rounded:,.{decimals}f}"
    return text.replace(",", "\x00").replace(".", decimal_separator).replace("\x00", thousands_separator)


def _iter(value: Any) -> list[Any]:
    if value is None:
        return []
    if isinstance(value, (list, tuple, set)):
        return list(value)
    return [value]


# --------------------------------------------------------------- transforms
def _trim(v: Any, ctx: TransformContext) -> Any:
    return v.strip() if isinstance(v, str) else v


def _collapse_spaces(v: Any, ctx: TransformContext) -> Any:
    return re.sub(r"\s+", " ", v).strip() if isinstance(v, str) else v


def _upper(v: Any, ctx: TransformContext) -> str:
    return _text(v).upper()


def _lower(v: Any, ctx: TransformContext) -> str:
    return _text(v).lower()


def _title_case(v: Any, ctx: TransformContext) -> str:
    return " ".join(w[:1].upper() + w[1:].lower() for w in _text(v).split())


def _capitalize(v: Any, ctx: TransformContext) -> str:
    t = _text(v)
    return t[:1].upper() + t[1:]


def _replace(v: Any, ctx: TransformContext, old: str, new: str = "") -> str:
    return _text(v).replace(old, new)


def _regex_replace(v: Any, ctx: TransformContext, pattern: str, repl: str = "") -> str:
    if len(pattern) > 200:
        raise TransformError("regex pattern too long (max 200 characters)")
    return re.sub(pattern, repl, _text(v))


def _regex_extract(v: Any, ctx: TransformContext, pattern: str, group: int = 0) -> Any:
    if len(pattern) > 200:
        raise TransformError("regex pattern too long (max 200 characters)")
    m = re.search(pattern, _text(v))
    return m.group(group) if m else None


def _slice(v: Any, ctx: TransformContext, start: int = 0, end: int | None = None) -> str:
    return _text(v)[start:end]


def _pad_left(v: Any, ctx: TransformContext, width: int, char: str = "0") -> str:
    return _text(v).rjust(width, char[:1] or " ")


def _pad_right(v: Any, ctx: TransformContext, width: int, char: str = " ") -> str:
    return _text(v).ljust(width, char[:1] or " ")


def _prefix(v: Any, ctx: TransformContext, text: str) -> str:
    return text + _text(v)


def _suffix(v: Any, ctx: TransformContext, text: str) -> str:
    return _text(v) + text


def _strip_accents(v: Any, ctx: TransformContext) -> str:
    return V.strip_accents(_text(v))


def _to_string(v: Any, ctx: TransformContext) -> str:
    return _text(v)


def _to_int(v: Any, ctx: TransformContext) -> int:
    return int(_decimal(v, ctx).to_integral_value(rounding=ROUND_HALF_UP)) if not isinstance(v, int) else v


def _to_decimal(v: Any, ctx: TransformContext) -> Decimal:
    return _decimal(v, ctx)


def _to_date(v: Any, ctx: TransformContext, format: str | None = None) -> dt.date:
    formats = (format,) if format else ctx.input_date_formats
    return V.to_date(v, formats)


def _to_bool(v: Any, ctx: TransformContext) -> bool:
    return V.to_bool(v)


def _format_date(v: Any, ctx: TransformContext, format: str | None = None) -> str:
    if isinstance(v, dt.datetime) and format is None:
        return v.strftime(ctx.date_format)
    date_value = v if isinstance(v, (dt.date, dt.datetime)) else V.to_date(v, ctx.input_date_formats)
    return date_value.strftime(format or ctx.date_format)


def _format_datetime(v: Any, ctx: TransformContext, format: str | None = None) -> str:
    value = V.to_datetime(v)
    return value.strftime(format or ctx.datetime_format)


def _format_number(
    v: Any,
    ctx: TransformContext,
    decimals: int = 0,
    decimal_separator: str | None = None,
    thousands_separator: str | None = None,
) -> str:
    return format_decimal(
        _decimal(v, ctx),
        decimals,
        ctx.decimal_separator if decimal_separator is None else decimal_separator,
        ctx.thousands_separator if thousands_separator is None else thousands_separator,
    )


def _format_currency(
    v: Any,
    ctx: TransformContext,
    symbol: str | None = None,
    decimals: int | None = None,
    position: str | None = None,
) -> str:
    number = format_decimal(
        _decimal(v, ctx),
        ctx.currency_decimals if decimals is None else decimals,
        ctx.decimal_separator,
        ctx.thousands_separator,
    )
    sym = ctx.currency_symbol if symbol is None else symbol
    pos = position or ctx.currency_position
    if not sym:
        return number
    return f"{sym}{number}" if pos == "prefix" else f"{number} {sym}"


def _format_percent(v: Any, ctx: TransformContext, decimals: int = 0, scale: bool = True) -> str:
    dec = _decimal(v, ctx)
    if scale:
        dec *= 100
    return format_decimal(dec, decimals, ctx.decimal_separator, ctx.thousands_separator) + "%"


def _format_bool(v: Any, ctx: TransformContext, true: str | None = None, false: str | None = None) -> str:
    flag = V.to_bool(v)
    return (
        (true if true is not None else ctx.true_text)
        if flag
        else (false if false is not None else ctx.false_text)
    )


def _number_to_words_vi(v: Any, ctx: TransformContext, suffix: str = "", capitalize: bool = True) -> str:
    return number_to_words_vi(_decimal(v, ctx), suffix=suffix, capitalize=capitalize)


def _default(v: Any, ctx: TransformContext, value: Any = "") -> Any:
    return value if V.is_empty(v) else v


def _map_values(
    v: Any,
    ctx: TransformContext,
    mapping: Mapping[str, Any],
    default: Any = None,
    keep_unmapped: bool = True,
) -> Any:
    key = V.canonical_key(v, casefold=False)
    normalized = {V.canonical_key(k, casefold=False): val for k, val in mapping.items()}
    if key in normalized:
        return normalized[key]
    if default is not None:
        return default
    return v if keep_unmapped else None


def _split(v: Any, ctx: TransformContext, separator: str = ",", trim: bool = True) -> list[str]:
    parts = _text(v).split(separator)
    return [p.strip() for p in parts if p.strip()] if trim else parts


def _join(v: Any, ctx: TransformContext, separator: str = ", ", skip_empty: bool = True) -> str:
    items = [_text(x) for x in _iter(v) if not (skip_empty and V.is_empty(x))]
    return separator.join(items)


def _first(v: Any, ctx: TransformContext) -> Any:
    items = _iter(v)
    return items[0] if items else None


def _last(v: Any, ctx: TransformContext) -> Any:
    items = _iter(v)
    return items[-1] if items else None


def _count(v: Any, ctx: TransformContext) -> int:
    return len([x for x in _iter(v) if not V.is_empty(x)])


def _sum(v: Any, ctx: TransformContext) -> Decimal:
    return sum((_decimal(x, ctx) for x in _iter(v) if not V.is_empty(x)), Decimal(0))


def _unique(v: Any, ctx: TransformContext) -> list[Any]:
    seen: set[str] = set()
    result = []
    for item in _iter(v):
        key = V.to_text(item)
        if key not in seen:
            seen.add(key)
            result.append(item)
    return result


def _round(v: Any, ctx: TransformContext, decimals: int = 0) -> Decimal:
    return _decimal(v, ctx).quantize(Decimal(1).scaleb(-decimals), rounding=ROUND_HALF_UP)


def _multiply(v: Any, ctx: TransformContext, by: float | int | str) -> Decimal:
    return _decimal(v, ctx) * V.to_decimal(by)


def _add(v: Any, ctx: TransformContext, value: float | int | str) -> Decimal:
    return _decimal(v, ctx) + V.to_decimal(value)


def _add_days(v: Any, ctx: TransformContext, days: int) -> dt.date:
    base = v if isinstance(v, dt.date) else V.to_date(v, ctx.input_date_formats)
    return base + dt.timedelta(days=days)


def _lines(v: Any, ctx: TransformContext) -> str:
    return "\n".join(_text(x) for x in _iter(v) if not V.is_empty(x))


def _is_checked(v: Any) -> bool:
    if V.is_empty(v):
        return False
    try:
        return V.to_bool(v)
    except V.ConversionError:
        return True  # any other non-empty mark ("✓", "x", a name) means checked


def _checkbox_symbol(v: Any, ctx: TransformContext, checked: str = "☒", unchecked: str = "☐") -> str:
    return checked if _is_checked(v) else unchecked


def _checkbox_value(v: Any, ctx: TransformContext) -> str:
    return "1" if _is_checked(v) else "0"


def _P(
    type_: str,
    required: bool = False,
    default: Any = None,
    description: str = "",
    choices: Iterable[str] = (),
) -> ParamSpec:
    return ParamSpec(type_, required, default, description, tuple(choices))


def build_default_registry() -> TransformRegistry:
    reg = TransformRegistry()
    defs = [
        TransformDef("trim", _trim, "Remove leading/trailing spaces"),
        TransformDef("collapse_spaces", _collapse_spaces, "Collapse repeated whitespace"),
        TransformDef("upper", _upper, "UPPER CASE"),
        TransformDef("lower", _lower, "lower case"),
        TransformDef("title_case", _title_case, "Title Case Each Word"),
        TransformDef("capitalize", _capitalize, "Capitalise first letter"),
        TransformDef(
            "replace", _replace, "Replace text", {"old": _P("str", True), "new": _P("str", default="")}
        ),
        TransformDef(
            "regex_replace",
            _regex_replace,
            "Replace by regular expression",
            {"pattern": _P("str", True), "repl": _P("str", default="")},
        ),
        TransformDef(
            "regex_extract",
            _regex_extract,
            "Extract by regular expression",
            {"pattern": _P("str", True), "group": _P("int", default=0)},
        ),
        TransformDef(
            "slice", _slice, "Substring by position", {"start": _P("int", default=0), "end": _P("int")}
        ),
        TransformDef(
            "pad_left",
            _pad_left,
            "Pad on the left",
            {"width": _P("int", True), "char": _P("str", default="0")},
        ),
        TransformDef(
            "pad_right",
            _pad_right,
            "Pad on the right",
            {"width": _P("int", True), "char": _P("str", default=" ")},
        ),
        TransformDef("prefix", _prefix, "Add text before", {"text": _P("str", True)}),
        TransformDef("suffix", _suffix, "Add text after", {"text": _P("str", True)}),
        TransformDef("strip_accents", _strip_accents, "Remove Vietnamese diacritics"),
        TransformDef("to_string", _to_string, "Convert to text", category="type"),
        TransformDef("to_int", _to_int, "Convert to whole number", category="type"),
        TransformDef("to_decimal", _to_decimal, "Convert to decimal number", category="type"),
        TransformDef("to_date", _to_date, "Parse a date", {"format": _P("str")}, category="type"),
        TransformDef("to_bool", _to_bool, "Parse yes/no", category="type"),
        TransformDef("format_date", _format_date, "Format a date", {"format": _P("str")}, category="format"),
        TransformDef(
            "format_datetime",
            _format_datetime,
            "Format date and time",
            {"format": _P("str")},
            category="format",
        ),
        TransformDef(
            "format_number",
            _format_number,
            "Format a number with separators",
            {
                "decimals": _P("int", default=0),
                "decimal_separator": _P("str"),
                "thousands_separator": _P("str"),
            },
            category="format",
        ),
        TransformDef(
            "format_currency",
            _format_currency,
            "Format money",
            {"symbol": _P("str"), "decimals": _P("int"), "position": _P("str", choices=("prefix", "suffix"))},
            category="format",
        ),
        TransformDef(
            "format_percent",
            _format_percent,
            "Format a percentage",
            {"decimals": _P("int", default=0), "scale": _P("bool", default=True)},
            category="format",
        ),
        TransformDef(
            "format_bool",
            _format_bool,
            "Yes/No text",
            {"true": _P("str"), "false": _P("str")},
            category="format",
        ),
        TransformDef(
            "checkbox_symbol",
            _checkbox_symbol,
            "Check box symbol from yes/no",
            {"checked": _P("str", default="☒"), "unchecked": _P("str", default="☐")},
            category="format",
            accepts_empty=True,
        ),
        TransformDef(
            "checkbox_value",
            _checkbox_value,
            "Check box state from yes/no",
            category="format",
            accepts_empty=True,
        ),
        TransformDef(
            "number_to_words_vi",
            _number_to_words_vi,
            "Read number in Vietnamese words",
            {"suffix": _P("str", default=""), "capitalize": _P("bool", default=True)},
            category="format",
        ),
        TransformDef(
            "default",
            _default,
            "Fallback when empty",
            {"value": _P("any", default="")},
            accepts_empty=True,
            category="logic",
        ),
        TransformDef(
            "map_values",
            _map_values,
            "Translate codes using a table",
            {"mapping": _P("mapping", True), "default": _P("any"), "keep_unmapped": _P("bool", default=True)},
            category="logic",
        ),
        TransformDef(
            "split",
            _split,
            "Split text into a list",
            {"separator": _P("str", default=","), "trim": _P("bool", default=True)},
            category="list",
        ),
        TransformDef(
            "join",
            _join,
            "Join a list into text",
            {"separator": _P("str", default=", "), "skip_empty": _P("bool", default=True)},
            category="list",
        ),
        TransformDef("lines", _lines, "Join a list with line breaks", category="list"),
        TransformDef("first", _first, "First item of a list", category="list"),
        TransformDef("last", _last, "Last item of a list", category="list"),
        TransformDef("count", _count, "Number of non-empty items", category="list", accepts_empty=True),
        TransformDef("sum", _sum, "Sum of numbers", category="list", accepts_empty=True),
        TransformDef("unique", _unique, "Remove duplicates", category="list"),
        TransformDef(
            "round", _round, "Round a number", {"decimals": _P("int", default=0)}, category="number"
        ),
        TransformDef("multiply", _multiply, "Multiply", {"by": _P("number", True)}, category="number"),
        TransformDef("add", _add, "Add", {"value": _P("number", True)}, category="number"),
        TransformDef("add_days", _add_days, "Shift a date", {"days": _P("int", True)}, category="number"),
    ]
    for d in defs:
        reg.register(d)
    return reg


DEFAULT_REGISTRY = build_default_registry()
