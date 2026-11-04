"""Value helpers: emptiness, text normalisation and type conversion.

All conversions are explicit and raise :class:`ConversionError` with a reason
instead of guessing, so callers can report the sheet/row/column precisely.
"""

from __future__ import annotations

import datetime as dt
import re
import unicodedata
from collections.abc import Iterable
from decimal import Decimal, InvalidOperation
from typing import Any

EXCEL_ERROR_VALUES = frozenset(
    {
        "#NULL!",
        "#DIV/0!",
        "#VALUE!",
        "#REF!",
        "#NAME?",
        "#NUM!",
        "#N/A",
        "#GETTING_DATA",
        "#SPILL!",
        "#CALC!",
        "#FIELD!",
        "#BLOCKED!",
        "#CONNECT!",
        "#BUSY!",
        "#UNKNOWN!",
    }
)
CURRENCY_SYMBOLS = ("VNĐ", "VND", "₫", "đ", "$", "€", "%")
_WS = re.compile(r"\s+")
_XML_INVALID = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]")
TRUE_WORDS = frozenset({"true", "yes", "y", "1", "x", "có", "co", "đúng", "dung", "t"})
FALSE_WORDS = frozenset({"false", "no", "n", "0", "không", "khong", "sai", "f", ""})


class ConversionError(ValueError):
    """A value could not be converted to the requested type."""


def is_empty(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return value.strip() == ""
    if isinstance(value, (list, tuple, set, dict)):
        return len(value) == 0
    return False


def strip_accents(text: str) -> str:
    """Remove Vietnamese (and other) diacritics: 'Đặng Thị Ánh' -> 'Dang Thi Anh'."""
    text = text.replace("đ", "d").replace("Đ", "D")
    decomposed = unicodedata.normalize("NFD", text)
    return unicodedata.normalize("NFC", "".join(ch for ch in decomposed if unicodedata.category(ch) != "Mn"))


def normalize_label(
    text: Any,
    *,
    casefold: bool = True,
    accents: bool = False,
    collapse_whitespace: bool = True,
) -> str:
    """Normalise a header/alias for matching. ``accents=True`` strips diacritics."""
    if text is None:
        return ""
    result = unicodedata.normalize("NFC", str(text)).strip()
    if collapse_whitespace:
        result = _WS.sub(" ", result)
    if accents:
        result = strip_accents(result)
    if casefold:
        result = result.casefold()
    return result


def sanitize_text(value: str) -> str:
    """NFC-normalise, remove characters invalid in XML 1.0 and normalise line endings.

    NFC matters for Vietnamese: Windows-1258 files and macOS store tones as combining
    marks ('a' + U+0303) while Excel normally stores precomposed letters ('ã').
    """
    if not value.isascii():
        value = unicodedata.normalize("NFC", value)
    value = value.replace("\r\n", "\n").replace("\r", "\n")
    value = value.replace("_x000D_", "")
    return _XML_INVALID.sub("", value)


def number_to_text(value: int | float | Decimal) -> str:
    """Render a number without float noise: 101.0 -> '101', 0.1 + 0.2 -> '0.3'."""
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, int):
        return str(value)
    dec = Decimal(repr(value)) if isinstance(value, float) else value
    if dec == dec.to_integral_value():
        return str(dec.quantize(Decimal(1)))
    text = format(dec.normalize(), "f")
    return text


def to_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, (int, float, Decimal)):
        return number_to_text(value)
    if isinstance(value, dt.datetime):
        return value.isoformat(sep=" ", timespec="seconds")
    if isinstance(value, dt.date):
        return value.isoformat()
    if isinstance(value, dt.time):
        return value.isoformat(timespec="seconds")
    return str(value)


def to_decimal(value: Any, *, decimal_separator: str = ".", thousands_separator: str = "") -> Decimal:
    if isinstance(value, bool):
        raise ConversionError("boolean is not a number")
    if isinstance(value, Decimal):
        return value
    if isinstance(value, int):
        return Decimal(value)
    if isinstance(value, float):
        return Decimal(repr(value))
    if isinstance(value, str):
        text = value.strip().replace(" ", "").replace(" ", "")
        if not text:
            raise ConversionError("empty text")
        for symbol in CURRENCY_SYMBOLS:
            text = text.replace(symbol, "")
        if thousands_separator:
            text = text.replace(thousands_separator, "")
        if decimal_separator != ".":
            text = text.replace(decimal_separator, ".")
        try:
            return Decimal(text)
        except InvalidOperation as exc:
            raise ConversionError(f"'{value}' is not a number") from exc
    raise ConversionError(f"cannot convert {type(value).__name__} to number")


def parse_localized_number(text: str, decimal_separator: str, thousands_separator: str) -> Decimal:
    """Parse '1.234.567,5' (vi) or '1,234,567.5' (en) without guessing wrongly on '12.5'.

    Grouped thousands are only accepted in the strict 3-digit pattern, so '1.234' with
    thousands '.' is 1234 while '12.5' stays 12.5 for a profile using ',' as decimal.
    """
    t = text.strip().replace("\u00a0", "").replace(" ", "")
    for symbol in CURRENCY_SYMBOLS:
        t = t.replace(symbol, "")
    th, dec = re.escape(thousands_separator), re.escape(decimal_separator)
    if thousands_separator and re.fullmatch(rf"[+-]?\d{{1,3}}({th}\d{{3}})+({dec}\d+)?", t):
        return to_decimal(t, decimal_separator=decimal_separator, thousands_separator=thousands_separator)
    if decimal_separator != "." and decimal_separator in t:
        return to_decimal(t, decimal_separator=decimal_separator, thousands_separator="")
    return to_decimal(t)


def to_int(value: Any, **kwargs: str) -> int:
    dec = to_decimal(value, **kwargs)
    if dec != dec.to_integral_value():
        raise ConversionError(f"'{value}' is not a whole number")
    return int(dec)


def to_date(value: Any, formats: Iterable[str] = ("%d/%m/%Y", "%Y-%m-%d")) -> dt.date:
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if isinstance(value, str):
        text = value.strip()
        for fmt in formats:
            try:
                return dt.datetime.strptime(text, fmt).date()
            except ValueError:
                continue
        try:
            return dt.date.fromisoformat(text[:10])
        except ValueError as exc:
            raise ConversionError(f"'{value}' does not match date formats {list(formats)}") from exc
    raise ConversionError(f"cannot convert {type(value).__name__} to date")


def to_datetime(value: Any, formats: Iterable[str] = ("%d/%m/%Y %H:%M", "%Y-%m-%d %H:%M:%S")) -> dt.datetime:
    if isinstance(value, dt.datetime):
        return value
    if isinstance(value, dt.date):
        return dt.datetime(value.year, value.month, value.day)
    if isinstance(value, str):
        text = value.strip()
        for fmt in formats:
            try:
                return dt.datetime.strptime(text, fmt)
            except ValueError:
                continue
        try:
            return dt.datetime.fromisoformat(text)
        except ValueError as exc:
            raise ConversionError(f"'{value}' is not a date/time") from exc
    raise ConversionError(f"cannot convert {type(value).__name__} to datetime")


def to_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float, Decimal)):
        return value != 0
    if isinstance(value, str):
        word = value.strip().casefold()
        if word in TRUE_WORDS:
            return True
        if word in FALSE_WORDS:
            return False
    raise ConversionError(f"'{value}' is not a yes/no value")


def canonical_key(
    value: Any,
    *,
    trim: bool = True,
    casefold: bool = False,
    numeric_text: bool = True,
    strip_leading_zeros: bool = False,
    accents: bool = False,
) -> str | None:
    """Canonical text used for lookup keys; ``None`` for empty values.

    With ``numeric_text`` an Excel number 101, 101.0 and the text '101' compare equal.
    """
    if value is None:
        return None
    if isinstance(value, (int, float, Decimal)) and not isinstance(value, bool):
        text = number_to_text(value)
    elif isinstance(value, dt.datetime):
        text = value.isoformat(sep=" ") if value.time() != dt.time() else value.date().isoformat()
    elif isinstance(value, dt.date):
        text = value.isoformat()
    else:
        text = unicodedata.normalize("NFC", str(value))
        if numeric_text:
            stripped = text.strip()
            if re.fullmatch(r"[+-]?\d+(\.\d+)?", stripped) and not (
                stripped.startswith("0") and len(stripped) > 1 and "." not in stripped
            ):
                text = number_to_text(Decimal(stripped))
    if trim:
        text = text.strip()
    if strip_leading_zeros and re.fullmatch(r"0+\d+", text):
        text = text.lstrip("0") or "0"
    if accents:
        text = strip_accents(text)
    if casefold:
        text = text.casefold()
    return text if text != "" else None
