"""Vietnamese number reading (đọc số thành chữ), common in contracts and invoices."""

from __future__ import annotations

from decimal import Decimal

_DIGITS = ["không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"]
_UNITS = ["", " nghìn", " triệu", " tỷ"]


def _read_triple(n: int, full: bool) -> str:
    """Read 0..999. *full* forces 'không trăm' / 'linh' for inner groups."""
    hundreds, rest = divmod(n, 100)
    tens, ones = divmod(rest, 10)
    words: list[str] = []
    if hundreds or full:
        words += [_DIGITS[hundreds], "trăm"]
    if tens == 0:
        if ones and (hundreds or full):
            words.append("linh")
    elif tens == 1:
        words.append("mười")
    else:
        words += [_DIGITS[tens], "mươi"]
    if ones:
        if ones == 1 and tens >= 2:
            words.append("mốt")
        elif ones == 5 and tens >= 1:
            words.append("lăm")
        elif ones == 4 and tens >= 2:
            words.append("tư")
        else:
            words.append(_DIGITS[ones])
    return " ".join(words)


def read_integer_vi(number: int) -> str:
    if number == 0:
        return "không"
    if number < 0:
        return "âm " + read_integer_vi(-number)
    # split into groups of 9 digits (tỷ) recursively, then thousands
    if number >= 1_000_000_000:
        high, low = divmod(number, 1_000_000_000)
        text = read_integer_vi(high) + " tỷ"
        if low:
            text += " " + _read_below_billion(low, full=True)
        return text
    return _read_below_billion(number, full=False)


def _read_below_billion(number: int, *, full: bool) -> str:
    groups: list[int] = []
    while number:
        number, part = divmod(number, 1000)
        groups.append(part)
    if not groups:
        return ""
    words: list[str] = []
    top = len(groups) - 1
    for index in range(top, -1, -1):
        part = groups[index]
        if part == 0:
            continue
        inner = full or index != top
        words.append(_read_triple(part, inner) + _UNITS[index])
    return " ".join(words)


def number_to_words_vi(value: int | float | Decimal, *, suffix: str = "", capitalize: bool = True) -> str:
    """Return Vietnamese words for *value* (fraction digits read one by one after 'phẩy')."""
    dec = value if isinstance(value, Decimal) else Decimal(repr(value) if isinstance(value, float) else value)
    negative = dec < 0
    dec = abs(dec)
    integer = int(dec)
    text = read_integer_vi(integer)
    fraction = dec - integer
    if fraction:
        digits = format(fraction.normalize(), "f").split(".")[1]
        text += " phẩy " + " ".join(_DIGITS[int(d)] for d in digits)
    if negative:
        text = "âm " + text
    if suffix:
        text += " " + suffix
    if capitalize:
        text = text[0].upper() + text[1:]
    return text
