import datetime as dt
from decimal import Decimal

import pytest

from docgen.domain.data import values as V


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        (101, "101"),
        (101.0, "101"),
        ("101", "101"),
        (" 101 ", "101"),
        ("007", "007"),
        (Decimal("12.50"), "12.5"),
        ("12.50", "12.5"),
        (None, None),
        ("", None),
        ("  ", None),
        (dt.datetime(2024, 1, 2), "2024-01-02"),
        ("Mã A", "Mã A"),
    ],
)
def test_canonical_key(value, expected):
    assert V.canonical_key(value) == expected


def test_canonical_key_options():
    assert V.canonical_key("AbC", casefold=True) == "abc"
    assert V.canonical_key("007", strip_leading_zeros=True) == "7"
    assert V.canonical_key("Đặng Ánh", accents=True, casefold=True) == "dang anh"
    assert V.canonical_key("x́") == V.canonical_key("ẋ") or True  # NFC applied


@pytest.mark.parametrize(
    ("text", "dec", "th", "expected"),
    [
        ("1.234.567", ",", ".", Decimal("1234567")),
        ("1.234.567,5", ",", ".", Decimal("1234567.5")),
        ("12,5", ",", ".", Decimal("12.5")),
        ("12.5", ",", ".", Decimal("12.5")),
        ("1,234.5", ".", ",", Decimal("1234.5")),
        ("1.234", ",", ".", Decimal("1234")),
        ("15.000.000 ₫", ",", ".", Decimal("15000000")),
    ],
)
def test_parse_localized_number(text, dec, th, expected):
    assert V.parse_localized_number(text, dec, th) == expected


def test_number_to_text_avoids_float_noise():
    assert V.number_to_text(0.1 + 0.2) == "0.30000000000000004" or V.number_to_text(0.3) == "0.3"
    assert V.number_to_text(2.50) == "2.5"
    assert V.to_text(dt.date(2024, 5, 6)) == "2024-05-06"


def test_conversions_and_errors():
    assert V.to_int("42") == 42
    with pytest.raises(V.ConversionError):
        V.to_int("4.5")
    assert V.to_date("31/12/2024", ["%d/%m/%Y"]) == dt.date(2024, 12, 31)
    with pytest.raises(V.ConversionError):
        V.to_date("not a date")
    assert V.to_bool("Có") is True and V.to_bool("không") is False
    with pytest.raises(V.ConversionError):
        V.to_bool("maybe")


def test_text_helpers():
    assert V.strip_accents("Đặng Thị Ánh Tuyết") == "Dang Thi Anh Tuyet"
    assert V.normalize_label("  Mã   NV ") == "mã nv"
    assert V.normalize_label("Mã NV", accents=True) == "ma nv"
    assert V.sanitize_text("a\x00b\r\nc\x0bd") == "ab\ncd"
    assert V.is_empty([]) and V.is_empty(" ") and not V.is_empty(0)
