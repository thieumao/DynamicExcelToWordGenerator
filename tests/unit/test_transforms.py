import datetime as dt
from decimal import Decimal

import pytest

from docgen.domain.mapping.transforms import DEFAULT_REGISTRY as REG
from docgen.domain.mapping.transforms import TransformContext, TransformError

CTX = TransformContext()


def apply(name, value, /, **params):
    return REG.apply(name, value, CTX, params)


def test_text_transforms():
    assert apply("trim", "  a ") == "a"
    assert apply("title_case", "nguyễn văn ánh") == "Nguyễn Văn Ánh"
    assert apply("upper", "đặng") == "ĐẶNG"
    assert apply("replace", "a-b", old="-", new="/") == "a/b"
    assert apply("pad_left", 7, width=3) == "007"
    assert apply("regex_extract", "HD-2024-001", pattern=r"\d{4}") == "2024"
    assert apply("strip_accents", "Hà Nội") == "Ha Noi"


def test_format_transforms_use_locale():
    assert apply("format_number", 1234567.891, decimals=2) == "1.234.567,89"
    assert apply("format_currency", 15000000) == "15.000.000 ₫"
    assert apply("format_currency", 15.5, symbol="$", decimals=2, position="prefix") == "$15,50"
    assert apply("format_date", dt.datetime(2024, 3, 9)) == "09/03/2024"
    assert apply("format_date", "2024-03-09", format="%Y/%m/%d") == "2024/03/09"
    assert apply("format_percent", 0.125, decimals=1) == "12,5%"
    assert apply("format_bool", "yes") == "Có"
    assert apply("number_to_words_vi", 21, suffix="đồng") == "Hai mươi mốt đồng"


def test_list_and_logic_transforms():
    assert apply("join", ["a", None, "b"], separator="; ") == "a; b"
    assert apply("split", "a, b,,c") == ["a", "b", "c"]
    assert apply("sum", ["1", 2, None]) == Decimal(3)
    assert apply("count", []) == 0
    assert apply("unique", ["x", "x", "y"]) == ["x", "y"]
    assert apply("default", None, value="N/A") == "N/A"
    assert apply("default", "v", value="N/A") == "v"
    assert apply("map_values", "M", mapping={"M": "Nam", "F": "Nữ"}) == "Nam"
    assert apply("map_values", 1, mapping={"1": "Một"}) == "Một"
    assert apply("map_values", "X", mapping={"M": "Nam"}, default="?") == "?"


def test_empty_values_pass_through():
    assert apply("format_date", None) is None
    assert apply("upper", "") == ""


def test_errors_are_reported():
    with pytest.raises(TransformError):
        apply("to_int", "abc")
    with pytest.raises(TransformError):
        apply("nope", "x")
    with pytest.raises(TransformError):
        apply("regex_replace", "x", pattern="a" * 300)


def test_validate_params():
    assert REG.validate_params("trim", {}) == []
    assert REG.validate_params("replace", {}) == ["transform 'replace' requires parameter 'old'"]
    assert "no parameter" in REG.validate_params("trim", {"x": 1})[0]
    assert "must be one of" in REG.validate_params("format_currency", {"position": "middle"})[0]
