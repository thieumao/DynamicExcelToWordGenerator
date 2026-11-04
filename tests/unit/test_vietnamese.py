import pytest

from docgen.domain.mapping.vietnamese import number_to_words_vi, read_integer_vi


@pytest.mark.parametrize(
    ("n", "words"),
    [
        (0, "không"),
        (5, "năm"),
        (10, "mười"),
        (11, "mười một"),
        (14, "mười bốn"),
        (15, "mười lăm"),
        (21, "hai mươi mốt"),
        (24, "hai mươi tư"),
        (25, "hai mươi lăm"),
        (101, "một trăm linh một"),
        (105, "một trăm linh năm"),
        (115, "một trăm mười lăm"),
        (1000, "một nghìn"),
        (1005, "một nghìn không trăm linh năm"),
        (1_000_000, "một triệu"),
        (125_000_000, "một trăm hai mươi lăm triệu"),
        (1_200_000_000, "một tỷ hai trăm triệu"),
        (1_000_000_000_000, "một nghìn tỷ"),
        (2_000_010, "hai triệu không trăm mười"),
    ],
)
def test_read_integer(n, words):
    assert read_integer_vi(n) == words


def test_number_to_words_with_suffix_and_fraction():
    assert number_to_words_vi(49_000_000, suffix="đồng") == "Bốn mươi chín triệu đồng"
    assert number_to_words_vi(12.5, capitalize=False) == "mười hai phẩy năm"
    assert number_to_words_vi(-3, capitalize=False) == "âm ba"
