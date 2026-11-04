import datetime as dt
from decimal import Decimal
from pathlib import Path

import pytest

from docgen.domain.diagnostics import Severity
from docgen.domain.profile.io import parse_profile_text
from docgen.infrastructure.excel.loader import WorkbookLoader
from docgen.infrastructure.excel.sources import column_index, column_letter, open_workbook
from docgen.testing.xlsx_builder import build_xls, build_xlsx, inject_cached_values


def profile(columns: str, sheet: str = "{name: Data}", extra: str = "", sheet_extra: str = ""):
    return parse_profile_text(f"""
profile: {{id: t, name: T}}
excel:
  primary: d
  {extra}
  sheets:
    - id: d
      match: {sheet}
      {sheet_extra}
      columns: {columns}
""")


def load(path: Path, prof):
    return WorkbookLoader(prof.excel, prof.formatting).load(path)


def codes(result):
    return {(i.code, i.severity) for i in result.issues}


def test_sheet_matching_by_alias_regex_index(tmp_path):
    wb = build_xlsx(tmp_path / "w.xlsx", {"Intro": [["x"]], "Nhân viên 2024": [["Mã"], ["E1"]]})
    cols = "[{id: code, header: Mã}]"
    for match in [
        '{name: Staff, aliases: ["Nhân viên 2024"]}',
        "{regex: 'nh[aâ]n vi[eê]n'}",
        "{index: 2}",
        '{name: "nhân viên 2024"}',
    ]:
        result = load(wb, profile(cols, match))
        assert result.dataset.table("d").rows == [("E1",)], match
    result = load(wb, profile(cols, "{name: Staff}"))
    assert ("sheet.not_found", Severity.ERROR) in codes(result)


def test_header_positions_and_missing_columns(tmp_path):
    wb = build_xlsx(
        tmp_path / "w.xlsx",
        {
            "Data": [
                ["Title"],
                [],
                [None, "Mã NV", "Họ và tên", "Ghi chú"],
                [None, "E1", "An", None],
                [None, None, None, None],
                [None, "E2", "Bình", "x"],
            ]
        },
    )
    prof = profile(
        """[{id: code, header: "Ma NV"}, {id: name, header: "Họ tên", aliases: ["Họ và tên"]},
               {id: phone, header: Phone, required: false}, {id: email, header: Email}]""",
        extra="header_matching: {strip_accents: true}",
    )
    result = load(wb, prof)
    info = result.sheets["d"]
    assert info.header_row == 3
    assert ("column.not_found", Severity.WARNING) in codes(result)  # optional phone
    assert ("column.not_found", Severity.ERROR) in codes(result)  # required email
    assert "d" not in result.dataset.tables


def test_values_rows_and_types(tmp_path):
    wb = build_xlsx(
        tmp_path / "w.xlsx",
        {
            "Data": [
                ["Code", "Amount", "Date", "Flag", "Text"],
                ["001", "1.234.567", "31/12/2024", "Có", "  Hà Nội "],
                [7, 12.5, dt.datetime(2024, 1, 2), "no", None],
                [None, None, None, None, None],
                ["X", "abc", "bad", "maybe", "#N/A"],
            ]
        },
    )
    prof = profile("""[{id: code, header: Code, type: string}, {id: amount, header: Amount, type: decimal},
               {id: date, header: Date, type: date}, {id: flag, header: Flag, type: boolean},
               {id: text, header: Text}]""")
    result = load(wb, prof)
    table = result.dataset.table("d")
    assert table.rows[0] == ("001", Decimal("1234567"), dt.date(2024, 12, 31), True, "  Hà Nội ")
    assert table.rows[1] == ("7", Decimal("12.5"), dt.date(2024, 1, 2), False, None)
    assert table.row_numbers == [2, 3, 5]  # blank row skipped, numbers preserved
    errors = [i for i in result.issues if i.code == "cell.type_mismatch"]
    assert {(i.row, i.column) for i in errors} == {(5, "Amount"), (5, "Date"), (5, "Flag")}
    assert ("cell.error_value", Severity.WARNING) in codes(result)
    assert (2, 1) in table.invalid and table.invalid[(2, 1)][0] == "cell.type_mismatch"


def test_position_fallback_and_headerless(tmp_path):
    wb = build_xlsx(tmp_path / "w.xlsx", {"Data": [["a", "b", "c"], [1, 2, 3]]})
    result = load(wb, profile("[{id: x, header: Missing, position: C}]"))
    assert result.dataset.table("d").rows[0] == (3,)
    result = load(wb, profile("[{id: x, position: 2}]", sheet_extra="header: {mode: none}"))
    assert result.dataset.table("d").rows == [("b",), (2,)]


def test_stop_after_blank_rows(tmp_path):
    rows = [["K"], ["a"], [None], [None], ["footer"]]
    wb = build_xlsx(tmp_path / "w.xlsx", {"Data": rows})
    result = load(wb, profile("[{id: k, header: K}]", sheet_extra="data: {stop_after_blank_rows: 2}"))
    assert result.dataset.table("d").rows == [("a",)]


def test_formula_policies(tmp_path):
    wb = build_xlsx(tmp_path / "w.xlsx", {"Data": [["A", "B", "Total"], [1, 2, "=A2+B2"], [3, 4, "=A3+B3"]]})
    cols = "[{id: a, header: A}, {id: total, header: Total}]"
    result = load(wb, profile(cols))
    assert sum(1 for i in result.issues if i.code == "formula.no_cached_value") == 2
    assert all(i.severity is Severity.ERROR for i in result.issues if i.code == "formula.no_cached_value")
    table = result.dataset.table("d")
    assert table.rows[0][1] is None and (0, 1) in table.invalid
    raw = load(wb, profile(cols, extra="formula_policy: raw_formula"))
    assert raw.dataset.table("d").rows[0][1] == "=A2+B2"
    warn = load(wb, profile(cols, extra="formula_policy: cached_warn"))
    assert all(i.severity is Severity.WARNING for i in warn.issues if i.code == "formula.no_cached_value")
    skip = load(wb, profile(cols, extra="formula_policy: skip_check"))
    assert not skip.issues
    inject_cached_values(wb, 1, {"C2": 3, "C3": 7})
    ok = load(wb, profile(cols))
    assert ok.dataset.table("d").rows == [(1, 3), (3, 7)]
    assert not [i for i in ok.issues if i.code == "formula.no_cached_value"]


def test_xls_reader(tmp_path):
    pytest.importorskip("xlwt")
    path = build_xls(tmp_path / "old.xls", {"Data": [["Mã", "Ngày", "Số"], ["Ê1", dt.date(2024, 2, 29), 10]]})
    result = load(
        path, profile("[{id: m, header: Mã}, {id: d, header: Ngày, type: date}, {id: n, header: Số}]")
    )
    assert result.format == "xls"
    assert result.dataset.table("d").rows == [("Ê1", dt.date(2024, 2, 29), 10)]


@pytest.mark.parametrize(("encoding", "delimiter"), [("utf-8-sig", ","), ("cp1258", ";"), ("utf-16", "\t")])
def test_csv_reader_encodings(tmp_path, encoding, delimiter):
    # Windows-1258 stores Vietnamese tones as combining marks (decomposed form)
    ma, ten, dang = (
        ("Ma\u0303", "T\u00ean", "\u0110\u0103\u0323ng Thu")
        if encoding == "cp1258"
        else ("Mã", "Tên", "Đặng Thu")
    )
    text = delimiter.join([ma, ten]) + "\n" + delimiter.join(["E1", dang]) + "\n"
    path = tmp_path / "data.csv"
    path.write_bytes(text.encode(encoding))
    result = load(path, profile("[{id: m, header: Mã}, {id: t, header: Tên}]", sheet="{index: 1}"))
    assert result.dataset.table("d").rows == [("E1", "Đặng Thu")]


def test_csv_folder_is_a_workbook(tmp_path):
    folder = tmp_path / "csvs"
    folder.mkdir()
    (folder / "Data.csv").write_text("K\nx\n", encoding="utf-8")
    (folder / "Other.csv").write_text("Z\n1\n", encoding="utf-8")
    source = open_workbook(folder)
    assert [s.name for s in source.sheets()] == ["Data", "Other"]
    assert load(folder, profile("[{id: k, header: K}]")).dataset.table("d").rows == [("x",)]


def test_hidden_sheet_and_unsupported_file(tmp_path):
    wb = build_xlsx(tmp_path / "w.xlsx", {"Visible": [["x"]], "Data": [["K"], ["v"]]}, hidden=["Data"])
    result = load(wb, profile("[{id: k, header: K}]"))
    assert ("sheet.hidden", Severity.INFO) in codes(result)
    bad = tmp_path / "bad.xlsx"
    bad.write_bytes(b"not a workbook")
    with pytest.raises(OSError):
        open_workbook(bad)


def test_column_letters():
    assert [column_letter(i) for i in (0, 25, 26, 701)] == ["A", "Z", "AA", "ZZ"]
    assert column_index("AA") == 26
