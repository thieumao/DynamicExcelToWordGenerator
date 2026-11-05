"""General-information sheets ("Thông tin công trình"): label/value layout and links without keys.

Construction workbooks keep the project details once (name, location, owner, contractor) and list
the work items to accept in another sheet; every document needs both.
"""

from __future__ import annotations

import datetime as dt
from pathlib import Path

from docgen.domain.diagnostics import Severity
from docgen.domain.mapping.engine import MappingEngine
from docgen.domain.profile.io import parse_profile_text
from docgen.domain.profile.validator import validate_profile
from docgen.infrastructure.excel.loader import WorkbookLoader
from docgen.testing.xlsx_builder import build_xlsx

SHEETS = {
    "Thông tin công trình": [
        ["THÔNG TIN CHUNG CỦA CÔNG TRÌNH"],  # title: one cell only, skipped
        [],
        ["Tên công trình:", "Nhà làm việc 5 tầng"],
        ["Địa điểm", None, "Thanh Xuân, Hà Nội"],  # value further right
        ["Ngày khởi công", dt.datetime(2024, 3, 1)],
        ["Ghi chú"],  # label without value: skipped
        ["Tổng mức đầu tư", "=1000*2"],
    ],
    "Danh mục": [
        ["STT", "Công việc"],
        [1, "Cốt thép cột tầng 1"],
        [2, "Bê tông cột tầng 1"],
    ],
    "Thành phần": [["Họ tên", "Đơn vị"], ["Nguyễn Văn A", "TVGS"], ["Trần Văn B", "Nhà thầu"]],
}

PROFILE = """
profile: {id: xd, name: XD}
excel:
  primary: dm
  formula_policy: %(policy)s
  sheets:
    - id: dm
      match: {name: "Danh mục"}
      columns:
        - {id: stt, header: STT}
        - {id: cv, header: "Công việc"}
    - id: ct
      match: {name: "Thông tin công trình"}
      layout: key_value
      columns:
        - {id: ten, header: "Tên công trình"}
        - {id: dia_diem, header: "Địa điểm"}
        - {id: khoi_cong, header: "Ngày khởi công", type: date}
        - {id: tmdt, header: "Tổng mức đầu tư", required: false}
    - id: tp
      match: {name: "Thành phần"}
      columns:
        - {id: ten, header: "Họ tên"}
        - {id: don_vi, header: "Đơn vị"}
relationships:
  - {id: chung, from: dm, to: ct}
  - {id: moi_nguoi, from: dm, to: tp, cardinality: one_to_many}
fields:
  - {name: cv.ten, source: cv}
  - {name: cong_trinh.ten, path: chung.ten}
  - {name: cong_trinh.dia_diem, path: chung.dia_diem}
  - {name: cong_trinh.khoi_cong, path: chung.khoi_cong, transforms: [{format_date: {format: "ngày %%d tháng %%m năm %%Y"}}]}
collections:
  - name: thanh_phan
    relationship: moi_nguoi
    fields: [{name: ten, source: ten}, {name: don_vi, source: don_vi}]
"""


def load(tmp_path: Path, sheets=None, policy: str = "cached_warn"):
    wb = build_xlsx(tmp_path / "w.xlsx", sheets or SHEETS)
    profile = parse_profile_text(PROFILE % {"policy": policy})
    return profile, WorkbookLoader(profile.excel, profile.formatting).load(wb)


def test_key_value_sheet_becomes_one_record(tmp_path):
    _, result = load(tmp_path)
    ct = result.dataset.tables["ct"]
    assert len(ct.rows) == 1
    assert ct.value(0, "ten") == "Nhà làm việc 5 tầng"
    assert ct.value(0, "dia_diem") == "Thanh Xuân, Hà Nội"
    assert ct.value(0, "khoi_cong") == dt.date(2024, 3, 1)
    assert ct.cell_ref(0, "dia_diem").column.startswith("Địa điểm")
    # the formula has no stored result (openpyxl wrote it): reported with its cell address
    formula = [i for i in result.issues if i.code == "formula.no_cached_value"]
    assert formula and formula[0].sheet == "Thông tin công trình" and "B7" in (formula[0].column or "")


def test_every_record_gets_the_general_information(tmp_path):
    profile, result = load(tmp_path)
    assert not [i for i in validate_profile(profile) if i.severity is Severity.ERROR]
    engine = MappingEngine(profile, result.dataset)
    assert not [i for i in engine.build_all_indexes() if i.severity is Severity.ERROR]
    for row in range(2):
        context = engine.map_record(row).context
        assert context["cong_trinh"] == {
            "ten": "Nhà làm việc 5 tầng",
            "dia_diem": "Thanh Xuân, Hà Nội",
            "khoi_cong": "ngày 01 tháng 03 năm 2024",
        }
        assert [p["ten"] for p in context["thanh_phan"]] == ["Nguyễn Văn A", "Trần Văn B"]


def test_missing_label_is_reported_in_vietnamese_terms(tmp_path):
    sheets = dict(SHEETS)
    sheets["Thông tin công trình"] = [r for r in SHEETS["Thông tin công trình"] if r[:1] != ["Địa điểm"]]
    _, result = load(tmp_path, sheets)
    missing = [i for i in result.issues if i.code == "column.not_found"]
    assert missing and missing[0].severity is Severity.ERROR and missing[0].column == "Địa điểm"


def test_general_information_must_be_one_row(tmp_path):
    profile = parse_profile_text(
        (PROFILE % {"policy": "cached_warn"}).replace("layout: key_value", "layout: table")
    )
    table_sheets = dict(SHEETS)
    table_sheets["Thông tin công trình"] = [
        ["Tên công trình", "Địa điểm", "Ngày khởi công", "Tổng mức đầu tư"],
        ["Nhà A", "Hà Nội", dt.datetime(2024, 3, 1), 1],
        ["Nhà B", "Huế", dt.datetime(2024, 4, 1), 2],
    ]
    wb = build_xlsx(tmp_path / "two.xlsx", table_sheets)
    result = WorkbookLoader(profile.excel, profile.formatting).load(wb)
    issues = MappingEngine(profile, result.dataset).build_all_indexes()
    assert [i.code for i in issues if i.severity is Severity.ERROR] == ["lookup.single_row_expected"]
    table_sheets["Thông tin công trình"] = table_sheets["Thông tin công trình"][:1]
    wb = build_xlsx(tmp_path / "none.xlsx", table_sheets)
    result = WorkbookLoader(profile.excel, profile.formatting).load(wb)
    engine = MappingEngine(profile, result.dataset)
    engine.build_all_indexes()
    record = engine.map_record(0)
    assert "lookup.single_row_missing" in [i.code for i in record.issues]
