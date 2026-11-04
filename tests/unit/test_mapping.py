"""Mapping engine: lookups, chains, composite keys, collections, aggregations, conditions."""

from __future__ import annotations

from decimal import Decimal
from pathlib import Path

import pytest

from docgen.domain.diagnostics import Severity
from docgen.domain.mapping.engine import ImageRef, MappingEngine
from docgen.domain.profile.io import parse_profile_text
from docgen.domain.validation.engine import ValidationEngine
from docgen.infrastructure.excel.loader import WorkbookLoader
from docgen.testing.xlsx_builder import build_xlsx

SHEETS = {
    "Orders": [
        ["OrderNo", "Customer", "Branch", "Dept", "Amount", "Status", "Photo"],
        ["O1", 101, "HN", "KD", 1500000, "paid", "img/o1.png"],
        ["O2", "102", "HCM", "KD", 250000, "open", None],
        ["O3", "999", "HN", "KT", None, "paid", None],
        ["O4", None, "HN", "XX", 10, "void", None],
    ],
    "Customers": [
        ["ID", "Name", "Tier", "RegionID"],
        [101.0, " An ", "gold", "R1"],
        ["102", "Bình", None, "R2"],
        ["103", "Châu", "silver", "R1"],
        ["103", "Châu (dup)", "silver", "R1"],
    ],
    "Regions": [["RegionID", "RegionName"], ["R1", "Miền Bắc"], ["R2", "Miền Nam"]],
    "Depts": [
        ["Branch", "Dept", "DeptName"],
        ["HN", "KD", "Kinh doanh HN"],
        ["HCM", "KD", "Kinh doanh HCM"],
        ["HN", "KT", "Kế toán HN"],
    ],
    "Lines": [
        ["OrderNo", "Item", "Qty", "Price", "Active"],
        ["O1", "Bút", 2, 5000, "y"],
        ["O1", "Vở", 10, 12000, "y"],
        ["O1", "Thước", 1, 3000, "n"],
        ["O2", "Bút", 1, 5000, "y"],
    ],
}

PROFILE = """
profile: {id: t, name: T}
excel:
  primary: o
  sheets:
    - id: o
      match: {name: Orders}
      columns:
        - {id: no, header: OrderNo}
        - {id: cust, header: Customer, required: false}
        - {id: branch, header: Branch}
        - {id: dept, header: Dept}
        - {id: amount, header: Amount, type: decimal}
        - {id: status, header: Status}
        - {id: photo, header: Photo}
    - id: c
      match: {name: Customers}
      columns: [{id: id, header: ID}, {id: name, header: Name}, {id: tier, header: Tier},
                {id: region, header: RegionID}]
    - id: r
      match: {name: Regions}
      columns: [{id: id, header: RegionID}, {id: name, header: RegionName}]
    - id: d
      match: {name: Depts}
      columns: [{id: branch, header: Branch}, {id: dept, header: Dept}, {id: name, header: DeptName}]
    - id: l
      match: {name: Lines}
      columns: [{id: no, header: OrderNo}, {id: item, header: Item}, {id: qty, header: Qty},
                {id: price, header: Price}, {id: active, header: Active}]
relationships:
  - {id: o_c, from: o, to: c, keys: [{from: cust, to: id}], on_missing: %(on_missing)s,
     default: "(khách lẻ)", on_duplicate: %(on_dup)s}
  - {id: c_r, from: c, to: r, keys: [{from: region, to: id}]}
  - {id: o_d, from: o, to: d, keys: [{from: branch, to: branch}, {from: dept, to: dept}], on_missing: warn}
  - {id: o_l, from: o, to: l, keys: [{from: no, to: no}], cardinality: one_to_many,
     filter: "active == 'y'", sort: [{by: price, order: desc}]}
fields:
  - {name: order.no, source: no}
  - {name: order.customer, path: o_c.name, transforms: [trim]}
  - {name: order.region, path: o_c.c_r.name, default: "?"}
  - {name: order.dept, path: o_d.name}
  - {name: order.tier, coalesce: [o_c.tier, status], transforms: [upper]}
  - {name: order.amount, source: amount, type: decimal, default: 0}
  - name: order.size
    when:
      - {if: "o.amount is none", value: unknown}
      - {if: "o.amount >= 1000000", value: big}
      - {else: small}
  - {name: order.label, expr: "order.no ~ ' / ' ~ order.customer"}
  - {name: order.items_text, path: o_l.item, aggregate: join, separator: " + "}
  - {name: order.photo, source: photo, type: image, image: {width_mm: 20}}
collections:
  - name: order.lines
    relationship: o_l
    filter: "item.qty > 1"
    fields:
      - {name: item, source: item}
      - {name: total, expr: "l.qty * l.price"}
aggregations:
  - {name: order.qty, relationship: o_l, op: sum, of: qty}
  - {name: order.max_price, relationship: o_l, op: max, of: price}
  - {name: order.count, relationship: o_l, op: count}
  - {name: order.items, relationship: o_l, op: join, of: item, separator: ", "}
validation:
  - {field: order.no, rule: unique}
  - {field: order.amount, rule: range, min: 0, max: 2000000, severity: warning}
"""


def build(tmp_path: Path, on_missing: str = "error", on_dup: str = "error", sheets=None):
    wb = build_xlsx(tmp_path / "w.xlsx", sheets or SHEETS)
    profile = parse_profile_text(PROFILE % {"on_missing": on_missing, "on_dup": on_dup})
    load = WorkbookLoader(profile.excel, profile.formatting).load(wb)
    assert not [i for i in load.issues if i.severity is Severity.ERROR], load.issues
    return MappingEngine(profile, load.dataset, trace=True), profile


def test_single_and_multi_level_lookup_and_composite_keys(tmp_path):
    engine, _ = build(tmp_path)
    r = engine.map_record(0)
    order = r.context["order"]
    assert order["customer"] == "An"  # 101 (int) matches 101.0 (float key)
    assert order["region"] == "Miền Bắc"  # Orders -> Customers -> Regions
    assert order["dept"] == "Kinh doanh HN"  # composite (branch, dept)
    assert order["tier"] == "GOLD"
    assert order["size"] == "big"
    assert order["label"] == "O1 / An"
    assert isinstance(order["photo"], ImageRef) and order["photo"].width_mm == 20
    assert not r.has_errors
    trace = r.traces["order.region"]
    assert [str(s) for s in trace.sources] == ["Regions!B2 [RegionName]"]
    assert any("o_c" in step for step in trace.steps) and any("c_r" in step for step in trace.steps)


def test_one_to_many_collections_and_aggregations(tmp_path):
    engine, _ = build(tmp_path)
    order = engine.map_record(0).context["order"]
    # filter active=='y' on the relationship, sorted by price desc; collection filter qty > 1
    assert order["lines"] == [{"item": "Vở", "total": 120000}, {"item": "Bút", "total": 10000}]
    assert order["items_text"] == "Vở + Bút"
    assert order["qty"] == Decimal(12)
    assert order["max_price"] == 12000
    assert order["count"] == 2
    assert order["items"] == "Vở, Bút"
    empty = engine.map_record(2).context["order"]
    assert empty["lines"] == [] and empty["count"] == 0 and empty["items"] == ""


def test_coalesce_falls_back_across_sheets(tmp_path):
    engine, _ = build(tmp_path)
    assert engine.map_record(1).context["order"]["tier"] == "OPEN"  # customer tier empty -> status


def test_missing_reference_policies(tmp_path):
    engine, _ = build(tmp_path)
    r = engine.map_record(2)  # customer 999 does not exist
    errors = [i for i in r.issues if i.severity is Severity.ERROR]
    assert errors and errors[0].code == "lookup.missing_reference"
    assert errors[0].sheet == "Orders" and errors[0].row == 4 and errors[0].rule == "o_c"
    assert r.context["order"]["size"] == "unknown"

    engine, _ = build(tmp_path, on_missing="default")
    r = engine.map_record(2)
    assert r.context["order"]["customer"] == "(khách lẻ)" and not r.has_errors

    engine, _ = build(tmp_path, on_missing="ignore")
    r = engine.map_record(2)
    assert r.context["order"]["customer"] is None and not r.issues


def test_empty_foreign_key_is_not_an_error_by_default(tmp_path):
    engine, _ = build(tmp_path)
    r = engine.map_record(3)  # Customer empty, dept XX missing -> warn
    assert not [i for i in r.issues if i.rule == "o_c"]
    assert [i.severity for i in r.issues if i.rule == "o_d"] == [Severity.WARNING]


def test_duplicate_keys(tmp_path):
    sheets = dict(SHEETS)
    sheets["Orders"] = [*SHEETS["Orders"], ["O5", "103", "HN", "KD", 1, "paid", None]]
    engine, _ = build(tmp_path, sheets=sheets)
    build_issues = engine.build_all_indexes()
    dup = [i for i in build_issues if i.code == "lookup.duplicate_key"]
    assert dup and dup[0].params["rows"] == [4, 5]
    r = engine.map_record(4)
    assert any(i.code == "lookup.ambiguous" for i in r.issues)
    assert r.context["order"]["customer"] is None  # never an arbitrary pick

    engine, _ = build(tmp_path, on_dup="warn", sheets=sheets)
    r = engine.map_record(4)
    assert r.context["order"]["customer"] == "Châu"
    assert [i.severity for i in r.issues if i.code == "lookup.ambiguous_first_used"] == [Severity.WARNING]


def test_validation_rules(tmp_path):
    engine, profile = build(tmp_path, on_missing="ignore")
    validator = ValidationEngine(profile.validation, engine.expressions)
    issues = []
    for i in range(4):
        issues += validator.check(engine.map_record(i))
    assert not [i for i in issues if i.code == "validation.unique"]
    issues = validator.check(engine.map_record(0))  # same order again -> duplicate
    assert [i.code for i in issues] == ["validation.unique"]


def test_untrusted_formula_cells_flag_records(tmp_path):
    sheets = {
        "Orders": SHEETS["Orders"],
        "Customers": [["ID", "Name", "Tier", "RegionID"], [101, '=UPPER("an")', "gold", "R1"]],
        "Regions": SHEETS["Regions"],
        "Depts": SHEETS["Depts"],
        "Lines": SHEETS["Lines"],
    }
    wb = build_xlsx(tmp_path / "w.xlsx", sheets)
    profile = parse_profile_text(PROFILE % {"on_missing": "ignore", "on_dup": "error"})
    load = WorkbookLoader(profile.excel, profile.formatting).load(wb)
    engine = MappingEngine(profile, load.dataset)
    r = engine.map_record(0)
    assert any(i.code == "mapping.untrusted_source" and i.severity is Severity.ERROR for i in r.issues)


def test_group_mapping(tmp_path):
    engine, _ = build(tmp_path, on_missing="ignore")
    recs = [engine.map_record(i) for i in (0, 2, 3)]
    group = engine.map_group(("HN",), recs, ["order.branch"])
    assert group.context["group"]["count"] == 3
    assert group.context["group"]["keys"]["order"]["branch"] == "HN"
    assert group.context["first"]["order"]["no"] == "O1"
    assert [r["order"]["no"] for r in group.context["records"]] == ["O1", "O3", "O4"]


@pytest.mark.parametrize(
    "op,expected",
    [
        ("sum", Decimal(6)),
        ("avg", Decimal(2)),
        ("min", 1),
        ("max", 3),
        ("count", 3),
        ("count_distinct", 2),
        ("first", 1),
        ("last", 3),
        ("join", "1, 2, 3"),
    ],
)
def test_aggregate_values(op, expected):
    from docgen.domain.mapping.engine import aggregate_values

    values = [1, None, 2, 3] if op != "count_distinct" else [1, 1, 2]
    assert aggregate_values(op, values) == expected
