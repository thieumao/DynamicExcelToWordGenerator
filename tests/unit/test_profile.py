import pytest

from docgen.domain.diagnostics import Severity
from docgen.domain.mapping.graph import Node, order_nodes
from docgen.domain.profile.io import ProfileLoadError, parse_profile_text, profile_to_text
from docgen.domain.profile.models import MatchMode, OnMissing
from docgen.domain.profile.validator import validate_profile

BASE = """
schema_version: 1
profile: {id: t, name: Test}
excel:
  primary: a
  sheets:
    - id: a
      match: {name: A}
      columns: [{id: k, header: K}, {id: v, header: V}, {id: fk, header: FK}]
    - id: b
      match: {name: B}
      columns: [{id: k, header: K}, {id: name, header: Name}]
relationships:
  - {id: a_b, from: a, to: b, keys: [{from: fk, to: k}]}
"""


def load(extra: str = ""):
    return parse_profile_text(BASE + extra)


def codes(profile):
    return [i.code for i in validate_profile(profile) if i.severity is Severity.ERROR]


def test_yaml12_booleans_and_aliases():
    p = load("""
fields:
  - name: out.flag
    when:
      - {if: "a.v == 'on'", value: yes}
      - {else: N}
  - {name: out.map, source: v, transforms: [{map_values: {mapping: {Y: Có, N: Không}}}]}
""")
    clause = p.fields[0].when
    assert clause[0].value == "yes" and clause[1].else_ == "N"
    assert p.fields[1].transforms[0].params["mapping"] == {"Y": "Có", "N": "Không"}
    assert p.relationships[0].keys[0].from_ == "fk"


def test_defaults_derived_from_cardinality():
    p = load("""
  - {id: a_many, from: a, to: b, keys: [{from: k, to: k}], cardinality: one_to_many}
""")
    assert p.relationships[0].match is MatchMode.UNIQUE
    assert p.relationships[1].match is MatchMode.ALL
    assert p.relationships[1].on_missing is OnMissing.IGNORE
    assert p.relationships[0].on_missing is OnMissing.ERROR


@pytest.mark.parametrize(
    ("snippet", "message"),
    [
        ("fields:\n  - {name: x.y}\n", "exactly one of"),
        ("fields:\n  - {name: x.y, source: v, expr: '1'}\n", "exactly one of"),
        ("fields:\n  - {name: 'bad name', source: v}\n", "String should match"),
        ("fields:\n  - name: x\n    when: [{else: 1}, {if: 'true', value: 2}]\n", "'else' must be the last"),
        ("generation: {unit: group}\n", "requires generation.group_by"),
    ],
)
def test_schema_errors(snippet, message):
    with pytest.raises(ProfileLoadError) as exc:
        load(snippet)
    assert message in str(exc.value)


def test_newer_schema_rejected():
    with pytest.raises(ProfileLoadError, match="newer"):
        parse_profile_text("schema_version: 99\n")


def test_round_trip_is_stable():
    p = load("""
fields:
  - {name: out.name, path: a_b.name, default: "-", transforms: [trim, {pad_left: {width: 5}}]}
  - name: out.level
    when: [{if: "a.v > 3", value: High}, {if: "a.v > 1", expr: "'Mid'"}, {else: Low}]
templates:
  - {id: t1, file: t1.docx, marker: {style: custom, open: "[", close: "]"}, bindings: {"Họ tên": out.name}}
""")
    text = profile_to_text(p)
    again = parse_profile_text(text)
    assert again.model_dump() == p.model_dump()
    assert parse_profile_text(profile_to_text(p, fmt="json"), fmt="json").model_dump() == p.model_dump()


def test_validator_reference_errors():
    p = load("""
fields:
  - {name: out.a, source: nope}
  - {name: out.b, path: missing_rel.name}
  - {name: out.c, path: a_b.nope}
  - {name: out.d, expr: "a.v >"}
  - {name: out.e, source: v, transforms: [{replace: {}}, unknown_tf]}
  - {name: report.x, const: 1}
  - {name: a.x, const: 1}
""")
    found = codes(p)
    for code in [
        "field.unknown_column",
        "field.unknown_relationship",
        "expression.syntax",
        "transform.invalid",
        "field.reserved_name",
        "field.shadows_sheet",
    ]:
        assert code in found, code


def test_validator_chain_and_to_many():
    p = load("""
  - {id: b_a, from: b, to: a, keys: [{from: k, to: fk}], cardinality: one_to_many}
fields:
  - {name: out.wrong, path: b_a.v}
  - {name: out.many, path: a_b.b_a.v}
  - {name: out.ok, path: a_b.b_a.v, aggregate: join}
""")
    issues = validate_profile(p)
    msgs = [(i.code, i.field) for i in issues if i.severity is Severity.ERROR]
    assert ("field.bad_chain", "out.wrong") in msgs
    assert ("field.to_many_without_aggregate", "out.many") in msgs
    assert not any(f == "out.ok" for _, f in msgs)


def test_cycles_and_path_conflicts():
    p = load("""
fields:
  - {name: out.a, expr: "out.b ~ 'x'"}
  - {name: out.b, expr: "out.a ~ 'y'"}
  - {name: out.self, expr: "out.self"}
  - {name: out.c, source: v}
  - {name: out.c.d, source: v}
""")
    issues = [i for i in validate_profile(p) if i.code == "field.cycle"]
    texts = " | ".join(i.message for i in issues)
    assert "out.a -> out.b -> out.a" in texts and "out.self -> out.self" in texts
    assert "field.path_conflict" in codes(p)


def test_dependency_order_is_deterministic():
    nodes = [
        Node("x.total", frozenset({("x", "a"), ("x", "b")})),
        Node("x.b", frozenset({("x", "a")})),
        Node("x.a", frozenset()),
        Node("x.free", frozenset()),
    ]
    order = order_nodes(nodes)
    names = [nodes[i].name for i in order.order]
    assert names == ["x.a", "x.b", "x.total", "x.free"] or names.index("x.a") < names.index("x.b")
    assert names.index("x.b") < names.index("x.total") and not order.cycles


def test_sample_profiles_are_valid(sample_a, sample_b, sample_c):
    for sample in (sample_a, sample_b, sample_c):
        assert codes(sample.profile) == []
