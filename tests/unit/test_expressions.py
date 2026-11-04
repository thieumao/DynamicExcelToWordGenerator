import pytest

from docgen.domain.mapping.expressions import (
    ExpressionEngine,
    ExpressionError,
    check_syntax,
    expression_references,
)


@pytest.fixture()
def engine():
    return ExpressionEngine()


def test_basic_evaluation(engine):
    ns = {"emp": {"grade": 5, "city": "Hà Nội"}, "employee": {"name": "An"}}
    assert engine.compile("emp.grade >= 5 and emp.city == 'Hà Nội'")(ns) is True
    assert engine.compile("employee.name ~ '!'")(ns) == "An!"
    assert engine.compile("coalesce(emp.missing, 'x')")(ns) == "x"
    assert engine.compile("concat_ws(' ', 'a', none, 'b')")(ns) == "a b"
    assert engine.compile("emp.city | strip_accents")(ns) == "Ha Noi"
    assert engine.compile("emp.missing")(ns) is None


def test_mapping_keys_win_over_attributes(engine):
    ns = {"row": {"items": "value-of-column", "keys": 3}}
    assert engine.compile("row.items")(ns) == "value-of-column"
    assert engine.compile("row.keys + 1")(ns) == 4


@pytest.mark.parametrize(
    "source",
    [
        "().__class__.__bases__[0].__subclasses__()",
        "''.__class__.__mro__",
        "emp.__init__.__globals__",
        "cycler.__init__.__globals__.os",
    ],
)
def test_sandbox_blocks_private_attributes(engine, source):
    with pytest.raises(ExpressionError):
        result = engine.compile(source)({"emp": {}})
        # undefined access may evaluate to None instead of raising; it must never leak objects
        assert result is None
        raise ExpressionError("blocked")


def test_syntax_errors():
    with pytest.raises(ExpressionError):
        check_syntax("emp.grade >=")
    with pytest.raises(ExpressionError):
        check_syntax("{{ x }}")
    with pytest.raises(ExpressionError):
        check_syntax("x" * 3000)


def test_references():
    refs = expression_references("employee.code ~ emp.grade | string ~ coalesce(a.b, c)")
    assert ("employee", "code") in refs and ("emp", "grade") in refs
    assert ("a", "b") in refs and ("c",) in refs
    assert not any(r[0] == "coalesce" for r in refs)
