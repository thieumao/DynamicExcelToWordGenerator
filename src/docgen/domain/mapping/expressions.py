"""Restricted expression language (a safe subset of Jinja2 expressions).

Expressions such as ``emp.grade >= 5 and employee.city == 'Hà Nội'`` are
compiled once by an immutable *sandboxed* Jinja environment. There is no
Python ``eval``; only whitelisted filters (the transform registry) and
functions are available, attribute access to private members is blocked by
the sandbox, and mapping lookups prefer keys over Python attributes so a
column called ``items`` never resolves to ``dict.items``.
"""

from __future__ import annotations

import datetime as dt
from collections.abc import Callable, Iterator, Mapping
from dataclasses import dataclass
from decimal import Decimal
from functools import lru_cache
from typing import Any

from jinja2 import ChainableUndefined, nodes
from jinja2.exceptions import TemplateError
from jinja2.sandbox import ImmutableSandboxedEnvironment

from docgen.domain.data import values as V
from docgen.domain.mapping.transforms import (
    DEFAULT_REGISTRY,
    TransformContext,
    TransformError,
    TransformRegistry,
)

MAX_EXPRESSION_LENGTH = 2000


class ExpressionError(ValueError):
    """Invalid expression syntax or evaluation failure."""


class TrackingUndefined(ChainableUndefined):
    """Undefined that records which names were missing during evaluation/rendering."""

    __slots__ = ()
    sink: list[str] | None = None  # set per environment subclass

    def _record(self) -> None:
        sink = type(self).sink
        if sink is not None:
            name = self._undefined_name
            if name is not None and len(sink) < 1000:
                sink.append(str(name))

    def __str__(self) -> str:
        self._record()
        return ""

    def __iter__(self) -> Iterator[Any]:
        self._record()
        return iter(())

    def __bool__(self) -> bool:
        self._record()
        return False


class DocGenEnvironment(ImmutableSandboxedEnvironment):
    """Sandboxed environment where mapping keys win over Python attributes."""

    def getattr(self, obj: Any, attribute: str) -> Any:
        if isinstance(obj, Mapping):
            try:
                return obj[attribute]
            except (KeyError, TypeError):
                return self.undefined(obj=obj, name=attribute)
        return super().getattr(obj, attribute)

    def getitem(self, obj: Any, argument: Any) -> Any:
        if isinstance(obj, Mapping) and isinstance(argument, str):
            try:
                return obj[argument]
            except KeyError:
                return self.undefined(obj=obj, name=argument)
        return super().getitem(obj, argument)


def _coalesce(*values: Any) -> Any:
    for v in values:
        if not V.is_empty(v) and not isinstance(v, ChainableUndefined):
            return v
    return None


def _concat_ws(separator: str, *values: Any) -> str:
    parts = []
    for v in values:
        if isinstance(v, ChainableUndefined) or V.is_empty(v):
            continue
        parts.append(V.to_text(v))
    return separator.join(parts)


def _today() -> dt.date:
    return dt.date.today()


def _now() -> dt.datetime:
    return dt.datetime.now().replace(microsecond=0)


def _days_between(a: Any, b: Any) -> int:
    return (V.to_date(b) - V.to_date(a)).days


def _num(value: Any) -> Decimal:
    return V.to_decimal(value)


def _is_empty(value: Any) -> bool:
    return isinstance(value, ChainableUndefined) or V.is_empty(value)


SAFE_FUNCTIONS: dict[str, Callable[..., Any]] = {
    "coalesce": _coalesce,
    "concat_ws": _concat_ws,
    "today": _today,
    "now": _now,
    "date": dt.date,
    "days_between": _days_between,
    "num": _num,
    "empty": _is_empty,
}


def make_environment(
    *,
    registry: TransformRegistry = DEFAULT_REGISTRY,
    ctx: TransformContext | None = None,
    sink: list[str] | None = None,
    autoescape: bool = False,
) -> DocGenEnvironment:
    """Create a sandboxed environment with transforms as filters and safe functions."""
    tctx = ctx or TransformContext()
    undefined_cls = type("BoundUndefined", (TrackingUndefined,), {"__slots__": (), "sink": sink})
    env = DocGenEnvironment(undefined=undefined_cls, autoescape=autoescape)
    for definition in registry.definitions():

        def make_filter(name: str) -> Callable[..., Any]:
            def _filter(value: Any, *args: Any, **kwargs: Any) -> Any:
                if isinstance(value, ChainableUndefined):
                    value = None
                spec = registry.get(name)
                assert spec is not None
                if args:
                    kwargs.update(zip(spec.params.keys(), args, strict=False))
                try:
                    return registry.apply(name, value, tctx, kwargs)
                except TransformError as exc:
                    raise ExpressionError(str(exc)) from exc

            return _filter

        # Jinja built-ins with the same name (default, join, first, last, sum, round, replace,
        # upper, lower, trim, unique, count) keep their documented Jinja semantics
        # (e.g. ``sum(attribute='x')``); YAML pipelines always use the registry directly.
        if definition.name not in env.filters:
            env.filters[definition.name] = make_filter(definition.name)
    env.globals.update(SAFE_FUNCTIONS)
    return env


@dataclass(frozen=True, slots=True)
class CompiledExpression:
    source: str
    _fn: Callable[..., Any]

    def __call__(self, namespace: Mapping[str, Any]) -> Any:
        try:
            result = self._fn(**namespace)
        except ExpressionError:
            raise
        except (TemplateError, TypeError, ValueError, ArithmeticError, AttributeError) as exc:
            raise ExpressionError(f"cannot evaluate '{self.source}': {exc}") from exc
        if isinstance(result, ChainableUndefined):
            return None
        return result


class ExpressionEngine:
    """Compiles and caches restricted expressions; also extracts their references."""

    def __init__(
        self,
        *,
        registry: TransformRegistry = DEFAULT_REGISTRY,
        ctx: TransformContext | None = None,
    ) -> None:
        self.env = make_environment(registry=registry, ctx=ctx)
        self._cache: dict[str, CompiledExpression] = {}

    def compile(self, source: str) -> CompiledExpression:
        cached = self._cache.get(source)
        if cached is not None:
            return cached
        check_syntax(source)
        try:
            fn = self.env.compile_expression(source, undefined_to_none=False)
        except TemplateError as exc:
            raise ExpressionError(f"invalid expression '{source}': {exc}") from exc
        compiled = CompiledExpression(source, fn)
        self._cache[source] = compiled
        return compiled


@lru_cache(maxsize=4096)
def _parse(source: str) -> nodes.Expr:
    env = ImmutableSandboxedEnvironment()
    try:
        template = env.parse("{{ (" + source + ") }}")
    except TemplateError as exc:
        raise ExpressionError(f"invalid expression '{source}': {exc}") from exc
    output = template.body[0]
    assert isinstance(output, nodes.Output)
    expr = output.nodes[0]
    assert isinstance(expr, nodes.Expr)
    return expr


def check_syntax(source: str) -> None:
    if len(source) > MAX_EXPRESSION_LENGTH:
        raise ExpressionError(f"expression longer than {MAX_EXPRESSION_LENGTH} characters")
    if "{{" in source or "{%" in source:
        raise ExpressionError("expressions must not contain template delimiters")
    _parse(source)


def dotted_chain(node: nodes.Node) -> tuple[str, ...] | None:
    """Return ('a','b','c') for the expression a.b.c / a['b'].c, else None."""
    parts: list[str] = []
    current: nodes.Node = node
    while True:
        if isinstance(current, nodes.Getattr):
            parts.append(current.attr)
            current = current.node
        elif (
            isinstance(current, nodes.Getitem)
            and isinstance(current.arg, nodes.Const)
            and isinstance(current.arg.value, str)
        ):
            parts.append(current.arg.value)
            current = current.node
        elif isinstance(current, nodes.Name):
            parts.append(current.name)
            return tuple(reversed(parts))
        else:
            return None


def expression_references(source: str) -> set[tuple[str, ...]]:
    """Longest dotted name chains used by an expression (``employee.code`` -> ('employee','code'))."""
    expr = _parse(source)
    refs: set[tuple[str, ...]] = set()
    _collect(expr, refs)
    return refs


def _collect(node: nodes.Node, refs: set[tuple[str, ...]]) -> None:
    chain = dotted_chain(node) if isinstance(node, (nodes.Getattr, nodes.Getitem, nodes.Name)) else None
    if chain is not None:
        if chain[0] not in SAFE_FUNCTIONS:
            refs.add(chain)
        return
    for child in node.iter_child_nodes():
        _collect(child, refs)
