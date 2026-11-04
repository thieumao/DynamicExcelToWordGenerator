"""Nested document context helpers and the context *schema* (known paths).

The mapping engine produces plain nested dicts/lists (picklable, JSON-friendly).
The schema lists every path a profile can produce, so template placeholders can
be validated before generation (``employee.projects[].name``).
"""

from __future__ import annotations

from collections.abc import Iterable, Mapping
from dataclasses import dataclass, field
from typing import Any

from docgen.domain.profile.models import CollectionSpec, Profile, UnitKind

_MISSING = object()


class PathConflictError(ValueError):
    """A value is assigned below a path that already holds a scalar (or vice versa)."""


def set_path(target: dict[str, Any], path: str, value: Any) -> None:
    parts = path.split(".")
    node = target
    for i, part in enumerate(parts[:-1]):
        child = node.get(part, _MISSING)
        if child is _MISSING:
            child = {}
            node[part] = child
        elif not isinstance(child, dict):
            raise PathConflictError(
                f"cannot set '{path}': '{'.'.join(parts[: i + 1])}' already holds a value"
            )
        node = child
    last = parts[-1]
    if isinstance(node.get(last), dict) and not isinstance(value, dict):
        raise PathConflictError(f"cannot set '{path}': it already contains nested fields")
    node[last] = value


def get_path(source: Mapping[str, Any], path: str, default: Any = None) -> Any:
    node: Any = source
    for part in path.split("."):
        if isinstance(node, Mapping) and part in node:
            node = node[part]
        else:
            return default
    return node


def flatten(ctx: Mapping[str, Any], prefix: str = "") -> dict[str, Any]:
    """{'a': {'b': 1}} -> {'a.b': 1}; lists are kept as values."""
    result: dict[str, Any] = {}
    for key, value in ctx.items():
        path = f"{prefix}{key}"
        if isinstance(value, Mapping):
            result.update(flatten(value, path + "."))
        else:
            result[path] = value
    return result


@dataclass(slots=True)
class ContextSchema:
    """Set of producible paths. List items are written as ``name[]``."""

    paths: set[str] = field(default_factory=set)
    open_prefixes: set[str] = field(default_factory=set)
    """Prefixes below which any key is acceptable (e.g. bindings 'b')."""

    def add(self, path: str) -> None:
        self.paths.add(path)

    def knows(self, path: str) -> bool:
        if path in self.paths:
            return True
        if any(
            path == p or path.startswith(p + ".") or path.startswith(p + "[]") for p in self.open_prefixes
        ):
            return True
        prefix = path + "."
        list_prefix = path + "[]"
        return any(p.startswith(prefix) or p.startswith(list_prefix) for p in self.paths)

    def suggestions(self, path: str, limit: int = 3) -> list[str]:
        import difflib

        return difflib.get_close_matches(path, sorted(self.paths), n=limit, cutoff=0.6)


REPORT_PATHS = (
    "report.generated_at",
    "report.generated_date",
    "report.job_id",
    "report.source_file",
    "report.profile.id",
    "report.profile.name",
    "report.profile.version",
)
UNIT_PATHS = ("unit.index", "unit.count", "unit.key", "unit.label")


def _collection_paths(spec: CollectionSpec, prefix: str, schema: ContextSchema) -> None:
    base = f"{prefix}{spec.name}[]"
    schema.add(base)
    for f in spec.fields:
        schema.add(f"{base}.{f.name}")
    for a in spec.aggregations:
        schema.add(f"{base}.{a.name}")
    for c in spec.collections:
        _collection_paths(c, base + ".", schema)


def record_paths(profile: Profile, prefix: str = "") -> ContextSchema:
    schema = ContextSchema()
    for f in profile.fields:
        schema.add(prefix + f.name)
    for a in profile.aggregations:
        schema.add(prefix + a.name)
    for c in profile.collections:
        _collection_paths(c, prefix, schema)
    return schema


def build_schema(profile: Profile, unit: UnitKind | None = None) -> ContextSchema:
    """Every path a template may reference for the given unit kind."""
    unit = unit or profile.generation.unit
    schema = ContextSchema()
    record = record_paths(profile)
    if unit is UnitKind.RECORD:
        schema.paths |= record.paths
    else:
        for prefix in ("group.records[].", "records[].", "first."):
            schema.paths |= record_paths(profile, prefix).paths
        schema.paths |= {"group.key", "group.count", "group.records[]", "records[]"}
        for gk in profile.generation.group_by:
            schema.add(f"group.keys.{gk}")
        for f in profile.group_fields:
            schema.add(f.name)
    schema.paths.update(REPORT_PATHS)
    schema.paths.update(UNIT_PATHS)
    schema.open_prefixes.add("b")
    return schema


def merge_paths(schemas: Iterable[ContextSchema]) -> ContextSchema:
    result = ContextSchema()
    for s in schemas:
        result.paths |= s.paths
        result.open_prefixes |= s.open_prefixes
    return result
