"""Profile (de)serialisation: YAML/JSON text <-> :class:`Profile`, plus migrations."""

from __future__ import annotations

import json
import re
from collections.abc import Callable
from pathlib import Path
from typing import Any

import yaml
from pydantic import ValidationError

from docgen.domain.diagnostics import Category, Issue, Severity, make_issue
from docgen.domain.profile.models import SCHEMA_VERSION, Profile


class _Yaml12SafeLoader(yaml.SafeLoader):
    """SafeLoader with YAML 1.2 booleans: only true/false (any case) are booleans.

    YAML 1.1 (PyYAML's default) turns on/off/yes/no/y/n into booleans, which silently
    breaks keys such as ``on:`` and value tables such as ``{Y: Có, N: Không}``.
    """


_Yaml12SafeLoader.yaml_implicit_resolvers = {
    key: [(tag, regexp) for tag, regexp in resolvers if tag != "tag:yaml.org,2002:bool"]
    for key, resolvers in yaml.SafeLoader.yaml_implicit_resolvers.items()
}
_Yaml12SafeLoader.add_implicit_resolver(
    "tag:yaml.org,2002:bool",
    re.compile(r"^(?:true|True|TRUE|false|False|FALSE)$"),
    list("tTfF"),
)


def yaml_load(text: str) -> object:
    return yaml.load(text, Loader=_Yaml12SafeLoader)


# schema_version -> function upgrading raw data to schema_version + 1
MIGRATIONS: dict[int, Callable[[dict[str, Any]], dict[str, Any]]] = {}


class ProfileLoadError(ValueError):
    def __init__(self, message: str, issues: list[Issue]) -> None:
        super().__init__(message)
        self.issues = issues


def migrate(data: dict[str, Any]) -> dict[str, Any]:
    version = int(data.get("schema_version", SCHEMA_VERSION))
    if version > SCHEMA_VERSION:
        raise ProfileLoadError(
            f"Profile schema_version {version} is newer than this application supports "
            f"({SCHEMA_VERSION}); please update DocGen.",
            [],
        )
    while version < SCHEMA_VERSION:
        data = MIGRATIONS[version](data)
        version += 1
        data["schema_version"] = version
    return data


def _pydantic_issues(exc: ValidationError) -> list[Issue]:
    issues = []
    for err in exc.errors():
        location = ".".join(str(p) for p in err["loc"])
        issues.append(
            make_issue(
                "profile.schema",
                Severity.ERROR,
                Category.CONFIG,
                f"{location}: {err['msg']}",
                rule=location,
                detail=err["msg"],
            )
        )
    return issues


def parse_profile_text(text: str, *, fmt: str = "yaml") -> Profile:
    try:
        data = json.loads(text) if fmt == "json" else yaml_load(text)
    except (yaml.YAMLError, json.JSONDecodeError) as exc:
        raise ProfileLoadError(f"Profile is not valid {fmt.upper()}: {exc}", []) from exc
    if not isinstance(data, dict):
        raise ProfileLoadError("Profile must be a mapping at the top level", [])
    data = migrate(data)
    try:
        return Profile.model_validate(data)
    except ValidationError as exc:
        issues = _pydantic_issues(exc)
        raise ProfileLoadError(
            f"Profile has {len(issues)} schema error(s): " + "; ".join(i.message for i in issues[:5]),
            issues,
        ) from exc


def load_profile(path: Path) -> Profile:
    fmt = "json" if path.suffix.lower() == ".json" else "yaml"
    return parse_profile_text(path.read_text(encoding="utf-8"), fmt=fmt)


def profile_to_text(profile: Profile, *, fmt: str = "yaml") -> str:
    data = profile.to_data()
    if fmt == "json":
        return json.dumps(data, ensure_ascii=False, indent=2)
    return yaml.safe_dump(data, allow_unicode=True, sort_keys=False, width=100)


def profile_json_schema() -> dict[str, Any]:
    return Profile.model_json_schema(by_alias=True)
