"""Per-user application directories (cross-platform, no hard-coded paths)."""

from __future__ import annotations

import os
from pathlib import Path

from platformdirs import user_config_dir, user_data_dir, user_log_dir

APP_NAME = "DocGen"
APP_AUTHOR = "DocGen"


def _override() -> Path | None:
    value = os.environ.get("DOCGEN_HOME")
    return Path(value) if value else None


def data_dir() -> Path:
    base = _override()
    path = base / "data" if base else Path(user_data_dir(APP_NAME, APP_AUTHOR))
    path.mkdir(parents=True, exist_ok=True)
    return path


def config_dir() -> Path:
    base = _override()
    path = base / "config" if base else Path(user_config_dir(APP_NAME, APP_AUTHOR))
    path.mkdir(parents=True, exist_ok=True)
    return path


def log_dir() -> Path:
    base = _override()
    path = base / "logs" if base else Path(user_log_dir(APP_NAME, APP_AUTHOR))
    path.mkdir(parents=True, exist_ok=True)
    return path


def profiles_dir() -> Path:
    path = data_dir() / "profiles"
    path.mkdir(parents=True, exist_ok=True)
    return path
