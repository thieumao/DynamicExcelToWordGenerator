"""Small SQLite store for application state: job history and recent files.

Only metadata is stored (no cell values) so confidential data does not end up
in application state.
"""

from __future__ import annotations

import json
import sqlite3
from contextlib import closing
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any

SCHEMA = """
CREATE TABLE IF NOT EXISTS jobs (
    job_id TEXT PRIMARY KEY,
    started TEXT NOT NULL,
    finished TEXT,
    status TEXT NOT NULL,
    profile_id TEXT,
    profile_version INTEGER,
    excel_file TEXT,
    output_dir TEXT,
    units INTEGER,
    documents INTEGER,
    errors INTEGER,
    warnings INTEGER,
    audit_path TEXT
);
CREATE TABLE IF NOT EXISTS recent (
    kind TEXT NOT NULL,
    value TEXT NOT NULL,
    used TEXT NOT NULL DEFAULT (datetime('now')),
    PRIMARY KEY (kind, value)
);
CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT);
"""


@dataclass(frozen=True, slots=True)
class JobRow:
    job_id: str
    started: str
    finished: str | None
    status: str
    profile_id: str | None
    profile_version: int | None
    excel_file: str | None
    output_dir: str | None
    units: int | None
    documents: int | None
    errors: int | None
    warnings: int | None
    audit_path: str | None


class AppDatabase:
    def __init__(self, path: Path) -> None:
        self.path = path
        path.parent.mkdir(parents=True, exist_ok=True)
        with closing(self._connect()) as con, con:
            con.executescript(SCHEMA)
            con.execute("INSERT OR IGNORE INTO meta(key, value) VALUES ('schema', '1')")

    def _connect(self) -> sqlite3.Connection:
        con = sqlite3.connect(self.path, timeout=5)
        con.execute("PRAGMA journal_mode=WAL")
        return con

    def record_job(self, result: Any, excel_file: str) -> None:
        with closing(self._connect()) as con, con:
            con.execute(
                "INSERT OR REPLACE INTO jobs VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (
                    result.job_id,
                    result.started,
                    result.finished,
                    result.status,
                    result.profile_id,
                    result.profile_version,
                    excel_file,
                    result.output_dir,
                    len(result.units),
                    len(result.documents),
                    result.error_count,
                    result.warning_count,
                    result.audit_path,
                ),
            )

    def jobs(self, limit: int = 50) -> list[JobRow]:
        with closing(self._connect()) as con:
            rows = con.execute("SELECT * FROM jobs ORDER BY started DESC LIMIT ?", (limit,)).fetchall()
        return [JobRow(*row) for row in rows]

    def touch_recent(self, kind: str, value: str, keep: int = 15) -> None:
        with closing(self._connect()) as con, con:
            con.execute(
                "INSERT INTO recent(kind, value, used) VALUES (?, ?, datetime('now')) "
                "ON CONFLICT(kind, value) DO UPDATE SET used = datetime('now')",
                (kind, value),
            )
            con.execute(
                "DELETE FROM recent WHERE kind = ? AND value NOT IN "
                "(SELECT value FROM recent WHERE kind = ? ORDER BY used DESC LIMIT ?)",
                (kind, kind, keep),
            )

    def recent(self, kind: str, limit: int = 10) -> list[str]:
        with closing(self._connect()) as con:
            rows = con.execute(
                "SELECT value FROM recent WHERE kind = ? ORDER BY used DESC LIMIT ?", (kind, limit)
            ).fetchall()
        return [r[0] for r in rows]


@dataclass(slots=True)
class Settings:
    language: str = "vi"
    default_output_dir: str = ""
    workers: int = 0
    """0 = automatic."""
    log_level: str = "INFO"
    log_cell_values: bool = False

    @classmethod
    def load(cls, path: Path) -> Settings:
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return cls()
        known = {k: v for k, v in data.items() if k in cls.__dataclass_fields__}
        return cls(**known)

    def save(self, path: Path) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp = path.with_suffix(".tmp")
        tmp.write_text(json.dumps(asdict(self), indent=2), encoding="utf-8")
        tmp.replace(path)
