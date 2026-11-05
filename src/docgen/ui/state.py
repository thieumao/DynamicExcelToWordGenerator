"""Shared UI state: current profile, workbook and services."""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QObject, Signal

from docgen.application.workbook import WorkbookSummary
from docgen.domain.diagnostics import Issue
from docgen.domain.profile.models import Profile
from docgen.infrastructure.storage import paths
from docgen.infrastructure.storage.app_db import AppDatabase, Settings
from docgen.infrastructure.storage.profile_repo import ProfileRepository


class AppState(QObject):
    profile_changed = Signal()
    """A different profile was opened (or the current one was replaced)."""
    profile_edited = Signal()
    """The current profile was modified in memory."""
    workbook_changed = Signal()
    issues_published = Signal(str, list)
    """(source label, list[Issue]) for the Logs page."""
    navigate = Signal(str)
    run_requested = Signal(str, str)
    """(Excel file, output folder): generate the current profile's documents now."""

    def __init__(self, data_root: Path | None = None) -> None:
        super().__init__()
        root = data_root
        self.repo = ProfileRepository(root / "profiles" if root else paths.profiles_dir())
        self.db = AppDatabase((root or paths.data_dir()) / "docgen.sqlite3")
        self.settings_path = (root / "settings.json") if root else paths.config_dir() / "settings.json"
        self.settings = Settings.load(self.settings_path)
        self.profile: Profile | None = None
        self.profile_dir: Path | None = None
        self.dirty = False
        self.excel_path: Path | None = None
        self.workbook: WorkbookSummary | None = None

    # ----------------------------------------------------------- profile
    def open_profile(self, profile_id: str) -> None:
        self.profile = self.repo.load(profile_id)
        self.profile_dir = self.repo.directory(profile_id)
        self.dirty = False
        self.db.touch_recent("profile", profile_id)
        self.profile_changed.emit()

    def set_profile(self, profile: Profile, directory: Path) -> None:
        self.profile = profile
        self.profile_dir = directory
        self.dirty = False
        self.profile_changed.emit()

    def mark_edited(self) -> None:
        self.dirty = True
        self.profile_edited.emit()

    def save_profile(self) -> Path | None:
        if self.profile is None:
            return None
        path = self.repo.save(self.profile, bump_version=self.dirty)
        self.dirty = False
        self.profile_edited.emit()
        return path

    # ---------------------------------------------------------- workbook
    def set_workbook(self, path: Path, summary: WorkbookSummary) -> None:
        self.excel_path = path
        self.workbook = summary
        self.db.touch_recent("excel", str(path))
        self.workbook_changed.emit()

    def publish_issues(self, source: str, issues: list[Issue]) -> None:
        self.issues_published.emit(source, issues)

    def save_settings(self) -> None:
        self.settings.save(self.settings_path)
