"""Customer profiles: create, open, duplicate, delete, import/export, validate, YAML editor."""

from __future__ import annotations

import re
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QFileDialog,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QPlainTextEdit,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from docgen.domain.profile.io import ProfileLoadError, parse_profile_text, profile_to_text
from docgen.domain.profile.validator import validate_profile
from docgen.i18n.catalog import tr
from docgen.infrastructure.storage.profile_repo import blank_profile
from docgen.ui.state import AppState
from docgen.ui.widgets.common import IssueView, button, confirm, error, info, page_title

_ID = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_.-]{0,63}$")


def ask_profile_id(parent: QWidget, title: str, default: str = "") -> tuple[str, str] | None:
    name, ok = QInputDialog.getText(parent, title, tr("Profile name (e.g. customer name):"), text=default)
    if not ok or not name.strip():
        return None
    from docgen.application.workbook import suggest_identifier

    suggestion = suggest_identifier(name, set())
    pid, ok = QInputDialog.getText(parent, title, tr("Profile ID (letters, digits, _ . -):"), text=suggestion)
    if not ok:
        return None
    pid = pid.strip()
    if not _ID.match(pid):
        error(parent, tr("Invalid profile ID: {pid}", pid=pid))
        return None
    return pid, name.strip()


class ProfilesPage(QWidget):
    def __init__(self, state: AppState) -> None:
        super().__init__()
        self.state = state
        layout = QVBoxLayout(self)
        layout.addWidget(
            page_title(
                tr("Customer profiles"),
                tr(
                    "A profile stores how one customer's Excel and Word files are mapped. Switch customers "
                    "without changing code."
                ),
            )
        )
        bar = QHBoxLayout()
        actions = [
            (tr("New"), self._new, True),
            (tr("Open"), self._open, False),
            (tr("Duplicate"), self._duplicate, False),
            (tr("Delete"), self._delete, False),
            (tr("Import…"), self._import, False),
            (tr("Export…"), self._export, False),
            (tr("Refresh"), self.refresh, False),
        ]
        for label, slot, primary in actions:
            btn = button(label, primary=primary)
            btn.clicked.connect(slot)
            bar.addWidget(btn)
        bar.addStretch(1)
        layout.addLayout(bar)

        splitter = QSplitter(Qt.Orientation.Vertical)
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels([tr("ID"), tr("Name"), tr("Version"), tr("Folder")])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.cellDoubleClicked.connect(lambda *_: self._open())
        splitter.addWidget(self.table)

        tabs = QTabWidget()
        editor_box = QWidget()
        ed_layout = QVBoxLayout(editor_box)
        self.current = QLabel()
        ed_layout.addWidget(self.current)
        self.editor = QPlainTextEdit()
        self.editor.setFont(QFont("Consolas, Menlo, monospace", 10))
        self.editor.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
        ed_layout.addWidget(self.editor, 1)
        ed_bar = QHBoxLayout()
        for label, slot in (
            (tr("Reload from profile"), self._load_editor),
            (tr("Validate"), self._validate),
            (tr("Apply YAML"), self._apply_yaml),
            (tr("Save profile"), self._save),
        ):
            btn = button(label, primary=label == tr("Save profile"))
            btn.clicked.connect(slot)
            ed_bar.addWidget(btn)
        ed_bar.addStretch(1)
        ed_layout.addLayout(ed_bar)
        tabs.addTab(editor_box, tr("Advanced YAML editor"))
        self.issues = IssueView(compact=True)
        tabs.addTab(self.issues, tr("Validation"))
        self.tabs = tabs
        splitter.addWidget(tabs)
        splitter.setSizes([220, 480])
        layout.addWidget(splitter, 1)
        state.profile_changed.connect(self._load_editor)
        state.profile_edited.connect(self._update_label)
        self.refresh()

    # ---------------------------------------------------------------
    def refresh(self) -> None:
        entries = self.state.repo.list()
        self.table.setRowCount(len(entries))
        for r, e in enumerate(entries):
            for c, value in enumerate(
                [e.id, e.name if not e.error else f"⚠ {e.error[:80]}", str(e.version), str(e.directory)]
            ):
                item = QTableWidgetItem(value)
                item.setData(Qt.ItemDataRole.UserRole, e.id)
                self.table.setItem(r, c, item)
        self.table.resizeColumnsToContents()
        self._update_label()

    def _selected_id(self) -> str | None:
        row = self.table.currentRow()
        item = self.table.item(row, 0) if row >= 0 else None
        return item.data(Qt.ItemDataRole.UserRole) if item else None

    def _update_label(self) -> None:
        p = self.state.profile
        if p is None:
            self.current.setText(tr("No profile open."))
        else:
            dirty = " *" + tr("unsaved changes") if self.state.dirty else ""
            self.current.setText(
                tr(
                    "Open profile: {name} ({pid}) v{version}",
                    name=p.profile.name,
                    pid=p.profile.id,
                    version=p.profile.version,
                )
                + dirty
            )

    def _load_editor(self) -> None:
        if self.state.profile is not None:
            self.editor.setPlainText(profile_to_text(self.state.profile))
        self._update_label()

    def _new(self) -> None:
        if not self.state.settings.advanced_mode:
            self.state.navigate.emit("quickstart")  # beginners: Excel + Word -> ready profile
            return
        answer = ask_profile_id(self, tr("New profile"))
        if not answer:
            return
        pid, name = answer
        try:
            profile = blank_profile(pid, name)
            self.state.repo.create(profile)
        except (FileExistsError, OSError, ValueError) as exc:
            error(self, str(exc))
            return
        self.state.open_profile(pid)
        self.refresh()
        info(self, tr("Profile created. Next: choose an Excel file on the Excel configuration page."))
        self.state.navigate.emit("excel")

    def _open(self) -> None:
        pid = self._selected_id()
        if not pid:
            return
        if self.state.dirty and not confirm(self, tr("Discard unsaved changes of the current profile?")):
            return
        try:
            self.state.open_profile(pid)
        except (ProfileLoadError, OSError) as exc:
            error(self, tr("Cannot open profile: {err}", err=exc))
            return
        if not self.state.settings.advanced_mode:
            self.state.navigate.emit("generate")

    def _duplicate(self) -> None:
        pid = self._selected_id()
        if not pid:
            return
        answer = ask_profile_id(self, tr("Duplicate profile"), default=f"{pid} copy")
        if not answer:
            return
        try:
            self.state.repo.duplicate(pid, *answer)
        except (FileExistsError, OSError, ProfileLoadError) as exc:
            error(self, str(exc))
        self.refresh()

    def _delete(self) -> None:
        pid = self._selected_id()
        if not pid or not confirm(self, tr("Move profile '{pid}' to the trash folder?", pid=pid)):
            return
        target = self.state.repo.delete(pid)
        if self.state.profile and self.state.profile.profile.id == pid:
            self.state.profile = None
            self.state.profile_changed.emit()
        info(self, tr("Profile moved to {path}", path=target))
        self.refresh()

    def _import(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, tr("Import profile"), "", tr("Profiles") + " (*.zip *.yaml *.yml *.json)"
        )
        if not path:
            return
        try:
            profile = self.state.repo.import_bundle(Path(path))
        except FileExistsError:
            answer = ask_profile_id(self, tr("Profile ID already exists – choose another"))
            if not answer:
                return
            try:
                profile = self.state.repo.import_bundle(Path(path), new_id=answer[0])
            except (ProfileLoadError, OSError, ValueError) as exc:
                error(self, str(exc))
                return
        except (ProfileLoadError, OSError, ValueError) as exc:
            error(self, str(exc))
            return
        self.refresh()
        info(self, tr("Imported profile {pid}", pid=profile.profile.id))

    def _export(self) -> None:
        pid = self._selected_id() or (self.state.profile.profile.id if self.state.profile else None)
        if not pid:
            return
        path, _ = QFileDialog.getSaveFileName(
            self,
            tr("Export profile"),
            f"{pid}.zip",
            tr("Bundle with templates") + " (*.zip);;YAML (*.yaml);;JSON (*.json)",
        )
        if path:
            self.state.repo.export_bundle(pid, Path(path))
            info(self, tr("Exported to {path}", path=path))

    def _parse_editor(self):  # type: ignore[no-untyped-def]
        try:
            return parse_profile_text(self.editor.toPlainText())
        except ProfileLoadError as exc:
            self.issues.set_issues(exc.issues)
            self.tabs.setCurrentWidget(self.issues)
            error(self, str(exc))
            return None

    def _validate(self) -> None:
        profile = self._parse_editor()
        if profile is None:
            return
        issues = validate_profile(profile)
        self.issues.set_issues(issues)
        self.state.publish_issues(tr("Profile validation"), issues)
        self.tabs.setCurrentWidget(self.issues)

    def _apply_yaml(self) -> None:
        profile = self._parse_editor()
        if profile is None or self.state.profile_dir is None:
            return
        if self.state.profile and profile.profile.id != self.state.profile.profile.id:
            error(self, tr("Changing the profile ID here is not supported; duplicate the profile instead."))
            return
        self.state.profile = profile
        self.state.mark_edited()
        self.state.profile_changed.emit()
        self._validate()

    def _save(self) -> None:
        if self.state.profile is None:
            return
        try:
            path = self.state.save_profile()
        except (OSError, ProfileLoadError, ValueError) as exc:
            error(self, tr("Cannot save: {err}", err=exc))
            return
        self._load_editor()
        self.refresh()
        info(self, tr("Saved {path}", path=path))
