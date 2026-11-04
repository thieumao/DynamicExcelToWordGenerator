"""Dashboard: recent jobs, recent profiles, last generated documents."""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from docgen.i18n.catalog import tr
from docgen.ui.state import AppState
from docgen.ui.widgets.common import button, page_title


class DashboardPage(QWidget):
    def __init__(self, state: AppState) -> None:
        super().__init__()
        self.state = state
        layout = QVBoxLayout(self)
        layout.addWidget(
            page_title(
                tr("Dashboard"),
                tr("Recent work and quick actions. Everything runs offline on this computer."),
            )
        )
        actions = QHBoxLayout()
        for label, page in (
            (tr("Customer profiles"), "profiles"),
            (tr("Excel configuration"), "excel"),
            (tr("Mapping designer"), "mapping"),
            (tr("Generate"), "generate"),
        ):
            btn = button(label, primary=page == "generate")
            btn.clicked.connect(lambda _=False, p=page: state.navigate.emit(p))
            actions.addWidget(btn)
        actions.addStretch(1)
        layout.addLayout(actions)

        grid = QGridLayout()
        jobs_box = QGroupBox(tr("Recent jobs"))
        jobs_layout = QVBoxLayout(jobs_box)
        self.jobs = QTableWidget(0, 6)
        self.jobs.setHorizontalHeaderLabels(
            [tr("Started"), tr("Status"), tr("Profile"), tr("Documents"), tr("Errors"), tr("Output folder")]
        )
        self.jobs.horizontalHeader().setStretchLastSection(True)
        self.jobs.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.jobs.cellDoubleClicked.connect(self._open_job_folder)
        jobs_layout.addWidget(self.jobs)
        grid.addWidget(jobs_box, 0, 0, 1, 2)

        prof_box = QGroupBox(tr("Recent customer profiles"))
        prof_layout = QVBoxLayout(prof_box)
        self.profiles = QListWidget()
        self.profiles.itemDoubleClicked.connect(self._open_profile)
        prof_layout.addWidget(self.profiles)
        grid.addWidget(prof_box, 1, 0)

        docs_box = QGroupBox(tr("Recent Excel files"))
        docs_layout = QVBoxLayout(docs_box)
        self.files = QListWidget()
        docs_layout.addWidget(self.files)
        grid.addWidget(docs_box, 1, 1)
        layout.addLayout(grid, 1)
        self.hint = QLabel()
        self.hint.setObjectName("Hint")
        layout.addWidget(self.hint)
        self.refresh()

    def refresh(self) -> None:
        jobs = self.state.db.jobs(30)
        self.jobs.setRowCount(len(jobs))
        for r, job in enumerate(jobs):
            values = [
                job.started.replace("T", " ")[:19],
                tr(job.status),
                f"{job.profile_id} v{job.profile_version}",
                str(job.documents or 0),
                str(job.errors or 0),
                job.output_dir or "",
            ]
            for c, value in enumerate(values):
                item = QTableWidgetItem(value)
                item.setData(Qt.ItemDataRole.UserRole, job.output_dir)
                self.jobs.setItem(r, c, item)
        self.jobs.resizeColumnsToContents()
        self.profiles.clear()
        for profile_id in self.state.db.recent("profile"):
            if self.state.repo.exists(profile_id):
                item = QListWidgetItem(profile_id)
                item.setData(Qt.ItemDataRole.UserRole, profile_id)
                self.profiles.addItem(item)
        self.files.clear()
        for path in self.state.db.recent("excel"):
            self.files.addItem(path)
        current = self.state.profile.profile.name if self.state.profile else tr("none")
        self.hint.setText(tr("Current profile: {name}", name=current))

    def _open_job_folder(self, row: int, _col: int) -> None:
        item = self.jobs.item(row, 0)
        folder = item.data(Qt.ItemDataRole.UserRole) if item else None
        if folder and Path(folder).exists():
            QDesktopServices.openUrl(QUrl.fromLocalFile(folder))

    def _open_profile(self, item: QListWidgetItem) -> None:
        self.state.open_profile(item.data(Qt.ItemDataRole.UserRole))
        self.state.navigate.emit("excel")
