"""Home page: three clear actions for first-time users, and the recent work."""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from docgen.application.starter import install_construction_kit
from docgen.i18n.catalog import tr
from docgen.ui.pages.quickstart import documents_folder
from docgen.ui.state import AppState
from docgen.ui.widgets.common import button, error, page_title

SAMPLE_FOLDER = "Mẫu - Hồ sơ nghiệm thu"


def card(title: str, text: str) -> QPushButton:
    btn = QPushButton(f"{title}\n{text}")
    btn.setObjectName("Card")
    btn.setCursor(Qt.CursorShape.PointingHandCursor)
    btn.setMinimumHeight(84)
    return btn


class DashboardPage(QWidget):
    def __init__(self, state: AppState) -> None:
        super().__init__()
        self.state = state
        self.sample_folder: Path | None = None
        layout = QVBoxLayout(self)
        layout.addWidget(
            page_title(
                tr("Dashboard"),
                tr(
                    "Make Word documents from your Excel data. Everything runs on this computer, no Internet needed."
                ),
            )
        )
        cards = QHBoxLayout()
        new = card(
            tr("1. Create a new document set"),
            tr("From your Excel file and Word forms –\nthe app matches the data for you."),
        )
        new.clicked.connect(lambda: state.navigate.emit("quickstart"))
        saved = card(
            tr("2. Use a saved document set"),
            tr("Choose one in the list below,\nthen click here to make the documents."),
        )
        saved.clicked.connect(self._open_selected)
        sample = card(
            tr("3. Try the example"),
            tr("Acceptance records for construction work:\nExcel + 2 Word forms, ready to run."),
        )
        sample.clicked.connect(self.try_sample)
        for c in (new, saved, sample):
            cards.addWidget(c)
        layout.addLayout(cards)
        guide = button(tr("📖 How to use DocGen (user guide)"))
        guide.clicked.connect(lambda: state.navigate.emit("help"))
        guide_row = QHBoxLayout()
        guide_row.addWidget(guide)
        guide_row.addStretch(1)
        layout.addLayout(guide_row)
        self.sample_hint = QLabel()
        self.sample_hint.setWordWrap(True)
        self.sample_hint.setObjectName("Hint")
        self.sample_open = button(tr("Open the example folder"))
        self.sample_open.clicked.connect(self._open_sample_folder)
        self.sample_open.hide()
        row = QHBoxLayout()
        row.addWidget(self.sample_hint, 1)
        row.addWidget(self.sample_open)
        layout.addLayout(row)

        lists = QHBoxLayout()
        sets_box = QGroupBox(tr("Saved document sets"))
        sets_layout = QVBoxLayout(sets_box)
        self.profiles = QListWidget()
        self.profiles.itemDoubleClicked.connect(lambda _item: self._open_selected())
        sets_layout.addWidget(self.profiles)
        lists.addWidget(sets_box, 1)
        jobs_box = QGroupBox(tr("Recent jobs"))
        jobs_layout = QVBoxLayout(jobs_box)
        self.jobs = QTableWidget(0, 5)
        self.jobs.setHorizontalHeaderLabels(
            [tr("Started"), tr("Status"), tr("Profile"), tr("Documents"), tr("Output folder")]
        )
        self.jobs.horizontalHeader().setStretchLastSection(True)
        self.jobs.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.jobs.cellDoubleClicked.connect(self._open_job_folder)
        self.jobs.setToolTip(tr("Double-click to open the folder"))
        jobs_layout.addWidget(self.jobs)
        lists.addWidget(jobs_box, 2)
        layout.addLayout(lists, 1)
        self.refresh()

    def refresh(self) -> None:
        recent = {pid: i for i, pid in enumerate(self.state.db.recent("profile"))}
        entries = sorted(
            (e for e in self.state.repo.list() if not e.error),
            key=lambda e: (recent.get(e.id, len(recent)), -e.modified),
        )
        self.profiles.clear()
        for entry in entries:
            item = QListWidgetItem(entry.name)
            item.setData(Qt.ItemDataRole.UserRole, entry.id)
            item.setToolTip(entry.id)
            self.profiles.addItem(item)
        if self.profiles.count():
            self.profiles.setCurrentRow(0)
        jobs = self.state.db.jobs(30)
        self.jobs.setRowCount(len(jobs))
        for r, job in enumerate(jobs):
            values = [
                job.started.replace("T", " ")[:16],
                tr(job.status),
                job.profile_id,
                str(job.documents or 0),
                job.output_dir or "",
            ]
            for c, value in enumerate(values):
                item = QTableWidgetItem(value)
                item.setData(Qt.ItemDataRole.UserRole, job.output_dir)
                self.jobs.setItem(r, c, item)
        self.jobs.resizeColumnsToContents()

    def _open_selected(self) -> None:
        item = self.profiles.currentItem()
        if item is None:
            self.state.navigate.emit("quickstart")
            return
        self.state.open_profile(item.data(Qt.ItemDataRole.UserRole))
        self.state.navigate.emit("generate")

    def try_sample(self) -> None:
        folder = documents_folder() / SAMPLE_FOLDER
        try:
            profile_id, excel = install_construction_kit(self.state.repo, folder)
        except (OSError, ValueError) as exc:
            error(self, tr("The example could not be created."), str(exc))
            return
        self.sample_folder = folder
        self.sample_hint.setText(
            tr(
                "The example is in {folder}: open data/HoSoNghiemThu.xlsx to see how the data is entered, "
                "and the templates folder for the Word forms with <<…>> marks.",
                folder=folder,
            )
        )
        self.sample_open.show()
        self.state.open_profile(profile_id)
        self.state.excel_path = excel
        self.state.workbook = None
        self.state.workbook_changed.emit()
        self.refresh()
        self.state.navigate.emit("generate")
        self.state.run_requested.emit(str(excel), str(folder / "Kết quả"))

    def _open_sample_folder(self) -> None:
        if self.sample_folder is not None:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.sample_folder)))

    def _open_job_folder(self, row: int, _col: int) -> None:
        item = self.jobs.item(row, 0)
        folder = item.data(Qt.ItemDataRole.UserRole) if item else None
        if folder and Path(folder).exists():
            QDesktopServices.openUrl(QUrl.fromLocalFile(folder))
