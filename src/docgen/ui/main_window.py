"""Main window: navigation sidebar + pages."""

from __future__ import annotations

from PySide6.QtCore import QSize, Qt
from PySide6.QtGui import QAction, QCloseEvent, QKeySequence
from PySide6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QStackedWidget,
    QStyle,
    QVBoxLayout,
    QWidget,
)

from docgen import __version__
from docgen.i18n.catalog import tr
from docgen.ui.pages.dashboard import DashboardPage
from docgen.ui.pages.excel import ExcelPage
from docgen.ui.pages.generate import GeneratePage
from docgen.ui.pages.logs import LogsPage
from docgen.ui.pages.mapping import MappingPage
from docgen.ui.pages.preview import PreviewPage
from docgen.ui.pages.profiles import ProfilesPage
from docgen.ui.pages.relationships import RelationshipsPage
from docgen.ui.pages.settings import SettingsPage
from docgen.ui.pages.templates import TemplatesPage
from docgen.ui.state import AppState
from docgen.ui.widgets.common import error

STYLE = """
QWidget { font-size: 10pt; }
#Sidebar { background: #1f2937; border: none; color: #e5e7eb; padding-top: 8px; }
#Sidebar::item { padding: 10px 14px; border-radius: 6px; margin: 2px 8px; }
#Sidebar::item:selected { background: #2563eb; color: white; }
#Sidebar::item:hover:!selected { background: #374151; }
#Brand { color: white; font-size: 14pt; font-weight: 600; padding: 14px 16px 4px 16px; background: #1f2937; }
#BrandSub { color: #9ca3af; padding: 0 16px 10px 16px; background: #1f2937; }
#PageTitle { font-size: 16pt; font-weight: 600; color: #111827; }
#PageSubtitle, #Hint { color: #6b7280; }
#Breadcrumb { color: #1d4ed8; font-family: Consolas, Menlo, monospace; }
QPushButton { padding: 6px 12px; border-radius: 6px; border: 1px solid #d1d5db; background: #ffffff; }
QPushButton:hover { background: #f3f4f6; }
QPushButton#Primary { background: #2563eb; color: white; border-color: #2563eb; }
QPushButton#Primary:hover { background: #1d4ed8; }
QPushButton:disabled { color: #9ca3af; background: #f9fafb; }
QGroupBox { font-weight: 600; border: 1px solid #e5e7eb; border-radius: 8px; margin-top: 12px; padding-top: 8px; }
QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; }
QTableView, QTableWidget, QTreeWidget, QListWidget { border: 1px solid #e5e7eb; border-radius: 6px; }
QProgressBar { border: 1px solid #d1d5db; border-radius: 6px; text-align: center; height: 18px; }
QProgressBar::chunk { background: #2563eb; border-radius: 6px; }
"""

PAGES = [
    ("dashboard", "Dashboard", QStyle.StandardPixmap.SP_DirHomeIcon),
    ("profiles", "Customer profiles", QStyle.StandardPixmap.SP_FileDialogContentsView),
    ("excel", "Excel configuration", QStyle.StandardPixmap.SP_FileDialogDetailedView),
    ("relationships", "Relationships", QStyle.StandardPixmap.SP_FileLinkIcon),
    ("mapping", "Mapping designer", QStyle.StandardPixmap.SP_FileDialogListView),
    ("templates", "Word templates", QStyle.StandardPixmap.SP_FileIcon),
    ("preview", "Preview", QStyle.StandardPixmap.SP_FileDialogInfoView),
    ("generate", "Generate", QStyle.StandardPixmap.SP_MediaPlay),
    ("logs", "Logs and errors", QStyle.StandardPixmap.SP_MessageBoxWarning),
    ("settings", "Settings", QStyle.StandardPixmap.SP_FileDialogDetailedView),
]


class MainWindow(QMainWindow):
    def __init__(self, state: AppState) -> None:
        super().__init__()
        self.state = state
        self.setWindowTitle(f"DocGen {__version__}")
        self.resize(1320, 860)
        self.setStyleSheet(STYLE)
        root = QWidget()
        layout = QHBoxLayout(root)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        side = QWidget()
        side.setFixedWidth(230)
        side.setStyleSheet("background: #1f2937;")
        sv = QVBoxLayout(side)
        sv.setContentsMargins(0, 0, 0, 0)
        brand = QLabel("DocGen")
        brand.setObjectName("Brand")
        sub = QLabel(tr("Excel → Word"))
        sub.setObjectName("BrandSub")
        sv.addWidget(brand)
        sv.addWidget(sub)
        self.nav = QListWidget()
        self.nav.setObjectName("Sidebar")
        self.nav.setIconSize(QSize(18, 18))
        sv.addWidget(self.nav, 1)
        self.profile_label = QLabel()
        self.profile_label.setObjectName("BrandSub")
        self.profile_label.setWordWrap(True)
        sv.addWidget(self.profile_label)
        layout.addWidget(side)
        self.stack = QStackedWidget()
        layout.addWidget(self.stack, 1)
        self.setCentralWidget(root)

        self.pages: dict[str, QWidget] = {
            "dashboard": DashboardPage(state),
            "profiles": ProfilesPage(state),
            "excel": ExcelPage(state),
            "relationships": RelationshipsPage(state),
            "mapping": MappingPage(state),
            "templates": TemplatesPage(state),
            "preview": PreviewPage(state),
            "generate": GeneratePage(state),
            "logs": LogsPage(state),
            "settings": SettingsPage(state),
        }
        for key, label, icon in PAGES:
            item = QListWidgetItem(self.style().standardIcon(icon), tr(label))
            item.setData(Qt.ItemDataRole.UserRole, key)
            self.nav.addItem(item)
            container = QWidget()
            cl = QVBoxLayout(container)
            cl.setContentsMargins(18, 14, 18, 14)
            cl.addWidget(self.pages[key])
            self.stack.addWidget(container)
        self.nav.currentRowChanged.connect(self._switch)
        self.nav.setCurrentRow(0)
        state.navigate.connect(self.go)
        state.profile_changed.connect(self._update_title)
        state.profile_edited.connect(self._update_title)

        save = QAction(tr("Save profile"), self)
        save.setShortcut(QKeySequence.StandardKey.Save)
        save.triggered.connect(self._save)
        self.addAction(save)
        self._update_title()

    def go(self, key: str) -> None:
        for i in range(self.nav.count()):
            if self.nav.item(i).data(Qt.ItemDataRole.UserRole) == key:
                self.nav.setCurrentRow(i)
                return

    def _switch(self, row: int) -> None:
        self.stack.setCurrentIndex(row)
        page = self.pages[PAGES[row][0]]
        refresh = getattr(page, "refresh", None)
        if PAGES[row][0] == "dashboard" and callable(refresh):
            refresh()

    def _update_title(self) -> None:
        p = self.state.profile
        dirty = " •" if self.state.dirty else ""
        self.setWindowTitle(f"DocGen {__version__}" + (f" – {p.profile.name}{dirty}" if p else ""))
        self.profile_label.setText(
            tr("Profile: {name}", name=p.profile.name) + dirty if p else tr("No profile open")
        )

    def _save(self) -> None:
        try:
            self.state.save_profile()
        except (OSError, ValueError) as exc:
            error(self, tr("Cannot save: {err}", err=exc))

    def closeEvent(self, event: QCloseEvent) -> None:
        if self.state.dirty:
            answer = QMessageBox.question(
                self,
                "DocGen",
                tr("Save changes to the profile before closing?"),
                QMessageBox.StandardButton.Save
                | QMessageBox.StandardButton.Discard
                | QMessageBox.StandardButton.Cancel,
            )
            if answer == QMessageBox.StandardButton.Cancel:
                event.ignore()
                return
            if answer == QMessageBox.StandardButton.Save:
                self._save()
        event.accept()
