"""Application settings: language, default output folder, workers, logging."""

from __future__ import annotations

from PySide6.QtCore import QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import QComboBox, QFormLayout, QLabel, QSpinBox, QVBoxLayout, QWidget

from docgen import __version__
from docgen.i18n.catalog import SUPPORTED, tr
from docgen.infrastructure.storage import paths
from docgen.ui.state import AppState
from docgen.ui.widgets.common import PathPicker, button, info, page_title


class SettingsPage(QWidget):
    def __init__(self, state: AppState) -> None:
        super().__init__()
        self.state = state
        s = state.settings
        layout = QVBoxLayout(self)
        layout.addWidget(page_title(tr("Settings")))
        form = QFormLayout()
        self.language = QComboBox()
        for code, label in SUPPORTED.items():
            self.language.addItem(label, code)
        self.language.setCurrentIndex(max(0, self.language.findData(s.language)))
        form.addRow(tr("Language"), self.language)
        self.output = PathPicker(folder=True)
        self.output.set_path(s.default_output_dir)
        form.addRow(tr("Default output folder"), self.output)
        self.workers = QSpinBox()
        self.workers.setRange(0, 16)
        self.workers.setSpecialValueText(tr("automatic"))
        self.workers.setValue(s.workers)
        form.addRow(tr("Parallel workers"), self.workers)
        layout.addLayout(form)
        save = button(tr("Save settings"), primary=True)
        save.clicked.connect(self._save)
        layout.addWidget(save)
        data = button(tr("Open application data folder"))
        data.clicked.connect(lambda: QDesktopServices.openUrl(QUrl.fromLocalFile(str(paths.data_dir()))))
        layout.addWidget(data)
        about = QLabel(
            tr(
                "DocGen {version} – offline document generator. No telemetry, no uploads.",
                version=__version__,
            )
        )
        about.setObjectName("Hint")
        layout.addWidget(about)
        layout.addStretch(1)

    def _save(self) -> None:
        s = self.state.settings
        changed_language = s.language != self.language.currentData()
        s.language = self.language.currentData()
        out = self.output.path()
        s.default_output_dir = str(out) if out else ""
        s.workers = self.workers.value()
        self.state.save_settings()
        info(
            self,
            tr("Settings saved.")
            + (" " + tr("Restart the application to change the language.") if changed_language else ""),
        )
