"""Help page: the Vietnamese user guide, inside the application (no Internet needed)."""

from __future__ import annotations

from PySide6.QtWidgets import QTextBrowser, QVBoxLayout, QWidget

from docgen.help import guide_text
from docgen.i18n.catalog import tr
from docgen.ui.state import AppState
from docgen.ui.widgets.common import page_title


class HelpPage(QWidget):
    def __init__(self, state: AppState) -> None:
        super().__init__()
        self.state = state
        layout = QVBoxLayout(self)
        layout.addWidget(page_title(tr("User guide"), tr("Step by step, for first-time users.")))
        self.browser = QTextBrowser()
        self.browser.setOpenExternalLinks(True)
        self.browser.setStyleSheet("QTextBrowser { padding: 12px 18px; font-size: 11pt; background: white; }")
        self.browser.setMarkdown(guide_text())
        layout.addWidget(self.browser, 1)
