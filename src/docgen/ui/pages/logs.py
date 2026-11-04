"""Logs & errors: every diagnostic from the latest operations with location and suggestion."""

from __future__ import annotations

from PySide6.QtWidgets import QComboBox, QHBoxLayout, QLabel, QVBoxLayout, QWidget

from docgen.domain.diagnostics import Issue
from docgen.i18n.catalog import tr
from docgen.ui.state import AppState
from docgen.ui.widgets.common import IssueView, page_title


class LogsPage(QWidget):
    def __init__(self, state: AppState) -> None:
        super().__init__()
        self.state = state
        self.sources: dict[str, list[Issue]] = {}
        layout = QVBoxLayout(self)
        layout.addWidget(
            page_title(
                tr("Logs and errors"),
                tr(
                    "Each problem shows the sheet, row, column, rule, placeholder and a suggested correction."
                ),
            )
        )
        bar = QHBoxLayout()
        bar.addWidget(QLabel(tr("Source")))
        self.source = QComboBox()
        self.source.currentTextChanged.connect(self._show)
        bar.addWidget(self.source, 1)
        layout.addLayout(bar)
        self.view = IssueView()
        layout.addWidget(self.view, 1)
        state.issues_published.connect(self._publish)

    def _publish(self, source: str, issues: list[Issue]) -> None:
        self.sources[source] = list(issues)
        if self.source.findText(source) < 0:
            self.source.addItem(source)
        self.source.setCurrentText(source)
        self._show(source)

    def _show(self, source: str) -> None:
        self.view.set_issues(self.sources.get(source, []))
