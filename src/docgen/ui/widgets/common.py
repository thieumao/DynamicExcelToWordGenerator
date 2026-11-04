"""Reusable widgets and helpers."""

from __future__ import annotations

import csv
import datetime
from collections.abc import Sequence
from pathlib import Path
from typing import Any

from PySide6.QtCore import QAbstractTableModel, QModelIndex, QPersistentModelIndex, QSortFilterProxyModel, Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QComboBox,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTableView,
    QVBoxLayout,
    QWidget,
)

from docgen.domain.diagnostics import Issue, Severity
from docgen.i18n.catalog import issue_text, tr

AnyIndex = QModelIndex | QPersistentModelIndex


def page_title(text: str, subtitle: str = "") -> QWidget:
    box = QWidget()
    layout = QVBoxLayout(box)
    layout.setContentsMargins(0, 0, 0, 8)
    title = QLabel(text)
    title.setObjectName("PageTitle")
    layout.addWidget(title)
    if subtitle:
        sub = QLabel(subtitle)
        sub.setObjectName("PageSubtitle")
        sub.setWordWrap(True)
        layout.addWidget(sub)
    return box


def button(text: str, *, primary: bool = False, tooltip: str = "") -> QPushButton:
    btn = QPushButton(text)
    if primary:
        btn.setObjectName("Primary")
    if tooltip:
        btn.setToolTip(tooltip)
    return btn


def info(parent: QWidget, text: str) -> None:
    QMessageBox.information(parent, "DocGen", text)


def error(parent: QWidget, text: str, details: str = "") -> None:
    box = QMessageBox(QMessageBox.Icon.Critical, "DocGen", text, parent=parent)
    if details:
        box.setDetailedText(details)
    box.exec()


def confirm(parent: QWidget, text: str) -> bool:
    return (
        QMessageBox.question(
            parent, "DocGen", text, QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        == QMessageBox.StandardButton.Yes
    )


class PathPicker(QWidget):
    """Line edit + browse button for files or folders."""

    def __init__(self, *, folder: bool = False, filters: str = "", placeholder: str = "") -> None:
        super().__init__()
        self.folder = folder
        self.filters = filters
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self.edit = QLineEdit()
        self.edit.setPlaceholderText(placeholder)
        browse = QPushButton(tr("Browse…"))
        browse.clicked.connect(self._browse)
        layout.addWidget(self.edit, 1)
        layout.addWidget(browse)

    def _browse(self) -> None:
        if self.folder:
            path = QFileDialog.getExistingDirectory(self, tr("Choose folder"), self.edit.text())
        else:
            path, _ = QFileDialog.getOpenFileName(self, tr("Choose file"), self.edit.text(), self.filters)
        if path:
            self.edit.setText(path)
            self.edit.editingFinished.emit()

    def path(self) -> Path | None:
        text = self.edit.text().strip()
        return Path(text) if text else None

    def set_path(self, path: Path | str | None) -> None:
        self.edit.setText(str(path) if path else "")


SEVERITY_COLORS = {
    Severity.ERROR: QColor("#c62828"),
    Severity.WARNING: QColor("#b26a00"),
    Severity.INFO: QColor("#1565c0"),
}


class IssueTableModel(QAbstractTableModel):
    COLUMNS = (
        "Severity",
        "Category",
        "Message",
        "Sheet",
        "Row",
        "Column",
        "Field",
        "Template",
        "Placeholder",
        "Rule",
        "Unit",
        "Suggestion",
    )

    def __init__(self) -> None:
        super().__init__()
        self.issues: list[Issue] = []

    def set_issues(self, issues: Sequence[Issue]) -> None:
        self.beginResetModel()
        self.issues = list(issues)
        self.endResetModel()

    def append(self, issues: Sequence[Issue]) -> None:
        if not issues:
            return
        start = len(self.issues)
        self.beginInsertRows(QModelIndex(), start, start + len(issues) - 1)
        self.issues.extend(issues)
        self.endInsertRows()

    def rowCount(self, parent: AnyIndex = QModelIndex()) -> int:  # noqa: B008
        return 0 if parent.isValid() else len(self.issues)

    def columnCount(self, parent: AnyIndex = QModelIndex()) -> int:  # noqa: B008
        return len(self.COLUMNS)

    def headerData(
        self, section: int, orientation: Qt.Orientation, role: int = Qt.ItemDataRole.DisplayRole
    ) -> Any:
        if role == Qt.ItemDataRole.DisplayRole and orientation == Qt.Orientation.Horizontal:
            return tr(self.COLUMNS[section])
        return None

    def row_values(self, issue: Issue) -> list[str]:
        message, suggestion = issue_text(issue)
        return [
            tr(issue.severity.value),
            issue.category.value,
            message,
            issue.sheet or "",
            "" if issue.row is None else str(issue.row),
            issue.column or "",
            issue.field or "",
            issue.template or "",
            issue.placeholder or "",
            issue.rule or "",
            issue.unit or "",
            suggestion,
        ]

    def data(self, index: AnyIndex, role: int = Qt.ItemDataRole.DisplayRole) -> Any:
        if not index.isValid():
            return None
        issue = self.issues[index.row()]
        if role in (Qt.ItemDataRole.DisplayRole, Qt.ItemDataRole.ToolTipRole):
            return self.row_values(issue)[index.column()]
        if role == Qt.ItemDataRole.ForegroundRole and index.column() == 0:
            return SEVERITY_COLORS.get(issue.severity)
        if role == Qt.ItemDataRole.UserRole:
            return issue
        return None


class IssueFilter(QSortFilterProxyModel):
    def __init__(self) -> None:
        super().__init__()
        self.min_severity = Severity.INFO
        self.text = ""

    def filterAcceptsRow(self, source_row: int, source_parent: AnyIndex) -> bool:
        model = self.sourceModel()
        assert isinstance(model, IssueTableModel)
        issue = model.issues[source_row]
        if issue.severity.rank < self.min_severity.rank:
            return False
        if self.text:
            blob = " ".join(model.row_values(issue)).casefold()
            return self.text.casefold() in blob
        return True


class IssueView(QWidget):
    """Filterable issue table with CSV export."""

    def __init__(self, *, compact: bool = False) -> None:
        super().__init__()
        self.model = IssueTableModel()
        self.proxy = IssueFilter()
        self.proxy.setSourceModel(self.model)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        bar = QHBoxLayout()
        self.severity = QComboBox()
        for sev in (Severity.INFO, Severity.WARNING, Severity.ERROR):
            self.severity.addItem(tr(f"{sev.value} and above"), sev)
        self.severity.currentIndexChanged.connect(self._apply)
        self.search = QLineEdit()
        self.search.setPlaceholderText(tr("Search messages, sheets, fields…"))
        self.search.textChanged.connect(self._apply)
        self.summary = QLabel()
        export = QPushButton(tr("Export CSV"))
        export.clicked.connect(self._export)
        bar.addWidget(self.severity)
        bar.addWidget(self.search, 1)
        bar.addWidget(self.summary)
        bar.addWidget(export)
        layout.addLayout(bar)
        self.table = QTableView()
        self.table.setModel(self.proxy)
        self.table.setSortingEnabled(True)
        self.table.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
        self.table.setWordWrap(False)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(True)
        self.table.setColumnWidth(2, 420)
        if compact:
            for col in (1, 6, 7, 8, 9, 10):
                self.table.setColumnHidden(col, True)
        layout.addWidget(self.table, 1)
        self._update_summary()

    def set_issues(self, issues: Sequence[Issue]) -> None:
        self.model.set_issues(issues)
        self._update_summary()

    def append(self, issues: Sequence[Issue]) -> None:
        self.model.append(issues)
        self._update_summary()

    def _apply(self) -> None:
        self.proxy.min_severity = self.severity.currentData()
        self.proxy.text = self.search.text().strip()
        self.proxy.invalidateFilter()
        self._update_summary()

    def _update_summary(self) -> None:
        issues = self.model.issues
        errors = sum(1 for i in issues if i.severity is Severity.ERROR)
        warnings = sum(1 for i in issues if i.severity is Severity.WARNING)
        self.summary.setText(tr("{errors} error(s), {warnings} warning(s)", errors=errors, warnings=warnings))

    def _export(self) -> None:
        path, _ = QFileDialog.getSaveFileName(self, tr("Export CSV"), "issues.csv", "CSV (*.csv)")
        if not path:
            return
        with Path(path).open("w", encoding="utf-8-sig", newline="") as fh:
            writer = csv.writer(fh)
            writer.writerow([tr(c) for c in IssueTableModel.COLUMNS])
            for issue in self.model.issues:
                writer.writerow(self.model.row_values(issue))


class RowsTableModel(QAbstractTableModel):
    """Read-only grid for sheet previews."""

    def __init__(
        self,
        rows: Sequence[Sequence[Any]] = (),
        headers: Sequence[str] = (),
        *,
        first_row: int = 1,
        highlight_row: int | None = None,
    ) -> None:
        super().__init__()
        self.rows = [list(r) for r in rows]
        self.width = max((len(r) for r in self.rows), default=0)
        self.headers = list(headers)
        self.first_row = first_row
        self.highlight_row = highlight_row

    def rowCount(self, parent: AnyIndex = QModelIndex()) -> int:  # noqa: B008
        return 0 if parent.isValid() else len(self.rows)

    def columnCount(self, parent: AnyIndex = QModelIndex()) -> int:  # noqa: B008
        return 0 if parent.isValid() else max(self.width, len(self.headers))

    def headerData(
        self, section: int, orientation: Qt.Orientation, role: int = Qt.ItemDataRole.DisplayRole
    ) -> Any:
        if role != Qt.ItemDataRole.DisplayRole:
            return None
        if orientation == Qt.Orientation.Vertical:
            return str(self.first_row + section)
        from docgen.infrastructure.excel.sources import column_letter

        if section < len(self.headers) and self.headers[section]:
            return f"{column_letter(section)} · {self.headers[section]}"
        return column_letter(section)

    def data(self, index: AnyIndex, role: int = Qt.ItemDataRole.DisplayRole) -> Any:
        if not index.isValid():
            return None
        row = self.rows[index.row()]
        if role == Qt.ItemDataRole.DisplayRole:
            value = row[index.column()] if index.column() < len(row) else None
            if isinstance(value, datetime.datetime) and value.time() == datetime.time():
                return value.date().isoformat()
            return "" if value is None else str(value)
        if role == Qt.ItemDataRole.BackgroundRole and self.highlight_row is not None:
            if self.first_row + index.row() == self.highlight_row:
                return QColor("#e3f2fd")
        return None
