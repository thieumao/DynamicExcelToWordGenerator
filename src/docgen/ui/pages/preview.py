"""Preview: normalised data, resolved mappings with traces, missing fields, generation plan."""

from __future__ import annotations

from typing import Any

from PySide6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from docgen.cli.main import parse_rows
from docgen.i18n.catalog import tr
from docgen.ui.state import AppState
from docgen.ui.widgets.common import IssueView, button, error, page_title
from docgen.ui.workers import run_in_background


def fill_tree(parent: QTreeWidgetItem | QTreeWidget, key: str, value: Any) -> None:
    if isinstance(value, dict):
        node = QTreeWidgetItem([key, ""])
        _add(parent, node)
        for k, v in value.items():
            fill_tree(node, str(k), v)
    elif isinstance(value, list):
        node = QTreeWidgetItem([key, tr("{n} item(s)", n=len(value))])
        _add(parent, node)
        for i, v in enumerate(value):
            fill_tree(node, f"[{i}]", v)
    else:
        _add(parent, QTreeWidgetItem([key, "" if value is None else str(value)]))


def _add(parent: QTreeWidgetItem | QTreeWidget, node: QTreeWidgetItem) -> None:
    if isinstance(parent, QTreeWidget):
        parent.addTopLevelItem(node)
    else:
        parent.addChild(node)


class PreviewPage(QWidget):
    def __init__(self, state: AppState) -> None:
        super().__init__()
        self.state = state
        layout = QVBoxLayout(self)
        layout.addWidget(
            page_title(
                tr("Preview"),
                tr(
                    "Check the resolved data before generating: values, where each value came from, problems, "
                    "and the list of documents that will be produced."
                ),
            )
        )
        bar = QHBoxLayout()
        self.rows = QLineEdit()
        self.rows.setPlaceholderText(tr("Excel rows, e.g. 5,7-9 (empty = first 5)"))
        run = button(tr("Preview"), primary=True)
        run.clicked.connect(self._run)
        bar.addWidget(self.rows, 1)
        bar.addWidget(run)
        layout.addLayout(bar)
        self.status = QLabel()
        layout.addWidget(self.status)
        tabs = QTabWidget()
        split = QSplitter()
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels([tr("Field"), tr("Value")])
        split.addWidget(self.tree)
        self.traces = QTableWidget(0, 4)
        self.traces.setHorizontalHeaderLabels([tr("Field"), tr("Value"), tr("Source cells"), tr("Steps")])
        self.traces.horizontalHeader().setStretchLastSection(True)
        split.addWidget(self.traces)
        tabs.addTab(split, tr("Data & traceability"))
        self.issues = IssueView()
        tabs.addTab(self.issues, tr("Problems"))
        self.plan = QTableWidget(0, 3)
        self.plan.setHorizontalHeaderLabels([tr("Unit"), tr("Excel rows"), tr("Output files")])
        self.plan.horizontalHeader().setStretchLastSection(True)
        tabs.addTab(self.plan, tr("Generation plan"))
        self.relationships = QTableWidget(0, 4)
        self.relationships.setHorizontalHeaderLabels(
            [tr("Relationship"), tr("Target sheet"), tr("Keys"), tr("Duplicated keys")]
        )
        tabs.addTab(self.relationships, tr("Lookup indexes"))
        layout.addWidget(tabs, 1)

    def _run(self) -> None:
        p, excel, folder = self.state.profile, self.state.excel_path, self.state.profile_dir
        if p is None or excel is None or folder is None:
            error(self, tr("Open a profile and load a workbook first."))
            return
        from docgen.application.preview import preview

        try:
            rows = parse_rows(self.rows.text())
        except ValueError:
            error(self, tr("Invalid row list"))
            return
        self.status.setText(tr("Working…"))
        run_in_background(
            preview,
            p.model_copy(deep=True),
            folder,
            excel,
            rows=rows,
            limit=5,
            on_done=self._show,
            on_error=lambda m, d: (self.status.setText(m), error(self, m, d)),
        )

    def _show(self, result: Any) -> None:
        self.tree.clear()
        self.traces.setRowCount(0)
        all_issues = list(result.issues)
        for rec in result.records:
            fill_tree(self.tree, tr("Row {row}", row=rec.excel_row), rec.context)
            all_issues.extend(rec.issues)
            for t in rec.traces:
                r = self.traces.rowCount()
                self.traces.insertRow(r)
                for c, v in enumerate(
                    [
                        f"[{rec.excel_row}] {t.path}",
                        repr(t.value),
                        ", ".join(str(s) for s in t.sources),
                        " → ".join(t.steps),
                    ]
                ):
                    self.traces.setItem(r, c, QTableWidgetItem(v))
        self.tree.expandToDepth(1)
        self.tree.resizeColumnToContents(0)
        self.traces.resizeColumnsToContents()
        self.issues.set_issues(all_issues)
        self.plan.setRowCount(len(result.plan))
        for r, (index, label, rows, names) in enumerate(result.plan):
            for c, v in enumerate([f"{index}. {label}", ", ".join(map(str, rows)), " | ".join(names)]):
                self.plan.setItem(r, c, QTableWidgetItem(v))
        self.plan.resizeColumnsToContents()
        self.relationships.setRowCount(len(result.relationships))
        for r, s in enumerate(result.relationships):
            for c, v in enumerate(
                [
                    s.relationship,
                    s.target_sheet,
                    str(s.keys),
                    f"{s.duplicate_keys} {s.sample_duplicates or ''}",
                ]
            ):
                self.relationships.setItem(r, c, QTableWidgetItem(v))
        errors = sum(1 for i in all_issues if i.severity.value == "error")
        self.status.setText(
            tr(
                "{records} record(s) previewed, {units} unit(s) planned, {errors} error(s)",
                records=len(result.records),
                units=len(result.plan),
                errors=errors,
            )
        )
        self.state.publish_issues(tr("Preview"), all_issues)
