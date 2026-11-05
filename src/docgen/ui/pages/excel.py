"""Excel configuration: choose workbook, detect sheets/headers, preview, configure sheets and columns."""

from __future__ import annotations

from pathlib import Path

from pydantic import ValidationError
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QSpinBox,
    QSplitter,
    QTableView,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from docgen.application.workbook import inspect_workbook, sheet_config_from_summary, suggest_identifier
from docgen.domain.profile.models import ColumnSpec, ColumnType, HeaderConfig, SheetConfig
from docgen.i18n.catalog import tr
from docgen.ui.state import AppState
from docgen.ui.widgets.common import PathPicker, RowsTableModel, button, error, info, page_title
from docgen.ui.workers import run_in_background

EXCEL_FILTER = "Excel (*.xlsx *.xlsm *.xls *.csv)"
COLUMN_HEADERS = ["ID", "Header", "Aliases", "Position", "Type", "Required"]


class ExcelPage(QWidget):
    def __init__(self, state: AppState) -> None:
        super().__init__()
        self.state = state
        layout = QVBoxLayout(self)
        layout.addWidget(
            page_title(
                tr("Excel configuration"),
                tr(
                    "Choose the customer's workbook, check the detected sheets and header rows, then declare the "
                    "sheets and columns the profile uses. The original file is never modified."
                ),
            )
        )
        row = QHBoxLayout()
        self.picker = PathPicker(
            filters=EXCEL_FILTER, placeholder=tr("Excel workbook (.xlsx, .xlsm, .xls, .csv)")
        )
        load = button(tr("Load workbook"), primary=True)
        load.clicked.connect(self._load)
        row.addWidget(self.picker, 1)
        row.addWidget(load)
        layout.addLayout(row)
        self.status = QLabel()
        layout.addWidget(self.status)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        left = QWidget()
        left_layout = QVBoxLayout(left)
        left_layout.addWidget(QLabel(tr("Sheets in workbook")))
        self.sheets = QListWidget()
        self.sheets.currentRowChanged.connect(self._show_sheet)
        left_layout.addWidget(self.sheets, 1)
        add = button(tr("Add selected sheet to profile"))
        add.clicked.connect(self._add_sheet)
        left_layout.addWidget(add)
        splitter.addWidget(left)
        self.preview = QTableView()
        splitter.addWidget(self.preview)
        splitter.setSizes([260, 900])
        layout.addWidget(splitter, 2)

        cfg_box = QGroupBox(tr("Profile sheets"))
        cfg = QHBoxLayout(cfg_box)
        side = QVBoxLayout()
        self.profile_sheets = QListWidget()
        self.profile_sheets.currentRowChanged.connect(self._show_profile_sheet)
        side.addWidget(self.profile_sheets, 1)
        for label, slot in (
            (tr("Set as primary sheet"), self._set_primary),
            (tr("Remove sheet"), self._remove_sheet),
        ):
            btn = button(label)
            btn.clicked.connect(slot)
            side.addWidget(btn)
        cfg.addLayout(side, 1)
        detail = QVBoxLayout()
        form = QFormLayout()
        self.match_name = QLabel()
        form.addRow(tr("Matches sheet"), self.match_name)
        self.header_mode = QComboBox()
        for mode in ("auto", "fixed", "none"):
            self.header_mode.addItem(tr(f"header: {mode}"), mode)
        self.header_row = QSpinBox()
        self.header_row.setRange(1, 1000)
        hr = QHBoxLayout()
        hr.addWidget(self.header_mode)
        hr.addWidget(QLabel(tr("row")))
        hr.addWidget(self.header_row)
        hr.addStretch(1)
        form.addRow(tr("Header row"), hr)
        detail.addLayout(form)
        self.columns = QTableWidget(0, len(COLUMN_HEADERS))
        self.columns.setHorizontalHeaderLabels([tr(h) for h in COLUMN_HEADERS])
        self.columns.horizontalHeader().setStretchLastSection(True)
        detail.addWidget(self.columns, 1)
        col_bar = QHBoxLayout()
        self.add_header = QComboBox()
        col_bar.addWidget(self.add_header, 1)
        for label, slot in (
            (tr("Add column"), self._add_column),
            (tr("Remove column"), self._remove_column),
            (tr("Apply sheet changes"), self._apply_sheet),
        ):
            btn = button(label, primary=label == tr("Apply sheet changes"))
            btn.clicked.connect(slot)
            col_bar.addWidget(btn)
        detail.addLayout(col_bar)
        cfg.addLayout(detail, 3)
        layout.addWidget(cfg_box, 2)

        state.profile_changed.connect(self._refresh_profile)
        state.workbook_changed.connect(self._refresh_workbook)
        self._refresh_profile()

    # ---------------------------------------------------------- workbook
    def _load(self) -> None:
        path = self.picker.path()
        if not path:
            return
        self.status.setText(tr("Reading {name}…", name=path.name))
        run_in_background(
            inspect_workbook,
            path,
            owner=self,
            on_done=lambda s: self.state.set_workbook(path, s),
            on_error=lambda msg, det: (self.status.setText(msg), error(self, msg, det)),
        )

    def _refresh_workbook(self) -> None:
        summary = self.state.workbook
        self.sheets.clear()
        if summary is None:
            return
        self.picker.set_path(self.state.excel_path)
        self.status.setText(
            tr(
                "{name}: {count} sheet(s), format {fmt}",
                name=summary.path.name,
                count=len(summary.sheets),
                fmt=summary.format,
            )
        )
        for sheet in summary.sheets:
            label = f"{sheet.index}. {sheet.name}"
            if sheet.header_row:
                label += "  " + tr("(header row {row})", row=sheet.header_row)
            if sheet.hidden:
                label += "  [" + tr("hidden") + "]"
            self.sheets.addItem(label)
        if summary.sheets:
            self.sheets.setCurrentRow(0)
        self._show_profile_sheet(self.profile_sheets.currentRow())

    def _show_sheet(self, row: int) -> None:
        summary = self.state.workbook
        if summary is None or row < 0:
            return
        sheet = summary.sheets[row]
        self.preview.setModel(RowsTableModel(sheet.preview, highlight_row=sheet.header_row))
        self.preview.resizeColumnsToContents()

    def _add_sheet(self) -> None:
        profile = self.state.profile
        summary = self.state.workbook
        row = self.sheets.currentRow()
        if profile is None or summary is None or row < 0:
            error(self, tr("Open a profile and load a workbook first."))
            return
        sheet = summary.sheets[row]
        taken = {s.id for s in profile.excel.sheets} | {r.id for r in profile.relationships}
        cfg = sheet_config_from_summary(sheet, suggest_identifier(sheet.name, taken))
        placeholder = (
            len(profile.excel.sheets) == 1
            and profile.excel.sheets[0].id == "main"
            and profile.excel.sheets[0].columns[0].id == "col_a"
        )
        if placeholder:
            profile.excel.sheets = [cfg]
            profile.excel.primary = cfg.id
        else:
            profile.excel.sheets.append(cfg)
        self.state.mark_edited()
        self._refresh_profile()
        info(
            self,
            tr(
                "Sheet '{name}' added as '{sid}' with {n} column(s). Review the columns below.",
                name=sheet.name,
                sid=cfg.id,
                n=len(cfg.columns),
            ),
        )

    # ----------------------------------------------------------- profile
    def _refresh_profile(self) -> None:
        self.profile_sheets.clear()
        profile = self.state.profile
        if profile is None:
            return
        for s in profile.excel.sheets:
            label = s.id + (" ★ " + tr("primary") if s.id == profile.excel.primary else "")
            item = QListWidgetItem(f"{label}  ←  {s.match.name or s.match.regex or s.match.index}")
            item.setData(Qt.ItemDataRole.UserRole, s.id)
            self.profile_sheets.addItem(item)
        if profile.excel.sheets:
            self.profile_sheets.setCurrentRow(0)

    def _current_sheet(self) -> SheetConfig | None:
        profile = self.state.profile
        item = self.profile_sheets.currentItem()
        if profile is None or item is None:
            return None
        return profile.excel.sheet(item.data(Qt.ItemDataRole.UserRole))

    def _show_profile_sheet(self, _row: int) -> None:
        cfg = self._current_sheet()
        self.columns.setRowCount(0)
        self.add_header.clear()
        if cfg is None:
            return
        self.match_name.setText(str(cfg.match.name or cfg.match.regex or f"#{cfg.match.index}"))
        self.header_mode.setCurrentIndex(["auto", "fixed", "none"].index(cfg.header.mode))
        self.header_row.setValue(cfg.header.row or 1)
        for col in cfg.columns:
            self._append_column_row(col)
        summary = self.state.workbook
        if summary:
            sheet = next((s for s in summary.sheets if s.name == cfg.match.name), None)
            if sheet:
                for h in sheet.headers:
                    if h:
                        self.add_header.addItem(h)

    def _append_column_row(self, col: ColumnSpec) -> None:
        r = self.columns.rowCount()
        self.columns.insertRow(r)
        values = [
            col.id,
            col.header or "",
            ", ".join(col.aliases),
            "" if col.position is None else str(col.position),
        ]
        for c, v in enumerate(values):
            self.columns.setItem(r, c, QTableWidgetItem(v))
        combo = QComboBox()
        for t in ColumnType:
            combo.addItem(t.value)
        combo.setCurrentText(col.type.value)
        self.columns.setCellWidget(r, 4, combo)
        check = QCheckBox()
        check.setChecked(col.required)
        self.columns.setCellWidget(r, 5, check)

    def _add_column(self) -> None:
        header = self.add_header.currentText().strip()
        cfg = self._current_sheet()
        if cfg is None:
            return
        taken = {
            self.columns.item(r, 0).text() for r in range(self.columns.rowCount()) if self.columns.item(r, 0)
        }
        self._append_column_row(
            ColumnSpec(
                id=suggest_identifier(header or "column", taken),
                header=header or None,
                position=None if header else "A",
            )
        )

    def _remove_column(self) -> None:
        row = self.columns.currentRow()
        if row >= 0:
            self.columns.removeRow(row)

    def _apply_sheet(self) -> None:
        cfg = self._current_sheet()
        profile = self.state.profile
        if cfg is None or profile is None:
            return
        columns = []
        try:
            for r in range(self.columns.rowCount()):

                def text(c: int, r: int = r) -> str:
                    item = self.columns.item(r, c)
                    return item.text().strip() if item else ""

                position: str | int | None = text(3) or None
                if isinstance(position, str) and position.isdigit():
                    position = int(position)
                type_combo = self.columns.cellWidget(r, 4)
                required = self.columns.cellWidget(r, 5)
                assert isinstance(type_combo, QComboBox) and isinstance(required, QCheckBox)
                columns.append(
                    ColumnSpec(
                        id=text(0),
                        header=text(1) or None,
                        aliases=[a.strip() for a in text(2).split(",") if a.strip()],
                        position=position,
                        type=ColumnType(type_combo.currentText()),
                        required=required.isChecked(),
                    )
                )
            mode = self.header_mode.currentData()
            header = HeaderConfig(
                mode=mode,
                row=self.header_row.value() if mode == "fixed" else None,
                rows=cfg.header.rows,
                scan_rows=cfg.header.scan_rows,
            )
            new_cfg = SheetConfig(
                id=cfg.id,
                match=cfg.match,
                header=header,
                data=cfg.data,
                columns=columns,
                key=[k for k in cfg.key if k in {c.id for c in columns}],
                formula_policy=cfg.formula_policy,
            )
        except (ValidationError, ValueError) as exc:
            error(self, tr("Invalid sheet configuration"), str(exc))
            return
        profile.excel.sheets = [new_cfg if s.id == cfg.id else s for s in profile.excel.sheets]
        self.state.mark_edited()
        info(self, tr("Sheet '{sid}' updated. Save the profile to keep the changes.", sid=cfg.id))

    def _set_primary(self) -> None:
        cfg = self._current_sheet()
        if cfg and self.state.profile:
            self.state.profile.excel.primary = cfg.id
            self.state.mark_edited()
            self._refresh_profile()

    def _remove_sheet(self) -> None:
        cfg = self._current_sheet()
        profile = self.state.profile
        if cfg is None or profile is None:
            return
        if cfg.id == profile.excel.primary or len(profile.excel.sheets) == 1:
            error(self, tr("The primary sheet cannot be removed."))
            return
        profile.excel.sheets = [s for s in profile.excel.sheets if s.id != cfg.id]
        self.state.mark_edited()
        self._refresh_profile()

    def set_excel(self, path: Path) -> None:
        self.picker.set_path(path)
        self._load()
