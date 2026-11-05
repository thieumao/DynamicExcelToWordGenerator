"""Quick start wizard for non-IT users: Excel + Word files -> documents, in four steps.

1. Choose the Excel file: the app recognises the list (one row = one set of documents), the
   general information and the lookup tables, and says so in plain sentences.
2. Choose the Word files (their order is the page order when they are combined).
3. Check the automatic matching: every input found in the Word files with the value it will
   get, an example, and a drop-down to change it.
4. Name it, choose how the files are made, and create the documents.

All the technical parts (profile, sheets, relationships, fields, bindings) are built behind the
scenes by :mod:`docgen.application.quickstart`; the result is a normal profile that can be opened
again later and refined in the advanced pages.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from PySide6.QtCore import QStandardPaths, Qt
from PySide6.QtGui import QBrush, QColor
from PySide6.QtWidgets import (
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QRadioButton,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from docgen.application import quickstart as q
from docgen.application.workbook import suggest_identifier
from docgen.i18n.catalog import tr
from docgen.ui.state import AppState
from docgen.ui.widgets.common import PathPicker, button, error, page_title
from docgen.ui.workers import Worker, run_in_background

STEPS = ("Choose the Excel file", "Choose the Word files", "Check the matching", "Create the documents")
ROLES = (
    (q.ROLE_LIST, "List – each row makes one set of documents"),
    (q.ROLE_INFO, "General information – entered once (project details…)"),
    (q.ROLE_LOOKUP, "Lookup table – found by a code in the list"),
    (q.ROLE_IGNORE, "Not used"),
)
KINDS = {
    "marker": "marker <<…>>",
    "blank": "blank after a label",
    "content_control": "Word input box",
    "form_field": "form field",
}
MATCHED = QColor("#dcfce7")
UNMATCHED = QColor("#fef9c3")


def documents_folder() -> Path:
    location = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.DocumentsLocation)
    return Path(location or Path.home()) / "DocGen"


class QuickStartPage(QWidget):
    def __init__(self, state: AppState) -> None:
        super().__init__()
        self.state = state
        self.analysis: q.QuickAnalysis | None = None
        self.options: list[q.ValueOption] = []
        self.word_files: list[Path] = []
        self.scans: list[q.WordScan] = []
        self.worker: Worker | None = None
        layout = QVBoxLayout(self)
        layout.addWidget(
            page_title(
                tr("Create documents quickly"),
                tr(
                    "Choose your Excel file and your Word forms: the app finds what to fill and matches it "
                    "with the Excel columns. Nothing has to be configured; your files are never changed."
                ),
            )
        )
        self.steps = QLabel()
        self.steps.setTextFormat(Qt.TextFormat.RichText)
        layout.addWidget(self.steps)
        self.stack = QStackedWidget()
        self.stack.addWidget(self._excel_step())
        self.stack.addWidget(self._word_step())
        self.stack.addWidget(self._match_step())
        self.stack.addWidget(self._finish_step())
        layout.addWidget(self.stack, 1)
        nav = QHBoxLayout()
        self.back_btn = button(tr("‹ Back"))
        self.back_btn.clicked.connect(self.back)
        self.next_btn = button(tr("Next ›"), primary=True)
        self.next_btn.clicked.connect(self.next)
        self.busy = QLabel()
        self.busy.setObjectName("Hint")
        nav.addWidget(self.back_btn)
        nav.addWidget(self.busy, 1)
        nav.addWidget(self.next_btn)
        layout.addLayout(nav)
        self._update_nav()

    # ------------------------------------------------------------ steps
    def _excel_step(self) -> QWidget:
        w = QWidget()
        v = QVBoxLayout(w)
        row = QHBoxLayout()
        pick = button(tr("Choose Excel file…"), primary=True)
        pick.clicked.connect(self._choose_excel)
        self.excel_label = QLabel(tr("No file chosen yet (.xlsx, .xls or .csv)."))
        row.addWidget(pick)
        row.addWidget(self.excel_label, 1)
        v.addLayout(row)
        self.found = QLabel()
        self.found.setWordWrap(True)
        v.addWidget(self.found)
        self.sheets = QTableWidget(0, 3)
        self.sheets.setHorizontalHeaderLabels([tr("Sheet"), tr("Used as"), tr("Rows")])
        self.sheets.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.sheets.verticalHeader().setVisible(False)
        v.addWidget(self.sheets, 1)
        hint = QLabel(
            tr(
                "Tip: keep the project details (name, location, owner, contractor…) in a sheet with two "
                "columns 'label | value'; every document can use them."
            )
        )
        hint.setObjectName("Hint")
        hint.setWordWrap(True)
        v.addWidget(hint)
        return w

    def _word_step(self) -> QWidget:
        w = QWidget()
        v = QVBoxLayout(w)
        row = QHBoxLayout()
        add = button(tr("Add Word files…"), primary=True)
        add.clicked.connect(self._choose_words)
        row.addWidget(add)
        for text, slot in ((tr("↑"), lambda: self._move(-1)), (tr("↓"), lambda: self._move(1))):
            b = button(text)
            b.clicked.connect(slot)
            row.addWidget(b)
        remove = button(tr("Remove"))
        remove.clicked.connect(self._remove_word)
        row.addWidget(remove)
        row.addStretch(1)
        v.addLayout(row)
        self.word_list = QListWidget()
        v.addWidget(self.word_list, 1)
        hint = QLabel(
            tr(
                "Mark what to fill by typing the name of the information between << and >> "
                "(for example <<Tên công trình>>), or keep your form as it is: blanks such as "
                "'Họ và tên: ......', Word input boxes and form fields are found too. "
                "The order above is the page order when the files are combined."
            )
        )
        hint.setObjectName("Hint")
        hint.setWordWrap(True)
        v.addWidget(hint)
        return w

    def _match_step(self) -> QWidget:
        w = QWidget()
        v = QVBoxLayout(w)
        self.match_info = QLabel()
        self.match_info.setWordWrap(True)
        v.addWidget(self.match_info)
        self.matches = QTableWidget(0, 5)
        self.matches.setHorizontalHeaderLabels(
            [
                tr("Word file"),
                tr("Place to fill"),
                tr("Text around it"),
                tr("Value from Excel"),
                tr("Example"),
            ]
        )
        self.matches.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.matches.verticalHeader().setVisible(False)
        v.addWidget(self.matches, 1)
        return w

    def _finish_step(self) -> QWidget:
        w = QWidget()
        v = QVBoxLayout(w)
        box = QGroupBox(tr("Save and create"))
        form = QFormLayout(box)
        self.name = QLineEdit()
        form.addRow(tr("Name of this document set"), self.name)
        self.per_row = QRadioButton(tr("One Word file per Excel row (all forms in one file)"))
        self.per_form = QRadioButton(tr("One Word file per form and per row"))
        self.per_row.setChecked(True)
        group = QButtonGroup(w)
        group.addButton(self.per_row)
        group.addButton(self.per_form)
        form.addRow(tr("Files"), self.per_row)
        form.addRow("", self.per_form)
        self.combined = QCheckBox(tr("Also make one file with everything, ready to print"))
        self.combined.setChecked(True)
        form.addRow("", self.combined)
        self.output = PathPicker(folder=True)
        form.addRow(tr("Save the documents in"), self.output)
        v.addWidget(box)
        self.summary = QLabel()
        self.summary.setWordWrap(True)
        v.addWidget(self.summary)
        v.addStretch(1)
        return w

    # ------------------------------------------------------------ navigation
    def _update_nav(self) -> None:
        index = self.stack.currentIndex()
        parts = []
        for i, step in enumerate(STEPS):
            text = f"{i + 1}. {tr(step)}"
            parts.append(
                f"<b style='color:#1d4ed8'>{text}</b>"
                if i == index
                else f"<span style='color:#6b7280'>{text}</span>"
            )
        self.steps.setText("&nbsp;&nbsp;›&nbsp;&nbsp;".join(parts))
        self.back_btn.setEnabled(index > 0 and self.worker is None)
        ready = self.worker is None and (
            (index == 0 and self.analysis is not None and self.analysis.primary is not None)
            or (index == 1 and bool(self.word_files))
            or index >= 2
        )
        self.next_btn.setEnabled(ready)
        self.next_btn.setText(tr("Save and create the documents") if index == 3 else tr("Next ›"))

    def back(self) -> None:
        self.stack.setCurrentIndex(max(0, self.stack.currentIndex() - 1))
        self._update_nav()

    def next(self) -> None:
        index = self.stack.currentIndex()
        if index == 1:
            self._scan_words()
            return
        if index == 2:
            self._prepare_finish()
        if index == 3:
            self.finish()
            return
        self.stack.setCurrentIndex(index + 1)
        self._update_nav()

    def _run(self, fn: Any, *args: Any, on_done: Any, message: str) -> None:
        self.busy.setText(message)
        self.worker = run_in_background(fn, *args, on_done=self._done(on_done), on_error=self._failed)
        self._update_nav()

    def _done(self, callback: Any) -> Any:
        def handle(result: Any) -> None:
            self.worker = None
            self.busy.clear()
            callback(result)
            self._update_nav()

        return handle

    def _failed(self, message: str, details: str) -> None:
        self.worker = None
        self.busy.clear()
        self._update_nav()
        error(self, message, details)

    # ------------------------------------------------------------ step 1
    def _choose_excel(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, tr("Choose Excel file"), "", "Excel (*.xlsx *.xlsm *.xls *.csv)"
        )
        if path:
            self.open_excel(Path(path))

    def open_excel(self, path: Path) -> None:
        self.excel_label.setText(path.name)
        self._run(q.analyze_workbook, path, on_done=self._analysed, message=tr("Reading the Excel file…"))

    def _analysed(self, analysis: q.QuickAnalysis) -> None:
        self.analysis = analysis
        self.sheets.setRowCount(len(analysis.sheets))
        for r, sheet in enumerate(analysis.sheets):
            self.sheets.setItem(r, 0, QTableWidgetItem(sheet.name))
            combo = QComboBox()
            for role, label in ROLES:
                combo.addItem(tr(label), role)
            combo.setCurrentIndex(max(0, combo.findData(sheet.role)))
            combo.currentIndexChanged.connect(
                lambda _i, n=sheet.name, c=combo: self._set_role(n, c.currentData())
            )
            self.sheets.setCellWidget(r, 1, combo)
            count = len(sheet.info) if sheet.role == q.ROLE_INFO else sheet.row_count
            self.sheets.setItem(r, 2, QTableWidgetItem(str(count)))
        self.sheets.resizeColumnToContents(0)
        self._describe()

    def _set_role(self, sheet_name: str, role: str) -> None:
        assert self.analysis is not None
        self.analysis.set_role(sheet_name, role)
        for r, sheet in enumerate(self.analysis.sheets):
            combo = self.sheets.cellWidget(r, 1)
            if isinstance(combo, QComboBox) and combo.currentData() != sheet.role:
                combo.blockSignals(True)
                combo.setCurrentIndex(combo.findData(sheet.role))
                combo.blockSignals(False)
        self._describe()
        self._update_nav()

    def _describe(self) -> None:
        assert self.analysis is not None
        lines = q.describe(self.analysis)
        if self.analysis.primary is None:
            lines.append(tr("Choose which sheet is the list (one row = one set of documents)."))
        self.found.setText("\n".join(f"✔ {line}" for line in lines))

    # ------------------------------------------------------------ step 2
    def _choose_words(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(self, tr("Add Word files"), "", "Word (*.docx)")
        self.add_words([Path(f) for f in files])

    def add_words(self, files: list[Path]) -> None:
        for path in files:
            if path not in self.word_files:
                self.word_files.append(path)
        self._refresh_words()

    def _refresh_words(self) -> None:
        self.word_list.clear()
        for i, path in enumerate(self.word_files, start=1):
            self.word_list.addItem(f"{i}. {path.name}")
        self._update_nav()

    def _move(self, delta: int) -> None:
        row = self.word_list.currentRow()
        target = row + delta
        if 0 <= row < len(self.word_files) and 0 <= target < len(self.word_files):
            self.word_files[row], self.word_files[target] = self.word_files[target], self.word_files[row]
            self._refresh_words()
            self.word_list.setCurrentRow(target)

    def _remove_word(self) -> None:
        row = self.word_list.currentRow()
        if 0 <= row < len(self.word_files):
            del self.word_files[row]
            self._refresh_words()

    def _scan_words(self) -> None:
        files = list(self.word_files)
        self._run(
            lambda: [q.scan_word(f) for f in files],
            on_done=self._scanned,
            message=tr("Looking for the places to fill…"),
        )

    # ------------------------------------------------------------ step 3
    def _scanned(self, scans: list[q.WordScan]) -> None:
        assert self.analysis is not None
        self.scans = scans
        self.options = q.value_options(self.analysis)
        rows = [(i, scan, item) for i, scan in enumerate(scans) for item in scan.inputs]
        self.matches.setRowCount(len(rows))
        matched = 0
        for r, (i, scan, item) in enumerate(rows):
            suggestion = q.suggest([item], self.options).get(item.name)
            matched += suggestion is not None
            cells = [scan.path.name, item.name, item.context]
            for c, text in enumerate(cells):
                cell = QTableWidgetItem(text)
                cell.setToolTip(f"{tr(KINDS.get(item.kind, item.kind))}: {item.context}")
                cell.setData(Qt.ItemDataRole.UserRole, (i, item.name))
                self.matches.setItem(r, c, cell)
            combo = QComboBox()
            combo.addItem(tr("— leave empty —") if item.kind == "marker" else tr("— leave as it is —"), None)
            group = ""
            for option in self.options:
                if option.group != group:
                    group = option.group
                    combo.insertSeparator(combo.count())
                combo.addItem(option.label, option.path)
            combo.setCurrentIndex(max(0, combo.findData(suggestion)) if suggestion else 0)
            combo.currentIndexChanged.connect(lambda _i, row=r: self._paint(row))
            self.matches.setCellWidget(r, 3, combo)
            self.matches.setItem(r, 4, QTableWidgetItem())
            self._paint(r)
        self.matches.resizeColumnToContents(0)
        self.matches.resizeColumnToContents(1)
        self.matches.setColumnWidth(3, 280)
        self.matches.setColumnWidth(4, 240)
        self.match_info.setText(
            tr(
                "{matched} of {total} places were matched automatically (green). Check them, and choose a "
                "value for the yellow ones if they must be filled; the others stay as they are in Word.",
                matched=matched,
                total=len(rows),
            )
            + (
                "\n" + tr("Files written with {{ }} tags should be set up in the advanced pages.")
                if any(s.uses_jinja for s in scans)
                else ""
            )
        )
        self.stack.setCurrentIndex(2)

    def _paint(self, row: int) -> None:
        combo = self.matches.cellWidget(row, 3)
        assert isinstance(combo, QComboBox)
        path = combo.currentData()
        option = next((o for o in self.options if o.path == path), None)
        example = self.matches.item(row, 4)
        if example is not None:
            example.setText(option.sample if option else "")
        color = MATCHED if path else UNMATCHED
        for c in (0, 1, 2, 4):
            cell = self.matches.item(row, c)
            if cell is not None:
                cell.setBackground(QBrush(color))

    def chosen(self) -> dict[tuple[int, str], str | None]:
        result: dict[tuple[int, str], str | None] = {}
        for r in range(self.matches.rowCount()):
            key = self.matches.item(r, 0).data(Qt.ItemDataRole.UserRole)
            combo = self.matches.cellWidget(r, 3)
            if isinstance(combo, QComboBox):
                result[tuple(key)] = combo.currentData()  # type: ignore[index]
        return result

    def choose(self, file_index: int, name: str, path: str | None) -> None:
        """Set the value for one place (used by the tests and the 'match' buttons)."""
        for r in range(self.matches.rowCount()):
            if tuple(self.matches.item(r, 0).data(Qt.ItemDataRole.UserRole)) == (file_index, name):
                combo = self.matches.cellWidget(r, 3)
                assert isinstance(combo, QComboBox)
                combo.setCurrentIndex(max(0, combo.findData(path)) if path else 0)

    # ------------------------------------------------------------ step 4
    def _prepare_finish(self) -> None:
        assert self.analysis is not None
        if not self.name.text().strip():
            self.name.setText(q.default_name(self.analysis))
        if self.output.path() is None:
            self.output.edit.setText(str(documents_folder() / _folder_name(self.name.text())))
        primary = self.analysis.primary
        chosen = self.chosen()
        self.summary.setText(
            tr(
                "{rows} row(s) of '{sheet}' × {files} Word file(s); {filled} place(s) will be filled.",
                rows=primary.row_count if primary else 0,
                sheet=primary.name if primary else "",
                files=len(self.scans),
                filled=sum(1 for v in chosen.values() if v),
            )
        )

    def finish(self) -> None:
        assert self.analysis is not None
        name = self.name.text().strip() or q.default_name(self.analysis)
        out = self.output.path() or documents_folder() / _folder_name(name)
        repo = self.state.repo
        taken = {e.id for e in repo.list()} | {p.name for p in repo.root.glob("*") if p.is_dir()}
        profile_id = suggest_identifier(name, taken)
        try:
            words = [q.WordChoice(repo.import_template(profile_id, scan.path), scan) for scan in self.scans]
            q.apply_bindings(words, self.chosen())
            profile = q.build_profile(
                self.analysis,
                words,
                profile_id=profile_id,
                name=name,
                one_file_per_row=self.per_row.isChecked(),
                combined=self.combined.isChecked(),
            )
            repo.save(profile)
        except (OSError, ValueError) as exc:
            error(self, tr("The document set could not be created."), str(exc))
            return
        self.state.open_profile(profile_id)
        self.state.excel_path = self.analysis.path
        self.state.workbook = None  # the Excel page reads it again when opened
        self.state.workbook_changed.emit()
        self.state.navigate.emit("generate")
        self.state.run_requested.emit(str(self.analysis.path), str(out))


def _folder_name(name: str) -> str:
    from docgen.domain.naming import sanitize_component

    return sanitize_component(name, max_length=80) or "DocGen"
