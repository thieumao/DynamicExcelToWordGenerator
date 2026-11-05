"""Generate: choose mode/records/templates/output, validate, run with progress, cancel safely."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QProgressBar,
    QSpinBox,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from docgen.application.generation.job import GenerationRequest, JobResult
from docgen.application.generation.service import DocumentGenerationService
from docgen.cli.main import parse_rows
from docgen.domain.diagnostics import Issue
from docgen.domain.profile.models import ErrorPolicy, OutputKind, UnitKind
from docgen.i18n.catalog import tr
from docgen.ui.state import AppState
from docgen.ui.widgets.common import IssueView, PathPicker, button, error, page_title
from docgen.ui.workers import Worker, run_in_background

MODES = [
    ("A", "A – One record → one document", UnitKind.RECORD, OutputKind.SEPARATE, True),
    ("B", "B – Batch: each record → its documents", UnitKind.RECORD, OutputKind.SEPARATE, False),
    ("C", "C – Grouped: related records → one document", UnitKind.GROUP, OutputKind.MERGED, False),
    ("D", "D – Merged: all records → one merged document", UnitKind.DATASET, OutputKind.MERGED, False),
    ("E", "E – Batch merged: each record/group → merged document", UnitKind.RECORD, OutputKind.MERGED, False),
]
STAGES = {
    "validate": "Validating",
    "read_excel": "Reading Excel",
    "index": "Building lookup indexes",
    "map": "Mapping records",
    "generate": "Generating documents",
}


class GeneratePage(QWidget):
    def __init__(self, state: AppState) -> None:
        super().__init__()
        self.state = state
        self.worker: Worker | None = None
        self.result: JobResult | None = None
        layout = QVBoxLayout(self)
        layout.addWidget(
            page_title(
                tr("Generate documents"),
                tr(
                    "Validate first, then generate. Outputs go to a separate folder; existing files are never "
                    "overwritten unless the profile says so."
                ),
            )
        )
        top = QHBoxLayout()
        box = QGroupBox(tr("Job"))
        form = QFormLayout(box)
        self.profile_label = QLabel()
        form.addRow(tr("Profile"), self.profile_label)
        self.excel = PathPicker(filters="Excel (*.xlsx *.xlsm *.xls *.csv)")
        form.addRow(tr("Excel input"), self.excel)
        self.output = PathPicker(folder=True)
        form.addRow(tr("Output folder"), self.output)
        self.mode = QComboBox()
        self.mode.addItem(tr("Use profile settings"), None)
        for code, label, *_ in MODES:
            self.mode.addItem(tr(label), code)
        form.addRow(tr("Generation mode"), self.mode)
        self.group_note = QLabel()
        form.addRow("", self.group_note)
        self.mode.currentIndexChanged.connect(self._mode_changed)
        self.rows = QLineEdit()
        self.rows.setPlaceholderText(tr("all records – or Excel rows like 5,7-9"))
        form.addRow(tr("Records"), self.rows)
        self.both = QCheckBox(tr("Also save individual template documents"))
        form.addRow("", self.both)
        self.fail_fast = QCheckBox(tr("Stop at the first error (fail-fast)"))
        form.addRow("", self.fail_fast)
        self.workers = QSpinBox()
        self.workers.setRange(0, 16)
        self.workers.setSpecialValueText(tr("automatic"))
        form.addRow(tr("Parallel workers"), self.workers)
        top.addWidget(box, 3)
        tbox = QGroupBox(tr("Word files"))
        tl = QVBoxLayout(tbox)
        self.templates = QListWidget()
        tl.addWidget(self.templates)
        add_word = button(tr("Add Word files…"))
        add_word.setToolTip(
            tr(
                "Choose the customer's Word files (.docx). They are copied into the profile; "
                "the originals are never modified."
            )
        )
        add_word.clicked.connect(self._add_word_files)
        tl.addWidget(add_word)
        self.word_hint = QLabel()
        self.word_hint.setWordWrap(True)
        self.word_hint.setObjectName("Hint")
        tl.addWidget(self.word_hint)
        top.addWidget(tbox, 2)
        layout.addLayout(top)

        actions = QHBoxLayout()
        self.validate_btn = button(tr("Validate"))
        self.generate_btn = button(tr("Generate"), primary=True)
        self.cancel_btn = button(tr("Cancel"))
        self.open_btn = button(tr("Open output folder"))
        self.cancel_btn.setEnabled(False)
        self.validate_btn.clicked.connect(lambda: self._start(dry_run=True))
        self.generate_btn.clicked.connect(lambda: self._start(dry_run=False))
        self.cancel_btn.clicked.connect(self._cancel)
        self.open_btn.clicked.connect(self._open_folder)
        for b in (self.validate_btn, self.generate_btn, self.cancel_btn, self.open_btn):
            actions.addWidget(b)
        actions.addStretch(1)
        layout.addLayout(actions)
        self.stage = QLabel()
        layout.addWidget(self.stage)
        self.progress = QProgressBar()
        layout.addWidget(self.progress)
        split = QSplitter(Qt.Orientation.Vertical)
        self.units = QTableWidget(0, 4)
        self.units.setHorizontalHeaderLabels([tr("Unit"), tr("Status"), tr("Excel rows"), tr("Outputs")])
        self.units.horizontalHeader().setStretchLastSection(True)
        self.units.cellDoubleClicked.connect(self._open_unit)
        split.addWidget(self.units)
        self.issues = IssueView()
        split.addWidget(self.issues)
        layout.addWidget(split, 1)
        state.profile_changed.connect(self.refresh)
        state.profile_edited.connect(self.refresh)
        state.workbook_changed.connect(lambda: self.excel.set_path(state.excel_path))
        self.refresh()

    def refresh(self) -> None:
        p = self.state.profile
        self.profile_label.setText(
            f"{p.profile.name} ({p.profile.id}) v{p.profile.version}" + (" *" if self.state.dirty else "")
            if p
            else tr("No profile open")
        )
        if self.state.excel_path and not self.excel.path():
            self.excel.set_path(self.state.excel_path)
        if not self.output.path() and self.state.settings.default_output_dir:
            self.output.set_path(self.state.settings.default_output_dir)
        self.templates.clear()
        if p:
            for t in p.ordered_templates():
                item = QListWidgetItem(t.id)
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                item.setCheckState(Qt.CheckState.Checked if t.enabled else Qt.CheckState.Unchecked)
                self.templates.addItem(item)
        self._mode_changed()

    def _add_word_files(self) -> None:
        p = self.state.profile
        if p is None:
            error(self, tr("Open a profile first."))
            return
        files, _ = QFileDialog.getOpenFileNames(self, tr("Add Word files"), "", "Word (*.docx)")
        if not files:
            return
        from docgen.application.workbook import suggest_identifier
        from docgen.domain.profile.models import TemplateSpec

        added = []
        for name in files:
            try:
                rel = self.state.repo.import_template(p.profile.id, Path(name))
            except (OSError, ValueError) as exc:
                error(self, str(exc))
                continue
            tid = suggest_identifier(Path(name).stem, {t.id for t in p.templates})
            p.templates.append(TemplateSpec(id=tid, file=rel))
            added.append(tid)
        if added:
            p.merge.order = [t.id for t in p.ordered_templates()]
            self.state.mark_edited()
            self.refresh()
            self.word_hint.setText(
                tr(
                    "Added: {ids}. Open the Word templates page to bind the inputs found in them "
                    "(blanks, content controls, form fields) to Excel fields.",
                    ids=", ".join(added),
                )
            )

    def _mode_changed(self) -> None:
        code = self.mode.currentData()
        p = self.state.profile
        if code == "C" and p is not None and not p.generation.group_by:
            self.group_note.setText(tr("⚠ Grouped mode needs generation.group_by in the profile."))
        else:
            self.group_note.setText("")

    def _request(self, dry_run: bool) -> GenerationRequest | None:
        p, folder = self.state.profile, self.state.profile_dir
        excel, out = self.excel.path(), self.output.path()
        if p is None or folder is None:
            error(self, tr("Open a profile first."))
            return None
        if excel is None or not excel.exists():
            error(self, tr("Choose an existing Excel file."))
            return None
        if out is None and not dry_run:
            error(self, tr("Choose an output folder."))
            return None
        try:
            rows = parse_rows(self.rows.text())
        except ValueError:
            error(self, tr("Invalid row list"))
            return None
        unit = output = None
        code = self.mode.currentData()
        for mcode, _, munit, mout, single in MODES:
            if mcode == code:
                unit, output = munit, mout
                if single and not rows:
                    error(self, tr("Mode A needs one Excel row in 'Records'."))
                    return None
        if self.both.isChecked():
            output = OutputKind.BOTH
        templates = [
            self.templates.item(i).text()
            for i in range(self.templates.count())
            if self.templates.item(i).checkState() == Qt.CheckState.Checked
        ]
        return GenerationRequest(
            profile=p.model_copy(deep=True),
            profile_dir=folder,
            excel_path=excel,
            output_dir=out or folder,
            unit=unit,
            output=output,
            rows=rows,
            template_ids=templates,
            workers=self.workers.value() or (self.state.settings.workers or None),
            error_policy=ErrorPolicy.FAIL_FAST if self.fail_fast.isChecked() else None,
            dry_run=dry_run,
        )

    def _start(self, *, dry_run: bool) -> None:
        if self.worker is not None:
            return
        request = self._request(dry_run)
        if request is None:
            return
        self.units.setRowCount(0)
        self.issues.set_issues([])
        self.progress.setRange(0, 0)
        self.stage.setText(tr("Starting…"))
        self._set_running(True)
        service = DocumentGenerationService()
        self.worker = run_in_background(
            service.run,
            request,
            with_reporter=True,
            on_done=lambda r: self._done(r, request),
            on_error=self._failed,
            on_progress=self._progress,
            on_stage=self._stage,
            on_issue=self._issue,
        )

    def _set_running(self, running: bool) -> None:
        self.validate_btn.setEnabled(not running)
        self.generate_btn.setEnabled(not running)
        self.cancel_btn.setEnabled(running)

    def _stage(self, name: str) -> None:
        self.stage.setText(tr(STAGES.get(name, name)) + "…")

    def _progress(self, done: int, total: int, message: str) -> None:
        if total > 0:
            self.progress.setRange(0, total)
            self.progress.setValue(done)
            self.stage.setText(f"{self.stage.text().split(' – ')[0]} – {done}/{total} {message}")

    def _issue(self, issue: Issue) -> None:
        self.issues.append([issue])

    def _cancel(self) -> None:
        if self.worker is not None:
            self.worker.cancel.cancel()
            self.stage.setText(tr("Cancelling after the current document…"))

    def _failed(self, message: str, details: str) -> None:
        self.worker = None
        self._set_running(False)
        self.progress.setRange(0, 1)
        error(self, message, details)

    def _done(self, result: JobResult, request: GenerationRequest) -> None:
        self.worker = None
        self.result = result
        self._set_running(False)
        self.progress.setRange(0, max(1, len(result.units)))
        self.progress.setValue(len(result.units))
        self.units.setRowCount(len(result.units))
        for r, u in enumerate(result.units):
            for c, v in enumerate(
                [
                    f"{u.index}. {u.label}",
                    tr(u.status),
                    ", ".join(map(str, u.rows[:20])),
                    " | ".join(Path(o).name for o in u.outputs),
                ]
            ):
                item = QTableWidgetItem(v)
                item.setData(Qt.ItemDataRole.UserRole, u.outputs[0] if u.outputs else None)
                self.units.setItem(r, c, item)
        self.units.resizeColumnsToContents()
        all_issues = result.all_issues()
        self.issues.set_issues(all_issues)
        self.stage.setText(
            tr(
                "{status}: {docs} document(s), {errors} error(s), {warnings} warning(s)",
                status=tr(result.status),
                docs=len(result.documents),
                errors=result.error_count,
                warnings=result.warning_count,
            )
        )
        self.state.publish_issues(tr("Generation"), all_issues)
        if not request.dry_run:
            try:
                self.state.db.record_job(result, str(request.excel_path))
            except Exception:  # history is optional; never fail the job because of it
                pass

    def _open_folder(self) -> None:
        out = self.output.path()
        if out and out.exists():
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(out)))

    def _open_unit(self, row: int, _col: int) -> None:
        item = self.units.item(row, 0)
        path: Any = item.data(Qt.ItemDataRole.UserRole) if item else None
        if path:
            QDesktopServices.openUrl(QUrl.fromLocalFile(path))
