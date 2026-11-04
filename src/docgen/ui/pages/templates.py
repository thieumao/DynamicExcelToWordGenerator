"""Word Template Manager: import, order, enable, markers, bindings, merge/section options."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from pydantic import ValidationError
from PySide6.QtCore import Qt
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
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from docgen.application.templates import PreparedTemplate, prepare_template
from docgen.application.workbook import suggest_identifier
from docgen.domain.mapping.context import build_schema
from docgen.domain.profile.models import MarkerConfig, MissingPolicy, SectionConfig, TemplateSpec
from docgen.i18n.catalog import tr
from docgen.ui.state import AppState
from docgen.ui.widgets.common import button, error, page_title
from docgen.ui.workers import run_in_background


class TemplatesPage(QWidget):
    def __init__(self, state: AppState) -> None:
        super().__init__()
        self.state = state
        layout = QVBoxLayout(self)
        layout.addWidget(
            page_title(
                tr("Word templates"),
                tr(
                    "Import the customer's .docx templates (copied into the profile), check detected placeholders, "
                    "bind them to fields and set the merge order. Original files are never modified."
                ),
            )
        )
        splitter = QSplitter(Qt.Orientation.Horizontal)
        left = QWidget()
        lv = QVBoxLayout(left)
        self.list = QListWidget()
        self.list.currentRowChanged.connect(self._show)
        self.list.itemChanged.connect(self._toggle)
        lv.addWidget(self.list, 1)
        row = QHBoxLayout()
        for label, slot in (
            (tr("Import…"), self._import),
            ("↑", lambda: self._move(-1)),
            ("↓", lambda: self._move(1)),
            (tr("Remove"), self._remove),
        ):
            btn = button(label, primary=label == tr("Import…"))
            btn.clicked.connect(slot)
            row.addWidget(btn)
        lv.addLayout(row)
        self.merge_enabled = QCheckBox(tr("Merge templates into one document"))
        self.merge_enabled.toggled.connect(self._merge_toggled)
        lv.addWidget(self.merge_enabled)
        splitter.addWidget(left)

        right = QWidget()
        rv = QVBoxLayout(right)
        box = QGroupBox(tr("Template settings"))
        form = QFormLayout(box)
        self.t_id = QLabel()
        self.t_file = QLabel()
        form.addRow(tr("ID"), self.t_id)
        form.addRow(tr("File"), self.t_file)
        self.marker = QComboBox()
        for style, label in (
            ("jinja", "Jinja {{ field }}"),
            ("custom", tr("Custom markers")),
            ("mergefield", tr("Word MERGEFIELD")),
        ):
            self.marker.addItem(label, style)
        self.m_open = QLineEdit("<<")
        self.m_close = QLineEdit(">>")
        mrow = QHBoxLayout()
        mrow.addWidget(self.marker)
        mrow.addWidget(QLabel(tr("open")))
        mrow.addWidget(self.m_open)
        mrow.addWidget(QLabel(tr("close")))
        mrow.addWidget(self.m_close)
        form.addRow(tr("Placeholder style"), mrow)
        self.include_if = QLineEdit()
        self.include_if.setPlaceholderText("employee.project_count > 0")
        form.addRow(tr("Include only if"), self.include_if)
        self.context_root = QLineEdit()
        self.context_root.setPlaceholderText(tr("(whole context)"))
        form.addRow(tr("Context root"), self.context_root)
        self.missing = QComboBox()
        self.missing.addItems([m.value for m in MissingPolicy])
        form.addRow(tr("Missing placeholder"), self.missing)
        self.section_mode = QComboBox()
        self.section_mode.addItem(tr("preserve (own page layout, headers/footers)"), "preserve")
        self.section_mode.addItem(tr("inherit (use first document's layout)"), "inherit")
        form.addRow(tr("When merged"), self.section_mode)
        self.break_before = QComboBox()
        self.break_before.addItems(["page", "section", "none"])
        form.addRow(tr("Break before"), self.break_before)
        self.restart_numbering = QCheckBox(tr("Restart page numbering"))
        form.addRow("", self.restart_numbering)
        apply = button(tr("Apply and re-scan"), primary=True)
        apply.clicked.connect(self._apply)
        form.addRow("", apply)
        rv.addWidget(box)
        rv.addWidget(QLabel(tr("Detected placeholders")))
        self.placeholders = QTableWidget(0, 4)
        self.placeholders.setHorizontalHeaderLabels(
            [tr("Placeholder"), tr("Kind"), tr("Status"), tr("Bound to field")]
        )
        self.placeholders.horizontalHeader().setStretchLastSection(True)
        rv.addWidget(self.placeholders, 1)
        self.scan_info = QLabel()
        self.scan_info.setWordWrap(True)
        rv.addWidget(self.scan_info)
        splitter.addWidget(right)
        splitter.setSizes([330, 870])
        layout.addWidget(splitter, 1)
        state.profile_changed.connect(self.refresh)
        self.refresh()

    # ----------------------------------------------------------------
    def refresh(self) -> None:
        self.list.blockSignals(True)
        self.list.clear()
        p = self.state.profile
        if p is not None:
            self.merge_enabled.setChecked(p.merge.enabled)
            for t in p.ordered_templates():
                item = QListWidgetItem(f"{t.id}  –  {Path(t.file).name}")
                item.setData(Qt.ItemDataRole.UserRole, t.id)
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                item.setCheckState(Qt.CheckState.Checked if t.enabled else Qt.CheckState.Unchecked)
                self.list.addItem(item)
        self.list.blockSignals(False)
        if self.list.count():
            self.list.setCurrentRow(0)

    def _current(self) -> TemplateSpec | None:
        item = self.list.currentItem()
        p = self.state.profile
        return p.template(item.data(Qt.ItemDataRole.UserRole)) if item and p else None

    def _sync_order(self) -> None:
        p = self.state.profile
        if p is None:
            return
        p.merge.order = [self.list.item(i).data(Qt.ItemDataRole.UserRole) for i in range(self.list.count())]
        self.state.mark_edited()

    def _toggle(self, item: QListWidgetItem) -> None:
        p = self.state.profile
        spec = p.template(item.data(Qt.ItemDataRole.UserRole)) if p else None
        if spec is not None:
            spec.enabled = item.checkState() == Qt.CheckState.Checked
            self.state.mark_edited()

    def _merge_toggled(self, checked: bool) -> None:
        if self.state.profile is not None and self.state.profile.merge.enabled != checked:
            self.state.profile.merge.enabled = checked
            self.state.mark_edited()

    def _move(self, delta: int) -> None:
        row = self.list.currentRow()
        target = row + delta
        if row < 0 or target < 0 or target >= self.list.count():
            return
        item = self.list.takeItem(row)
        self.list.insertItem(target, item)
        self.list.setCurrentRow(target)
        self._sync_order()

    def _import(self) -> None:
        p = self.state.profile
        if p is None:
            error(self, tr("Open a profile first."))
            return
        files, _ = QFileDialog.getOpenFileNames(self, tr("Import Word templates"), "", "Word (*.docx)")
        for name in files:
            try:
                rel = self.state.repo.import_template(p.profile.id, Path(name))
            except (OSError, ValueError) as exc:
                error(self, str(exc))
                continue
            taken = {t.id for t in p.templates}
            tid = suggest_identifier(Path(name).stem, taken)
            p.templates.append(TemplateSpec(id=tid, file=rel))
        if files:
            self.state.mark_edited()
            self._sync_order()
            self.refresh()

    def _remove(self) -> None:
        spec = self._current()
        p = self.state.profile
        if spec is None or p is None:
            return
        p.templates = [t for t in p.templates if t.id != spec.id]
        p.merge.order = [i for i in p.merge.order if i != spec.id]
        self.state.mark_edited()
        self.refresh()

    def _show(self, _row: int) -> None:
        spec = self._current()
        if spec is None:
            return
        self.t_id.setText(spec.id)
        self.t_file.setText(spec.file)
        self.marker.setCurrentIndex(self.marker.findData(spec.marker.style.value))
        self.m_open.setText(spec.marker.open)
        self.m_close.setText(spec.marker.close)
        self.include_if.setText(spec.include_if or "")
        self.context_root.setText(spec.context_root or "")
        self.missing.setCurrentText(spec.missing_placeholder.value)
        self.section_mode.setCurrentIndex(self.section_mode.findData(spec.section.mode))
        self.break_before.setCurrentText(spec.section.break_before)
        self.restart_numbering.setChecked(spec.section.restart_page_numbering)
        self._scan(spec)

    def _scan(self, spec: TemplateSpec) -> None:
        p, folder = self.state.profile, self.state.profile_dir
        if p is None or folder is None:
            return
        self.scan_info.setText(tr("Scanning…"))
        snapshot = p.model_copy(deep=True)
        run_in_background(
            prepare_template,
            snapshot,
            folder,
            spec.model_copy(deep=True),
            on_done=lambda r: self._show_scan(spec, r),
            on_error=lambda m, d: error(self, m, d),
        )

    def _show_scan(self, spec: TemplateSpec, result: tuple[PreparedTemplate | None, list[Any]]) -> None:
        prepared, issues = result
        self.placeholders.setRowCount(0)
        p = self.state.profile
        if p is None:
            return
        fields = sorted(build_schema(p).paths)
        if prepared is None:
            self.scan_info.setText("; ".join(i.message for i in issues))
            return
        rows: list[tuple[str, str, str, str | None]] = []
        for name in prepared.inspection.binding_names:
            bound = prepared.bindings.get(name)
            rows.append((name, tr("marker"), tr("bound") if bound else tr("unbound"), bound or ""))
        seen = set()
        for ref in prepared.inspection.refs:
            if ref.binding_name is not None or (ref.path, ref.kind) in seen:
                continue
            seen.add((ref.path, ref.kind))
            status = tr("unknown field") if ref.path in prepared.unknown_paths else tr("ok")
            rows.append((ref.path, tr(ref.kind.value), status, None))
        for name, kind, status, bound in rows:
            r = self.placeholders.rowCount()
            self.placeholders.insertRow(r)
            self.placeholders.setItem(r, 0, QTableWidgetItem(name))
            self.placeholders.setItem(r, 1, QTableWidgetItem(kind))
            st = QTableWidgetItem(status)
            if status in (tr("unbound"), tr("unknown field")):
                st.setForeground(Qt.GlobalColor.red)
            self.placeholders.setItem(r, 2, st)
            if bound is not None:
                combo = QComboBox()
                combo.addItem("", "")
                for f in fields:
                    combo.addItem(f, f)
                combo.setCurrentIndex(max(0, combo.findData(bound)))
                combo.currentIndexChanged.connect(lambda _i, n=name, c=combo: self._bind(n, c.currentData()))
                self.placeholders.setCellWidget(r, 3, combo)
        self.placeholders.resizeColumnsToContents()
        problems = [i for i in issues if i.severity.value != "info"]
        self.scan_info.setText(
            tr("{n} placeholder(s); {p} problem(s)", n=len(rows), p=len(problems))
            + ("\n" + "\n".join(f"• {i.message}" for i in problems[:8]) if problems else "")
        )
        self.state.publish_issues(tr("Template {tid}", tid=spec.id), issues)

    def _bind(self, name: str, path: str) -> None:
        spec = self._current()
        if spec is None:
            return
        if path:
            spec.bindings[name] = path
        else:
            spec.bindings.pop(name, None)
        self.state.mark_edited()

    def _apply(self) -> None:
        spec = self._current()
        p = self.state.profile
        if spec is None or p is None:
            return
        try:
            new = TemplateSpec(
                id=spec.id,
                file=spec.file,
                enabled=spec.enabled,
                description=spec.description,
                marker=MarkerConfig(
                    style=self.marker.currentData(),
                    open=self.m_open.text() or "<<",
                    close=self.m_close.text() or ">>",
                ),
                bindings=spec.bindings,
                context_root=self.context_root.text().strip() or None,
                include_if=self.include_if.text().strip() or None,
                missing_placeholder=MissingPolicy(self.missing.currentText()),
                section=SectionConfig(
                    mode=self.section_mode.currentData(),
                    break_before=self.break_before.currentText(),  # type: ignore[arg-type]
                    restart_page_numbering=self.restart_numbering.isChecked(),
                ),
            )
        except ValidationError as exc:
            error(self, tr("Invalid template settings"), str(exc))
            return
        p.templates = [new if t.id == spec.id else t for t in p.templates]
        self.state.mark_edited()
        self._scan(new)
