"""Relationships between sheets: keys (composite), cardinality, missing/duplicate policies, test."""

from __future__ import annotations

from pydantic import ValidationError
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QComboBox,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QTableWidget,
    QVBoxLayout,
    QWidget,
)

from docgen.domain.profile.models import (
    Cardinality,
    JoinPair,
    MatchMode,
    OnDuplicate,
    OnMissing,
    Relationship,
)
from docgen.i18n.catalog import tr
from docgen.ui.state import AppState
from docgen.ui.widgets.common import button, error, page_title
from docgen.ui.workers import run_in_background


def _combo(values: list[str], current: str | None = None) -> QComboBox:
    combo = QComboBox()
    combo.addItems(values)
    if current is not None and current in values:
        combo.setCurrentText(current)
    return combo


class RelationshipsPage(QWidget):
    def __init__(self, state: AppState) -> None:
        super().__init__()
        self.state = state
        layout = QVBoxLayout(self)
        layout.addWidget(
            page_title(
                tr("Relationships"),
                tr(
                    "Link the IDs in one sheet to rows in another sheet, e.g. Employees.DepartmentID → "
                    "Departments.DepartmentID. Several key pairs form a composite key."
                ),
            )
        )
        body = QHBoxLayout()
        left = QVBoxLayout()
        self.list = QListWidget()
        self.list.currentRowChanged.connect(self._show)
        left.addWidget(self.list, 1)
        for label, slot in (
            (tr("New relationship"), self._new),
            (tr("Delete"), self._delete),
            (tr("Test all relationships"), self._test),
        ):
            btn = button(label)
            btn.clicked.connect(slot)
            left.addWidget(btn)
        body.addLayout(left, 1)

        box = QGroupBox(tr("Relationship"))
        form = QFormLayout(box)
        self.rel_id = QLineEdit()
        form.addRow(tr("ID"), self.rel_id)
        self.from_sheet = QComboBox()
        self.to_sheet = QComboBox()
        self.from_sheet.currentTextChanged.connect(self._sheets_changed)
        self.to_sheet.currentTextChanged.connect(self._sheets_changed)
        form.addRow(tr("From sheet"), self.from_sheet)
        form.addRow(tr("To sheet"), self.to_sheet)
        self.keys = QTableWidget(0, 2)
        self.keys.setHorizontalHeaderLabels([tr("From column"), tr("To column")])
        self.keys.horizontalHeader().setStretchLastSection(True)
        self.keys.setMinimumHeight(110)
        form.addRow(tr("Key columns"), self.keys)
        kb = QHBoxLayout()
        add_key, del_key = button(tr("Add key pair")), button(tr("Remove key pair"))
        add_key.clicked.connect(lambda: self._add_key_row())
        del_key.clicked.connect(lambda: self.keys.removeRow(self.keys.currentRow()))
        kb.addWidget(add_key)
        kb.addWidget(del_key)
        kb.addStretch(1)
        form.addRow("", kb)
        self.cardinality = _combo([c.value for c in Cardinality])
        form.addRow(tr("Cardinality"), self.cardinality)
        self.match = _combo(["(default)", *[m.value for m in MatchMode]])
        form.addRow(tr("When several rows match"), self.match)
        self.on_missing = _combo([m.value for m in OnMissing])
        form.addRow(tr("When no row matches"), self.on_missing)
        self.on_duplicate = _combo([d.value for d in OnDuplicate])
        form.addRow(tr("Duplicate keys"), self.on_duplicate)
        self.default = QLineEdit()
        form.addRow(tr("Default value (on_missing: default)"), self.default)
        self.filter = QLineEdit()
        self.filter.setPlaceholderText("status == 'Active'")
        form.addRow(tr("Filter target rows"), self.filter)
        self.sort = QLineEdit()
        self.sort.setPlaceholderText("column, -other_column")
        form.addRow(tr("Sort target rows"), self.sort)
        apply = button(tr("Apply relationship"), primary=True)
        apply.clicked.connect(self._apply)
        form.addRow("", apply)
        self.result = QLabel()
        self.result.setWordWrap(True)
        form.addRow(tr("Test result"), self.result)
        body.addWidget(box, 3)
        layout.addLayout(body, 1)
        state.profile_changed.connect(self.refresh)
        self.refresh()

    # ------------------------------------------------------------------
    def _sheet_ids(self) -> list[str]:
        return [s.id for s in self.state.profile.excel.sheets] if self.state.profile else []

    def _columns(self, sheet_id: str) -> list[str]:
        if not self.state.profile:
            return []
        cfg = self.state.profile.excel.sheet(sheet_id)
        return [c.id for c in cfg.columns] if cfg else []

    def refresh(self) -> None:
        self.list.clear()
        sheets = self._sheet_ids()
        for combo in (self.from_sheet, self.to_sheet):
            combo.blockSignals(True)
            combo.clear()
            combo.addItems(sheets)
            combo.blockSignals(False)
        if self.state.profile is None:
            return
        for rel in self.state.profile.relationships:
            pairs = ", ".join(f"{p.from_}→{p.to}" for p in rel.keys)
            item = QListWidgetItem(f"{rel.id}: {rel.from_} → {rel.to} ({pairs}) [{rel.cardinality.value}]")
            item.setData(Qt.ItemDataRole.UserRole, rel.id)
            self.list.addItem(item)
        if self.list.count():
            self.list.setCurrentRow(0)

    def _add_key_row(self, src: str | None = None, dst: str | None = None) -> None:
        r = self.keys.rowCount()
        self.keys.insertRow(r)
        self.keys.setCellWidget(r, 0, _combo(self._columns(self.from_sheet.currentText()), src))
        self.keys.setCellWidget(r, 1, _combo(self._columns(self.to_sheet.currentText()), dst))

    def _sheets_changed(self) -> None:
        pairs = self._pairs()
        self.keys.setRowCount(0)
        for src, dst in pairs or [(None, None)]:
            self._add_key_row(src, dst)

    def _pairs(self) -> list[tuple[str, str]]:
        pairs = []
        for r in range(self.keys.rowCount()):
            a, b = self.keys.cellWidget(r, 0), self.keys.cellWidget(r, 1)
            if isinstance(a, QComboBox) and isinstance(b, QComboBox) and a.currentText() and b.currentText():
                pairs.append((a.currentText(), b.currentText()))
        return pairs

    def _show(self, row: int) -> None:
        if self.state.profile is None or row < 0:
            return
        rel = self.state.profile.relationships[row]
        self.rel_id.setText(rel.id)
        for combo, value in ((self.from_sheet, rel.from_), (self.to_sheet, rel.to)):
            combo.blockSignals(True)
            combo.setCurrentText(value)
            combo.blockSignals(False)
        self.keys.setRowCount(0)
        for pair in rel.keys:
            self._add_key_row(pair.from_, pair.to)
        self.cardinality.setCurrentText(rel.cardinality.value)
        self.match.setCurrentText(rel.effective_match.value)
        self.on_missing.setCurrentText(rel.on_missing.value)
        self.on_duplicate.setCurrentText(rel.on_duplicate.value)
        self.default.setText("" if rel.default is None else str(rel.default))
        self.filter.setText(rel.filter or "")
        self.sort.setText(", ".join(("-" if s.order == "desc" else "") + s.by for s in rel.sort))

    def _new(self) -> None:
        self.list.clearSelection()
        self.rel_id.setText(
            f"rel_{len(self.state.profile.relationships) + 1}" if self.state.profile else "rel_1"
        )
        self.keys.setRowCount(0)
        self._add_key_row()
        self.default.clear()
        self.filter.clear()
        self.sort.clear()

    def _apply(self) -> None:
        profile = self.state.profile
        if profile is None:
            return
        from docgen.domain.profile.models import SortKey

        sort = []
        for token in [t.strip() for t in self.sort.text().split(",") if t.strip()]:
            sort.append(SortKey(by=token.lstrip("-"), order="desc" if token.startswith("-") else "asc"))
        data = {
            "id": self.rel_id.text().strip(),
            "from": self.from_sheet.currentText(),
            "to": self.to_sheet.currentText(),
            "keys": [JoinPair.model_validate({"from": a, "to": b}) for a, b in self._pairs()],
            "cardinality": self.cardinality.currentText(),
            "on_missing": self.on_missing.currentText(),
            "on_duplicate": self.on_duplicate.currentText(),
            "default": self.default.text() or None,
            "filter": self.filter.text().strip() or None,
            "sort": sort,
        }
        if self.match.currentText() != "(default)":
            data["match"] = self.match.currentText()
        try:
            rel = Relationship.model_validate(data)
        except ValidationError as exc:
            error(self, tr("Invalid relationship"), str(exc))
            return
        others = [r for r in profile.relationships if r.id != rel.id]
        existing = [r.id for r in profile.relationships]
        if rel.id in existing:
            profile.relationships = [rel if r.id == rel.id else r for r in profile.relationships]
        else:
            profile.relationships = [*others, rel]
        self.state.mark_edited()
        self.refresh()

    def _delete(self) -> None:
        item = self.list.currentItem()
        if item is None or self.state.profile is None:
            return
        rid = item.data(Qt.ItemDataRole.UserRole)
        self.state.profile.relationships = [r for r in self.state.profile.relationships if r.id != rid]
        self.state.mark_edited()
        self.refresh()

    def _test(self) -> None:
        profile, excel = self.state.profile, self.state.excel_path
        if profile is None or excel is None or self.state.profile_dir is None:
            error(self, tr("Open a profile and load a workbook first."))
            return
        from docgen.application.preview import preview

        self.result.setText(tr("Testing…"))

        def done(result) -> None:  # type: ignore[no-untyped-def]
            if not result.ok:
                self.result.setText(tr("Cannot test: see the Logs page."))
                self.state.publish_issues(tr("Relationship test"), result.issues)
                return
            lines = [
                tr(
                    "{rel} → {sheet}: {keys} key(s), {dups} duplicated {sample}",
                    rel=s.relationship,
                    sheet=s.target_sheet,
                    keys=s.keys,
                    dups=s.duplicate_keys,
                    sample=s.sample_duplicates or "",
                )
                for s in result.relationships
            ]
            self.result.setText("\n".join(lines) or tr("No relationships."))
            self.state.publish_issues(tr("Relationship test"), result.issues)

        run_in_background(
            preview,
            profile.model_copy(deep=True),
            self.state.profile_dir,
            excel,
            limit=0,
            with_plan=False,
            on_done=done,
            on_error=lambda m, d: error(self, m, d),
        )
