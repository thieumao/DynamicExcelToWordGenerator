"""Mapping Designer: visual editor for output fields, lookup chains, transforms and conditions."""

from __future__ import annotations

import json
from typing import Any

from pydantic import ValidationError
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QPlainTextEdit,
    QSpinBox,
    QSplitter,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from docgen.domain.mapping.transforms import DEFAULT_REGISTRY
from docgen.domain.profile.models import (
    AggregationSpec,
    CollectionSpec,
    FieldRule,
    FieldType,
    Profile,
    TransformCall,
)
from docgen.i18n.catalog import issue_text, tr
from docgen.ui.state import AppState
from docgen.ui.widgets.common import button, error, page_title
from docgen.ui.workers import run_in_background

KINDS = [
    ("source", "Column of this sheet"),
    ("path", "Lookup in other sheets"),
    ("expr", "Expression"),
    ("when", "Conditional"),
    ("coalesce", "First non-empty of"),
    ("const", "Constant value"),
]
AGG_OPS = ["count", "count_distinct", "sum", "min", "max", "avg", "join", "first", "last", "list"]


def _header(sheet: Any, column_id: str) -> str:
    """Show the Excel header the user knows ('DepartmentID') rather than the internal id."""
    spec = sheet.column(column_id) if sheet is not None else None
    return spec.header if spec is not None and spec.header else column_id


class TransformList(QWidget):
    """Ordered transform pipeline with parameters (JSON object per step)."""

    def __init__(self) -> None:
        super().__init__()
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self.table = QTableWidget(0, 2)
        self.table.setHorizontalHeaderLabels([tr("Transform"), tr("Parameters (JSON)")])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setMinimumHeight(110)
        layout.addWidget(self.table)
        bar = QHBoxLayout()
        self.choice = QComboBox()
        for d in DEFAULT_REGISTRY.definitions():
            params = ", ".join(f"{k}{'*' if s.required else ''}" for k, s in d.params.items())
            self.choice.addItem(
                f"{d.name}  –  {tr(d.description)}" + (f"  ({params})" if params else ""), d.name
            )
        bar.addWidget(self.choice, 1)
        for label, slot in (
            ("+", self._add),
            ("↑", lambda: self._move(-1)),
            ("↓", lambda: self._move(1)),
            ("✕", lambda: self.table.removeRow(self.table.currentRow())),
        ):
            btn = button(label)
            btn.setMaximumWidth(36)
            btn.clicked.connect(slot)
            bar.addWidget(btn)
        layout.addLayout(bar)

    def _add(self) -> None:
        name = self.choice.currentData()
        definition = DEFAULT_REGISTRY.get(name)
        params = {k: s.default for k, s in definition.params.items() if s.required} if definition else {}
        self._append(name, params)

    def _append(self, name: str, params: dict[str, Any]) -> None:
        r = self.table.rowCount()
        self.table.insertRow(r)
        self.table.setItem(r, 0, QTableWidgetItem(name))
        self.table.setItem(r, 1, QTableWidgetItem(json.dumps(params, ensure_ascii=False) if params else ""))

    def _move(self, delta: int) -> None:
        r = self.table.currentRow()
        t = r + delta
        if r < 0 or t < 0 or t >= self.table.rowCount():
            return
        a = [self.table.item(r, c).text() for c in range(2)]
        b = [self.table.item(t, c).text() for c in range(2)]
        for c in range(2):
            self.table.item(r, c).setText(b[c])
            self.table.item(t, c).setText(a[c])
        self.table.setCurrentCell(t, 0)

    def set_calls(self, calls: list[TransformCall]) -> None:
        self.table.setRowCount(0)
        for call in calls:
            self._append(call.name, dict(call.params))

    def calls(self) -> list[TransformCall]:
        result = []
        for r in range(self.table.rowCount()):
            name = self.table.item(r, 0).text().strip()
            raw = self.table.item(r, 1).text().strip() if self.table.item(r, 1) else ""
            params = json.loads(raw) if raw else {}
            if not isinstance(params, dict):
                raise ValueError(f"parameters of '{name}' must be a JSON object")
            result.append(TransformCall(name=name, params=params))
        return result


class ChainBuilder(QWidget):
    """relationship → relationship → column, filtered so each step starts at the current sheet."""

    def __init__(self) -> None:
        super().__init__()
        self.profile: Profile | None = None
        self.base = ""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self.steps: list[QComboBox] = []
        self.steps_box = QVBoxLayout()
        layout.addLayout(self.steps_box)
        row = QHBoxLayout()
        add, remove = button(tr("+ step")), button(tr("− step"))
        add.clicked.connect(self._add_step)
        remove.clicked.connect(self._remove_step)
        self.column = QComboBox()
        row.addWidget(add)
        row.addWidget(remove)
        row.addWidget(QLabel(tr("return column")))
        row.addWidget(self.column, 1)
        layout.addLayout(row)
        self.breadcrumb = QLabel()
        self.breadcrumb.setObjectName("Breadcrumb")
        self.breadcrumb.setWordWrap(True)
        layout.addWidget(self.breadcrumb)
        self.column.currentTextChanged.connect(self._update_crumb)

    def configure(self, profile: Profile, base_sheet: str) -> None:
        self.profile = profile
        self.base = base_sheet
        for combo in self.steps:
            combo.deleteLater()
        self.steps = []

    def _sheet_after(self, n: int) -> str:
        sheet = self.base
        assert self.profile is not None
        for combo in self.steps[:n]:
            rel = self.profile.relationship(combo.currentData() or "")
            if rel is None:
                break
            sheet = rel.to
        return sheet

    def _add_step(self, rel_id: str | None = None) -> None:
        if self.profile is None:
            return
        sheet = self._sheet_after(len(self.steps))
        combo = QComboBox()
        for rel in self.profile.relationships:
            if rel.from_ == sheet:
                label = f"{rel.id}: {rel.from_} → {rel.to}" + (" (1-n)" if rel.cardinality.to_many else "")
                combo.addItem(label, rel.id)
        if rel_id:
            idx = combo.findData(rel_id)
            if idx >= 0:
                combo.setCurrentIndex(idx)
        combo.currentIndexChanged.connect(self._refresh_columns)
        self.steps.append(combo)
        self.steps_box.addWidget(combo)
        self._refresh_columns()

    def _remove_step(self) -> None:
        if self.steps:
            self.steps.pop().deleteLater()
            self._refresh_columns()

    def _refresh_columns(self) -> None:
        if self.profile is None:
            return
        current = self.column.currentText()
        sheet = self.profile.excel.sheet(self._sheet_after(len(self.steps)))
        self.column.blockSignals(True)
        self.column.clear()
        if sheet:
            self.column.addItems([c.id for c in sheet.columns])
        if current:
            self.column.setCurrentText(current)
        self.column.blockSignals(False)
        self._update_crumb()

    def _update_crumb(self) -> None:
        if self.profile is None:
            return
        parts = []
        sheet_id = self.base
        for combo in self.steps:
            rel = self.profile.relationship(combo.currentData() or "")
            if rel is None:
                continue
            src = self.profile.excel.sheet(rel.from_)
            dst = self.profile.excel.sheet(rel.to)
            src_name = src.match.name if src and src.match.name else rel.from_
            dst_name = dst.match.name if dst and dst.match.name else rel.to
            keys = ", ".join(_header(src, p.from_) for p in rel.keys)
            tkeys = ", ".join(_header(dst, p.to) for p in rel.keys)
            parts.append(f"{src_name}.{keys} → {dst_name}.{tkeys}")
            sheet_id = rel.to
        final = self.profile.excel.sheet(sheet_id)
        final_name = final.match.name if final and final.match.name else sheet_id
        parts.append(f"{final_name}.{_header(final, self.column.currentText())}")
        self.breadcrumb.setText(" → ".join(parts))

    def set_path(self, path: str) -> None:
        if not path:
            self._refresh_columns()
            return
        *rels, column = path.split(".")
        for rel in rels:
            self._add_step(rel)
        self._refresh_columns()
        self.column.setCurrentText(column)

    def path(self) -> str:
        rels = [c.currentData() for c in self.steps if c.currentData()]
        return ".".join([*rels, self.column.currentText()])

    def to_many(self) -> bool:
        assert self.profile is not None
        return any(
            (rel := self.profile.relationship(c.currentData() or "")) is not None and rel.cardinality.to_many
            for c in self.steps
        )


class FieldEditor(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.profile: Profile | None = None
        self.base = ""
        form = QFormLayout(self)
        self.name = QLineEdit()
        self.name.setPlaceholderText("employee.department_name")
        form.addRow(tr("Output field"), self.name)
        self.kind = QComboBox()
        for key, label in KINDS:
            self.kind.addItem(tr(label), key)
        form.addRow(tr("Value from"), self.kind)
        self.stack = QStackedWidget()
        self.source = QComboBox()
        self.chain = ChainBuilder()
        chain_box = QWidget()
        cb = QVBoxLayout(chain_box)
        cb.setContentsMargins(0, 0, 0, 0)
        cb.addWidget(self.chain)
        agg_row = QHBoxLayout()
        self.aggregate = QComboBox()
        self.aggregate.addItem(tr("(single value)"), None)
        for op in AGG_OPS:
            self.aggregate.addItem(op, op)
        self.separator = QLineEdit(", ")
        self.separator.setMaximumWidth(60)
        agg_row.addWidget(QLabel(tr("If several rows")))
        agg_row.addWidget(self.aggregate)
        agg_row.addWidget(QLabel(tr("separator")))
        agg_row.addWidget(self.separator)
        agg_row.addStretch(1)
        cb.addLayout(agg_row)
        self.expr = QPlainTextEdit()
        self.expr.setPlaceholderText("emp.grade >= 5 and employee.city == 'Hà Nội'")
        self.expr.setMaximumHeight(80)
        self.when = QTableWidget(0, 2)
        self.when.setHorizontalHeaderLabels([tr("If (expression)"), tr("Value")])
        self.when.horizontalHeader().setStretchLastSection(True)
        when_box = QWidget()
        wb = QVBoxLayout(when_box)
        wb.setContentsMargins(0, 0, 0, 0)
        wb.addWidget(self.when)
        wrow = QHBoxLayout()
        add_c, del_c = button(tr("Add condition")), button(tr("Remove condition"))
        add_c.clicked.connect(lambda: self._when_row("", ""))
        del_c.clicked.connect(lambda: self.when.removeRow(self.when.currentRow()))
        self.else_value = QLineEdit()
        wrow.addWidget(add_c)
        wrow.addWidget(del_c)
        wrow.addWidget(QLabel(tr("Else")))
        wrow.addWidget(self.else_value, 1)
        wb.addLayout(wrow)
        self.coalesce = QLineEdit()
        self.coalesce.setPlaceholderText("rel_a.name, rel_b.name, local_column")
        self.const = QLineEdit()
        for widget in (self.source, chain_box, self.expr, when_box, self.coalesce, self.const):
            self.stack.addWidget(widget)
        self.kind.currentIndexChanged.connect(self.stack.setCurrentIndex)
        form.addRow("", self.stack)
        self.type = QComboBox()
        self.type.addItems([t.value for t in FieldType])
        form.addRow(tr("Type"), self.type)
        self.transforms = TransformList()
        form.addRow(tr("Transforms"), self.transforms)
        self.default = QLineEdit()
        form.addRow(tr("Default if empty"), self.default)
        self.required = QCheckBox(tr("Required (empty value is an error)"))
        form.addRow("", self.required)
        hint = QLabel(
            tr(
                "Expressions use Jinja syntax. Sheet columns: <sheet id>.<column id>; computed fields: "
                "their full name; functions: coalesce(), concat_ws(), today(), days_between()."
            )
        )
        hint.setWordWrap(True)
        hint.setObjectName("Hint")
        form.addRow("", hint)

    def _when_row(self, cond: str, value: str) -> None:
        r = self.when.rowCount()
        self.when.insertRow(r)
        self.when.setItem(r, 0, QTableWidgetItem(cond))
        self.when.setItem(r, 1, QTableWidgetItem(value))

    def load(self, profile: Profile, base_sheet: str, rule: FieldRule | None) -> None:
        self.profile, self.base = profile, base_sheet
        sheet = profile.excel.sheet(base_sheet)
        self.source.clear()
        if sheet:
            self.source.addItems([c.id for c in sheet.columns])
        self.chain.configure(profile, base_sheet)
        rule = rule or FieldRule(name="new.field", source=sheet.columns[0].id if sheet else "x")
        self.name.setText(rule.name)
        self.kind.setCurrentIndex([k for k, _ in KINDS].index(rule.kind))
        self.stack.setCurrentIndex(self.kind.currentIndex())
        if rule.source:
            self.source.setCurrentText(rule.source.split(".")[-1])
        self.chain.set_path(rule.path or "")
        idx = self.aggregate.findData(rule.aggregate)
        self.aggregate.setCurrentIndex(max(0, idx))
        self.separator.setText(rule.separator)
        self.expr.setPlainText(rule.expr or "")
        self.when.setRowCount(0)
        self.else_value.clear()
        for clause in rule.when or ():
            if clause.is_else:
                self.else_value.setText("" if clause.else_ is None else str(clause.else_))
            else:
                self._when_row(
                    clause.if_ or "", (clause.expr and f"={clause.expr}") or str(clause.value or "")
                )
        self.coalesce.setText(", ".join(rule.coalesce or ()))
        self.const.setText("" if rule.const is None else str(rule.const))
        self.type.setCurrentText(rule.type.value)
        self.transforms.set_calls(list(rule.transforms))
        self.default.setText("" if rule.default is None else str(rule.default))
        self.required.setChecked(rule.required)

    def rule(self) -> FieldRule:
        kind = self.kind.currentData()
        data: dict[str, Any] = {
            "name": self.name.text().strip(),
            "type": self.type.currentText(),
            "transforms": self.transforms.calls(),
            "default": self.default.text() or None,
            "required": self.required.isChecked(),
        }
        if kind == "source":
            data["source"] = self.source.currentText()
        elif kind == "path":
            data["path"] = self.chain.path()
            if self.aggregate.currentData():
                data["aggregate"] = self.aggregate.currentData()
                data["separator"] = self.separator.text()
            elif self.chain.to_many():
                data["aggregate"] = "join"
        elif kind == "expr":
            data["expr"] = self.expr.toPlainText().strip()
        elif kind == "when":
            clauses: list[dict[str, Any]] = []
            for r in range(self.when.rowCount()):
                cond = self.when.item(r, 0).text().strip() if self.when.item(r, 0) else ""
                value = self.when.item(r, 1).text() if self.when.item(r, 1) else ""
                if not cond:
                    continue
                clauses.append(
                    {"if": cond, "expr": value[1:]} if value.startswith("=") else {"if": cond, "value": value}
                )
            clauses.append({"else": self.else_value.text() or None})
            data["when"] = clauses
        elif kind == "coalesce":
            data["coalesce"] = [p.strip() for p in self.coalesce.text().split(",") if p.strip()]
        else:
            data["const"] = self.const.text()
        return FieldRule.model_validate(data)


class MappingPage(QWidget):
    def __init__(self, state: AppState) -> None:
        super().__init__()
        self.state = state
        layout = QVBoxLayout(self)
        layout.addWidget(
            page_title(
                tr("Mapping designer"),
                tr(
                    "Define the output fields used by Word templates. Follow relationships to read values from "
                    "other sheets, add transforms and conditions, then test with a real record."
                ),
            )
        )
        splitter = QSplitter(Qt.Orientation.Horizontal)
        left = QWidget()
        lv = QVBoxLayout(left)
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels([tr("Output"), tr("Source")])
        self.tree.currentItemChanged.connect(self._select)
        lv.addWidget(self.tree, 1)
        grid = QHBoxLayout()
        for label, slot in (
            (tr("+ Field"), self._add_field),
            (tr("+ Collection"), self._add_collection),
            (tr("+ Aggregation"), self._add_aggregation),
            (tr("Delete"), self._delete),
        ):
            btn = button(label)
            btn.clicked.connect(slot)
            grid.addWidget(btn)
        lv.addLayout(grid)
        splitter.addWidget(left)

        right = QWidget()
        rv = QVBoxLayout(right)
        self.editor_stack = QStackedWidget()
        self.empty = QLabel(tr("Select or add an output field."))
        self.field_editor = FieldEditor()
        self.collection_editor = self._build_collection_editor()
        self.aggregation_editor = self._build_aggregation_editor()
        for w in (self.empty, self.field_editor, self.collection_editor, self.aggregation_editor):
            self.editor_stack.addWidget(w)
        rv.addWidget(self.editor_stack, 1)
        apply = button(tr("Apply"), primary=True)
        apply.clicked.connect(self._apply)
        rv.addWidget(apply)
        test_box = QGroupBox(tr("Test with a sample record"))
        tl = QVBoxLayout(test_box)
        tr_row = QHBoxLayout()
        self.test_row = QSpinBox()
        self.test_row.setRange(1, 1_048_576)
        self.test_row.setPrefix(tr("Excel row "))
        test = button(tr("Test mapping"))
        test.clicked.connect(self._test)
        tr_row.addWidget(self.test_row)
        tr_row.addWidget(test)
        tr_row.addStretch(1)
        tl.addLayout(tr_row)
        self.test_output = QListWidget()
        tl.addWidget(self.test_output)
        rv.addWidget(test_box, 1)
        splitter.addWidget(right)
        splitter.setSizes([380, 820])
        layout.addWidget(splitter, 1)
        self._selection: tuple[Any, ...] | None = None
        state.profile_changed.connect(self.refresh)
        state.workbook_changed.connect(lambda: self.test_row.setValue(self._first_data_row()))
        self.refresh()

    # ------------------------------------------------------------ editors
    def _build_collection_editor(self) -> QWidget:
        w = QWidget()
        form = QFormLayout(w)
        self.c_name = QLineEdit()
        self.c_rel = QLineEdit()
        self.c_rel.setPlaceholderText("emp_projects  (or rel_a.rel_b)")
        self.c_filter = QLineEdit()
        self.c_filter.setPlaceholderText("item.hours > 0")
        self.c_sort = QLineEdit()
        self.c_sort.setPlaceholderText("name, -hours")
        self.c_limit = QSpinBox()
        self.c_limit.setRange(0, 100000)
        self.c_limit.setSpecialValueText(tr("no limit"))
        form.addRow(tr("List name"), self.c_name)
        form.addRow(tr("Relationship (one-to-many)"), self.c_rel)
        form.addRow(tr("Filter items"), self.c_filter)
        form.addRow(tr("Sort items"), self.c_sort)
        form.addRow(tr("Limit"), self.c_limit)
        note = QLabel(
            tr(
                "Item fields are listed under the collection; select one to edit it, or use "
                "'+ Field' while the collection is selected."
            )
        )
        note.setWordWrap(True)
        form.addRow("", note)
        return w

    def _build_aggregation_editor(self) -> QWidget:
        w = QWidget()
        form = QFormLayout(w)
        self.a_name = QLineEdit()
        self.a_rel = QLineEdit()
        self.a_op = QComboBox()
        self.a_op.addItems(AGG_OPS)
        self.a_of = QLineEdit()
        self.a_filter = QLineEdit()
        self.a_sep = QLineEdit(", ")
        self.a_transforms = TransformList()
        self.a_default = QLineEdit()
        for label, widget in (
            (tr("Output field"), self.a_name),
            (tr("Relationship (one-to-many)"), self.a_rel),
            (tr("Operation"), self.a_op),
            (tr("Column of related rows"), self.a_of),
            (tr("Filter rows"), self.a_filter),
            (tr("Separator (join)"), self.a_sep),
            (tr("Transforms"), self.a_transforms),
            (tr("Default if empty"), self.a_default),
        ):
            form.addRow(label, widget)
        return w

    # --------------------------------------------------------------- tree
    def refresh(self) -> None:
        self.tree.clear()
        p = self.state.profile
        if p is None:
            return
        self.test_row.setValue(self._first_data_row())
        sections = [
            (
                tr("Fields"),
                [
                    (("field", i), f.name, f"{f.kind}: {f.source or f.path or f.expr or ''}")
                    for i, f in enumerate(p.fields)
                ],
            ),
            (
                tr("Collections (lists)"),
                [(("collection", i), c.name, c.relationship) for i, c in enumerate(p.collections)],
            ),
            (
                tr("Aggregations"),
                [
                    (("aggregation", i), a.name, f"{a.op}({a.of or ''}) via {a.relationship}")
                    for i, a in enumerate(p.aggregations)
                ],
            ),
            (tr("Group fields"), [(("group", i), f.name, f.kind) for i, f in enumerate(p.group_fields)]),
        ]
        for title, items in sections:
            top = QTreeWidgetItem([title])
            top.setData(0, Qt.ItemDataRole.UserRole, ("section", title))
            self.tree.addTopLevelItem(top)
            for key, name, source in items:
                item = QTreeWidgetItem([name, source])
                item.setData(0, Qt.ItemDataRole.UserRole, key)
                top.addChild(item)
                if key[0] == "collection":
                    for fi, f in enumerate(p.collections[key[1]].fields):
                        child = QTreeWidgetItem([f.name, f.kind])
                        child.setData(0, Qt.ItemDataRole.UserRole, ("item", key[1], fi))
                        item.addChild(child)
            top.setExpanded(True)
        self.tree.resizeColumnToContents(0)
        self.editor_stack.setCurrentWidget(self.empty)

    def _first_data_row(self) -> int:
        p, wb = self.state.profile, self.state.workbook
        if p is None or wb is None:
            return 2
        primary = p.excel.sheet(p.excel.primary)
        name = primary.match.name if primary else None
        sheet = next((s for s in wb.sheets if s.name == name), None)
        return (sheet.header_row or 1) + 1 if sheet else 2

    def _collection_sheet(self, ci: int) -> str:
        p = self.state.profile
        assert p is not None
        rel = p.relationship(p.collections[ci].relationship.split(".")[-1])
        return rel.to if rel else p.excel.primary

    def _select(self, item: QTreeWidgetItem | None, _prev: Any = None) -> None:
        p = self.state.profile
        if item is None or p is None:
            return
        key = item.data(0, Qt.ItemDataRole.UserRole)
        self._selection = key
        kind = key[0]
        if kind in ("field", "group"):
            rules = p.fields if kind == "field" else p.group_fields
            self.field_editor.load(p, p.excel.primary, rules[key[1]])
            self.editor_stack.setCurrentWidget(self.field_editor)
        elif kind == "item":
            self.field_editor.load(p, self._collection_sheet(key[1]), p.collections[key[1]].fields[key[2]])
            self.editor_stack.setCurrentWidget(self.field_editor)
        elif kind == "collection":
            c = p.collections[key[1]]
            self.c_name.setText(c.name)
            self.c_rel.setText(c.relationship)
            self.c_filter.setText(c.filter or "")
            self.c_sort.setText(", ".join(("-" if s.order == "desc" else "") + s.by for s in c.sort))
            self.c_limit.setValue(c.limit or 0)
            self.editor_stack.setCurrentWidget(self.collection_editor)
        elif kind == "aggregation":
            a = p.aggregations[key[1]]
            self.a_name.setText(a.name)
            self.a_rel.setText(a.relationship)
            self.a_op.setCurrentText(a.op)
            self.a_of.setText(a.of or "")
            self.a_filter.setText(a.filter or "")
            self.a_sep.setText(a.separator)
            self.a_transforms.set_calls(list(a.transforms))
            self.a_default.setText("" if a.default is None else str(a.default))
            self.editor_stack.setCurrentWidget(self.aggregation_editor)
        else:
            self.editor_stack.setCurrentWidget(self.empty)

    def _add_field(self) -> None:
        p = self.state.profile
        if p is None:
            return
        sel = self._selection
        if sel and sel[0] in ("collection", "item"):
            ci = sel[1]
            self._selection = ("item", ci, len(p.collections[ci].fields))
            self.field_editor.load(p, self._collection_sheet(ci), None)
            self.field_editor.name.setText("value")
        else:
            self._selection = ("field", len(p.fields))
            self.field_editor.load(p, p.excel.primary, None)
        self.editor_stack.setCurrentWidget(self.field_editor)

    def _add_collection(self) -> None:
        p = self.state.profile
        if p is None:
            return
        self._selection = ("collection", len(p.collections))
        for w in (self.c_name, self.c_rel, self.c_filter, self.c_sort):
            w.clear()
        self.c_name.setText("record.items")
        self.c_limit.setValue(0)
        self.editor_stack.setCurrentWidget(self.collection_editor)

    def _add_aggregation(self) -> None:
        p = self.state.profile
        if p is None:
            return
        self._selection = ("aggregation", len(p.aggregations))
        self.a_name.setText("record.total")
        self.a_transforms.set_calls([])
        self.editor_stack.setCurrentWidget(self.aggregation_editor)

    def _delete(self) -> None:
        p, sel = self.state.profile, self._selection
        if p is None or not sel:
            return
        kind = sel[0]
        if kind == "field" and sel[1] < len(p.fields):
            p.fields = [f for i, f in enumerate(p.fields) if i != sel[1]]
        elif kind == "group" and sel[1] < len(p.group_fields):
            p.group_fields = [f for i, f in enumerate(p.group_fields) if i != sel[1]]
        elif kind == "collection" and sel[1] < len(p.collections):
            p.collections = [c for i, c in enumerate(p.collections) if i != sel[1]]
        elif kind == "aggregation" and sel[1] < len(p.aggregations):
            p.aggregations = [a for i, a in enumerate(p.aggregations) if i != sel[1]]
        elif kind == "item":
            coll = p.collections[sel[1]]
            coll.fields = [f for i, f in enumerate(coll.fields) if i != sel[2]]
        else:
            return
        self.state.mark_edited()
        self.refresh()

    def _apply(self) -> None:
        p, sel = self.state.profile, self._selection
        if p is None or not sel:
            return
        try:
            kind = sel[0]
            if kind in ("field", "group", "item"):
                rule = self.field_editor.rule()
                if kind == "item":
                    target = p.collections[sel[1]].fields
                else:
                    target = p.fields if kind == "field" else p.group_fields
                index = sel[-1]
                if index < len(target):
                    target[index] = rule
                else:
                    target.append(rule)
            elif kind == "collection":
                from docgen.domain.profile.models import SortKey

                sort = [
                    SortKey(by=t.strip().lstrip("-"), order="desc" if t.strip().startswith("-") else "asc")
                    for t in self.c_sort.text().split(",")
                    if t.strip()
                ]
                old = p.collections[sel[1]] if sel[1] < len(p.collections) else None
                spec = CollectionSpec(
                    name=self.c_name.text().strip(),
                    relationship=self.c_rel.text().strip(),
                    filter=self.c_filter.text().strip() or None,
                    sort=sort,
                    limit=self.c_limit.value() or None,
                    fields=old.fields if old else [],
                    collections=old.collections if old else [],
                    aggregations=old.aggregations if old else [],
                )
                if old:
                    p.collections[sel[1]] = spec
                else:
                    p.collections.append(spec)
            elif kind == "aggregation":
                spec_a = AggregationSpec(
                    name=self.a_name.text().strip(),
                    relationship=self.a_rel.text().strip(),
                    op=self.a_op.currentText(),
                    of=self.a_of.text().strip() or None,  # type: ignore[arg-type]
                    filter=self.a_filter.text().strip() or None,
                    separator=self.a_sep.text(),
                    transforms=self.a_transforms.calls(),
                    default=self.a_default.text() or None,
                )
                if sel[1] < len(p.aggregations):
                    p.aggregations[sel[1]] = spec_a
                else:
                    p.aggregations.append(spec_a)
        except (ValidationError, ValueError, json.JSONDecodeError) as exc:
            error(self, tr("Invalid mapping"), str(exc))
            return
        self.state.mark_edited()
        keep = self._selection
        self.refresh()
        self._selection = keep
        from docgen.domain.profile.validator import validate_profile

        issues = [i for i in validate_profile(p) if i.severity.value == "error"]
        self.state.publish_issues(tr("Mapping validation"), issues)
        if issues:
            self.test_output.clear()
            for issue in issues[:20]:
                self.test_output.addItem(f"⚠ {issue_text(issue)[0]}")

    def _test(self) -> None:
        p, excel, folder = self.state.profile, self.state.excel_path, self.state.profile_dir
        if p is None or excel is None or folder is None:
            error(self, tr("Open a profile and load a workbook first."))
            return
        from docgen.application.preview import preview

        self.test_output.clear()
        self.test_output.addItem(tr("Testing…"))
        row = self.test_row.value()

        def done(result: Any) -> None:
            self.test_output.clear()
            if not result.records:
                self.test_output.addItem(
                    tr(
                        "No data row {row} in the primary sheet. Available rows: {rows}",
                        row=row,
                        rows=result.row_numbers[:10],
                    )
                )
            for issue in result.issues:
                if issue.severity.value != "info":
                    self.test_output.addItem(f"⚠ {issue_text(issue)[0]}")
            for rec in result.records:
                for trace in rec.traces:
                    value = trace.value
                    flag = "  ⚠ " + tr("untrusted source") if trace.untrusted else ""
                    item = QListWidgetItem(f"{trace.path} = {value!r}{flag}")
                    item.setToolTip(trace.describe())
                    self.test_output.addItem(item)
                    if trace.sources or trace.steps:
                        self.test_output.addItem(f"      ↳ {trace.describe()}")
                for issue in rec.issues:
                    self.test_output.addItem(f"⚠ {issue_text(issue)[0]}")
            self.state.publish_issues(
                tr("Mapping test"), [*result.issues, *(i for r in result.records for i in r.issues)]
            )

        run_in_background(
            preview,
            p.model_copy(deep=True),
            folder,
            excel,
            rows=[row],
            with_plan=False,
            owner=self,
            on_done=done,
            on_error=lambda m, d: error(self, m, d),
        )
