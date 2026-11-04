"""Extract profile-declared sheets/columns from a workbook into a :class:`DataSet`."""

from __future__ import annotations

import difflib
import re
from collections.abc import Callable, Sequence
from dataclasses import dataclass, field
from itertools import islice
from pathlib import Path
from typing import Any

from docgen.domain.data import values as V
from docgen.domain.data.table import ColumnBinding, DataSet, SheetTable
from docgen.domain.diagnostics import Category, Issue, Severity, make_issue
from docgen.domain.profile.models import (
    ColumnSpec,
    ColumnType,
    ExcelConfig,
    FormattingConfig,
    FormulaPolicy,
    SheetConfig,
)
from docgen.infrastructure.excel.header_detect import combine_header_rows, detect_header_row
from docgen.infrastructure.excel.sources import (
    SheetInfo,
    WorkbookSource,
    column_index,
    column_letter,
    open_workbook,
)

Progress = Callable[[str], None]


@dataclass(slots=True)
class SheetLoadInfo:
    sheet_id: str
    sheet_name: str | None
    header_row: int | None
    header_confidence: float
    rows: int
    headers: list[str] = field(default_factory=list)


@dataclass(slots=True)
class LoadResult:
    dataset: DataSet
    issues: list[Issue]
    sheets: dict[str, SheetLoadInfo]
    available_sheets: list[SheetInfo]
    format: str

    @property
    def has_errors(self) -> bool:
        return any(i.severity is Severity.ERROR for i in self.issues)


def match_sheet(
    cfg: SheetConfig, available: Sequence[SheetInfo], *, casefold: bool = True
) -> SheetInfo | None:
    """Resolve a sheet config to a real sheet: exact name, aliases, regex, then index."""
    names = [cfg.match.name, *cfg.match.aliases] if cfg.match.name else list(cfg.match.aliases)
    for wanted in names:
        for info in available:
            if info.name == wanted:
                return info
    for wanted in names:
        norm = V.normalize_label(wanted, casefold=casefold)
        for info in available:
            if V.normalize_label(info.name, casefold=casefold) == norm:
                return info
    if cfg.match.regex:
        pattern = re.compile(cfg.match.regex, re.IGNORECASE)
        for info in available:
            if pattern.search(info.name):
                return info
    if cfg.match.index is not None and cfg.match.index <= len(available):
        return available[cfg.match.index - 1]
    return None


class WorkbookLoader:
    def __init__(self, excel: ExcelConfig, formatting: FormattingConfig | None = None) -> None:
        self.excel = excel
        self.formatting = formatting or FormattingConfig()

    def load(
        self,
        path: Path,
        *,
        sheet_ids: Sequence[str] | None = None,
        progress: Progress | None = None,
        header_overrides: dict[str, int] | None = None,
    ) -> LoadResult:
        source = open_workbook(
            path, csv_encodings=self.excel.csv_encodings, csv_delimiter=self.excel.csv_delimiter
        )
        try:
            return self._load(source, sheet_ids, progress, header_overrides or {})
        finally:
            source.close()

    # ------------------------------------------------------------------
    def _load(
        self,
        source: WorkbookSource,
        sheet_ids: Sequence[str] | None,
        progress: Progress | None,
        header_overrides: dict[str, int],
    ) -> LoadResult:
        issues: list[Issue] = []
        available = source.sheets()
        dataset = DataSet(source_path=str(source.path))
        infos: dict[str, SheetLoadInfo] = {}
        wanted = [s for s in self.excel.sheets if sheet_ids is None or s.id in sheet_ids]
        self._full_calc = source.supports_formulas and source.full_calc_on_load()
        self._full_calc_reported = False
        for cfg in wanted:
            if progress:
                progress(f"Reading sheet {cfg.id}")
            info = match_sheet(cfg, available, casefold=self.excel.header_matching.casefold)
            if info is None:
                names = [s.name for s in available]
                wanted_name = cfg.match.name or (cfg.match.aliases[0] if cfg.match.aliases else "")
                close = difflib.get_close_matches(wanted_name, names, n=3, cutoff=0.5)
                issues.append(
                    make_issue(
                        "sheet.not_found",
                        Severity.ERROR,
                        Category.SHEET,
                        f"Sheet for '{cfg.id}' not found (looked for {wanted_name!r}). Available: {names}",
                        suggestion=(
                            f"Did you mean {close}? Add it to match.aliases."
                            if close
                            else "Check the sheet name or set match.aliases / match.index."
                        ),
                        sheet=wanted_name,
                        rule=cfg.id,
                    )
                )
                infos[cfg.id] = SheetLoadInfo(cfg.id, None, None, 0.0, 0)
                continue
            if info.hidden:
                issues.append(
                    make_issue(
                        "sheet.hidden",
                        Severity.INFO,
                        Category.SHEET,
                        f"Sheet '{info.name}' is hidden in the workbook; it is still read.",
                        sheet=info.name,
                    )
                )
            table, load_info, sheet_issues = self._load_sheet(source, cfg, info, header_overrides.get(cfg.id))
            issues.extend(sheet_issues)
            infos[cfg.id] = load_info
            if table is not None:
                dataset.tables[cfg.id] = table
        return LoadResult(dataset, issues, infos, available, source.format)

    def _load_sheet(
        self,
        source: WorkbookSource,
        cfg: SheetConfig,
        info: SheetInfo,
        header_override: int | None,
    ) -> tuple[SheetTable | None, SheetLoadInfo, list[Issue]]:
        issues: list[Issue] = []
        hm = self.excel.header_matching
        header_cfg = cfg.header
        header_row: int | None
        confidence = 1.0
        if header_override is not None:
            header_row = header_override
        elif header_cfg.mode == "fixed":
            header_row = header_cfg.row
        elif header_cfg.mode == "none":
            header_row = None
        else:
            scan = list(islice(source.iter_rows(info.name, 1, header_cfg.scan_rows), header_cfg.scan_rows))
            expected = [c.header for c in cfg.columns if c.header] + [
                a for c in cfg.columns for a in c.aliases
            ]
            detection = detect_header_row(
                scan, expected=expected, casefold=hm.casefold, accents=hm.strip_accents
            )
            header_row, confidence = detection.row, detection.confidence
            if header_row is None:
                issues.append(
                    make_issue(
                        "header.not_detected",
                        Severity.ERROR,
                        Category.HEADER,
                        f"Could not detect a header row in the first {header_cfg.scan_rows} rows "
                        f"of sheet '{info.name}'",
                        suggestion="Set header.mode: fixed and header.row to the header row number.",
                        sheet=info.name,
                    )
                )
                return None, SheetLoadInfo(cfg.id, info.name, None, 0.0, 0), issues
            if confidence < 0.5 and expected:
                issues.append(
                    make_issue(
                        "header.low_confidence",
                        Severity.WARNING,
                        Category.HEADER,
                        f"Header row {header_row} in sheet '{info.name}' was detected with low "
                        f"confidence ({confidence:.0%})",
                        suggestion="Verify the header row in the Excel page or fix it manually.",
                        sheet=info.name,
                        row=header_row,
                    )
                )

        if header_row is not None:
            header_rows = list(
                islice(
                    source.iter_rows(info.name, header_row, header_row + header_cfg.rows - 1),
                    header_cfg.rows,
                )
            )
            headers = combine_header_rows(
                header_rows, joiner=header_cfg.joiner, fill_merged=header_cfg.fill_merged
            )
            data_start = header_row + header_cfg.rows
        else:
            headers = []
            data_start = 1
        if cfg.data.start_row is not None:
            data_start = cfg.data.start_row

        bindings, bind_issues = self._bind_columns(cfg, info.name, headers)
        issues.extend(bind_issues)
        if any(i.severity is Severity.ERROR for i in bind_issues):
            return None, SheetLoadInfo(cfg.id, info.name, header_row, confidence, 0, headers), issues

        table = SheetTable(cfg.id, info.name, bindings, header_row=header_row)
        positions = [b.index for b in bindings]
        blank_run = 0
        max_row = cfg.data.end_row
        for offset, raw in enumerate(source.iter_rows(info.name, data_start, max_row)):
            excel_row = data_start + offset
            projected = tuple((raw[p] if p is not None and p < len(raw) else None) for p in positions)
            if all(V.is_empty(v) for v in projected):
                blank_run += 1
                if cfg.data.stop_after_blank_rows and blank_run >= cfg.data.stop_after_blank_rows:
                    break
                if cfg.data.skip_blank_rows:
                    continue
            else:
                blank_run = 0
            converted = []
            row_index = len(table.rows)
            for ci, (spec, binding, value) in enumerate(zip(cfg.columns, bindings, projected, strict=True)):
                before = len(issues)
                converted.append(self._convert(value, spec, info.name, excel_row, binding, issues))
                if len(issues) > before and issues[-1].code == "cell.type_mismatch":
                    table.invalid[(row_index, ci)] = ("cell.type_mismatch", "error")
            table.rows.append(tuple(converted))
            table.row_numbers.append(excel_row)

        self._apply_formula_policy(source, cfg, table, issues)
        return (
            table,
            SheetLoadInfo(cfg.id, info.name, header_row, confidence, len(table.rows), headers),
            issues,
        )

    def _bind_columns(
        self, cfg: SheetConfig, sheet_name: str, headers: list[str]
    ) -> tuple[list[ColumnBinding], list[Issue]]:
        hm = self.excel.header_matching
        norm = [V.normalize_label(h, casefold=hm.casefold, accents=hm.strip_accents) for h in headers]
        issues: list[Issue] = []
        bindings: list[ColumnBinding] = []
        for spec in cfg.columns:
            index: int | None = None
            labels = [spec.header] if spec.header else []
            labels += spec.aliases
            if labels and headers:
                wanted = {
                    V.normalize_label(lbl, casefold=hm.casefold, accents=hm.strip_accents) for lbl in labels
                }
                hits = [i for i, h in enumerate(norm) if h and h in wanted]
                if len(hits) > 1:
                    pos = self._position(spec)
                    if pos is not None and pos in hits:
                        index = pos
                    else:
                        issues.append(
                            make_issue(
                                "column.ambiguous",
                                Severity.ERROR,
                                Category.COLUMN,
                                f"Column '{spec.id}' matches several headers in sheet "
                                f"'{sheet_name}': {[column_letter(h) for h in hits]}",
                                suggestion="Set 'position' to pick one, or rename the headers.",
                                sheet=sheet_name,
                                column=spec.header,
                                rule=f"{cfg.id}.{spec.id}",
                            )
                        )
                        continue
                elif hits:
                    index = hits[0]
            if index is None and (pos := self._position(spec)) is not None:
                index = pos
                if labels and headers:
                    found = headers[pos] if pos < len(headers) else ""
                    issues.append(
                        make_issue(
                            "column.position_fallback",
                            Severity.WARNING if found else Severity.INFO,
                            Category.COLUMN,
                            f"Column '{spec.id}' header {spec.header!r} not found; using position "
                            f"{column_letter(pos)} (header there: {found!r})",
                            sheet=sheet_name,
                            column=spec.header,
                            rule=f"{cfg.id}.{spec.id}",
                        )
                    )
            if index is None:
                severity = Severity.ERROR if spec.required else Severity.WARNING
                close = difflib.get_close_matches(spec.header or "", headers, n=3, cutoff=0.5)
                issues.append(
                    make_issue(
                        "column.not_found",
                        severity,
                        Category.COLUMN,
                        f"Column {spec.header or spec.id!r} ({cfg.id}.{spec.id}) not found in "
                        f"sheet '{sheet_name}'",
                        suggestion=(
                            f"Did you mean {close}? Add it to the column aliases."
                            if close
                            else "Check the header row or add aliases / a position."
                        ),
                        sheet=sheet_name,
                        column=spec.header or spec.id,
                        rule=f"{cfg.id}.{spec.id}",
                    )
                )
                bindings.append(ColumnBinding(spec.id, spec.header or spec.id, None, ""))
                continue
            header_text = (
                headers[index] if index < len(headers) and headers[index] else (spec.header or spec.id)
            )
            bindings.append(ColumnBinding(spec.id, header_text, index, column_letter(index)))
        return bindings, issues

    @staticmethod
    def _position(spec: ColumnSpec) -> int | None:
        if spec.position is None:
            return None
        if isinstance(spec.position, int):
            return spec.position - 1
        return column_index(spec.position)

    def _convert(
        self,
        value: Any,
        spec: ColumnSpec,
        sheet: str,
        row: int,
        binding: ColumnBinding,
        issues: list[Issue],
    ) -> Any:
        if isinstance(value, str):
            value = V.sanitize_text(value)
            if value in V.EXCEL_ERROR_VALUES:
                issues.append(
                    make_issue(
                        "cell.error_value",
                        Severity.WARNING,
                        Category.FORMULA,
                        f"Cell contains the Excel error {value}",
                        suggestion="Fix the formula in Excel.",
                        sheet=sheet,
                        row=row,
                        column=binding.header,
                    )
                )
                return None
        if value is None or spec.type is ColumnType.AUTO:
            return value
        fmt = self.formatting
        try:
            if spec.type is ColumnType.STRING:
                return V.to_text(value)
            if V.is_empty(value):
                return None
            if spec.type in (ColumnType.INTEGER, ColumnType.DECIMAL):
                number = (
                    V.parse_localized_number(value, fmt.decimal_separator, fmt.thousands_separator)
                    if isinstance(value, str)
                    else V.to_decimal(value)
                )
                if spec.type is ColumnType.DECIMAL:
                    return number
                return V.to_int(number)
            if spec.type is ColumnType.DATE:
                return V.to_date(value, fmt.input_date_formats)
            if spec.type is ColumnType.DATETIME:
                return V.to_datetime(value)
            if spec.type is ColumnType.BOOLEAN:
                return V.to_bool(value)
        except V.ConversionError as exc:
            issues.append(
                make_issue(
                    "cell.type_mismatch",
                    Severity.ERROR,
                    Category.TYPE,
                    f"Value {value!r} is not a valid {spec.type.value}: {exc}",
                    suggestion=f"Correct the cell or change the column type of '{spec.id}'.",
                    sheet=sheet,
                    row=row,
                    column=binding.header,
                    rule=spec.id,
                )
            )
            return None
        return value

    def _apply_formula_policy(
        self, source: WorkbookSource, cfg: SheetConfig, table: SheetTable, issues: list[Issue]
    ) -> None:
        policy = cfg.formula_policy or self.excel.formula_policy
        if policy is FormulaPolicy.SKIP_CHECK or not table.rows:
            return
        if not source.supports_formulas:
            return
        columns = {b.index for b in table.columns if b.index is not None}
        first, last = table.row_numbers[0], table.row_numbers[-1]
        formulas = source.formula_cells(table.sheet_name, columns, first, last)
        if not formulas:
            return
        if self._full_calc and not self._full_calc_reported and policy is not FormulaPolicy.RAW_FORMULA:
            self._full_calc_reported = True
            issues.append(
                make_issue(
                    "formula.full_calc_on_load",
                    Severity.WARNING,
                    Category.FORMULA,
                    "The workbook contains formulas and asks Excel to recalculate them on open, "
                    "so stored results may be stale.",
                    suggestion="Open the file in Excel, let it recalculate, save, then retry.",
                    sheet=table.sheet_name,
                )
            )
        col_pos = {b.index: i for i, b in enumerate(table.columns) if b.index is not None}
        row_lookup = {r: i for i, r in enumerate(table.row_numbers)}
        missing_reported = 0
        for (excel_row, col), formula in formulas.items():
            ri = row_lookup.get(excel_row)
            if ri is None:
                continue
            ci = col_pos[col]
            cached = table.rows[ri][ci]
            binding = table.columns[ci]
            if policy is FormulaPolicy.RAW_FORMULA:
                row_values = list(table.rows[ri])
                row_values[ci] = formula
                table.rows[ri] = tuple(row_values)
                continue
            if cached is None:
                severity = Severity.ERROR if policy is FormulaPolicy.CACHED_REQUIRED else Severity.WARNING
                table.invalid[(ri, ci)] = ("formula.no_cached_value", severity.value)
                if missing_reported < 500:
                    issues.append(
                        make_issue(
                            "formula.no_cached_value",
                            severity,
                            Category.FORMULA,
                            f"Formula {formula} has no calculated value stored in the file "
                            "(openpyxl does not calculate formulas)",
                            suggestion="Open and save the workbook in Excel (or LibreOffice) so "
                            "results are stored, or use formula_policy: cached_warn.",
                            sheet=table.sheet_name,
                            row=excel_row,
                            column=f"{binding.header} ({binding.letter})",
                        )
                    )
                missing_reported += 1
        if missing_reported > 500:
            issues.append(
                make_issue(
                    "formula.no_cached_value_more",
                    Severity.WARNING,
                    Category.FORMULA,
                    f"{missing_reported - 500} more formula cells without stored results in "
                    f"sheet '{table.sheet_name}'",
                    sheet=table.sheet_name,
                )
            )
