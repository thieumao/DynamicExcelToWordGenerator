"""Customer profile schema (version 1).

A profile fully describes how one customer's workbook is read, how values are
resolved across sheets, and how Word templates are rendered and merged. It is
stored as YAML (or JSON) and validated with Pydantic. Structural validation
lives here; cross-reference validation (unknown sheets, cycles, transform
parameters, expression syntax) lives in :mod:`docgen.domain.profile.validator`.
"""

from __future__ import annotations

from enum import StrEnum
from typing import Annotated, Any, Literal

from pydantic import (
    BaseModel,
    ConfigDict,
    Field,
    StringConstraints,
    field_validator,
    model_serializer,
    model_validator,
)

SCHEMA_VERSION: Literal[1] = 1

Identifier = Annotated[
    str, StringConstraints(pattern=r"^[A-Za-z_][A-Za-z0-9_]*$", min_length=1, max_length=64)
]
FieldPath = Annotated[
    str,
    StringConstraints(
        pattern=r"^[A-Za-z_][A-Za-z0-9_]*(\.[A-Za-z_][A-Za-z0-9_]*)*$", min_length=1, max_length=200
    ),
]
ProfileId = Annotated[
    str, StringConstraints(pattern=r"^[A-Za-z0-9][A-Za-z0-9_.-]*$", min_length=1, max_length=64)
]

RESERVED_ROOTS = frozenset(
    {"report", "unit", "group", "records", "first", "b", "template", "loop", "item", "this"}
)


class _Model(BaseModel):
    model_config = ConfigDict(extra="forbid", populate_by_name=True, validate_assignment=True)


# --------------------------------------------------------------------- enums
class FormulaPolicy(StrEnum):
    CACHED_REQUIRED = "cached_required"
    """Use cached results; a formula cell without a cached result is an error."""
    CACHED_WARN = "cached_warn"
    """Use cached results; missing cached results produce a warning and an empty value."""
    RAW_FORMULA = "raw_formula"
    """Return the formula text (e.g. '=SUM(A1:A3)') instead of a value."""
    SKIP_CHECK = "skip_check"
    """Fastest: read cached values without verifying formula cells (documented risk)."""


class ColumnType(StrEnum):
    AUTO = "auto"
    STRING = "string"
    INTEGER = "integer"
    DECIMAL = "decimal"
    DATE = "date"
    DATETIME = "datetime"
    BOOLEAN = "boolean"


class FieldType(StrEnum):
    AUTO = "auto"
    STRING = "string"
    INTEGER = "integer"
    DECIMAL = "decimal"
    DATE = "date"
    DATETIME = "datetime"
    BOOLEAN = "boolean"
    IMAGE = "image"


class Cardinality(StrEnum):
    ONE_TO_ONE = "one_to_one"
    MANY_TO_ONE = "many_to_one"
    ONE_TO_MANY = "one_to_many"

    @property
    def to_many(self) -> bool:
        return self is Cardinality.ONE_TO_MANY


class MatchMode(StrEnum):
    FIRST = "first"
    UNIQUE = "unique"
    ALL = "all"


class OnMissing(StrEnum):
    ERROR = "error"
    WARN = "warn"
    DEFAULT = "default"
    IGNORE = "ignore"


class OnDuplicate(StrEnum):
    ERROR = "error"
    WARN = "warn"


class SeverityName(StrEnum):
    ERROR = "error"
    WARNING = "warning"
    INFO = "info"


class MarkerStyle(StrEnum):
    JINJA = "jinja"
    CUSTOM = "custom"
    MERGEFIELD = "mergefield"


class MissingPolicy(StrEnum):
    ERROR = "error"
    WARN = "warn"
    EMPTY = "empty"


class UnitKind(StrEnum):
    RECORD = "record"
    GROUP = "group"
    DATASET = "dataset"


class OutputKind(StrEnum):
    SEPARATE = "separate"
    MERGED = "merged"
    BOTH = "both"


class ErrorPolicy(StrEnum):
    COLLECT = "collect"
    FAIL_FAST = "fail_fast"


class ConflictPolicy(StrEnum):
    SUFFIX = "suffix"
    SKIP = "skip"
    ERROR = "error"
    OVERWRITE = "overwrite"


# ------------------------------------------------------------------- general
class ProfileInfo(_Model):
    id: ProfileId
    name: str = Field(min_length=1, max_length=200)
    version: int = Field(default=1, ge=1)
    description: str = ""
    locale: str = "vi_VN"


class FormattingConfig(_Model):
    date_format: str = "%d/%m/%Y"
    datetime_format: str = "%d/%m/%Y %H:%M"
    decimal_separator: str = Field(default=",", max_length=1)
    thousands_separator: str = Field(default=".", max_length=1)
    currency_symbol: str = "₫"
    currency_position: Literal["prefix", "suffix"] = "suffix"
    currency_decimals: int = Field(default=0, ge=0, le=6)
    true_text: str = "Có"
    false_text: str = "Không"
    input_date_formats: list[str] = Field(
        default_factory=lambda: ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y"]
    )


# --------------------------------------------------------------------- excel
class SheetMatch(_Model):
    name: str | None = None
    aliases: list[str] = Field(default_factory=list)
    regex: str | None = None
    index: int | None = Field(default=None, ge=1, description="1-based sheet position")

    @model_validator(mode="after")
    def _one_criterion(self) -> SheetMatch:
        if self.name is None and self.regex is None and self.index is None and not self.aliases:
            raise ValueError("sheet match needs at least one of name, aliases, regex or index")
        return self


class HeaderConfig(_Model):
    mode: Literal["auto", "fixed", "none"] = "auto"
    row: int | None = Field(default=None, ge=1, description="1-based header row (fixed mode)")
    rows: int = Field(default=1, ge=1, le=5, description="number of header rows to combine")
    scan_rows: int = Field(default=30, ge=1, le=500)
    joiner: str = " "
    fill_merged: bool = True

    @model_validator(mode="after")
    def _fixed_needs_row(self) -> HeaderConfig:
        if self.mode == "fixed" and self.row is None:
            raise ValueError("header.mode 'fixed' requires header.row")
        return self


class DataRange(_Model):
    start_row: int | None = Field(default=None, ge=1)
    end_row: int | None = Field(default=None, ge=1)
    stop_after_blank_rows: int = Field(default=0, ge=0, description="0 = read to the end")
    skip_blank_rows: bool = True


class ColumnSpec(_Model):
    id: Identifier
    header: str | None = None
    aliases: list[str] = Field(default_factory=list)
    position: str | int | None = Field(
        default=None, description="Column letter ('C') or 1-based index; fallback or primary"
    )
    type: ColumnType = ColumnType.AUTO
    required: bool = True
    description: str = ""

    @field_validator("position")
    @classmethod
    def _check_position(cls, value: str | int | None) -> str | int | None:
        if isinstance(value, str):
            if not value.isalpha() or len(value) > 3:
                raise ValueError("position must be a column letter like 'C' or a 1-based number")
            return value.upper()
        if isinstance(value, int) and value < 1:
            raise ValueError("position index is 1-based")
        return value

    @model_validator(mode="after")
    def _header_or_position(self) -> ColumnSpec:
        if self.header is None and self.position is None and not self.aliases:
            raise ValueError(f"column '{self.id}' needs a header, aliases or a position")
        return self


class SheetConfig(_Model):
    id: Identifier
    match: SheetMatch
    layout: Literal["table", "key_value"] = Field(
        default="table",
        description="table: one header row, one record per row; key_value: a label cell with its "
        "value in the next filled cell to the right (e.g. 'Tên công trình | Nhà A'), read as one row",
    )
    header: HeaderConfig = Field(default_factory=HeaderConfig)
    data: DataRange = Field(default_factory=DataRange)
    columns: list[ColumnSpec] = Field(min_length=1)
    key: list[Identifier] = Field(default_factory=list)
    formula_policy: FormulaPolicy | None = None
    description: str = ""

    @model_validator(mode="after")
    def _unique_columns(self) -> SheetConfig:
        ids = [c.id for c in self.columns]
        dupes = sorted({i for i in ids if ids.count(i) > 1})
        if dupes:
            raise ValueError(f"duplicate column ids in sheet '{self.id}': {dupes}")
        missing = [k for k in self.key if k not in ids]
        if missing:
            raise ValueError(f"key columns not declared in sheet '{self.id}': {missing}")
        return self

    def column(self, column_id: str) -> ColumnSpec | None:
        return next((c for c in self.columns if c.id == column_id), None)


class HeaderMatching(_Model):
    casefold: bool = True
    strip_accents: bool = False
    collapse_whitespace: bool = True


class ExcelConfig(_Model):
    primary: Identifier
    sheets: list[SheetConfig] = Field(min_length=1)
    formula_policy: FormulaPolicy = FormulaPolicy.CACHED_REQUIRED
    header_matching: HeaderMatching = Field(default_factory=HeaderMatching)
    csv_encodings: list[str] = Field(default_factory=lambda: ["utf-8-sig", "utf-16", "cp1258"])
    csv_delimiter: str | None = None

    @model_validator(mode="after")
    def _check(self) -> ExcelConfig:
        ids = [s.id for s in self.sheets]
        dupes = sorted({i for i in ids if ids.count(i) > 1})
        if dupes:
            raise ValueError(f"duplicate sheet ids: {dupes}")
        if self.primary not in ids:
            raise ValueError(f"primary sheet '{self.primary}' is not declared in excel.sheets")
        return self

    def sheet(self, sheet_id: str) -> SheetConfig | None:
        return next((s for s in self.sheets if s.id == sheet_id), None)


# ------------------------------------------------------------- relationships
class KeyNormalize(_Model):
    trim: bool = True
    casefold: bool = False
    numeric_text: bool = True
    strip_leading_zeros: bool = False
    strip_accents: bool = False


class JoinPair(_Model):
    from_: Identifier = Field(alias="from")
    to: Identifier


class SortKey(_Model):
    by: str
    order: Literal["asc", "desc"] = "asc"
    nulls: Literal["first", "last"] = "last"


class Relationship(_Model):
    id: Identifier
    from_: Identifier = Field(alias="from", description="source sheet id")
    to: Identifier = Field(description="target sheet id")
    keys: list[JoinPair] = Field(
        default_factory=list,
        description="several pairs = composite key; none = every row uses the target sheet's only "
        "row (general information such as the project details), or all its rows for one_to_many",
    )
    cardinality: Cardinality = Cardinality.MANY_TO_ONE
    match: MatchMode | None = None
    on_missing: OnMissing = OnMissing.ERROR
    on_duplicate: OnDuplicate = OnDuplicate.ERROR
    empty_key_is_missing: bool = False
    default: Any = None
    key_normalize: KeyNormalize = Field(default_factory=KeyNormalize)
    filter: str | None = Field(default=None, description="expression over target row columns")
    sort: list[SortKey] = Field(default_factory=list)
    description: str = ""

    @model_validator(mode="after")
    def _check(self) -> Relationship:
        if self.cardinality.to_many and "on_missing" not in self.model_fields_set:
            # having no related rows (e.g. no projects) is normal for one-to-many
            self.on_missing = OnMissing.IGNORE
        if self.match is None:
            self.match = MatchMode.ALL if self.cardinality.to_many else MatchMode.UNIQUE
        if self.cardinality.to_many and self.match is not MatchMode.ALL:
            raise ValueError(f"relationship '{self.id}': one_to_many requires match 'all'")
        if not self.cardinality.to_many and self.match is MatchMode.ALL:
            raise ValueError(f"relationship '{self.id}': match 'all' needs cardinality one_to_many")
        return self

    @property
    def effective_match(self) -> MatchMode:
        assert self.match is not None
        return self.match


# -------------------------------------------------------------------- fields
class TransformCall(_Model):
    """A transform step. YAML accepts ``trim`` or ``{format_date: {format: '%d/%m/%Y'}}``."""

    name: Identifier
    params: dict[str, Any] = Field(default_factory=dict)

    @model_validator(mode="before")
    @classmethod
    def _coerce(cls, data: Any) -> Any:
        if isinstance(data, str):
            return {"name": data, "params": {}}
        if isinstance(data, dict) and "name" not in data and len(data) == 1:
            ((name, params),) = data.items()
            if params is None:
                params = {}
            if not isinstance(params, dict):
                params = {"value": params}
            return {"name": name, "params": params}
        return data

    @model_serializer(mode="plain")
    def _serialize(self) -> Any:
        return self.name if not self.params else {self.name: dict(self.params)}


class WhenClause(_Model):
    if_: str | None = Field(default=None, alias="if")
    else_: Any = Field(default=None, alias="else")
    value: Any = None
    expr: str | None = None

    @model_validator(mode="after")
    def _check(self) -> WhenClause:
        has_if = self.if_ is not None
        has_else = "else_" in self.model_fields_set
        if has_if == has_else:
            raise ValueError("a 'when' clause needs exactly one of 'if' or 'else'")
        if has_else and ("value" in self.model_fields_set or self.expr is not None):
            raise ValueError("an 'else' clause carries its value directly: {else: VALUE}")
        if has_if and self.expr is not None and "value" in self.model_fields_set:
            raise ValueError("use either 'value' or 'expr' in a 'when' clause, not both")
        return self

    @property
    def is_else(self) -> bool:
        return self.if_ is None

    @model_serializer(mode="plain")
    def _serialize(self) -> dict[str, Any]:
        if self.is_else:
            return {"else": self.else_}
        data: dict[str, Any] = {"if": self.if_}
        if self.expr is not None:
            data["expr"] = self.expr
        else:
            data["value"] = self.value
        return data


class ImageOptions(_Model):
    base_dir: str | None = Field(default=None, description="directory that image paths resolve in")
    width_mm: float | None = Field(default=None, gt=0, le=1000)
    height_mm: float | None = Field(default=None, gt=0, le=1000)
    max_bytes: int = Field(default=20 * 1024 * 1024, gt=0)


AggregateOp = Literal["count", "count_distinct", "sum", "min", "max", "avg", "join", "first", "last", "list"]


class FieldRule(_Model):
    """One output value. Exactly one value source must be set."""

    name: FieldPath
    source: str | None = Field(default=None, description="column id of the base sheet")
    path: str | None = Field(default=None, description="relationship chain: 'rel1.rel2.column'")
    expr: str | None = None
    const: Any = None
    when: list[WhenClause] | None = None
    coalesce: list[str] | None = Field(
        default=None, description="columns/paths tried in order; first non-empty wins"
    )
    aggregate: AggregateOp | None = Field(default=None, description="collapse a to-many path into one value")
    separator: str = ", "
    type: FieldType = FieldType.AUTO
    transforms: list[TransformCall] = Field(default_factory=list)
    default: Any = None
    required: bool = False
    image: ImageOptions | None = None
    description: str = ""

    @model_validator(mode="after")
    def _one_source(self) -> FieldRule:
        sources = [
            name for name in ("source", "path", "expr", "when", "coalesce") if getattr(self, name) is not None
        ]
        if "const" in self.model_fields_set:
            sources.append("const")
        if len(sources) != 1:
            raise ValueError(
                f"field '{self.name}' needs exactly one of source/path/expr/const/when/coalesce,"
                f" got {sources or 'none'}"
            )
        if self.when is not None:
            if not self.when:
                raise ValueError(f"field '{self.name}': 'when' needs at least one clause")
            else_positions = [i for i, c in enumerate(self.when) if c.is_else]
            if else_positions and else_positions != [len(self.when) - 1]:
                raise ValueError(f"field '{self.name}': 'else' must be the last clause")
        if self.type is FieldType.IMAGE and self.image is None:
            self.image = ImageOptions()
        return self

    @property
    def kind(self) -> str:
        for name in ("source", "path", "expr", "when", "coalesce"):
            if getattr(self, name) is not None:
                return name
        return "const"


class CollectionSpec(_Model):
    """A list of related records (one-to-many) with per-item fields."""

    name: FieldPath
    relationship: str = Field(description="relationship id or chain 'rel1.rel2' ending to-many")
    filter: str | None = Field(default=None, description="expression over the item row")
    sort: list[SortKey] = Field(default_factory=list)
    limit: int | None = Field(default=None, ge=1)
    distinct_by: list[str] = Field(default_factory=list)
    fields: list[FieldRule] = Field(default_factory=list)
    collections: list[CollectionSpec] = Field(default_factory=list)
    aggregations: list[AggregationSpec] = Field(default_factory=list)
    description: str = ""


class AggregationSpec(_Model):
    name: FieldPath
    relationship: str = Field(description="relationship id or chain ending to-many")
    op: AggregateOp
    of: str | None = Field(default=None, description="column id of the target sheet")
    filter: str | None = None
    separator: str = ", "
    distinct: bool = False
    transforms: list[TransformCall] = Field(default_factory=list)
    default: Any = None
    description: str = ""

    @model_validator(mode="after")
    def _needs_of(self) -> AggregationSpec:
        if self.op not in ("count",) and self.of is None:
            raise ValueError(f"aggregation '{self.name}': op '{self.op}' needs 'of'")
        return self


CollectionSpec.model_rebuild()


class ValidationRuleSpec(_Model):
    field: FieldPath
    rule: Literal["required", "regex", "range", "one_of", "length", "unique", "expr"]
    pattern: str | None = None
    min: float | int | str | None = None
    max: float | int | str | None = None
    values: list[Any] | None = None
    expr: str | None = None
    severity: SeverityName = SeverityName.ERROR
    message: str | None = None

    @model_validator(mode="after")
    def _params(self) -> ValidationRuleSpec:
        needs = {"regex": "pattern", "one_of": "values", "expr": "expr"}
        required = needs.get(self.rule)
        if required and getattr(self, required) is None:
            raise ValueError(f"validation rule '{self.rule}' on '{self.field}' needs '{required}'")
        if self.rule in ("range", "length") and self.min is None and self.max is None:
            raise ValueError(f"validation rule '{self.rule}' needs 'min' and/or 'max'")
        return self


# ----------------------------------------------------------------- templates
class MarkerConfig(_Model):
    style: MarkerStyle = MarkerStyle.JINJA
    open: str = Field(default="<<", min_length=1, max_length=10)
    close: str = Field(default=">>", min_length=1, max_length=10)
    content_controls: bool = Field(
        default=True, description="fill Word content controls (by tag/title) that are bound"
    )
    form_fields: bool = Field(
        default=True, description="fill legacy FORMTEXT/FORMCHECKBOX fields (by name) that are bound"
    )
    blanks: bool = Field(
        default=True,
        description="detect 'Label: ......' blanks and empty cells next to a label cell; "
        "only blanks with an explicit binding are filled",
    )
    blank_min: int = Field(default=5, ge=3, le=50, description="minimum dots/underscores for a blank")


class SectionConfig(_Model):
    mode: Literal["inherit", "preserve"] = "preserve"
    break_before: Literal["none", "page", "section"] = "page"
    start: Literal["nextPage", "continuous", "oddPage", "evenPage"] = "nextPage"
    restart_page_numbering: bool = False
    inherit_missing_headers: bool = False


class TemplateSpec(_Model):
    id: Identifier
    file: str = Field(min_length=1, description="path relative to the profile directory")
    enabled: bool = True
    description: str = ""
    marker: MarkerConfig = Field(default_factory=MarkerConfig)
    bindings: dict[str, str] = Field(default_factory=dict, description="placeholder name -> context path")
    context_root: str | None = None
    include_if: str | None = None
    missing_placeholder: MissingPolicy = MissingPolicy.ERROR
    section: SectionConfig = Field(default_factory=SectionConfig)


class MergeConfig(_Model):
    enabled: bool = True
    order: list[Identifier] = Field(default_factory=list)
    preserve_styles: bool = False
    update_fields_on_open: bool = True


class GenerationConfig(_Model):
    unit: UnitKind = UnitKind.RECORD
    output: OutputKind = OutputKind.MERGED
    group_by: list[FieldPath] = Field(default_factory=list)
    filter: str | None = Field(default=None, description="expression over the record context")
    sort: list[SortKey] = Field(default_factory=list)
    error_policy: ErrorPolicy = ErrorPolicy.COLLECT
    max_errors: int = Field(default=1000, ge=1)
    workers: int | Literal["auto"] = "auto"
    verify_output: bool = True

    @model_validator(mode="after")
    def _group(self) -> GenerationConfig:
        if self.unit is UnitKind.GROUP and not self.group_by:
            raise ValueError("generation.unit 'group' requires generation.group_by")
        return self


class OutputConfig(_Model):
    merged_filename: str = "{profile.id}_{unit.index}.docx"
    individual_filename: str = "{template.id}_{unit.index}.docx"
    subfolder: str | None = None
    on_conflict: ConflictPolicy = ConflictPolicy.SUFFIX
    max_path_length: int = Field(default=240, ge=60, le=32000)
    write_audit: bool = True


# ------------------------------------------------------------------- profile
class Profile(_Model):
    schema_version: Literal[1] = SCHEMA_VERSION
    profile: ProfileInfo
    formatting: FormattingConfig = Field(default_factory=FormattingConfig)
    excel: ExcelConfig
    relationships: list[Relationship] = Field(default_factory=list)
    fields: list[FieldRule] = Field(default_factory=list)
    collections: list[CollectionSpec] = Field(default_factory=list)
    aggregations: list[AggregationSpec] = Field(default_factory=list)
    group_fields: list[FieldRule] = Field(default_factory=list)
    validation: list[ValidationRuleSpec] = Field(default_factory=list)
    templates: list[TemplateSpec] = Field(default_factory=list)
    merge: MergeConfig = Field(default_factory=MergeConfig)
    generation: GenerationConfig = Field(default_factory=GenerationConfig)
    output: OutputConfig = Field(default_factory=OutputConfig)

    @model_validator(mode="after")
    def _unique_ids(self) -> Profile:
        for label, ids in (
            ("relationship", [r.id for r in self.relationships]),
            ("template", [t.id for t in self.templates]),
        ):
            dupes = sorted({i for i in ids if ids.count(i) > 1})
            if dupes:
                raise ValueError(f"duplicate {label} ids: {dupes}")
        return self

    # convenience lookups -------------------------------------------------
    def relationship(self, rel_id: str) -> Relationship | None:
        return next((r for r in self.relationships if r.id == rel_id), None)

    def template(self, template_id: str) -> TemplateSpec | None:
        return next((t for t in self.templates if t.id == template_id), None)

    def ordered_templates(self) -> list[TemplateSpec]:
        """Templates in merge order (merge.order first, then remaining declaration order)."""
        by_id = {t.id: t for t in self.templates}
        ordered = [by_id[i] for i in self.merge.order if i in by_id]
        ordered += [t for t in self.templates if t.id not in self.merge.order]
        return ordered

    def to_data(self) -> dict[str, Any]:
        """Serialisable dict using YAML-facing aliases ('from', 'if', 'else')."""
        data: dict[str, Any] = self.model_dump(mode="json", by_alias=True, exclude_none=False)
        pruned = _prune_defaults(data)
        assert isinstance(pruned, dict)
        return pruned


def _prune_defaults(data: Any) -> Any:
    """Drop None values and empty containers so saved YAML stays readable."""
    if isinstance(data, dict):
        result = {}
        for key, value in data.items():
            pruned = _prune_defaults(value)
            if pruned is None or pruned == [] or pruned == {}:
                if key in ("else",):
                    result[key] = pruned
                continue
            result[key] = pruned
        return result
    if isinstance(data, list):
        return [_prune_defaults(v) for v in data]
    return data
