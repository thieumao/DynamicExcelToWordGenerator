# Data mapping design and profile reference (schema v1)

A **customer profile** describes everything that differs between customers. Code never contains
customer-specific sheet names, columns or placeholders. Profiles are YAML (or JSON) files validated
by Pydantic (`src/docgen/domain/profile/models.py`); the JSON Schema can be exported with
`docgen profile schema -o profile.schema.json` for editor auto-completion.

Complete working examples: `samples/customer_a/profile.yaml` (Jinja templates, 5 sheets, chained
lookups, one-to-many) and `samples/customer_b/profile.yaml` (Vietnamese headers, header on row 4,
composite key, `<<...>>` markers, MERGEFIELDs).

> YAML note: DocGen reads YAML with **YAML 1.2 booleans** – only `true`/`false` are booleans.
> `yes`, `no`, `on`, `off`, `Y`, `N` stay text (PyYAML's default YAML 1.1 rules would silently turn
> them into booleans and break keys such as `on:` or code tables such as `{Y: Có, N: Không}`).

## 1. Top-level structure

```yaml
schema_version: 1                 # migrations upgrade older files automatically
profile: {id: customer_a, name: "Customer A", version: 3, locale: vi_VN, description: ""}
formatting: {...}                 # locale defaults for dates, numbers, currency, booleans
excel: {...}                      # sheets, columns, header detection, formula policy
relationships: [...]              # how sheets link to each other
fields: [...]                     # scalar output values
collections: [...]                # lists (one-to-many) with per-item fields
aggregations: [...]               # sum/count/join/... over related rows
group_fields: [...]               # values computed per group (unit: group/dataset)
validation: [...]                 # data rules
templates: [...]                  # Word templates, markers, bindings, section behaviour
merge: {...}                      # merge order and options
generation: {...}                 # unit kind (record/group/dataset), output kind, policies
output: {...}                     # file name patterns, conflict policy
```

## 2. Excel

```yaml
excel:
  primary: emp                         # id of the sheet whose rows drive generation
  formula_policy: cached_required      # cached_required | cached_warn | raw_formula | skip_check
  header_matching: {casefold: true, strip_accents: false, collapse_whitespace: true}
  csv_encodings: [utf-8-sig, utf-16, cp1258]
  sheets:
    - id: emp                          # stable logical id used everywhere else
      match: {name: Employees, aliases: ["Nhân viên"], regex: null, index: null}
      header: {mode: auto, row: null, rows: 1, scan_rows: 30, joiner: " ", fill_merged: true}
      data: {start_row: null, end_row: null, stop_after_blank_rows: 0, skip_blank_rows: true}
      key: [employee_id]               # optional: declared primary key
      columns:
        - {id: employee_id, header: EmployeeID, aliases: ["Mã NV"], position: null,
           type: string, required: true}
```

* **Sheet matching** – exact name, aliases, case/space-insensitive name, regex, then 1-based index.
  Unknown sheets produce an error listing the available sheets and close matches.
* **Header detection** (`mode: auto`) scores the first `scan_rows` rows: expected header/alias hits
  (weighted ×10), number of text cells, label uniqueness, and whether the next row looks like data.
  Titles and blank rows above the header are skipped. `mode: fixed` + `row` overrides;
  `mode: none` means headerless (columns by `position`). `rows: 2` combines multi-row headers,
  forward-filling merged cells (`"Thông tin" + "Họ tên" → "Thông tin Họ tên"`).
* **Column binding** – by header/aliases (normalised: NFC, trimmed, collapsed spaces, case-folded,
  optionally accent-insensitive), else by `position` (`"C"` or `3`). Several matching headers →
  ambiguity error unless `position` disambiguates. A missing required column is an error with
  "did you mean" suggestions; a missing optional column yields empty values and a warning.
* **Types** – `auto` keeps Excel's value; `string`, `integer`, `decimal` (locale-aware text parsing:
  `1.234.567,5`), `date`, `datetime`, `boolean` (Có/Không, yes/no, x, 1/0). Conversion failures are
  errors with sheet/row/column; the cell becomes *untrusted* (see §6).
* **Only declared columns are kept** (projection); blank rows are skipped; original Excel row numbers
  are retained for diagnostics.

### Formula policy

openpyxl never calculates formulas. DocGen reads the *cached* result Excel stored when saving.

| Policy | Behaviour |
|---|---|
| `cached_required` (default) | Formula cell without a stored result → error; records using it fail. |
| `cached_warn` | Same detection, but warning; value is empty. |
| `raw_formula` | Value is the formula text (`=SUM(A1:A3)`), for documentation use cases. |
| `skip_check` | No formula scan (fastest). Use only when the customer guarantees no formulas. |

Detection is a streamed scan of the sheet XML (`<f>` elements), preceded by a byte-level check so
sheets without formulas cost almost nothing. If the workbook requests `fullCalcOnLoad` and contains
formulas, a warning says stored results may be stale. `.xls` and `.csv` have no formula information
(.xls exposes only results). Recalculating requires Excel or LibreOffice and is deliberately not
done automatically.

## 3. Relationships (lookups)

```yaml
relationships:
  - id: emp_dept
    from: emp                    # source sheet
    to: dept                     # target sheet
    keys:                        # one pair = simple key, several pairs = composite key
      - {from: department_id, to: department_id}
    cardinality: many_to_one     # one_to_one | many_to_one | one_to_many
    match: unique                # first | unique | all   (default: unique, or all for one_to_many)
    on_missing: error            # error | warn | default | ignore (default ignore for one_to_many)
    on_duplicate: error          # error | warn            (only with match: unique)
    default: "(không rõ)"        # value used when on_missing: default
    empty_key_is_missing: false  # an empty foreign key normally means "no relation"
    key_normalize: {trim: true, casefold: false, numeric_text: true, strip_leading_zeros: false,
                    strip_accents: false}
    filter: "status == 'Active'" # expression over target row columns
    sort: [{by: valid_from, order: desc}]
```

* **Indexes** – one hash index per relationship and job (`dict[key tuple → row list]`), built once,
  so each lookup is O(1).
* **Key normalisation** – `numeric_text` makes Excel's `101`, `101.0` and the text `"101"` equal while
  keeping `"007"` distinct from `7` (unless `strip_leading_zeros`).
* **Duplicates are never resolved silently.** With `match: unique` duplicate keys are reported at
  index time (all row numbers). A record hitting a duplicated key gets an error and **no value**
  (`on_duplicate: error`), or a warning and the first row in sheet/sort order (`warn`).
  `match: first` is an explicit, documented choice of the first row.
* **Cardinality** – `one_to_many` returns all matching rows (for collections/aggregations).

## 4. Output rules

Every output has a dotted `name` (`employee.department.name`), producing a nested context.
Names must not start with reserved roots (`report`, `unit`, `group`, `records`, `first`, `b`,
`template`, `loop`, `item`, `this`) or with a sheet/relationship id.

### Fields – exactly one value source

```yaml
fields:
  - {name: employee.code, source: employee_id}                  # column of the base sheet
  - {name: employee.dept, path: emp_dept.name}                  # relationship chain → column
  - {name: employee.manager, path: emp_dept.dept_mgr.name}      # multi-level lookup
  - {name: employee.projects_text, path: emp_proj.project, aggregate: join, separator: "; "}
  - {name: employee.contact, coalesce: [email, emp_dept.email]} # first non-empty across sheets
  - {name: employee.label, expr: "employee.full_name ~ ' (' ~ employee.code ~ ')'"}
  - {name: employee.kind, const: "Chính thức"}
  - name: employee.level
    when:                                                      # conditional mapping
      - {if: "emp.grade >= 5", value: "Cao cấp"}
      - {if: "emp.grade >= 3", expr: "'Chuyên viên ' ~ emp.grade"}
      - {else: "Nhân viên"}
    type: string                       # optional final type conversion (also: image)
    transforms: [trim, title_case, {format_currency: {symbol: "VNĐ"}}]
    default: "—"                       # used when the value is empty
    required: true                     # empty after default → error
```

A `path` whose chain passes a one-to-many relationship must declare `aggregate`
(`join | first | last | count | count_distinct | sum | min | max | avg | list`); otherwise the
validator reports an error, because a scalar placeholder would otherwise receive a list.

### Collections (repeating data)

```yaml
collections:
  - name: employee.projects
    relationship: emp_proj               # or a chain "a.b" ending in one_to_many
    filter: "item.hours > 0"             # expression over the item row (item.<column>)
    sort: [{by: hours, order: desc}]     # item field or target column
    limit: 20
    distinct_by: [name]
    fields:                              # evaluated with the target sheet as base
      - {name: name, source: project}
      - {name: client, path: proj_client.name}   # lookups continue from the item row
    collections: []                      # nested lists are allowed
    aggregations: []
```

In Word: `{%tr for p in employee.projects %} … {{ p.name }} … {%tr endfor %}`.

### Aggregations

```yaml
aggregations:
  - {name: employee.total_hours, relationship: emp_proj, op: sum, of: hours,
     filter: "item.billable == 'Y'", transforms: [{format_number: {decimals: 1}}], default: "0"}
  - {name: employee.project_count, relationship: emp_proj, op: count}
```

### Group fields (unit: group / dataset)

Evaluated once per group with `group`, `records` and `first` in scope:

```yaml
group_fields:
  - {name: dept_total, expr: "records | sum(attribute='employee.salary_value')"}
```

## 5. Expressions

Expressions are **Jinja expressions** evaluated in an immutable sandbox (no Python `eval`):

* Names: the base sheet id gives the current row (`emp.grade`), computed fields by full path
  (`employee.full_name`), `report.*`, and in collections `item.*`, `this.*`, `parent.*`.
* Operators: `== != < <= > >= and or not in ~ + - * / // %`, inline `a if cond else b`.
* Filters: every transform (`| title_case`, `| format_date('%d/%m/%Y')`) plus Jinja built-ins.
* Functions: `coalesce()`, `concat_ws(sep, …)`, `today()`, `now()`, `date(y, m, d)`,
  `days_between(a, b)`, `num(x)`, `empty(x)`.
* Mapping keys win over Python attributes: a column named `items` or `keys` works as expected.
* Private attributes and Python internals are blocked by the sandbox; expressions longer than 2,000
  characters are rejected.

## 6. Determinism, dependencies, traceability

* **Dependency graph** – references in `expr`/`when` create edges to the fields they use
  (`employee.label` depends on `employee.full_name`). Evaluation follows a topological order with
  declaration order as the tie-breaker, so results never depend on dict/set ordering.
* **Cycle detection** – cycles are reported with their path (`out.a -> out.b -> out.a`), including
  self-references; such a profile cannot be used for generation.
* **Traceability** – each field records its source cells (`Departments!B5 [DepartmentName]`,
  including cells read by expressions/conditions) and the steps applied
  (`emp_dept: Employees -> Departments rows [5] -> trim -> title_case`). Shown on the Preview page,
  in the Mapping designer test panel and by `docgen preview --trace`.
* **Untrusted cells** – formula cells without stored results and cells that failed type conversion
  are tracked; any field that reads one produces a record-level issue, so documents are never built
  from unverified values silently.

## 7. Transforms

| Category | Transforms |
|---|---|
| Text | `trim`, `collapse_spaces`, `upper`, `lower`, `title_case`, `capitalize`, `replace{old,new}`, `regex_replace{pattern,repl}`, `regex_extract{pattern,group}`, `slice{start,end}`, `pad_left{width,char}`, `pad_right`, `prefix{text}`, `suffix{text}`, `strip_accents` |
| Type | `to_string`, `to_int`, `to_decimal`, `to_date{format}`, `to_bool` |
| Format | `format_date{format}`, `format_datetime`, `format_number{decimals,…}`, `format_currency{symbol,decimals,position}`, `format_percent`, `format_bool{true,false}`, `number_to_words_vi{suffix,capitalize}` |
| Logic | `default{value}`, `map_values{mapping,default,keep_unmapped}` |
| Lists | `split`, `join`, `lines`, `first`, `last`, `count`, `sum`, `unique` |
| Numbers/dates | `round{decimals}`, `multiply{by}`, `add{value}`, `add_days{days}` |

Formatting defaults come from `formatting:` (`dd/mm/yyyy`, `1.234.567,5`, `15.000.000 ₫`).
Parameters are validated statically (unknown transform, unknown/missing parameter, invalid choice).

## 8. Generation units and modes

```yaml
generation:
  unit: record          # record | group | dataset
  output: merged        # separate | merged | both
  group_by: [employee.department.name]   # unit: group
  filter: "employee.status != 'Nghỉ việc'"
  sort: [{by: employee.code}]
  error_policy: collect # collect | fail_fast
  max_errors: 1000
  workers: auto
  verify_output: true
```

| Mode | Settings |
|---|---|
| A – single document | `unit: record`, one selected row (`--rows 5` / Records field) |
| B – batch documents | `unit: record`, `output: separate` |
| C – grouped document | `unit: group`, `group_by: [...]` → templates use `group.*`, `records`, `first` |
| D – merged document | several templates, `output: merged` (any unit; `unit: dataset` = everything in one) |
| E – batch merged | `unit: record` or `group`, `output: merged` |

## 9. Output naming

```yaml
output:
  merged_filename: "{employee.department.name}/HoSo_{employee.code}_{report.generated_at:%Y%m%d}.docx"
  individual_filename: "{template.id}_{employee.code}.docx"
  subfolder: null
  on_conflict: suffix   # suffix | skip | error | overwrite
  max_path_length: 240
```

`/` in the *pattern* creates sub-folders; `/` or `\` inside *values* is replaced so data cannot
create folders. Names are sanitised for Windows (reserved names such as `CON`, invalid characters,
trailing dots/spaces), Unicode NFC, shortened with a hash suffix when the path is too long. Existing
files are never overwritten unless `on_conflict: overwrite`; even then two units of one job never
overwrite each other, and a file locked by Word is written next to it with a numbered name.
