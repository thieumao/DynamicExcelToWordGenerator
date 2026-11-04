# Test strategy

Run everything: `uv run pytest` (set `QT_QPA_PLATFORM=offscreen` on headless machines).
CI runs the suite on Windows, macOS and Linux (`.github/workflows/ci.yml`), and the Windows
packaging workflow re-runs it before building and smoke-tests the built executables.

## Layers

| Suite | Path | What it proves |
|---|---|---|
| Unit – values, transforms, Vietnamese numbers, expressions | `tests/unit/test_values.py`, `test_transforms.py`, `test_vietnamese.py`, `test_expressions.py` | Key normalisation (`101` = `101.0` = `"101"`, `"007"` kept), locale number parsing, all transform families, sandbox blocks `__class__`/globals access, reference extraction. |
| Unit – profile schema & validator | `tests/unit/test_profile.py` | YAML 1.2 booleans, aliases (`from`/`if`/`else`), defaults by cardinality, schema errors, round-trip stability (YAML and JSON), unknown sheets/columns/relationships, wrong chains, to-many without aggregate, reserved names, cycles incl. self-reference, path conflicts, deterministic order. |
| Unit – Excel | `tests/unit/test_excel_loader.py`, `test_header_detect.py` | Sheet matching by name/alias/regex/index/case; header below title rows; multi-row headers with merged cells; accent-insensitive headers; missing required/optional columns; position fallback; headerless sheets; empty cells/rows; type conversion errors with row/column; Excel error values; `stop_after_blank_rows`; formula policies (missing cached values, cached values injected as Excel would save them, raw formulas, warn/skip); `.xls`; CSV in UTF-8-BOM / Windows-1258 (combining tones) / UTF-16; CSV folders; hidden sheets; invalid files. |
| Unit – mapping | `tests/unit/test_mapping.py` | One-level and multi-level lookups, composite keys, coalesce across sheets, conditional mapping, expressions, image fields, one-to-many collections with relationship filter/sort and item filter, aggregations (sum/max/count/join/avg/min/first/last/count_distinct), missing-reference policies (error/default/ignore/warn), empty foreign keys, duplicate keys (error with no arbitrary pick vs warn), validation rules incl. uniqueness, untrusted formula cells, group contexts, traces with source cells. |
| Unit – Word | `tests/unit/test_word.py` | Custom markers split across runs (formatting kept), headers, unsafe names; MERGEFIELD simple/complex, unsupported regions; inspector paths for loops/conditions/headers; syntax errors; foreign marker warning; rendering of values, table-row loops, conditional paragraphs, escaping of `<&>`, line breaks, date/decimal/bool formatting, header/footer placeholders; missing-placeholder policies; images (allowed folder, missing, outside); verifier (corrupt ZIP, leftover markers, broken relationships). |
| Unit – files | `tests/unit/test_files.py` | Filename patterns, Windows reserved names, long names, path traversal, conflict policies (suffix/skip/error/overwrite/reserved), zip-bomb detection. |
| PoC – merging | `tests/poc/test_merge_poc.py` | Section/header/footer preservation, inheritance options, section breaks, page-number restart; optional LibreOffice rendering check (page count, per-page headers, numbering restart). |
| Regression – formatting | `tests/regression/test_formatting_preservation.py` | Template vs output: fonts, run formatting, styles, tables, merged cells, numbering, images, header images, page breaks, section layout, fields, hyperlinks; and after merging two rendered documents. |
| Integration | `tests/integration/test_pipeline.py`, `test_cli.py` | Excel → mapping → context → Word for modes A–E; customer B with the same code; process pool; no overwrite; cancellation (no partial/temp files); collect vs fail-fast; error location; dry run; missing template; unknown placeholder caught before generation; inputs unchanged (hash); unwritable output; CLI commands. |
| UI | `tests/ui/test_ui.py` (pytest-qt) | All pages render; mapping designer builds a lookup chain + transform and tests it on a record; generation from the UI with **event-loop heartbeat** (max stall asserted < 0.5 s while 30 merged documents are produced); cancel from the UI; Vietnamese catalog. |
| Architecture | `tests/test_architecture.py` | Core layers do not import Qt; domain does not import infrastructure; no `eval`/`exec`. |
| Packaging | `packaging/smoke_test.py` | Built `docgen-cli.exe`: version, sample generation, profile validation, 25-document batch through the process pool inside the frozen app, customer B; `DocGen.exe --smoke-test` starts and quits. |

Fixtures are **synthetic and generated in code** (`src/docgen/testing/`): workbooks via openpyxl/xlwt
(plus an XML patch that inserts cached formula results as Excel does), documents via python-docx with
real Word features (headers with images, PAGE fields, landscape sections, numbering, hyperlinks,
merged cells, MERGEFIELDs, runs split character by character). This keeps tests reproducible on every
OS without binary fixtures in git.

## Real customer documents

Synthetic fixtures do not prove compatibility with every customer file. For each new customer:

1. Put anonymised copies of their workbook and templates in `tests/fixtures/real/<customer>/`
   (git-ignored – may contain confidential data) together with their `profile.yaml`.
2. Run `docgen profile validate` and `docgen generate --dry-run`, then a real generation.
3. Open the outputs in **Microsoft Word on Windows** and walk through the checklist below.
4. Add a regression test (structure snapshot like `test_formatting_preservation.py`) for every
   feature of their templates that is not yet covered.

## Manual checklist (Windows + Microsoft Word) – required before calling a customer "supported"

- [ ] Output opens without the "unreadable content" repair prompt.
- [ ] Fonts (incl. Vietnamese diacritics), styles, colours, table borders/widths match the template.
- [ ] Headers/footers per section (first-page/even headers if used), page numbers continue/restart as
      configured, landscape sections correct.
- [ ] Numbered/bulleted lists restart or continue as expected per part.
- [ ] Images (body and header) appear with the right size; hyperlinks work.
- [ ] Table of contents / cross-references update when prompted.
- [ ] Repeating tables have the right number of rows; conditional sections appear only when expected.
- [ ] File names, sub-folders and Unicode names are correct in Explorer; no overwritten files.
- [ ] The audit JSON/CSV list every unit, error and output path.
- [ ] UI stays responsive during a large batch; Cancel stops after the current document.

## What has *not* been tested yet (honest status)

* Opening outputs in Microsoft Word (no Word in the development container) – LibreOffice rendering
  and structural checks were used instead.
* Real customer files (none provided yet).
* The Windows executable on a physical Windows 10/11 machine: the workflow builds and smoke-tests it
  on GitHub's `windows-latest` runner once the repository can be pushed.
* Very large templates (hundreds of pages), charts, SmartArt, content controls, comments, tracked
  changes.
