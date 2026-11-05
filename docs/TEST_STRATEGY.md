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
| Unit – Word input spots | `tests/unit/test_spots.py` | Detection of blanks after labels (dots, underscores, ellipsis; repeated labels `tháng (2)`; header parts; short ellipses ignored), empty label/value table cells, content controls (text/drop-down/check box/date/block, tag vs title naming) and legacy FORMTEXT/FORMCHECKBOX fields; detection alone changes nothing; only bound inputs are filled and land in place with their formatting; unbound signature/date lines untouched; check box states both ways; filled cells never become labels; custom-XML binding and stale date removed from filled controls; options switch detectors off. |
| PoC – merging | `tests/poc/test_merge_poc.py` | Section/header/footer preservation, inheritance options, section breaks, page-number restart; optional LibreOffice rendering check (page count, per-page headers, numbering restart). |
| Regression – formatting | `tests/regression/test_formatting_preservation.py` | Template vs output: fonts, run formatting, styles, tables, merged cells, numbering, images, header images, page breaks, section layout, fields, hyperlinks; and after merging two rendered documents. |
| Regression – Office conformance | `tests/regression/test_ooxml_conformance.py`, `tools/ooxml-validator/` | All outputs of customers A, B and C (separate + merged, filled form controls) and the rich template in `preserve`/`section`/`inherit` merges are validated with **Microsoft's Open XML SDK `OpenXmlValidator`** (DocumentFormat.OpenXml 3.4.1) against Office 2010, Office 2019 and Microsoft 365 – 0 errors required. Needs the .NET 8 SDK (preinstalled on GitHub runners; CI sets `DOCGEN_REQUIRE_OOXML=1` so it cannot be skipped there). |
| Regression – other editors | `tests/regression/test_libreoffice_templates.py` | Sample templates re-saved by LibreOffice (docx → odt → docx, different OOXML producer) still generate complete documents; unbound MERGEFIELDs flattened to `«Name»` text by the editor are filled. |
| Integration | `tests/integration/test_pipeline.py`, `test_cli.py` | Excel → mapping → context → Word for modes A–E; customer B with the same code; customer C filling an existing Word form (blanks, controls, form fields, table cells; unbound inputs untouched); process pool; no overwrite; cancellation (no partial/temp files); collect vs fail-fast; error location; dry run; missing template; unknown placeholder caught before generation; inputs unchanged (hash); unwritable output; CLI commands. |
| UI | `tests/ui/test_ui.py` (pytest-qt) | All pages render; mapping designer builds a lookup chain + transform and tests it on a record; generation from the UI with **event-loop heartbeat** (max stall asserted < 0.5 s while 30 merged documents are produced); cancel from the UI; Templates page lists form inputs of customer C and binding a blank from its combo box re-scans and marks it "will be filled"; Vietnamese catalog. |
| Architecture | `tests/test_architecture.py` | Core layers do not import Qt; domain does not import infrastructure; no `eval`/`exec`. |
| Packaging | `packaging/smoke_test.py` | Built `docgen-cli.exe`: version, sample generation, profile validation, 25-document batch through the process pool inside the frozen app, customer B, customer C (existing form), Open XML SDK validation of the frozen app's outputs; `DocGen.exe --smoke-test` starts and quits. |

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
- [ ] Existing forms: filled blanks keep the label's font/underline; filled content controls show the
      value (not the grey placeholder) and can still be edited; check boxes show ☒/☐ and toggle when
      clicked; legacy form fields keep working after "Restrict Editing → Filling in forms";
      unbound blanks, signature lines and office-use boxes are unchanged.

## What has *not* been tested yet (honest status)

* Opening outputs in Microsoft Word itself (no Word in the development container). Substitutes in
  place: Microsoft's Open XML SDK validator (0 errors for Office 2010/2019/365 on every generated
  file), LibreOffice rendering, structural verification. Validation against Office 2007 reports only
  `w:tblLook` attributes (`firstRow`, `noHBand`, …) that python-docx's default template – and Word
  2010+ itself – writes; Word 2007 is not a target.
* Real customer files (none provided yet).
* The Windows executable on a physical Windows 10/11 machine: the workflow builds and smoke-tests it
  on GitHub's `windows-latest` runner once the repository can be pushed.
* Very large templates (hundreds of pages), charts, SmartArt, comments, tracked changes. Content
  controls and legacy form fields are covered by synthetic forms only (python-docx + raw XML as Word
  writes it), not yet by forms saved by Word itself.
