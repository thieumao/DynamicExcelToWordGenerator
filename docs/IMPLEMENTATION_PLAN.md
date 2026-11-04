# Implementation plan

## Problem

Customers send multi-sheet Excel workbooks: one primary sheet with IDs/codes and several lookup and
related sheets. Resolved values must fill placeholders in several Word templates, each drawing on
different sheets; outputs are individual documents, merged documents, or both. Every customer uses
different sheet names, columns, header positions and template conventions, so all of it must be
configuration (customer profiles), not code. Target users run Windows 10/11; development happens
on macOS.

## Functional requirements (summary)

1. Read .xlsx/.xlsm/.xls/.csv; detect sheets, header rows (auto + manual), columns by header,
   aliases or position.
2. Relationships between sheets: exact and composite keys, multi-level chains, one-to-one,
   many-to-one, one-to-many; first/unique/all matching; explicit missing/duplicate policies.
3. Output fields: column, lookup chain, coalesce across sheets, expression, condition, constant;
   transforms (text, type, date/number/currency, Vietnamese number words); defaults; required flags;
   collections (nested lists), aggregations, group fields; filtering/sorting; validation rules.
4. Customer profiles (YAML/JSON, versioned, schema-validated): create, duplicate, delete
   (recoverable), import/export (with templates), edit in UI or YAML, validate, preview.
5. Word templates: Jinja placeholders, custom markers and MERGEFIELDs; repeating tables, conditional
   sections, nested data, images; placeholder detection and binding; missing-placeholder policies.
6. Merging with per-template order, enablement, include conditions, page/section breaks and section
   preservation (orientation, margins, headers/footers).
7. Generation modes A–E, configurable file names, no accidental overwrites, audit report.
8. Desktop UI: dashboard, profiles, Excel configuration, relationships, mapping designer, templates,
   preview, generate (progress/cancel), logs, settings; Vietnamese and English.
9. CLI for automation and tests.

## Non-functional requirements

* Offline, no telemetry, no Office dependency; same core on macOS and Windows.
* Responsive UI (all heavy work in background threads/processes); cancellation between safe stages.
* Correctness first: never choose arbitrary matches, never trust formula cells without stored
  results, never deliver unverified or partially written files, never modify inputs.
* Traceability: every value can be traced to sheet/row/column and transformation steps.
* Security: sandboxed expressions, safe XML/ZIP handling, path validation, minimal logging of data.
* Maintainability: layered architecture, typed code (mypy strict on the core), automated tests.

## Milestones and status

| # | Milestone | Status |
|---|---|---|
| M0 | Project skeleton (uv, ruff, mypy, pytest), architecture & plan documents | done |
| M1 | Proof of concept for high-risk areas: DOCX merge with section/header preservation (+ LibreOffice rendering check), docxtpl split runs/loops/escaping, custom markers & MERGEFIELD, formula cached-value detection | done |
| M2 | Vertical slice: YAML profile → xlsx → lookup → context → render → merge → CLI | done |
| M3 | Full core: header detection, aliases/positions, composite & chained relationships, collections, aggregations, conditions, expressions, transforms, validation, cycle detection, traces, .xls/.csv, formula policies, naming, atomic output, verification, audit, modes A–E, cancellation, process pool | done |
| M4 | Desktop UI (all pages), background workers, vi/en catalog, UI tests incl. responsiveness | done |
| M5 | Benchmarks (1k–100k rows, 5/20 lookups, 10 templates, 100/1,000 documents), profiling, measured optimisations | done (see PERFORMANCE_PLAN.md) |
| M6 | PyInstaller onedir spec, Windows CI build + smoke test, release docs | done in repo; first Windows run pending repository push access |

Remaining work is tracked in [PROGRESS.md](PROGRESS.md).

## Acceptance criteria (MVP definition of done) and evidence

| Criterion | Evidence |
|---|---|
| Import a workbook, detect sheets | Excel page; `docgen inspect`; loader tests |
| Select primary sheet, configure relationships | Excel & Relationships pages; profile tests |
| Mapping engine resolves values correctly | `tests/unit/test_mapping.py`, integration tests |
| Import several Word templates, detect placeholders, map them | Templates page; inspector tests |
| Validate required mappings | Static validator + template schema check before generation |
| Individual and merged documents, batch processing | Integration tests modes A–E, process pool |
| UI stays responsive | `test_generate_from_ui_stays_responsive` (heartbeat) |
| Errors clearly reported | Issue model with location + suggestion; Logs page; audit CSV/JSON |
| Customer configurations saved and reused | Profile repository with history; samples A and B |
| Core works on macOS and Windows; Windows executable | CI matrix; `build-windows.yml` + smoke test (needs push access to run) |
| Automated tests pass for supported features | 166 tests passing locally, incl. Open XML SDK conformance |

## Risks

| Risk | Impact | Mitigation |
|---|---|---|
| Merge loses section/header formatting | High | Custom section-preserving merge on top of docxcompose; PoC + LibreOffice check; Word checklist; alternatives documented. |
| Customer templates use unsupported Word features | Medium | Inspector warnings, verification, regression tests per customer, `separate` output fallback. |
| Formulas without cached values | High (wrong documents) | Formula policy, untrusted-cell propagation to records, clear message to recalculate in Excel. |
| Ambiguous lookups | High | Duplicate detection at index time; no arbitrary picks; explicit `match: first`. |
| Large workbooks (100k+ rows) | Medium | Streaming reads, projection, indexes, two-pass planning, bounded process pool; measured. |
| Off-thread Qt object destruction | High (crash) | GC pinned to GUI thread; covered by UI tests. |
| LGPL obligations (Qt, docxtpl) | Medium | Onedir build, docxtpl as plain .py, notices; legal review before commercial release. |
| No Word available for automated tests | Medium | LibreOffice + structural checks automated; manual Windows checklist mandatory. |
