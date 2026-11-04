# Architecture

DocGen is a desktop application (PySide6) with a UI-independent core that is also exposed through a
CLI. All processing is local and offline; Microsoft Office is not required.

## Layers

```
┌──────────────────────────────── ui (PySide6) ────────────────────────────────┐
│ pages: dashboard · profiles · excel · relationships · mapping · templates ·  │
│        preview · generate · logs · settings                                  │
│ workers.py (QThreadPool) · state.py · gc_guard.py · widgets                  │
└───────────────┬──────────────────────────────────────────────┬───────────────┘
                │                                              │
┌───────────────▼──────────── application (use cases) ─────────▼───────────────┐
│ generation/service.py  DocumentGenerationService (pipeline, executor, audit)  │
│ generation/planner.py  pass 1: map + validate + units (modes A–E)            │
│ generation/worker.py   render + merge + verify one unit (process-safe)       │
│ templates.py  preview.py  workbook.py                                        │
└───────┬───────────────────────────────────────────────────────┬──────────────┘
        │                                                       │
┌───────▼───────────── domain (pure Python) ──────┐   ┌─────────▼──── infrastructure ─────────────┐
│ profile/  models (Pydantic), validator, io      │   │ excel/  sources (openpyxl/xlrd/csv),       │
│ mapping/  engine, lookup, transforms,           │   │         header_detect, loader              │
│           expressions (sandboxed Jinja), graph, │   │ word/   markers, inspector, renderer       │
│           context, vietnamese                   │   │         (docxtpl), merger (docxcompose),   │
│ validation/ engine                              │   │         verifier                           │
│ data/  table (SheetTable), values               │   │ fs/     atomic finalisation, zip safety     │
│ diagnostics.py  naming.py                       │   │ storage/ profile repo, SQLite, paths       │
└─────────────────────────────────────────────────┘   └────────────────────────────────────────────┘
cli/main.py — same services as the UI (used by tests, CI smoke tests and automation)
```

Rules (enforced by `tests/test_architecture.py`):

* `domain`, `application`, `infrastructure` and `cli` never import PySide6.
* `domain` never imports `application` or `infrastructure`.
* No `eval`/`exec` anywhere; expressions run in Jinja's immutable sandbox.

## Processing pipeline

```
Excel input
 → snapshot copy + SHA-256           (original never opened for writing; stable if Excel keeps it open)
 → ZIP safety check                  (size, entry count, compression ratio, encryption)
 → workbook inspection               (sheet matching by name/alias/regex/index; header detection)
 → profile validation                (schema, references, expression syntax, transforms, cycles)
 → template preparation              (marker conversion on a copy, placeholder inspection,
                                      binding resolution, unknown placeholders → errors before work)
 → data extraction                   (only declared columns; type conversion; NFC text)
 → formula policy                    (streamed <f> scan; cells without cached results are untrusted)
 → lookup index build                (one hash index per relationship; duplicates reported)
 → PASS 1: map + validate + plan     (all records; contexts discarded; units formed)
 → PASS 2 per unit:
      context → template contexts (bindings, context_root, include_if)
      → render each template (docxtpl, sandbox, autoescape) → merge (docxcompose + section logic)
      → verify (ZIP, rels, reopen, leftover markers) → atomic finalise (temp → os.replace)
 → audit JSON + per-unit CSV report
 → cleanup of temporary folders
```

Fatal problems (missing sheets/columns, bad configuration, unknown placeholders, unreadable
templates, output folder not writable) stop the job before any document is written. Record-level
problems (missing references, ambiguous keys, untrusted cells, required empty fields) mark the
affected unit as failed; with `error_policy: collect` the other units are still generated, with
`fail_fast` nothing is generated.

## Main components

| Component | Module | Responsibility |
|---|---|---|
| ExcelReader / WorkbookInspector | `infrastructure/excel/sources.py`, `application/workbook.py` | Read .xlsx/.xlsm (openpyxl read-only), .xls (xlrd), .csv/CSV folders; list sheets; previews |
| SheetSchemaDetector | `infrastructure/excel/header_detect.py`, `loader.py` | Header row scoring, multi-row headers, column binding by header/alias/position |
| LookupEngine | `domain/mapping/lookup.py` | Hash indexes, composite keys, key normalisation, match/missing/duplicate policies |
| MappingEngine | `domain/mapping/engine.py` | Fields, chained paths, coalesce, conditions, expressions, collections, aggregations, traces |
| TransformationEngine | `domain/mapping/transforms.py` | Registry shared by YAML pipelines and Jinja filters |
| ValidationEngine | `domain/validation/engine.py`, `domain/profile/validator.py` | Data rules; static profile validation incl. cycle detection |
| ContextBuilder | `domain/mapping/context.py` | Nested contexts, context schema for placeholder validation |
| WordTemplateInspector | `infrastructure/word/inspector.py` | Placeholder/loop/condition paths from docx parts |
| WordRenderer | `infrastructure/word/renderer.py` | docxtpl rendering in the sandbox, images, missing-placeholder policy |
| WordMergeEngine | `infrastructure/word/merger.py` | docxcompose with `inherit`/`preserve` section modes |
| DocumentGenerationService / JobManager | `application/generation/service.py` | Orchestration, executor, cancellation, finalisation, audit |
| ProfileRepository | `infrastructure/storage/profile_repo.py` | YAML profiles + templates folder, history, trash, import/export bundles |
| ExportService | `application/generation/audit.py`, `infrastructure/fs/atomic.py` | Audit/report files, safe output naming |

## Concurrency model

* **UI thread** only builds widgets and receives signals. Every Excel/Word operation runs in a
  `QThreadPool` worker (`ui/workers.py`); progress is throttled (≥100 ms between updates) and
  issues are capped so the event loop is never flooded.
* **Cancellation** uses a `CancellationToken` checked between safe stages (every 200 records in
  pass 1, between units in pass 2). In-flight units finish or are discarded; temporary files are
  deleted, so cancelled jobs never leave partial `.docx` files.
* **Process pool** (`ProcessPoolExecutor`, `spawn` start method on every OS) renders units in
  parallel when a job has ≥20 runnable units and more than one worker is configured. Mapping stays
  in the parent (it needs the dataset); workers receive plain picklable contexts. At most
  `2 × workers` units are in flight, which bounds memory. Default workers: `min(4, CPUs − 1)`.
* **Garbage collection** is pinned to the GUI thread (`ui/gc_guard.py`). Without it Python can run
  cyclic GC inside a worker thread and destroy Qt widgets off-thread, which crashes the process (this
  was reproduced in the UI tests and fixed).

## Data model

* `Profile` (Pydantic, `schema_version: 1`) – see [DATA_MAPPING_DESIGN.md](DATA_MAPPING_DESIGN.md).
* `SheetTable` – projected rows as tuples + original Excel row numbers + untrusted-cell map.
* `RecordResult` – nested context dict + `FieldTrace` per field (source cells, steps) + issues.
* `Issue` – code, severity, category, message, params, sheet/row/column/field/template/placeholder/
  rule/unit, suggestion. Messages are localised by code (`i18n/vi.json`).
* `JobResult` – units with status/outputs/issues, stage timings, input hashes → audit report.

## Security

* Offline only; no telemetry; no network calls (the expression sandbox exposes no I/O functions).
* No `eval`/`exec`; expressions and templates run in `ImmutableSandboxedEnvironment` with
  whitelisted filters/functions; mapping keys win over Python attributes.
* Profiles are parsed with a `SafeLoader` subclass (YAML 1.2 booleans).
* OOXML inputs: ZIP bomb checks before opening; lxml parsers with `resolve_entities=False`,
  `no_network=True`; `defusedxml` installed so openpyxl uses it.
* Image paths from Excel must resolve inside the configured base folder; size and type limited.
* Output names are sanitised (Windows reserved names, invalid characters, length) and values cannot
  create sub-folders (`/` and `\` in values are replaced).
* Log files contain job metadata only (ids, counts, durations, file names), never cell values.
* Original Excel and Word files are only read; templates are copied into the profile on import.
