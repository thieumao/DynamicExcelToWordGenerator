# DocGen – Dynamic Excel-to-Word Document Generator

Offline desktop application (Windows-first, developed on macOS) that reads customer Excel workbooks
with many sheets, resolves IDs/codes through lookup sheets, and fills Word templates – as individual
documents or merged documents – driven entirely by **customer profiles** (YAML), not code.

Tiếng Việt: xem [README.vi.md](README.vi.md).

```
Excel (.xlsx/.xlsm/.xls/.csv) ──► profile: sheets · relationships · fields · transforms · rules
                                   │
                                   ▼
                       normalised, traceable data context
                                   │
          Word templates (Jinja / <<markers>> / MERGEFIELD) ──► render ──► merge ──► verify ──► .docx
```

## Highlights

* Any sheet names, header positions (auto-detected or fixed), column aliases or positions.
* Relationships with composite keys, multi-level chains, one-to-many collections, aggregations,
  conditional mapping, expressions (sandboxed), Vietnamese number-to-words, locale formatting.
* Explicit policies for missing references and duplicate keys – never an arbitrary pick.
* Formula cells without stored results are detected and never used silently.
* Works with the customer's **existing Word forms**: blanks after labels (`Họ và tên: ......`),
  empty label/value table cells, content controls (incl. check boxes, drop-downs and photo frames)
  and legacy form fields are detected and – once
  bound to a field – filled; unbound inputs (signatures, office-use boxes) stay untouched.
* Multiple templates per document; merge keeps each template's orientation, margins and
  headers/footers; repeating tables, conditional sections, images.
* Generation modes: single, batch, grouped, merged, batch merged; safe file naming, no overwrites,
  atomic writes, output verification, audit report (JSON + CSV).
* Desktop UI (PySide6) in Vietnamese and English with a visual mapping designer, previews with
  per-value traceability, progress and cancel; CLI for automation.

## Quick start (developer)

```bash
uv sync --extra gui
uv run docgen samples samples                      # (re)create sample customers A, B and C
uv run docgen-gui                                  # desktop app
uv run docgen generate --profile samples/customer_a/profile.yaml \
  --excel samples/customer_a/data/HR.xlsx --out output/customer_a
uv run pytest                                      # QT_QPA_PLATFORM=offscreen on headless Linux
```

Windows users receive a ready-to-run folder (`DocGen.exe`, no Python needed) built by
`.github/workflows/build-windows.yml` – see [docs/BUILD_WINDOWS.md](docs/BUILD_WINDOWS.md).

## Documentation

| Document | Content |
|---|---|
| [docs/IMPLEMENTATION_PLAN.md](docs/IMPLEMENTATION_PLAN.md) | Requirements, milestones, acceptance criteria, risks |
| [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) | Layers, pipeline, components, concurrency, security |
| [docs/TECH_STACK.md](docs/TECH_STACK.md) | Libraries, versions, trade-offs |
| [docs/DATA_MAPPING_DESIGN.md](docs/DATA_MAPPING_DESIGN.md) | Profile schema reference: sheets, relationships, fields, expressions, transforms |
| [docs/WORD_PROCESSING_DESIGN.md](docs/WORD_PROCESSING_DESIGN.md) | Templates, markers, rendering, merging PoC results, limitations |
| [docs/PERFORMANCE_PLAN.md](docs/PERFORMANCE_PLAN.md) | Design measures and measured benchmarks |
| [docs/TEST_STRATEGY.md](docs/TEST_STRATEGY.md) | Test suites, manual Windows/Word checklist, untested areas |
| [docs/HUONG_DAN_SU_DUNG.md](docs/HUONG_DAN_SU_DUNG.md) | Hướng dẫn sử dụng & soạn template (tiếng Việt) |
| [docs/DEV_SETUP_MACOS.md](docs/DEV_SETUP_MACOS.md), [docs/BUILD_WINDOWS.md](docs/BUILD_WINDOWS.md), [docs/RELEASE.md](docs/RELEASE.md) | Development, packaging, release |
| [docs/LICENSES.md](docs/LICENSES.md) | Dependency licenses (LGPL notes) |
| [docs/PROGRESS.md](docs/PROGRESS.md) | Completed and remaining work |

## Project layout

```
src/docgen/domain          pure logic: profile schema/validator, mapping engine, transforms, expressions
src/docgen/application     use cases: generation service/planner/worker, preview, templates, workbook
src/docgen/infrastructure  Excel readers, Word inspector/renderer/merger/verifier, storage, file safety
src/docgen/ui              PySide6 desktop application
src/docgen/cli             command-line interface
src/docgen/testing         synthetic fixture builders and sample customers
samples/                   customer_a (Jinja), customer_b (<<markers>>, MERGEFIELD), customer_c (existing form)
tests/                     unit, integration, PoC, regression, UI, architecture tests
benchmarks/                reproducible performance measurements
packaging/                 PyInstaller spec, entry points, smoke test
```
