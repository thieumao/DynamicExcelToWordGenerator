# Technology stack and trade-offs

Versions verified against PyPI on 2026-11-04; the lock file (`uv.lock`) pins exact builds.

| Concern | Choice | Version range | License | Why |
|---|---|---|---|---|
| Language | Python | 3.12 (builds), 3.12–3.13 tested | PSF | Mature ecosystem for OOXML; same code on macOS and Windows. |
| Desktop UI | **PySide6-Essentials** (Qt Widgets) | ≥6.8,<6.13 (6.12.0 tested) | LGPL-3.0 | Native-looking, data-heavy forms/tables; the *Essentials* wheel omits WebEngine/3D/etc. (smaller install). |
| Excel .xlsx/.xlsm | openpyxl (read-only, streaming) | 3.1.x | MIT | Pure Python, no Excel needed. Does **not** calculate formulas – see formula policy. |
| Excel .xls | xlrd | 2.0.x | BSD | Only maintained reader for BIFF .xls (results only, no formulas). |
| CSV | stdlib `csv` | – | PSF | Encoding detection (UTF-8 BOM, UTF-16, Windows-1258), delimiter sniffing. |
| Word templates | docxtpl | 0.20.x | LGPL-2.1 | Jinja2 inside .docx incl. table-row/paragraph loops, images, sub-documents. |
| Word manipulation | python-docx, lxml | 1.2.x, ≥5 | MIT, BSD | Part/relationship access for markers, merging, verification. |
| Word merging | docxcompose | 2.2.x | MIT | Copies styles, numbering, images, footnotes; DocGen adds section/header preservation. |
| Expressions | Jinja2 `ImmutableSandboxedEnvironment` | 3.1.x | BSD | Safe, documented syntax identical to template syntax; no `eval`. |
| Configuration | Pydantic v2 + PyYAML (`SafeLoader`, YAML 1.2 booleans) | 2.11+, 6.x | MIT | Typed models, precise validation errors, JSON Schema export. |
| App state | SQLite (stdlib) | – | Public domain | Job history and recent files only; profiles stay human-readable YAML files. |
| Paths | platformdirs | 4.x | MIT | Correct per-user folders on Windows/macOS/Linux. |
| XML hardening | defusedxml | 0.7 | PSF | openpyxl uses it automatically when installed. |
| Packaging | PyInstaller (onedir) | 6.x | GPL-2.0 + bootloader exception | Users do not install Python; onedir starts faster and keeps Qt DLLs replaceable (LGPL). |
| Tooling | uv, pytest, pytest-qt, ruff, mypy (strict), psutil | – | – | Reproducible environments, UI tests, measurement. |

## Decisions and rejected alternatives

* **No pandas / Polars / DuckDB.** Data is extracted once into compact tuples (only declared
  columns) and joined through hash indexes. Benchmarks (see PERFORMANCE_PLAN.md) show mapping
  100,000 rows with 5 lookups in ~10 s and reading dominated by openpyxl's XML parsing, which a
  dataframe library would not remove (it would still need an .xlsx reader). Reconsider only if
  measurements show join/aggregation as the bottleneck for real customer files.
* **No Electron / web stack.** A Python core is required for OOXML anyway; a Qt UI in the same process
  avoids a server, IPC and a second runtime, and works fully offline.
* **No Microsoft Word / COM automation in the core.** Generation must work on machines without Office
  and on macOS/CI. Optional Word-based features (e.g. PDF export, pixel-perfect merge) would be an
  isolated plug-in, never a dependency.
* **No LibreOffice dependency.** It is used only in an optional test (independent rendering check)
  and could be an optional "recalculate formulas"/"PDF preview" feature later.
* **Jinja sandbox for expressions** instead of a custom parser: one syntax for users (templates and
  rules), well-tested sandbox; risk mitigated by immutable environment, whitelisted filters/globals,
  no I/O functions, length limits and mapping-first attribute resolution.
* **YAML for profiles** (JSON also accepted) – customers' consultants can read and diff them; the UI
  edits them without exposing YAML for common tasks.

## Licensing notes (summary – see LICENSES.md)

* PySide6 (LGPL-3.0) and docxtpl (LGPL-2.1) are LGPL. Distributing the onedir build keeps them as
  separate, replaceable files/modules; ship the license texts and a notice on how to obtain/replace
  them. Obtain legal review before commercial distribution.
* All other runtime dependencies are permissive (MIT/BSD/PSF).
