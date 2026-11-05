# Progress record

Last update: 2026-11-05 (version 0.1.0, branch `main`).

## Completed

- [x] Requirements analysis, architecture, plan, risk assessment (docs/*)
- [x] PoC: section/header-preserving DOCX merge (incl. LibreOffice rendering check), docxtpl
      behaviour, custom markers (split runs), MERGEFIELD conversion, formula cached-value detection
- [x] Profile schema v1 (Pydantic), YAML 1.2 loader, JSON Schema export, migrations hook,
      static validator (references, chains, to-many, reserved names, transforms, expressions, cycles)
- [x] Excel: .xlsx/.xlsm (openpyxl read-only), .xls (xlrd), .csv and CSV folders; sheet matching;
      header detection (multi-row, merged); column binding by header/alias/position; type
      conversion; formula policies with streamed scan; ZIP safety; input snapshot + hash
- [x] Mapping engine: lookup indexes, composite/chained relationships, cardinalities, match/missing/
      duplicate policies, coalesce, conditions, sandboxed expressions, transforms (incl. Vietnamese
      number words, locale formats), collections (nested), aggregations, group fields, validation
      rules, dependency ordering, traces, untrusted-cell propagation
- [x] Word: inspector (paths, loops, conditions, headers/footers), renderer (sandbox, autoescape,
      locale finalize, images, missing-placeholder policies), merger (inherit/preserve, breaks,
      page numbering, update fields), verifier
- [x] Generation service: modes A–E, two-pass planning, bounded recycled process pool, cancellation,
      fail-fast/collect, naming + sanitisation, atomic finalisation, audit JSON + CSV
- [x] CLI (inspect, templates scan, profile validate/schema, preview, generate, samples)
- [x] Desktop UI (10 pages), background workers, GUI-thread GC, vi/en catalogs
- [x] Profile repository (history, trash, import/export bundles), SQLite job history, settings
- [x] Tests: 283 automated (unit, integration, PoC, regression incl. Office conformance, UI, architecture)
- [x] Benchmarks with measured results; two measured optimisations
- [x] PyInstaller onedir spec (GUI + CLI), smoke test (validated with a Linux build),
      GitHub Actions CI matrix and Windows build workflow
- [x] Sample customers A and B; English and Vietnamese documentation
- [x] Microsoft Open XML SDK conformance validation of all outputs (Office 2010/2019/365: 0 errors),
      in tests, CI and the packaging smoke test
- [x] Templates re-saved by LibreOffice: found and fixed silently unfilled «Name» merge fields;
      style-specific leftover-marker verification
- [x] Existing customer Word forms: detection and filling of blanks after labels, empty label/value
      table cells, content controls (text/date/drop-down/check box/picture) and legacy form fields
      (text/check box/drop-down); only
      bound inputs are filled, unbound ones stay untouched; Templates page lists every input with
      its surrounding text and a field picker; "Add Word files…" on the Generate page; sample
      customer C; `templates scan` lists inputs
- [x] Every error/warning message (87 kinds) and its hint is available in Vietnamese with its
      details; a test fails when a new message has no Vietnamese text or misses a parameter;
      the CSV report and the Templates/Mapping pages show the Vietnamese text
- [x] Forms re-saved by LibreOffice keep every input detectable and fillable (regression test)
- [x] GitHub Actions: CI matrix (Windows/macOS/Linux) and the Windows build workflow run on `main`

## Remaining / next steps

- [ ] Download and test the `DocGen-windows-x64` artifact on Windows 10 and 11.
- [ ] Try the input-spot detection on real customer forms; tune `blank_min` defaults if needed.
- [ ] Manual acceptance in Microsoft Word with the checklist in TEST_STRATEGY.md (automated Open XML
      SDK validation passes; opening in Word itself still needs a Windows machine with Office).
- [ ] Real (anonymised) customer workbooks/templates → per-customer regression tests.
- [ ] Benchmarks on the reference Windows machine; confirm targets.
- [ ] Code signing, THIRD_PARTY_NOTICES.txt, legal review of LGPL obligations, installer (Inno Setup).
- [ ] UI polish: drag-and-drop ordering, editing nested collections/aggregations inside collections
      in forms (YAML editor covers them today), per-field "insert path" picker in expression editor.
- [ ] Optional features (isolated, not core): PDF/visual preview via LibreOffice or Word,
      formula recalculation via LibreOffice, "skeleton" master-template merge mode.
- [ ] Profile schema migrations when v2 is introduced.
