# Progress record

Last update: 2026-11-04 (version 0.1.0, branch `claude/trusting-newton-8v5rwv`).

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
- [x] Tests: 166 automated (unit, integration, PoC, regression incl. Office conformance, UI, architecture)
- [x] Benchmarks with measured results; two measured optimisations
- [x] PyInstaller onedir spec (GUI + CLI), smoke test (validated with a Linux build),
      GitHub Actions CI matrix and Windows build workflow
- [x] Sample customers A and B; English and Vietnamese documentation
- [x] Microsoft Open XML SDK conformance validation of all outputs (Office 2010/2019/365: 0 errors),
      in tests, CI and the packaging smoke test
- [x] Templates re-saved by LibreOffice: found and fixed silently unfilled «Name» merge fields;
      style-specific leftover-marker verification

## Remaining / next steps

- [ ] **Push access to GitHub** for this session (push was rejected with 403), then run CI and the
      Windows build workflow; download and test `DocGen-windows-x64.zip` on Windows 10 and 11.
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
