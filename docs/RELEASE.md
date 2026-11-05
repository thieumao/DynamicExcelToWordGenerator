# Release process

1. Update `version` in `pyproject.toml`, `src/docgen/__init__.py` and `packaging/version_info.txt`.
2. Update `docs/PROGRESS.md` and the change log section below.
3. Run locally: `uv run ruff check src tests && uv run mypy && uv run pytest`.
4. Run benchmarks on the reference Windows machine (`uv run python benchmarks/bench.py all`) and add
   the JSON to `benchmarks/results/`; update `docs/PERFORMANCE_PLAN.md` if numbers changed materially.
5. Tag: `git tag v0.1.0 && git push origin v0.1.0` → *Build Windows* workflow builds, tests,
   smoke-tests, builds and tests the installer; download `DocGen-Setup-<version>-x64.exe` (for users)
   and `DocGen-windows-x64.zip` (for IT) from the run.
6. Manual acceptance on Windows 10 and 11 with the checklist in `docs/TEST_STRATEGY.md`, including at
   least one real (anonymised) customer profile.
7. Code-sign the executables and the installer, add `THIRD_PARTY_NOTICES.txt` (see
   `docs/LICENSES.md`), publish internally with the SHA-256 of the files. Give users the installer and
   `docs/HUONG_DAN_SU_DUNG.md` (part A is written for non-IT users).

Profiles are forward-compatible within `schema_version: 1`. When the schema changes, add a migration
in `src/docgen/domain/profile/io.py` (`MIGRATIONS`) and a test that loads an old profile.

## Change log

### 0.1.0 (unreleased) – first release, construction companies first
* For non-IT users: Windows installer, home page with three actions, "Tạo hồ sơ nhanh" (Excel +
  Word files → automatic matching → documents), simple mode by default, plain Vietnamese wording and
  Vietnamese messages for every error and warning.
* Construction example: acceptance records (phiếu yêu cầu + biên bản nghiệm thu) with project details
  entered once, people looked up by code, one file per work item and one combined file to print.
* Existing Word forms: blanks after labels, content controls (text, date, drop-down, check box,
  photo frame) and legacy form fields (text, check box, drop-down) are found and filled.
* Engine: profile system, Excel readers (.xlsx/.xlsm/.xls/.csv), general-information sheets, lookups
  (composite, chained, without keys), mapping engine, Word rendering/merging, verification, audit.
* CLI and CI (Windows/macOS/Linux, Open XML SDK conformance, installer test).
