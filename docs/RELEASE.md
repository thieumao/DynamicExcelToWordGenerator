# Release process

1. Update `version` in `pyproject.toml`, `src/docgen/__init__.py` and `packaging/version_info.txt`.
2. Update `docs/PROGRESS.md` and the change log section below.
3. Run locally: `uv run ruff check src tests && uv run mypy && uv run pytest`.
4. Run benchmarks on the reference Windows machine (`uv run python benchmarks/bench.py all`) and add
   the JSON to `benchmarks/results/`; update `docs/PERFORMANCE_PLAN.md` if numbers changed materially.
5. Tag: `git tag v0.1.0 && git push origin v0.1.0` → *Build Windows* workflow builds, tests and
   smoke-tests; download `DocGen-windows-x64.zip` from the run.
6. Manual acceptance on Windows 10 and 11 with the checklist in `docs/TEST_STRATEGY.md`, including at
   least one real (anonymised) customer profile.
7. Code-sign the executables, add `THIRD_PARTY_NOTICES.txt` (see `docs/LICENSES.md`), zip, publish
   internally with the SHA-256 of the archive.

Profiles are forward-compatible within `schema_version: 1`. When the schema changes, add a migration
in `src/docgen/domain/profile/io.py` (`MIGRATIONS`) and a test that loads an old profile.

## Change log

### 0.1.0 (unreleased)
* First MVP: profile system, Excel readers, mapping engine, Word rendering/merging, desktop UI,
  CLI, Windows packaging workflow.
