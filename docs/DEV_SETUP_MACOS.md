# Development setup (macOS, also works on Linux/Windows)

Requirements: macOS 12+, [uv](https://docs.astral.sh/uv/) (`brew install uv`), Git. uv installs
Python 3.12 automatically if needed.

```bash
git clone <repo-url> DynamicExcelToWordGenerator
cd DynamicExcelToWordGenerator
uv sync --extra gui            # creates .venv with runtime + dev dependencies
```

Run the desktop app:

```bash
uv run docgen-gui
```

Create the sample customers and try the CLI:

```bash
uv run docgen samples samples
uv run docgen generate --profile samples/customer_a/profile.yaml --excel samples/customer_a/data/HR.xlsx --out output/a
```

Quality checks (same as CI):

```bash
uv run ruff check src tests && uv run ruff format --check src tests
uv run mypy
uv run pytest
```

Notes

* UI tests use pytest-qt; on macOS they run with the native platform plugin. On headless Linux set
  `QT_QPA_PLATFORM=offscreen` (and install `libegl1 libgl1 libxkbcommon0`).
* The optional LibreOffice rendering test runs when `soffice` and `pdftotext` are on `PATH`
  (`brew install --cask libreoffice poppler`), otherwise it is skipped.
* Do not build the Windows executable on macOS: PyInstaller builds for the OS it runs on. Push to
  GitHub and use the *Build Windows* workflow (see BUILD_WINDOWS.md). `uv run pyinstaller
  packaging/docgen.spec` on macOS produces a macOS folder build for local testing only.
* `DOCGEN_HOME=/tmp/docgen-home uv run docgen-gui` keeps profiles/settings of a test session separate
  from your normal ones.
* Paths in profiles are relative to the profile folder and use `/`; never hard-code `/Users/...`.
