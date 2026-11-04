# Building the Windows application

The Windows build is a **PyInstaller one-folder (onedir)** distribution containing:

```
DocGen/
  DocGen.exe        desktop application (no console window)
  docgen-cli.exe    command-line tool (automation, support, smoke tests)
  _internal/        Python runtime, Qt DLLs, libraries, data files (i18n, docx templates)
```

Users unzip the folder anywhere (e.g. `C:\Program Files\DocGen` or their profile) and start
`DocGen.exe`. Python is not required. Profiles, settings and history are stored per user in
`%LOCALAPPDATA%\DocGen\DocGen` (platformdirs), never inside the program folder.

## Automated build (recommended)

`.github/workflows/build-windows.yml` runs on `windows-latest` for pushes to `main`/`claude/**`,
tags `v*` and manual dispatch:

1. `uv sync --extra gui --group build` (Python 3.12)
2. full test suite (UI tests offscreen)
3. `pyinstaller packaging/docgen.spec --noconfirm --clean`
4. `packaging/smoke_test.py dist/DocGen` – runs the **built** executables: version, sample
   generation, profile validation, a 25-document batch through the process pool inside the frozen
   app (`multiprocessing.freeze_support()` path), customer B, and `DocGen.exe --smoke-test`
5. uploads `DocGen-windows-x64.zip` as a workflow artifact (30 days)

## Manual build on a Windows machine

```powershell
winget install --id=astral-sh.uv -e
git clone <repo-url>; cd DynamicExcelToWordGenerator
uv sync --extra gui --group build --python 3.12
$env:QT_QPA_PLATFORM="offscreen"; uv run pytest -q; Remove-Item Env:QT_QPA_PLATFORM
uv run pyinstaller packaging/docgen.spec --noconfirm --clean
uv run python packaging/smoke_test.py dist/DocGen
```

## Notes and known considerations

* **Long paths:** Windows limits paths to 260 characters unless "LongPathsEnabled" is set. DocGen
  shortens output names to `output.max_path_length` (default 240) with a hash suffix; prefer short
  output folders such as `D:\HoSo`.
* **Locked files:** if an output file is open in Word, DocGen writes a numbered copy next to it (with
  `on_conflict: overwrite`) or reports the lock. Input workbooks open in Excel can be read (DocGen
  copies them to a temporary snapshot first).
* **Antivirus/SmartScreen:** unsigned executables may trigger SmartScreen warnings. Sign `DocGen.exe`
  and `docgen-cli.exe` with the company's code-signing certificate before distribution (`signtool sign
  /fd SHA256 /tr <timestamp-url> /td SHA256 ...`).
* **Installer:** evaluated separately. Inno Setup (free) is the recommended option: per-user install
  without admin rights, Start-menu shortcut, uninstaller, `.docgenprofile` association later. MSIX
  needs signing and store-like packaging; WiX/MSI suits corporate deployment via GPO/Intune. The
  onedir folder is the input for any of them.
* **Size:** PySide6-Essentials plus unused Qt modules excluded in the spec; expect roughly 150–250 MB
  unpacked (measure on the first CI run).
* **Python version:** builds use 3.12 (stable with PyInstaller 6 and PySide6 6.8–6.12).
