# Building the Windows application

The Windows build is a **PyInstaller one-folder (onedir)** distribution containing:

```
DocGen/
  DocGen.exe        desktop application (no console window)
  docgen-cli.exe    command-line tool (automation, support, smoke tests)
  _internal/        Python runtime, Qt DLLs, libraries, data files (i18n, docx templates)
```

End users get **`DocGen-Setup-<version>-x64.exe`** (Inno Setup, see below): double-click, *Next*,
done – per-user install without administrator rights, Start-menu entry "DocGen – Tạo hồ sơ Word từ
Excel", optional desktop icon, uninstaller. The same folder is also published as a zip for IT staff
who prefer to copy it. Python is not required. Profiles, settings and history are stored per user in
`%LOCALAPPDATA%\DocGen\DocGen` (platformdirs), never inside the program folder.

## Automated build (recommended)

`.github/workflows/build-windows.yml` runs on `windows-latest` for pushes to `main`/`claude/**`,
tags `v*` and manual dispatch:

1. `uv sync --extra gui --group build` (Python 3.12)
2. full test suite (UI tests offscreen)
3. `pyinstaller packaging/docgen.spec --noconfirm --clean`
4. `packaging/smoke_test.py dist/DocGen` – runs the **built** executables: version, sample
   generation, profile validation, a 25-document batch through the process pool inside the frozen
   app (`multiprocessing.freeze_support()` path), customers B and C, the construction kit (with its
   combined print file), Open XML SDK validation of every output, and `DocGen.exe --smoke-test`
5. zips the folder (`DocGen-windows-x64.zip`)
6. builds the installer: `choco install innosetup`, then `ISCC.exe /DAppVersion=<version>
   packaging/installer.iss`. Installer screens use Inno Setup's unofficial Vietnamese texts when
   they compile; otherwise English.
7. **tests the installer**: silent install (`/VERYSILENT /CURRENTUSER /DIR=…`), the full smoke test
   against the *installed* folder, silent uninstall, and a check that nothing is left behind
8. uploads both artifacts for 30 days: `DocGen-windows-x64` (zip) and `DocGen-Setup-windows-x64`
   (the installer)

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
* **Antivirus/SmartScreen:** the executables and the installer are **not signed yet**, so Windows
  shows "Windows protected your PC" (users click *More info → Run anyway*; the Vietnamese guide says
  so). Sign `DocGen.exe`, `docgen-cli.exe` and the setup with the company's code-signing certificate
  before wide distribution (`signtool sign /fd SHA256 /tr <timestamp-url> /td SHA256 ...`; Inno
  Setup can sign the uninstaller via `SignTool=`).
* **Installer:** `packaging/installer.iss` (Inno Setup 6). Per-user by default
  (`PrivilegesRequired=lowest`), all-users selectable; fixed `AppId` so new versions upgrade in
  place; user data in `%LOCALAPPDATA%` and generated documents are never removed by the uninstaller.
  WiX/MSI remains an option for corporate deployment via GPO/Intune; the onedir folder is the input
  for any of them.
* **Size:** PySide6-Essentials plus unused Qt modules excluded in the spec; expect roughly 150–250 MB
  unpacked (measure on the first CI run).
* **Python version:** builds use 3.12 (stable with PyInstaller 6 and PySide6 6.8–6.12).
