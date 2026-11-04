# PyInstaller spec: one-folder (onedir) build with two executables sharing one runtime.
#   DocGen.exe      windowed desktop application
#   docgen-cli.exe  console tool (automation, smoke tests)
# Build on the target OS (Windows builds must run on Windows):
#   uv run pyinstaller packaging/docgen.spec --noconfirm
# ruff: noqa
from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

ROOT = Path(SPECPATH).parent
SRC = ROOT / "src"

datas = []
for package in ("docgen", "docx", "docxcompose", "docxtpl", "openpyxl", "babel"):
    datas += collect_data_files(package)

hiddenimports = (
    collect_submodules("docgen")
    + collect_submodules("lxml")
    + ["xlrd", "defusedxml", "platformdirs", "yaml", "jinja2.ext"]
)

# Qt modules the application does not use (keeps the folder smaller)
excludes = [
    "tkinter", "PySide6.QtQml", "PySide6.QtQuick", "PySide6.QtQuickWidgets", "PySide6.QtNetwork",
    "PySide6.QtOpenGL", "PySide6.QtOpenGLWidgets", "PySide6.QtSql", "PySide6.QtTest", "PySide6.QtPdf",
    "PySide6.QtPdfWidgets", "PySide6.QtSvg", "PySide6.QtDBus", "PySide6.QtConcurrent", "PySide6.QtXml",
    "PySide6.QtDesigner", "PySide6.QtHelp", "PySide6.QtUiTools", "PySide6.QtPrintSupport",
    "pytest", "pytestqt", "mypy", "ruff", "xlwt",
]

# LGPL: keep docxtpl as plain .py files next to the executable so it can be replaced by the user.
# python-docx builds paths like "docx/parts/../templates/x.xml"; its package folders must exist on
# disk (pyz+py), otherwise non-Windows file systems cannot resolve the ".." segment.
common = dict(pathex=[str(SRC)], datas=datas, hiddenimports=hiddenimports, excludes=excludes,
              noarchive=False, module_collection_mode={"docxtpl": "py", "docx": "pyz+py"})

gui = Analysis([str(ROOT / "packaging" / "entry_gui.py")], **common)
cli = Analysis([str(ROOT / "packaging" / "entry_cli.py")], **common)

gui_pyz = PYZ(gui.pure)
cli_pyz = PYZ(cli.pure)

version_file = ROOT / "packaging" / "version_info.txt"

gui_exe = EXE(
    gui_pyz, gui.scripts, [], exclude_binaries=True, name="DocGen", console=False,
    upx=False, version=str(version_file) if version_file.exists() else None,
)
cli_exe = EXE(
    cli_pyz, cli.scripts, [], exclude_binaries=True, name="docgen-cli", console=True,
    upx=False, version=str(version_file) if version_file.exists() else None,
)

coll = COLLECT(
    gui_exe, gui.binaries, gui.datas,
    cli_exe, cli.binaries, cli.datas,
    strip=False, upx=False, name="DocGen",
)
