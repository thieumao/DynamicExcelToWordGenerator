# Dependency licenses

Runtime dependencies shipped in the Windows build (verify with `uv tree` / `pip-licenses` for the
exact versions of a release):

| Package | License | Notes |
|---|---|---|
| PySide6-Essentials, shiboken6 (Qt 6) | LGPL-3.0 (also GPL/commercial) | Shipped as separate DLLs/.pyd in the onedir folder → user-replaceable. Include the LGPL text and Qt notices. |
| docxtpl | LGPL-2.1 | Collected as plain `.py` files (`module_collection_mode={"docxtpl": "py"}` in the spec) so it can be replaced. |
| python-docx | MIT | |
| docxcompose | MIT | Depends on Babel (BSD-3-Clause). |
| openpyxl, et-xmlfile | MIT | |
| xlrd | BSD | |
| lxml | BSD-3-Clause | Bundles libxml2/libxslt (MIT). |
| Jinja2, MarkupSafe | BSD-3-Clause | |
| pydantic, pydantic-core, annotated-types, typing-extensions, typing-inspection | MIT / PSF | |
| PyYAML | MIT | |
| platformdirs | MIT | |
| defusedxml | PSF-2.0 | |
| Python runtime | PSF-2.0 | Embedded by PyInstaller. |
| PyInstaller bootloader | GPL-2.0 with bootloader exception | The exception allows distributing the generated executables under any license. |

Development-only tools (pytest, pytest-qt, ruff, mypy, psutil, xlwt, PyInstaller itself) are not
shipped.

**Third-party notices:** the Windows build writes `THIRD_PARTY_NOTICES.txt` into the program
folder (`packaging/third_party_notices.py`). It lists every shipped Python distribution – the runtime
dependencies of `docgen[gui]`, resolved from the build environment – with version, license and the
full license texts found in its metadata, plus the LGPL statement (replaceable files, where to get
the source). The smoke test fails if the file is missing; the installer ships it, and Settings has a
button to open it.

**Still needed before commercial distribution:** a legal review of the LGPL obligations (offer of
source/relinking for Qt and docxtpl, no static linking, no prohibition of reverse engineering for
debugging LGPL modifications) and of the notices file's wording.
