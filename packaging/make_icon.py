"""Write packaging/docgen.ico (used by PyInstaller and Inno Setup). Run before building:

uv run python packaging/make_icon.py
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtGui import QGuiApplication

from docgen.ui.icon import ico_bytes


def main() -> int:
    target = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).with_name("docgen.ico")
    app = QGuiApplication([])  # QPainter needs an application object
    target.write_bytes(ico_bytes())
    print(f"wrote {target} ({target.stat().st_size} bytes)")
    del app
    return 0


if __name__ == "__main__":
    sys.exit(main())
