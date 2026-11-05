"""The user guide shipped with the application (a copy of docs/HUONG_DAN_SU_DUNG.md)."""

from __future__ import annotations

from importlib import resources

GUIDE = "HUONG_DAN_SU_DUNG.md"


def guide_text() -> str:
    """The Vietnamese user guide (Markdown), or a short note if it was not packaged."""
    try:
        return resources.files(__name__).joinpath(GUIDE).read_text(encoding="utf-8")
    except (FileNotFoundError, OSError):
        return "# Hướng dẫn sử dụng DocGen\n\nKhông tìm thấy file hướng dẫn trong bản cài đặt này."
