"""Ready-made examples users can try, open and copy (construction acceptance records first)."""

from __future__ import annotations

import shutil
from pathlib import Path

from docgen.domain.profile.io import load_profile
from docgen.infrastructure.storage.profile_repo import ProfileRepository
from docgen.testing.construction import KIT_ID, write_construction_kit

EXCEL_NAME = "HoSoNghiemThu.xlsx"


def install_construction_kit(repo: ProfileRepository, folder: Path) -> tuple[str, Path]:
    """Put the example in *folder* (Excel + Word forms the user can open) and register its profile.

    Files the user already has there are kept: the example is only written when it is missing,
    so edits to the sample Excel file are never lost. Returns (profile id, Excel file).
    """
    excel = folder / "data" / EXCEL_NAME
    templates = folder / "templates"
    if not excel.exists() or not (folder / "profile.yaml").exists() or not any(templates.glob("*.docx")):
        keep = excel.read_bytes() if excel.exists() else None
        write_construction_kit(folder)
        if keep is not None:
            excel.write_bytes(keep)
    if not repo.exists(KIT_ID):
        target = repo.directory(KIT_ID) / "templates"
        target.mkdir(parents=True, exist_ok=True)
        for docx in templates.glob("*.docx"):
            shutil.copy2(docx, target / docx.name)
        repo.save(load_profile(folder / "profile.yaml"))
    return KIT_ID, excel
