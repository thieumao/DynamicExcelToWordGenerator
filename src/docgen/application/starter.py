"""Ready-made examples users can try, open and copy (construction acceptance records first)."""

from __future__ import annotations

import shutil
from pathlib import Path

from docgen.domain.profile.io import load_profile
from docgen.infrastructure.storage.profile_repo import ProfileRepository
from docgen.testing.construction import (
    KIT_ID,
    MATERIALS_KIT_ID,
    MATERIALS_PROFILE_FILE,
    write_construction_kit,
)

EXCEL_NAME = "HoSoNghiemThu.xlsx"
KITS = ((KIT_ID, "profile.yaml"), (MATERIALS_KIT_ID, MATERIALS_PROFILE_FILE))
"""Profiles of the example, all using the same workbook (work items, incoming materials)."""


def install_construction_kit(repo: ProfileRepository, folder: Path) -> tuple[str, Path]:
    """Put the example in *folder* (Excel + Word forms the user can open) and register its profiles.

    Files the user already has there are kept: the example is only written when something is
    missing, so edits to the sample Excel file are never lost. Returns (profile id of the work-item
    records, Excel file).
    """
    excel = folder / "data" / EXCEL_NAME
    templates = folder / "templates"
    complete = (
        excel.exists() and all((folder / name).exists() for _, name in KITS) and any(templates.glob("*.docx"))
    )
    if not complete:
        keep = excel.read_bytes() if excel.exists() else None
        write_construction_kit(folder)
        if keep is not None:
            excel.write_bytes(keep)
    for profile_id, name in KITS:
        if repo.exists(profile_id):
            continue
        profile = load_profile(folder / name)
        target = repo.directory(profile_id) / "templates"
        target.mkdir(parents=True, exist_ok=True)
        for template in profile.templates:
            source = folder / template.file
            shutil.copy2(source, target / source.name)
        repo.save(profile)
    return KIT_ID, excel
