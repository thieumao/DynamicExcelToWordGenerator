"""Customer profile repository (folders of YAML + templates, with history).

Layout::

    <profiles>/<profile id>/profile.yaml
    <profiles>/<profile id>/templates/*.docx      (copies; originals never touched)
    <profiles>/<profile id>/history/profile-v3-20241102-101500.yaml
    <profiles>/.trash/<profile id>-<timestamp>/   (deleted profiles are recoverable)

Saves are atomic (temp file + replace) and the previous version is kept in
``history`` so user configuration is never lost.
"""

from __future__ import annotations

import datetime as dt
import io
import os
import shutil
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path

from docgen.domain.profile.io import ProfileLoadError, load_profile, parse_profile_text, profile_to_text
from docgen.domain.profile.models import (
    ColumnSpec,
    ExcelConfig,
    Profile,
    ProfileInfo,
    SheetConfig,
    SheetMatch,
)
from docgen.infrastructure.fs.safety import UnsafeFileError, ZipLimits, check_zip

PROFILE_FILE = "profile.yaml"
HISTORY_KEEP = 30


@dataclass(frozen=True, slots=True)
class ProfileEntry:
    id: str
    name: str
    version: int
    directory: Path
    modified: float
    error: str | None = None


def _stamp() -> str:
    return dt.datetime.now().strftime("%Y%m%d-%H%M%S")


def blank_profile(profile_id: str, name: str) -> Profile:
    sheet = SheetConfig(
        id="main",
        match=SheetMatch(index=1),
        columns=[ColumnSpec(id="col_a", position="A", required=False)],
    )
    return Profile(
        profile=ProfileInfo(id=profile_id, name=name),
        excel=ExcelConfig(primary="main", sheets=[sheet]),
    )


class ProfileRepository:
    def __init__(self, root: Path) -> None:
        self.root = root
        self.root.mkdir(parents=True, exist_ok=True)

    def directory(self, profile_id: str) -> Path:
        return self.root / profile_id

    def list(self) -> list[ProfileEntry]:
        entries = []
        for folder in sorted(p for p in self.root.iterdir() if p.is_dir() and not p.name.startswith(".")):
            path = folder / PROFILE_FILE
            if not path.is_file():
                continue
            try:
                profile = load_profile(path)
                entries.append(
                    ProfileEntry(
                        profile.profile.id,
                        profile.profile.name,
                        profile.profile.version,
                        folder,
                        path.stat().st_mtime,
                    )
                )
            except (ProfileLoadError, OSError) as exc:
                entries.append(
                    ProfileEntry(folder.name, folder.name, 0, folder, path.stat().st_mtime, str(exc))
                )
        return entries

    def exists(self, profile_id: str) -> bool:
        return (self.directory(profile_id) / PROFILE_FILE).is_file()

    def load(self, profile_id: str) -> Profile:
        return load_profile(self.directory(profile_id) / PROFILE_FILE)

    def save(self, profile: Profile, *, bump_version: bool = False) -> Path:
        if bump_version:
            profile.profile.version += 1
        folder = self.directory(profile.profile.id)
        folder.mkdir(parents=True, exist_ok=True)
        target = folder / PROFILE_FILE
        if target.exists():
            history = folder / "history"
            history.mkdir(exist_ok=True)
            try:
                old_version = load_profile(target).profile.version
            except (ProfileLoadError, OSError):
                old_version = 0
            shutil.copy2(target, history / f"profile-v{old_version}-{_stamp()}.yaml")
            backups = sorted(history.glob("profile-*.yaml"), key=lambda p: p.stat().st_mtime)
            for old in backups[:-HISTORY_KEEP]:
                old.unlink(missing_ok=True)
        text = profile_to_text(profile)
        parse_profile_text(text)  # never write something that cannot be read back
        fd, tmp = tempfile.mkstemp(prefix=".profile-", suffix=".yaml", dir=folder)
        with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as fh:
            fh.write(text)
        Path(tmp).replace(target)
        return target

    def create(self, profile: Profile) -> Path:
        if self.exists(profile.profile.id):
            raise FileExistsError(f"Profile '{profile.profile.id}' already exists")
        (self.directory(profile.profile.id) / "templates").mkdir(parents=True, exist_ok=True)
        return self.save(profile)

    def duplicate(self, profile_id: str, new_id: str, new_name: str) -> Profile:
        if self.exists(new_id):
            raise FileExistsError(f"Profile '{new_id}' already exists")
        src = self.directory(profile_id)
        dst = self.directory(new_id)
        shutil.copytree(src, dst, ignore=shutil.ignore_patterns("history"))
        profile = self.load(profile_id)
        profile.profile.id = new_id
        profile.profile.name = new_name
        profile.profile.version = 1
        self.save(profile)
        return profile

    def delete(self, profile_id: str) -> Path:
        """Move the profile folder to .trash (recoverable) and return its new location."""
        trash = self.root / ".trash"
        trash.mkdir(exist_ok=True)
        target = trash / f"{profile_id}-{_stamp()}"
        shutil.move(str(self.directory(profile_id)), str(target))
        return target

    def import_template(self, profile_id: str, source: Path) -> str:
        """Copy a .docx into the profile; return the relative path for TemplateSpec.file."""
        check_zip(source, ZipLimits(max_file_bytes=200 * 1024 * 1024))
        folder = self.directory(profile_id) / "templates"
        folder.mkdir(parents=True, exist_ok=True)
        target = folder / source.name
        n = 2
        while target.exists() and target.read_bytes() != source.read_bytes():
            target = folder / f"{source.stem} ({n}){source.suffix}"
            n += 1
        shutil.copy2(source, target)
        return f"templates/{target.name}"

    # ------------------------------------------------------- import/export
    def export_bundle(self, profile_id: str, destination: Path) -> Path:
        folder = self.directory(profile_id)
        if destination.suffix.lower() in (".yaml", ".yml", ".json"):
            fmt = "json" if destination.suffix.lower() == ".json" else "yaml"
            destination.write_text(profile_to_text(self.load(profile_id), fmt=fmt), encoding="utf-8")
            return destination
        with zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.write(folder / PROFILE_FILE, PROFILE_FILE)
            for tpl in sorted((folder / "templates").glob("*.docx")):
                zf.write(tpl, f"templates/{tpl.name}")
        return destination

    def import_bundle(self, source: Path, *, new_id: str | None = None) -> Profile:
        """Import a profile from .zip (profile + templates) or a plain .yaml/.json file."""
        if source.suffix.lower() == ".zip":
            check_zip(source, ZipLimits(max_file_bytes=500 * 1024 * 1024))
            with zipfile.ZipFile(source) as zf:
                names = zf.namelist()
                if PROFILE_FILE not in names:
                    raise ProfileLoadError("The bundle does not contain profile.yaml", [])
                profile = parse_profile_text(zf.read(PROFILE_FILE).decode("utf-8"))
                templates = {
                    n: zf.read(n)
                    for n in names
                    if n.startswith("templates/")
                    and n.lower().endswith(".docx")
                    and "/" not in n[10:]
                    and ".." not in n
                }
        else:
            fmt = "json" if source.suffix.lower() == ".json" else "yaml"
            profile = parse_profile_text(source.read_text(encoding="utf-8"), fmt=fmt)
            templates = {}
        if new_id:
            profile.profile.id = new_id
        if self.exists(profile.profile.id):
            raise FileExistsError(f"Profile '{profile.profile.id}' already exists")
        folder = self.directory(profile.profile.id)
        (folder / "templates").mkdir(parents=True, exist_ok=True)
        for name, data in templates.items():
            check_zip_bytes(data)
            (folder / name).write_bytes(data)
        self.save(profile)
        return profile


def check_zip_bytes(data: bytes) -> None:
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            if sum(i.file_size for i in zf.infolist()) > 1024 * 1024 * 1024:
                raise UnsafeFileError("template too large when uncompressed")
    except zipfile.BadZipFile as exc:
        raise UnsafeFileError(f"invalid .docx in bundle: {exc}") from exc
