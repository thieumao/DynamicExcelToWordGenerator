"""Defensive checks for untrusted OOXML/ZIP inputs (zip bombs, oversized files)."""

from __future__ import annotations

import zipfile
from dataclasses import dataclass
from pathlib import Path


class UnsafeFileError(ValueError):
    """The input looks malicious or exceeds configured resource limits."""


@dataclass(frozen=True, slots=True)
class ZipLimits:
    max_file_bytes: int = 500 * 1024 * 1024
    max_total_uncompressed: int = 2 * 1024 * 1024 * 1024
    max_entries: int = 50_000
    max_ratio: float = 200.0
    ratio_min_member_bytes: int = 20 * 1024 * 1024


DEFAULT_LIMITS = ZipLimits()


def check_zip(path: Path, limits: ZipLimits = DEFAULT_LIMITS) -> None:
    """Raise :class:`UnsafeFileError` if *path* is not a sane ZIP container."""
    size = path.stat().st_size
    if size > limits.max_file_bytes:
        raise UnsafeFileError(
            f"{path.name}: file is {size / 1e6:.0f} MB, above the {limits.max_file_bytes / 1e6:.0f} MB limit"
        )
    try:
        with zipfile.ZipFile(path) as zf:
            infos = zf.infolist()
    except zipfile.BadZipFile as exc:
        raise UnsafeFileError(f"{path.name}: not a valid Office Open XML (ZIP) file: {exc}") from exc
    if len(infos) > limits.max_entries:
        raise UnsafeFileError(f"{path.name}: too many ZIP entries ({len(infos)})")
    total = 0
    for info in infos:
        if info.flag_bits & 0x1:
            raise UnsafeFileError(f"{path.name}: encrypted content is not supported ({info.filename})")
        total += info.file_size
        if total > limits.max_total_uncompressed:
            raise UnsafeFileError(
                f"{path.name}: uncompressed size exceeds {limits.max_total_uncompressed / 1e9:.1f} GB"
            )
        if info.file_size >= limits.ratio_min_member_bytes and info.compress_size > 0:
            ratio = info.file_size / info.compress_size
            if ratio > limits.max_ratio:
                raise UnsafeFileError(
                    f"{path.name}: suspicious compression ratio {ratio:.0f}:1 in {info.filename}"
                )


def is_within(path: Path, base: Path) -> bool:
    """True if *path* resolves inside *base* (prevents ../ traversal)."""
    try:
        path.resolve().relative_to(base.resolve())
    except ValueError:
        return False
    return True
