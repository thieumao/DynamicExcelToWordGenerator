"""Safe output finalisation.

Documents are written to a temporary file in the destination directory and
moved into place with :func:`os.replace` (atomic on the same volume), so a
crash or cancellation never leaves a half-written final .docx behind.
"""

from __future__ import annotations

import os
import shutil
import tempfile
import uuid
from contextlib import suppress
from dataclasses import dataclass
from pathlib import Path

from docgen.domain.profile.models import ConflictPolicy


class OutputConflictError(FileExistsError):
    pass


class OutputLockedError(PermissionError):
    pass


@dataclass(frozen=True, slots=True)
class Finalized:
    path: Path | None
    skipped: bool = False
    renamed_from: Path | None = None


def temp_path_for(target: Path) -> Path:
    target.parent.mkdir(parents=True, exist_ok=True)
    return target.parent / f".~docgen-{uuid.uuid4().hex[:12]}{target.suffix}.tmp"


def _with_suffix_number(target: Path, n: int) -> Path:
    return target.with_name(f"{target.stem} ({n}){target.suffix}")


def finalize(
    temp: Path, target: Path, policy: ConflictPolicy, *, reserved: set[Path] | None = None
) -> Finalized:
    """Move *temp* to *target* according to the conflict *policy*.

    *reserved* holds targets already produced by the current job so two units
    rendering to the same name never overwrite each other, even with 'overwrite'.
    """
    reserved = reserved if reserved is not None else set()
    target.parent.mkdir(parents=True, exist_ok=True)
    final = target
    exists = target.exists() or target in reserved
    if exists:
        if policy is ConflictPolicy.SKIP:
            with suppress(OSError):
                temp.unlink()
            return Finalized(None, skipped=True)
        if policy is ConflictPolicy.ERROR:
            with suppress(OSError):
                temp.unlink()
            raise OutputConflictError(f"Output already exists: {target}")
        if policy is ConflictPolicy.SUFFIX or target in reserved:
            n = 1
            while final.exists() or final in reserved:
                n += 1
                final = _with_suffix_number(target, n)
    try:
        os.replace(temp, final)
    except PermissionError as exc:
        if policy is ConflictPolicy.OVERWRITE and final == target:
            # typically: file is open in Word on Windows -> write next to it instead
            n = 2
            alt = _with_suffix_number(target, n)
            while alt.exists() or alt in reserved:
                n += 1
                alt = _with_suffix_number(target, n)
            try:
                os.replace(temp, alt)
            except OSError:
                with suppress(OSError):
                    temp.unlink()
                raise OutputLockedError(
                    f"{target.name} is locked (open in Word?) and no alternative name could be used"
                ) from exc
            reserved.add(alt)
            return Finalized(alt, renamed_from=target)
        with suppress(OSError):
            temp.unlink()
        raise OutputLockedError(f"Cannot write {final}: {exc}") from exc
    reserved.add(final)
    return Finalized(final, renamed_from=target if final != target else None)


def check_writable_dir(directory: Path) -> str | None:
    """Return a problem description if *directory* cannot receive output files."""
    try:
        directory.mkdir(parents=True, exist_ok=True)
        fd, name = tempfile.mkstemp(prefix=".~docgen-probe-", dir=directory)
        os.close(fd)
        Path(name).unlink()
    except OSError as exc:
        return f"Output folder is not writable: {directory} ({exc})"
    return None


def safe_rmtree(path: Path) -> None:
    shutil.rmtree(path, ignore_errors=True)
