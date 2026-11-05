"""Write THIRD_PARTY_NOTICES.txt for the packaged application.

Lists every Python distribution the application ships (the runtime dependencies of `docgen` with
the `gui` extra, resolved recursively in the build environment) with its version, license and the
full license texts found in its metadata. Run after PyInstaller:

    uv run python packaging/third_party_notices.py dist/DocGen/THIRD_PARTY_NOTICES.txt
"""

from __future__ import annotations

import re
import sys
from importlib import metadata
from pathlib import Path

from packaging.requirements import Requirement

ROOT_PACKAGE = "docgen"
EXTRAS = {"gui"}
LICENSE_FILE = re.compile(r"(^|/)(licen[cs]e|copying|notice|authors)[^/]*$", re.IGNORECASE)

HEADER = """THIRD-PARTY SOFTWARE NOTICES – DocGen

DocGen includes the open-source components listed below. Each is distributed under its own
license; the full texts follow. Components under the GNU Lesser General Public License (LGPL),
notably Qt for Python (PySide6, shiboken6) and docxtpl, are shipped as separate, replaceable files
in the program folder (DLL/.pyd files and plain .py files); you may replace them with modified
versions. The corresponding source code is available from the projects' official sources
(https://code.qt.io, https://pypi.org/project/PySide6-Essentials/, https://pypi.org/project/docxtpl/)
or on request from the distributor of this copy of DocGen.

The Python runtime (Python Software Foundation License) is included as well.
"""


def _norm(name: str) -> str:
    return re.sub(r"[-_.]+", "-", name).lower()


def runtime_distributions() -> list[metadata.Distribution]:
    """`docgen[gui]` and everything it needs at run time, in alphabetical order."""
    seen: dict[str, metadata.Distribution] = {}
    queue: list[tuple[str, set[str]]] = [(ROOT_PACKAGE, set(EXTRAS))]
    while queue:
        name, extras = queue.pop()
        try:
            dist = metadata.distribution(name)
        except metadata.PackageNotFoundError:
            continue
        key = _norm(dist.metadata["Name"])
        if key in seen:
            continue
        seen[key] = dist
        for spec in dist.requires or []:
            req = Requirement(spec)
            envs = [{"extra": e} for e in extras] or [{"extra": ""}]
            if req.marker is None or any(req.marker.evaluate(env) for env in envs):
                queue.append((req.name, set(req.extras)))
    seen.pop(_norm(ROOT_PACKAGE), None)
    return sorted(seen.values(), key=lambda d: _norm(d.metadata["Name"]))


def license_of(dist: metadata.Distribution) -> str:
    meta = dist.metadata
    expression = meta.get("License-Expression")
    if expression:
        return expression
    classifiers = [
        c.split("::")[-1].strip() for c in meta.get_all("Classifier") or [] if c.startswith("License ::")
    ]
    text = (meta.get("License") or "").strip()
    if classifiers:
        return "; ".join(classifiers)
    return text.splitlines()[0][:120] if text else "see license text below"


def license_texts(dist: metadata.Distribution) -> list[tuple[str, str]]:
    texts = []
    for file in dist.files or []:
        if LICENSE_FILE.search(str(file)):
            try:
                texts.append(
                    (
                        str(file),
                        Path(str(dist.locate_file(file))).read_text(encoding="utf-8", errors="replace"),
                    )
                )
            except OSError:
                continue
    if not texts:
        body = (dist.metadata.get("License") or "").strip()
        if len(body) > 200:  # the whole license text is sometimes in the metadata
            texts.append(("METADATA License", body))
    return texts


def render() -> str:
    dists = runtime_distributions()
    parts = [HEADER, "", "Components", "=" * 10, ""]
    for dist in dists:
        meta = dist.metadata
        parts.append(f"- {meta['Name']} {dist.version} – {license_of(dist)}")
    for dist in dists:
        meta = dist.metadata
        parts += ["", "", "=" * 78, f"{meta['Name']} {dist.version}", f"License: {license_of(dist)}"]
        home = meta.get("Home-page") or next(
            (u.split(",", 1)[1].strip() for u in meta.get_all("Project-URL") or [] if "," in u), ""
        )
        if home:
            parts.append(f"Project: {home}")
        parts.append("=" * 78)
        texts = license_texts(dist)
        if not texts:
            parts.append("(No license file in the package metadata; see the project page.)")
        for name, text in texts:
            parts += ["", f"--- {name} ---", text.strip()]
    return "\n".join(parts) + "\n"


def main() -> int:
    target = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("THIRD_PARTY_NOTICES.txt")
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(render(), encoding="utf-8")
    print(f"wrote {target} ({target.stat().st_size // 1024} KB)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
