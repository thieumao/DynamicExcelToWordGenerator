"""Smoke-test a built distribution folder: CLI + sample generation + GUI start/stop.

Usage: python packaging/smoke_test.py dist/DocGen
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


def run(cmd: list[str], **kwargs) -> subprocess.CompletedProcess[str]:
    print("$", " ".join(cmd), flush=True)
    result = subprocess.run(
        cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=600, **kwargs
    )
    print(result.stdout[-3000:])
    if result.stderr.strip() and result.returncode == 0:
        print(result.stderr[-1000:])
    if result.returncode != 0:
        print(result.stderr[-3000:])
        raise SystemExit(f"command failed with exit code {result.returncode}")
    return result


def validate_outputs(folders: list[Path], work: Path) -> None:
    """Validate documents produced by the *frozen* app with Microsoft's Open XML SDK (if dotnet exists)."""
    if shutil.which("dotnet") is None:
        if os.environ.get("DOCGEN_REQUIRE_OOXML") == "1":
            raise SystemExit("dotnet required for OOXML validation")
        print("dotnet not found: skipping OOXML validation")
        return
    project = Path(__file__).resolve().parents[1] / "tools" / "ooxml-validator"
    out = work / "ooxmlval"
    run(["dotnet", "build", str(project), "-c", "Release", "-o", str(out)])
    for fmt in ("Office2010", "Office2019"):
        run(["dotnet", str(out / "OoxmlValidator.dll"), "--format", fmt, *map(str, folders)])


def main() -> int:
    dist = Path(sys.argv[1])
    exe_suffix = ".exe" if os.name == "nt" else ""
    cli = str(dist / f"docgen-cli{exe_suffix}")
    gui = str(dist / f"DocGen{exe_suffix}")
    work = Path(tempfile.mkdtemp(prefix="docgen-smoke-"))
    run([cli, "--version"])
    run([cli, "samples", str(work / "samples"), "--employees", "25"])
    a = work / "samples" / "customer_a"
    run([cli, "profile", "validate", str(a / "profile.yaml"), "--excel", str(a / "data" / "HR.xlsx")])
    # 25 units >= pool threshold: exercises the process pool inside the frozen executable
    run(
        [
            cli,
            "generate",
            "--profile",
            str(a / "profile.yaml"),
            "--excel",
            str(a / "data" / "HR.xlsx"),
            "--out",
            str(work / "out_a"),
            "--workers",
            "2",
        ]
    )
    docs = list((work / "out_a").glob("*.docx"))
    assert len(docs) == 25, f"expected 25 documents, got {len(docs)}"
    b = work / "samples" / "customer_b"
    run(
        [
            cli,
            "generate",
            "--profile",
            str(b / "profile.yaml"),
            "--excel",
            str(b / "data" / "HopDong.xlsx"),
            "--out",
            str(work / "out_b"),
        ]
    )
    assert len(list((work / "out_b").rglob("*.docx"))) == 9
    validate_outputs([work / "out_a", work / "out_b"], work)
    env = dict(
        os.environ,
        QT_QPA_PLATFORM=os.environ.get("QT_QPA_PLATFORM", "offscreen"),
        DOCGEN_HOME=str(work / "home"),
    )
    run([gui, "--smoke-test"], env=env)
    print("SMOKE TEST PASSED")
    return 0


if __name__ == "__main__":
    sys.exit(main())
