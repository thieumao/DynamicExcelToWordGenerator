"""Reproducible benchmarks with synthetic data.

Usage::

    uv run python benchmarks/bench.py --list
    uv run python benchmarks/bench.py R1 R2 G1           # selected scenarios
    uv run python benchmarks/bench.py all --out benchmarks/results

Each scenario reports wall time per stage (from the job's stage timer), peak
RSS (process + worker children, sampled every 50 ms) and Python heap peak
(tracemalloc, mapping scenarios only). Results are appended as JSON with the
machine description so numbers are never presented without their context.
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import platform
import random
import shutil
import sys
import tempfile
import threading
import time
import tracemalloc
from collections.abc import Callable
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any

import psutil
from docx import Document

from docgen.application.generation.job import GenerationRequest
from docgen.application.generation.service import DocumentGenerationService
from docgen.domain.profile.io import parse_profile_text
from docgen.testing.xlsx_builder import build_xlsx_streaming


@dataclass
class Scenario:
    key: str
    description: str
    rows: int
    lookups: int
    templates: int = 0
    documents: int = 0
    workers: int = 1
    formula_check: bool = True


SCENARIOS = {
    s.key: s
    for s in [
        Scenario("R1", "1,000 rows, 5 lookup sheets (read + index + map)", 1_000, 5),
        Scenario("R2", "10,000 rows, 5 lookup sheets", 10_000, 5),
        Scenario("R3", "100,000 rows, 5 lookup sheets", 100_000, 5),
        Scenario(
            "R3s", "100,000 rows, 5 lookup sheets, formula check skipped", 100_000, 5, formula_check=False
        ),
        Scenario("R4", "10,000 rows, 20 lookup sheets", 10_000, 20),
        Scenario("R5", "100,000 rows, 20 lookup sheets", 100_000, 20),
        Scenario("G1", "100 merged documents x 10 templates, 1 worker", 100, 5, 10, 100, 1),
        Scenario("G2", "100 merged documents x 10 templates, 3 worker processes", 100, 5, 10, 100, 3),
        Scenario("G3", "1,000 merged documents x 3 templates, 3 worker processes", 1_000, 5, 3, 1_000, 3),
    ]
}

LOOKUP_ROWS = 2_000


@dataclass
class Result:
    scenario: str
    description: str
    status: str
    seconds_total: float
    stage_seconds: dict[str, float]
    peak_rss_mb: float
    python_heap_peak_mb: float | None
    documents: int
    records: int
    notes: list[str] = field(default_factory=list)


def machine() -> dict[str, Any]:
    cpu = platform.processor() or ""
    try:
        for line in Path("/proc/cpuinfo").read_text().splitlines():
            if line.startswith("model name"):
                cpu = line.split(":", 1)[1].strip()
                break
    except OSError:
        pass
    return {
        "os": f"{platform.system()} {platform.release()}",
        "python": sys.version.split()[0],
        "cpu": cpu,
        "logical_cpus": os.cpu_count(),
        "ram_gb": round(psutil.virtual_memory().total / 1e9, 1),
        "date": dt.datetime.now().isoformat(timespec="seconds"),
    }


class PeakRss:
    def __init__(self) -> None:
        self.peak = 0
        self._stop = threading.Event()
        self._thread = threading.Thread(target=self._run, daemon=True)

    def _run(self) -> None:
        proc = psutil.Process()
        while not self._stop.is_set():
            try:
                total = proc.memory_info().rss
                for child in proc.children(recursive=True):
                    try:
                        total += child.memory_info().rss
                    except psutil.Error:
                        pass
                self.peak = max(self.peak, total)
            except psutil.Error:
                pass
            time.sleep(0.05)

    def __enter__(self) -> PeakRss:
        self._thread.start()
        return self

    def __exit__(self, *exc: object) -> None:
        self._stop.set()
        self._thread.join()


def make_workbook(path: Path, rows: int, lookups: int, seed: int = 1) -> None:
    rng = random.Random(seed)

    def main_rows():  # type: ignore[no-untyped-def]
        yield ["ID", "Name", "Amount", "Date", *[f"FK{i}" for i in range(1, lookups + 1)]]
        base = dt.datetime(2020, 1, 1)
        for n in range(1, rows + 1):
            yield [
                f"R{n:06d}",
                f"nguyễn văn {n}",
                rng.randrange(1, 10_000) * 1000,
                base + dt.timedelta(days=n % 1500),
                *[f"K{rng.randrange(1, LOOKUP_ROWS + 1):05d}" for _ in range(lookups)],
            ]

    sheets: dict[str, Any] = {"Main": main_rows()}
    for i in range(1, lookups + 1):
        sheets[f"L{i}"] = [["Code", "Label", "Extra"]] + [
            [f"K{k:05d}", f"Nhãn {i}-{k}", k * i] for k in range(1, LOOKUP_ROWS + 1)
        ]
    build_xlsx_streaming(path, sheets)


def make_profile(lookups: int, templates: int, formula_check: bool) -> str:
    lines = [
        "profile: {id: bench, name: Benchmark}",
        "excel:",
        "  primary: m",
        f"  formula_policy: {'cached_required' if formula_check else 'skip_check'}",
        "  sheets:",
        "    - id: m",
        "      match: {name: Main}",
        "      header: {mode: fixed, row: 1}",
        "      columns:",
        "        - {id: id, header: ID}",
        "        - {id: name, header: Name}",
        "        - {id: amount, header: Amount, type: decimal}",
        "        - {id: date, header: Date, type: date}",
    ]
    lines += [f"        - {{id: fk{i}, header: FK{i}}}" for i in range(1, lookups + 1)]
    for i in range(1, lookups + 1):
        lines += [
            f"    - id: l{i}",
            f"      match: {{name: L{i}}}",
            "      header: {mode: fixed, row: 1}",
            "      columns: [{id: code, header: Code}, {id: label, header: Label}, {id: extra, header: Extra}]",
        ]
    lines.append("relationships:")
    lines += [
        f"  - {{id: r{i}, from: m, to: l{i}, keys: [{{from: fk{i}, to: code}}]}}"
        for i in range(1, lookups + 1)
    ]
    lines += [
        "fields:",
        "  - {name: rec.id, source: id}",
        "  - {name: rec.name, source: name, transforms: [title_case]}",
        "  - {name: rec.amount, source: amount, transforms: [format_currency]}",
        "  - {name: rec.words, source: amount, transforms: [number_to_words_vi]}",
        "  - {name: rec.date, source: date, transforms: [format_date]}",
    ]
    lines += [f"  - {{name: rec.l{i}, path: r{i}.label}}" for i in range(1, lookups + 1)]
    if templates:
        lines.append("templates:")
        lines += [
            f"  - {{id: t{i}, file: t{i}.docx, section: {{mode: preserve}}}}" for i in range(1, templates + 1)
        ]
    lines += [
        "generation: {unit: record, output: merged, verify_output: true}",
        "output: {merged_filename: 'Doc_{rec.id}.docx'}",
    ]
    return "\n".join(lines) + "\n"


def make_template(path: Path, index: int, lookups: int) -> None:
    doc = Document()
    doc.sections[0].header.paragraphs[0].text = f"Mẫu {index} – {{{{ rec.id }}}}"
    doc.add_heading(f"Phần {index}: {{{{ rec.name }}}}", level=1)
    for i in range(1, min(lookups, 5) + 1):
        doc.add_paragraph(f"Thuộc tính {i}: {{{{ rec.l{i} }}}}")
    doc.add_paragraph("Số tiền {{ rec.amount }} ({{ rec.words }}), ngày {{ rec.date }}.")
    table = doc.add_table(rows=2, cols=3)
    table.style = "Table Grid"
    for c, text in enumerate(["Mã", "Tên", "Ngày"]):
        table.cell(0, c).text = text
    for c, text in enumerate(["{{ rec.id }}", "{{ rec.name }}", "{{ rec.date }}"]):
        table.cell(1, c).text = text
    doc.save(str(path))


def run_scenario(s: Scenario, work: Path, *, heap: bool = False) -> Result:
    folder = work / s.key
    shutil.rmtree(folder, ignore_errors=True)
    folder.mkdir(parents=True)
    wb = folder / "data.xlsx"
    t0 = time.perf_counter()
    make_workbook(wb, s.rows, s.lookups)
    build_seconds = time.perf_counter() - t0
    for i in range(1, s.templates + 1):
        make_template(folder / f"t{i}.docx", i, s.lookups)
    profile = parse_profile_text(make_profile(s.lookups, s.templates, s.formula_check))
    service = DocumentGenerationService(pool_threshold=2)
    rows = None
    if s.documents and s.documents < s.rows:
        rows = list(range(2, 2 + s.documents))
    request = GenerationRequest(
        profile, folder, wb, folder / "out", rows=rows, workers=s.workers, dry_run=s.templates == 0
    )
    heap_peak = None
    if heap:
        tracemalloc.start()
    start = time.perf_counter()
    with PeakRss() as rss:
        result = service.run(request)
    total = time.perf_counter() - start
    if heap:
        heap_peak = tracemalloc.get_traced_memory()[1] / 1e6
        tracemalloc.stop()
    notes = [
        f"synthetic workbook written in {build_seconds:.1f}s (not included)",
        f"workbook size {wb.stat().st_size / 1e6:.1f} MB",
    ]
    if heap:
        notes.append("tracemalloc active: timings include its overhead")
    return Result(
        s.key,
        s.description,
        result.status,
        round(total, 2),
        {k: round(v, 2) for k, v in result.stage_seconds.items()},
        round(rss.peak / 1e6, 1),
        None if heap_peak is None else round(heap_peak, 1),
        len(result.documents),
        result.records_selected,
        notes,
    )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("scenarios", nargs="*", default=["R1", "R2", "G1"])
    parser.add_argument("--out", default="benchmarks/results")
    parser.add_argument("--work", default=None, help="working folder (default: temp)")
    parser.add_argument("--list", action="store_true")
    parser.add_argument("--heap", action="store_true", help="also measure Python heap peak (slower)")
    args = parser.parse_args(argv)
    if args.list:
        for s in SCENARIOS.values():
            print(f"{s.key:4} {s.description}")
        return 0
    keys = list(SCENARIOS) if args.scenarios == ["all"] else args.scenarios
    work = Path(args.work) if args.work else Path(tempfile.mkdtemp(prefix="docgen-bench-"))
    info = machine()
    print(json.dumps(info, indent=2))
    results = []
    for key in keys:
        print(f"\n== {key}: {SCENARIOS[key].description}", flush=True)
        r = run_scenario(SCENARIOS[key], work, heap=args.heap)
        results.append(r)
        print(json.dumps(asdict(r), ensure_ascii=False, indent=2), flush=True)
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    target = out / f"bench-{stamp}-{platform.system().lower()}.json"
    target.write_text(
        json.dumps({"machine": info, "results": [asdict(r) for r in results]}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"\nSaved {target}")
    if not args.work:
        shutil.rmtree(work, ignore_errors=True)
    return 0


Runner = Callable[..., int]

if __name__ == "__main__":
    sys.exit(main())
