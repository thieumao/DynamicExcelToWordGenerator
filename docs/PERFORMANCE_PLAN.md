# Performance plan and measured results

Correctness and document integrity come first; throughput is optimised only where measurements show a
bottleneck. All numbers below were **measured**, with the scripts in `benchmarks/`; raw JSON is in
`benchmarks/results/`.

## Design measures

| Requirement | Implementation |
|---|---|
| Never block the UI thread | All Excel/Word work in `QThreadPool` workers; throttled progress (≤10 updates/s); capped issue streaming. |
| CPU-heavy work in processes | Rendering/merging in a `spawn` `ProcessPoolExecutor` for jobs with ≥20 units; mapping stays in the parent. |
| Bounded concurrency/memory | Default workers `min(4, CPUs − 1)`; at most `2 × workers` units in flight; worker processes recycled every 50 units (`max_tasks_per_child`). |
| Cancellation | Token checked every 200 records (pass 1) and between units (pass 2); temporary files removed. |
| Load only what is needed | openpyxl read-only streaming; only declared sheets and columns kept (projection). |
| No repeated parsing | One load per job; input snapshot; formula scan only where `<f>` tags exist (byte-level pre-check). |
| Indexes instead of scans | One hash index per relationship and job. |
| Incremental batches | Pass 1 keeps only row indices and sort/group keys; pass 2 maps, renders and writes one unit at a time; no document is kept in memory. |
| Temp files | Per-job temp folder on the output volume (atomic `os.replace`), always deleted. |
| Logs | Ring-buffered issue lists in the UI; log files rotate at 2 MB × 3. |

## Reference machine for the numbers below

The development container, **not** the target Windows PC:

* Linux 6.18 (container), Python 3.12.3
* Intel Xeon @ 2.10 GHz, **4 logical CPUs**, 16.9 GB RAM
* Date: 2026-11-04, DocGen 0.1.0, openpyxl 3.1.5, docxtpl 0.20.2, docxcompose 2.2.0

Expect different absolute numbers on Windows (different allocator, file system, antivirus scanning
of new files). Re-run `uv run python benchmarks/bench.py all` and
`uv run python benchmarks/ui_responsiveness.py` on the customer's reference Windows machine and add
the JSON to `benchmarks/results/` before committing to SLAs.

## Synthetic workloads

* Primary sheet: N rows × (4 value columns + K foreign keys); K lookup sheets × 2,000 rows × 3
  columns; every lookup is resolved for every row (`fields` = 5 + K, incl. number-to-words and
  currency formatting). Workbooks get a `<dimension>` element like files saved by Excel.
* Templates: heading, 5 lookup values, amount/words/date paragraph, 2×3 table, header with
  placeholder; merged in `preserve` section mode; every output verified.

## Results (measured)

### Reading + indexing + mapping (pass 1, dry run)

| Scenario | Rows | Lookup sheets | Total | read Excel | index | map | Peak RSS |
|---|---:|---:|---:|---:|---:|---:|---:|
| R1 | 1,000 | 5 | **0.71 s** | 0.55 s | 0.02 s | 0.09 s | 61 MB |
| R2 | 10,000 | 5 | **2.69 s** | 1.59 s | 0.02 s | 1.03 s | 75 MB |
| R4 | 10,000 | 20 | **7.36 s** | 4.59 s | 0.12 s | 2.59 s | 107 MB |
| R3 | 100,000 | 5 | **23.4 s** | 12.8 s | 0.02 s | 10.5 s | 537 MB |
| R3s | 100,000 | 5 (formula check skipped) | 21.7 s | 11.3 s | 0.05 s | 10.4 s | 538 MB |
| R5 | 100,000 | 20 | **60.7 s** | 32.3 s | 0.10 s | 28.2 s | 635 MB |

Mapping costs ≈ 100 µs per record with 10 fields / 5 lookups and ≈ 280 µs with 25 fields /
20 lookups. Reading is dominated by openpyxl's XML parsing (pure Python).

### Rendering + merging + verification

| Scenario | Documents | Templates per document | Workers | Total | Per document | Peak RSS (all processes) |
|---|---:|---:|---:|---:|---:|---:|
| G1 | 100 | 10 | 1 process | **52.6 s** | 0.53 s | 381 MB |
| G2 | 100 | 10 | 3 processes | **18.8 s** | 0.19 s | 1,086 MB* |
| G3 | 1,000 | 3 | 3 processes | **59.2 s** | 0.059 s | 726 MB |

\* G2 was measured before worker recycling was introduced (see below); G1/G3 rows are after.

### UI responsiveness

`benchmarks/ui_responsiveness.py --employees 100 --workers 2` (sample customer A, 3 templates
merged per document, 10 ms heartbeat on the GUI thread while generating):

| Documents | Duration | Max event-loop gap | p95 gap | Median gap |
|---:|---:|---:|---:|---:|
| 100 | 7.5 s | **31.6 ms** | 10.9 ms | 10.5 ms |

`tests/ui/test_ui.py::test_generate_from_ui_stays_responsive` asserts a maximum gap below 0.5 s in CI.

## Optimisations made because of measurements

1. **Formula check** – the first implementation re-opened the workbook with openpyxl
   (`data_only=False`) for every sheet. Profiling R1 showed 4.8 s of 7.0 s spent there
   (re-parsing all sheets per call). Replaced by a single streamed lxml scan per sheet plus a
   byte-level `<f>` pre-check that skips formula-free sheets entirely. R1 load: 7.0 s → 1.8 s
   (under cProfile); 100k rows now pay < 2 s for the check (R3 vs R3s).
2. **Worker memory** – a 1,000-document run grew to 1.59 GB because native memory used while
   merging (lxml/docxcompose; Python objects verified stable with `tracemalloc`/`gc`) is not fully
   returned to the OS (~1 MB per merge on glibc). Large jobs now always render in worker processes
   recycled every 50 units: G3 peak 1,594 MB → **726 MB**, time 54.4 s → 59.2 s (+9 %).
3. **Tracemalloc excluded from timings** (it inflated mapping times 1.5–3×); heap peaks can be
   measured separately with `--heap`.

## Targets (proposed, to be confirmed on the Windows reference machine)

| Target | Measured here | Status |
|---|---|---|
| UI event loop never blocked > 100 ms during generation | 31.6 ms max | met |
| 10,000 rows × 5 lookups validated/mapped < 5 s | 2.7 s | met |
| 100,000 rows × 20 lookups < 2 min, < 1 GB | 61 s, 635 MB | met |
| Merged document (3 templates) ≤ 0.2 s per document with 3 workers | 0.06 s | met |
| 1,000 documents < 5 min, < 1.5 GB total | 59 s, 726 MB | met |
| Cancellation stops within one document | covered by tests | met |

## Known limits and next steps (only if real data requires it)

* openpyxl parsing is the dominant cost for very large workbooks (~0.13 ms per row with 9 columns).
  Options if real customers exceed this: parallel sheet loading in processes, or a faster reader
  (e.g. `python-calamine`, Rust-based) behind the existing `WorkbookSource` interface – to be
  benchmarked for exact type/date fidelity before adoption.
* Mapping is single-threaded Python; for > 500k rows consider chunked mapping in worker processes.
* Pass 1 + pass 2 map every record twice (by design, to keep memory flat); mapping is cheap compared
  with rendering.
* Windows Defender scanning freshly written .docx files can add latency per file; measure on target.
