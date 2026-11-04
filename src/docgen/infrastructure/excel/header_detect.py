"""Header row detection.

Customers put headers on different rows (titles, logos, blank lines above).
Each candidate row in the scan window is scored by: how many cells match the
profile's expected headers/aliases, how many non-empty text cells it has, how
unique those labels are, and whether the following row looks like data. The
best row wins (ties -> earliest row). Users can always override manually.
"""

from __future__ import annotations

from collections.abc import Iterable, Sequence
from dataclasses import dataclass, field
from typing import Any

from docgen.domain.data.values import is_empty, normalize_label


@dataclass(frozen=True, slots=True)
class HeaderCandidate:
    row: int
    score: float
    expected_hits: int
    non_empty: int


@dataclass(slots=True)
class HeaderDetection:
    row: int | None
    confidence: float
    candidates: list[HeaderCandidate] = field(default_factory=list)


def _is_text(value: Any) -> bool:
    return isinstance(value, str) and not is_empty(value)


def detect_header_row(
    rows: Sequence[Sequence[Any]],
    *,
    first_row: int = 1,
    expected: Iterable[str] = (),
    casefold: bool = True,
    accents: bool = False,
) -> HeaderDetection:
    """Pick the most header-like row from *rows* (which start at Excel row *first_row*)."""
    expected_norm = {normalize_label(e, casefold=casefold, accents=accents) for e in expected if e}
    candidates: list[HeaderCandidate] = []
    for i, row in enumerate(rows):
        cells = [c for c in row if not is_empty(c)]
        if not cells:
            continue
        texts = [c for c in cells if _is_text(c)]
        text_ratio = len(texts) / len(cells)
        if text_ratio < 0.5:
            continue
        labels = [normalize_label(t, casefold=casefold, accents=accents) for t in texts]
        uniqueness = len(set(labels)) / len(labels) if labels else 0
        hits = len(expected_norm.intersection(labels))
        next_row = rows[i + 1] if i + 1 < len(rows) else ()
        next_non_empty = sum(1 for c in next_row if not is_empty(c))
        data_like = 1.0 if next_non_empty >= max(1, len(cells) // 2) else 0.0
        if len(cells) == 1 and not hits:
            # a single text cell is usually a title, not a header
            score = 0.1
        else:
            score = hits * 10 + len(texts) * text_ratio * uniqueness + data_like
        candidates.append(HeaderCandidate(first_row + i, score, hits, len(cells)))
    if not candidates:
        return HeaderDetection(None, 0.0, [])
    best = max(candidates, key=lambda c: (c.score, -c.row))
    if expected_norm:
        confidence = min(1.0, best.expected_hits / max(1, len(expected_norm)))
    else:
        # without expectations: how clearly the best row beats the runner-up
        others = sorted((c.score for c in candidates if c is not best), reverse=True)
        runner_up = others[0] if others else 0.0
        confidence = best.score / (best.score + runner_up) if best.score > 0 else 0.0
    ranked = sorted(candidates, key=lambda c: (-c.score, c.row))[:5]
    return HeaderDetection(best.row, round(confidence, 3), ranked)


def combine_header_rows(
    rows: Sequence[Sequence[Any]], *, joiner: str = " ", fill_merged: bool = True
) -> list[str]:
    """Combine multi-row headers. Merged cells appear as a value followed by Nones;
    with *fill_merged* the value is forward-filled across the merged span on all but
    the last header row (e.g. 'Thông tin | Họ tên' -> 'Thông tin Họ tên')."""
    width = max((len(r) for r in rows), default=0)
    filled: list[list[Any]] = []
    for depth, row in enumerate(rows):
        values = list(row) + [None] * (width - len(row))
        if fill_merged and depth < len(rows) - 1:
            last: Any = None
            for i, v in enumerate(values):
                if is_empty(v):
                    values[i] = last
                else:
                    last = v
        filled.append(values)
    headers = []
    for col in range(width):
        parts: list[str] = []
        for row in filled:
            v = row[col]
            if not is_empty(v):
                text = str(v).strip()
                if not parts or parts[-1] != text:
                    parts.append(text)
        headers.append(joiner.join(parts))
    return headers
