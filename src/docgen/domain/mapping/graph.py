"""Deterministic dependency ordering with cycle detection.

Nodes are output paths (fields, collections, aggregations). A node depends on
another when one of its expressions references the other's path (or a parent
object containing it). Evaluation order is a topological sort that falls back
to declaration order for independent nodes, so results never depend on dict
or set iteration order.
"""

from __future__ import annotations

import heapq
from collections.abc import Iterable, Sequence
from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class Node:
    name: str
    refs: frozenset[tuple[str, ...]]


@dataclass(slots=True)
class Ordering:
    order: list[int]
    cycles: list[list[str]]
    edges: dict[int, set[int]]


def resolve_ref(chain: tuple[str, ...], names: Sequence[str]) -> set[int]:
    """Map a dotted reference to node indices.

    ``employee.department.name`` matches a node with that exact name, a node that
    is a prefix of it (``employee.department`` when it is a collection) and all
    nodes below it (when the reference is an object such as ``employee``).
    """
    ref = ".".join(chain)
    hits: set[int] = set()
    for i, name in enumerate(names):
        if name == ref or ref.startswith(name + ".") or name.startswith(ref + "."):
            hits.add(i)
    return hits


def order_nodes(nodes: Sequence[Node]) -> Ordering:
    names = [n.name for n in nodes]
    edges: dict[int, set[int]] = {i: set() for i in range(len(nodes))}  # i depends on edges[i]
    for i, node in enumerate(nodes):
        for chain in node.refs:
            for j in resolve_ref(chain, names):
                if j != i or ".".join(chain) == node.name:
                    edges[i].add(j)
    dependents: dict[int, set[int]] = {i: set() for i in range(len(nodes))}
    indegree = {i: 0 for i in range(len(nodes))}
    for i, deps in edges.items():
        for j in deps:
            if j == i:
                continue
            dependents[j].add(i)
            indegree[i] += 1
    heap = [i for i, d in indegree.items() if d == 0 and i not in edges[i]]
    heapq.heapify(heap)
    order: list[int] = []
    while heap:
        i = heapq.heappop(heap)
        order.append(i)
        for k in sorted(dependents[i]):
            indegree[k] -= 1
            if indegree[k] == 0 and k not in edges[k]:
                heapq.heappush(heap, k)
    remaining = [i for i in range(len(nodes)) if i not in set(order)]
    cycles = _find_cycles(remaining, edges, names) if remaining else []
    return Ordering(order, cycles, edges)


def _find_cycles(
    remaining: Iterable[int], edges: dict[int, set[int]], names: Sequence[str]
) -> list[list[str]]:
    remaining_set = set(remaining)
    cycles: list[list[str]] = []
    seen_cycles: set[frozenset[int]] = set()
    for start in sorted(remaining_set):
        stack: list[int] = []
        on_stack: set[int] = set()

        def dfs(n: int, stack: list[int] = stack, on_stack: set[int] = on_stack) -> bool:
            stack.append(n)
            on_stack.add(n)
            for m in sorted(edges[n] & remaining_set):
                if m in on_stack:
                    cycle = stack[stack.index(m) :]
                    key = frozenset(cycle)
                    if key not in seen_cycles:
                        seen_cycles.add(key)
                        cycles.append([names[c] for c in cycle] + [names[m]])
                    return True
                if dfs(m):
                    return True
            stack.pop()
            on_stack.discard(n)
            return False

        dfs(start)
    return cycles
