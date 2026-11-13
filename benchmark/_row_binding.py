"""
Bind a study to the grid rows it uses, not to the whole grid file (#302).

A study that hashes a live grid file goes stale whenever anyone appends an
unrelated row, and each promotion then forces a metadata-only regeneration
(#23, #289, #327). A study's real input is a set of rows. row_digest() hashes
exactly those:

  * the rows are selected by function name;
  * each is identified by _grid_safety.canonical_key (IEEE-754 argument bits,
    regime, evidence set), so "0.85" and "0.84999999999999998" are one row;
  * the digest covers, for each selected row in file order, the key and the
    claim, metric, reference, observation and expected_error cells;
  * rows outside the selection do not affect it, wherever they are.

Appending or inserting a row of another function leaves the digest unchanged.
Adding, removing, reordering or changing a selected row, including its
observation, changes it. The order of the selected rows is bound because a
study replays its rows in file order and records them in that order.
Standard library only.
"""
import csv
import hashlib

from _grid_safety import canonical_key

CELLS = ("claim", "metric", "reference", "observed_vba", "expected_error")


def row_digest(path, functions):
    """(row count, SHA-256) of the rows of `path` whose function is in `functions`."""
    wanted = frozenset(functions)
    with open(path, newline="", encoding="utf-8-sig") as f:
        rows = [r for r in csv.DictReader(f) if r.get("function") in wanted]
    lines = []
    for r in rows:
        key = canonical_key(r)
        cells = tuple((r.get(c) or "").strip() for c in CELLS)
        lines.append(repr((key, cells)))
    if len({canonical_key(r) for r in rows}) != len(rows):
        raise ValueError(f"{path}: duplicate canonical key among the bound rows")
    digest = hashlib.sha256("\n".join(lines).encode("utf-8")).hexdigest()
    return len(rows), digest
