"""
Shared safeguards for tools that read or write the accuracy grids (#17).

The committed grids hold Excel observations that cannot be regenerated from
Python. This module is the single definition of the rules every grid tool
applies before it writes anything:

  * the canonical row identity - function, four IEEE-754 binary64 argument
    bit patterns, regime and evidence_set - so "0.85" and
    "0.84999999999999998" are the same row;
  * duplicate canonical identities are an error, never a tie to break;
  * a reference-only output must never land on the committed grid, nor on
    any other file that already carries observations.

reconcile_grid.py, promote_grid_rows.py, migrate_references.py and
check_promotion.py keep their own byte-identical key() for now; unifying the
helpers across the whole generator belongs to #32. test_grid_regeneration.py
pins all of them to canonical_key() so they cannot drift apart silently.
"""
import csv
import os
import struct
from collections import Counter

HERE = os.path.dirname(os.path.abspath(__file__))
COMMITTED_GRID = os.path.join(HERE, "probability_accuracy_grid.csv")
COMMITTED_HOLDOUT = os.path.join(HERE, "holdout", "holdout_grid.csv")

GRID_FIELDS = ["function", "vba_kernel", "claim", "metric",
               "arg1", "arg2", "arg3", "arg4", "reference", "observed_vba",
               "regime", "evidence_set", "expected_error"]


def f64_bits(text):
    """IEEE-754 bit pattern of an argument cell; '' for blank, '?...' if unparsable."""
    t = (text or "").strip()
    if t == "":
        return ""
    try:
        return struct.pack(">d", float(t)).hex()
    except (ValueError, OverflowError):
        return "?" + t


def canonical_key(row):
    return (row["function"], f64_bits(row["arg1"]), f64_bits(row["arg2"]),
            f64_bits(row["arg3"]), f64_bits(row.get("arg4", "")),
            row["regime"], row.get("evidence_set", ""))


def duplicate_keys(rows):
    """{canonical_key: count} for every identity that occurs more than once."""
    return {k: n for k, n in Counter(canonical_key(r) for r in rows).items() if n > 1}


def describe_key(k):
    args = ",".join(a for a in k[1:5] if a)
    return f"{k[0]}  args=({args})  regime={k[5]}  set={k[6]}"


def read_grid(path):
    with open(path, newline="", encoding="utf-8") as fh:
        return list(csv.DictReader(fh))


def observation_count(path):
    """Rows with a non-blank observed_vba cell; 0 if the file or column is absent."""
    if not os.path.exists(path):
        return 0
    try:
        rows = read_grid(path)
    except (OSError, UnicodeDecodeError, csv.Error):
        # Unreadable as a grid: not provably observation-free, so treat it as
        # carrying evidence rather than as safe to overwrite.
        return -1
    return sum(1 for r in rows if (r.get("observed_vba") or "").strip())


def _same_file(a, b):
    if not (os.path.exists(a) and os.path.exists(b)):
        return os.path.realpath(a) == os.path.realpath(b)
    return os.path.samefile(a, b)


def refuse_reference_output(out_path, *authoritative):
    """
    Why `out_path` must not receive reference-only rows, or None if it may.

    Refuses the committed main grid and holdout grid, any path given in
    `authoritative`, and any existing file that already carries observations
    - a copy of the grid under another name is still evidence.
    """
    for grid in (COMMITTED_GRID, COMMITTED_HOLDOUT) + tuple(authoritative):
        if grid and _same_file(out_path, grid):
            return (f"{out_path} is the authoritative grid {grid}; reference-only "
                    f"output would blank its observations")
    n = observation_count(out_path)
    if n < 0:
        return (f"{out_path} exists and cannot be read as a grid, so it cannot be "
                f"shown to be free of observations")
    if n:
        return (f"{out_path} already carries {n} observation(s); reference-only "
                f"output would discard them")
    return None


def write_grid(path, rows, fields=None):
    """Write rows with LF endings, matching the repository's eol=lf policy."""
    fields = fields or GRID_FIELDS
    with open(path, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=fields, lineterminator="\n",
                           extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)
