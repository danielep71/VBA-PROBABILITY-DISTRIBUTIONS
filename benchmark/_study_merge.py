"""
Reference-only merge of a regenerated study grid into the committed grid (#159).

A study generator writes a fresh grid with every observation column blank,
because the observations come from Excel. Running it over the committed grid
would discard that evidence. This helper is the safe path instead: it takes the
rows the fixed generator emits and changes the `reference` cell of the
committed grid, and nothing else.

It refuses, writing nothing, when:

  * the generator and the committed grid do not hold the same rows: a key is
    missing, extra or duplicated on either side;
  * any column other than the reference and the observation columns differs
    for a matched row (an argument token, a regime, a claim);
  * a reference to be changed is not a finite number.

After the patch it asserts that every observation cell, the row order and every
other column are byte-for-byte what they were. Rows can be held back by a
predicate, for a row that a binding outside the study protects; a held row
keeps its committed reference and is listed in the report.

Every change is recorded in study_reference_rebase.csv (grid, row, old and new
reference, relative change), so the old values survive in the tree and not only
in Git history. Report-only unless write=True.
"""
import csv
import io
import math
import os

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
RECORD = os.path.join(HERE, "study_reference_rebase.csv")
RECORD_FIELDS = ["grid", "row", "function", "arguments", "regime", "old_reference",
                 "new_reference", "relative_change", "issue"]


class MergeError(RuntimeError):
    pass


def _read(path, delimiter):
    with open(path, newline="", encoding="utf-8") as fh:
        text = fh.read()
    reader = csv.DictReader(io.StringIO(text), delimiter=delimiter)
    return list(reader.fieldnames or []), list(reader), text


def _render(fields, rows, delimiter):
    buf = io.StringIO()
    w = csv.DictWriter(buf, fieldnames=fields, delimiter=delimiter, lineterminator="\n")
    w.writeheader()
    w.writerows(rows)
    return buf.getvalue()


def relative_change(old, new):
    with mp.workdps(60):
        o, n = mp.mpf(old), mp.mpf(new)
        if n == 0:
            return "0" if o == 0 else "inf"
        return mp.nstr(abs(o - n) / abs(n), 3)


def merge_references(grid_path, generated, key, observation_columns, *,
                     delimiter=",", hold=None, write=False, issue="#159",
                     describe=None, record=RECORD):
    """Patch references of grid_path from generated rows. Returns the changes.

    key(row) identifies a row; it must be unique on both sides.
    observation_columns are the Excel columns: never compared, never written.
    hold(row) -> reason or None keeps a row's committed reference.
    describe(row) -> (function, arguments, regime) for the record."""
    fields, committed, original = _read(grid_path, delimiter)
    rel = os.path.relpath(grid_path, HERE).replace(os.sep, "/")
    if _render(fields, committed, delimiter) != original:
        raise MergeError(f"{rel}: the committed file does not round-trip through "
                         "the CSV writer; refusing to rewrite it")

    def index(rows, side):
        out = {}
        for r in rows:
            k = key(r)
            if k in out:
                raise MergeError(f"{rel}: duplicate key on the {side} side: {k}")
            out[k] = r
        return out

    have, want = index(committed, "committed"), index(generated, "generator")
    missing, extra = set(have) - set(want), set(want) - set(have)
    if missing or extra:
        raise MergeError(f"{rel}: generator and grid disagree on the row set "
                         f"({len(missing)} committed row(s) not generated, "
                         f"{len(extra)} generated row(s) not committed); first: "
                         f"{sorted(missing or extra)[:2]}")
    compare = [f for f in fields if f != "reference" and f not in observation_columns]
    for k, cur in have.items():
        new = want[k]
        diff = [f for f in compare if (cur.get(f) or "") != (new.get(f) or "")]
        if diff:
            raise MergeError(f"{rel}: row {k} differs in {diff}; a reference merge "
                             "changes only the reference")

    changes, held = [], []
    for i, cur in enumerate(committed, 2):
        new_ref = want[key(cur)]["reference"]
        if new_ref == cur["reference"]:
            continue
        try:
            if not math.isfinite(float(new_ref)):
                raise ValueError
        except ValueError:
            raise MergeError(f"{rel}: generated reference {new_ref!r} is not a finite number")
        reason = hold(cur) if hold else None
        fn, args, regime = describe(cur) if describe else ("", "", "")
        entry = {"grid": rel, "row": str(i), "function": fn, "arguments": args,
                 "regime": regime, "old_reference": cur["reference"],
                 "new_reference": new_ref,
                 "relative_change": relative_change(cur["reference"], new_ref),
                 "issue": issue}
        if reason:
            held.append((entry, reason))
        else:
            changes.append(entry)

    if not write:
        return changes, held

    before = [{f: r.get(f) for f in fields} for r in committed]
    patch = {(c["row"]): c["new_reference"] for c in changes}
    for i, cur in enumerate(committed, 2):
        if str(i) in patch:
            cur["reference"] = patch[str(i)]
    for b, a in zip(before, committed):
        for f in fields:
            if f != "reference" and b[f] != a.get(f):
                raise MergeError(f"{rel}: column {f} changed during the merge")
    assert len(before) == len(committed)
    if changes:
        with open(grid_path, "w", newline="", encoding="utf-8") as fh:
            fh.write(_render(fields, committed, delimiter))
        _record(record, rel, changes)
    return changes, held


def _record(path, rel, changes):
    """Add this grid's changes to the committed record, keeping earlier ones."""
    rows = []
    if os.path.exists(path):
        with open(path, newline="", encoding="utf-8") as fh:
            rows = list(csv.DictReader(fh))
    done = {(r["grid"], r["row"], r["old_reference"], r["new_reference"]) for r in rows}
    for c in changes:
        if (c["grid"], c["row"], c["old_reference"], c["new_reference"]) not in done:
            rows.append(c)
    rows.sort(key=lambda r: (r["grid"], int(r["row"])))
    with open(path, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=RECORD_FIELDS, lineterminator="\n")
        w.writeheader()
        w.writerows(rows)


def print_report(label, changes, held, write):
    print(f"{label}: {len(changes)} reference(s) "
          f"{'changed' if write else 'would change'}, {len(held)} held")
    worst = max((float(c["relative_change"]) for c in changes), default=0.0)
    if changes:
        print(f"  largest relative change {worst:.3g}")
    for entry, reason in held:
        print(f"  HELD {entry['function']} ({entry['arguments']}): {reason}")
    if not write:
        print("  report only; pass --merge-references --write to patch the committed grid")


def study_key(r):
    """Function, argument tokens and regime: the identity of a study row."""
    return (r.get("function", ""), r.get("arg1", ""), r.get("arg2", ""),
            r.get("arg3", ""), r.get("arg4", "") or "", r.get("regime", "") or "",
            r.get("evidence_set", "") or "")


def study_describe(r):
    args = ";".join(r.get(f"arg{i}", "") for i in range(1, 5) if r.get(f"arg{i}"))
    return r.get("function", ""), args, r.get("regime", "") or ""
