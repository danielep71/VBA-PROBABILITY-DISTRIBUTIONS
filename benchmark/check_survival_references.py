"""
Fail-closed audit of every committed survival reference (#159).

Each survival reference in every grid that carries one is recomputed by a direct
upper-tail route at the exact binary64 value of its argument tokens
(_survival_oracle.py), and must agree with it to the 25-digit floor of the
reference audits (relative difference at most SER_FLOOR = 1E-24). The check
fails when:

  * a reference misses the floor and is not registered as debt;
  * a debt entry no longer matches a failing row: the row was fixed, changed or
    removed, so the entry is stale and must be deleted;
  * the oracle cannot certify a value (its two precisions disagree);
  * a row names a survival function the oracle has no route for;
  * a CSV under benchmark/ carries survival rows but is not in GRIDS, so a new
    grid cannot escape the check by not being listed;
  * a listed grid is missing or no longer carries a survival row.

survival_reference_debt.json registers the rows this audit knows to be below
the floor but that cannot be re-referenced yet, each with the blocker that
prevents it (for example a hash-bound manifest that only a fresh Excel export
may rebind). An entry names the grid, function, argument tokens and the exact
committed reference text, so it covers that one cell and nothing else.

Every row is evaluated: there is no sampling. A full run takes about 100 CPU
seconds with mpmath 1.3.0 and runs in parallel (--jobs, default: every CPU).

    python check_survival_references.py              # the committed grids
    python check_survival_references.py --jobs 1     # serial
    python check_survival_references.py --verbose    # one line per row
"""
import argparse
import csv
import io
import json
import os
import sys
from multiprocessing import Pool

import mpmath as mp
from mpmath.libmp import NoConvergence

import _survival_oracle as oracle

HERE = os.path.dirname(os.path.abspath(__file__))
DEBT = "survival_reference_debt.json"

# Every grid that carries a survival reference, relative to benchmark/.
GRIDS = (
    "probability_accuracy_grid.csv",
    "holdout/holdout_grid.csv",
    "holdout_older/holdout_older_grid.csv",
    "beta_f_unbalanced/beta_f_unbalanced_grid.csv",
    "cdf_large_shape/cdf_large_shape_grid.csv",
    "f_envelope/f_envelope_grid.csv",
    "f_envelope/f_envelope_gap_grid.csv",
    "f_envelope/f_envelope_bothlarge_grid.csv",
    "survival_boundary/survival_boundary_grid.csv",
    "tchi_large_df/tchi_large_df_grid.csv",
    "excel_comparison/excel_comparison_grid.csv",
    "gamma_lower_q/gamma_lower_q_grid.csv",
)


class AuditError(RuntimeError):
    pass


def _read_text(path):
    with open(path, newline="", encoding="utf-8-sig") as fh:
        return fh.read()


def survival_rows(path, rel):
    """(grid, line, function, tokens, reference) for every survival row.

    Comma grids carry function/argN/reference columns. The pipe-delimited
    excel_comparison grid names the call instead, as K_STATS_Name(a,b,c)."""
    text = _read_text(path)
    first = text.split("\n", 1)[0]
    out = []
    if "|" in first and "kstats_call" in first:
        for line, r in enumerate(csv.DictReader(io.StringIO(text), delimiter="|"), 2):
            call = (r.get("kstats_call") or "").strip()
            if not call.startswith("K_STATS_") or "(" not in call:
                continue
            name, args = call[len("K_STATS_"):].split("(", 1)
            if name.endswith("_Survival"):
                toks = tuple(t.strip() for t in args.rstrip(")").split(","))
                out.append((rel, line, name, toks, (r.get("reference") or "").strip()))
        return out
    reader = csv.DictReader(io.StringIO(text))
    fields = reader.fieldnames or []
    if "function" not in fields or "reference" not in fields:
        return out
    for line, r in enumerate(reader, 2):
        fn = (r.get("function") or "").strip()
        if fn.endswith("_Survival") and "Inverse" not in fn:
            toks = tuple((r.get(f"arg{i}") or "").strip() for i in range(1, 5))
            out.append((rel, line, fn, toks, (r.get("reference") or "").strip()))
    return out


def discover(root):
    """Every CSV under root that carries a survival row, relative to root."""
    found = []
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = sorted(d for d in dirnames if not d.startswith((".", "__")))
        for name in sorted(filenames):
            if name.endswith(".csv"):
                path = os.path.join(dirpath, name)
                rel = os.path.relpath(path, root).replace(os.sep, "/")
                try:
                    if survival_rows(path, rel):
                        found.append(rel)
                except (csv.Error, UnicodeDecodeError):
                    continue
    return found


def _evaluate(item):
    grid, line, fn, toks, ref = item
    try:
        value, agree, route = oracle.certify(fn, toks)
        rel = oracle.relative_error(ref, value)
        return item, "ok" if rel <= oracle.SER_FLOOR else "below", rel, value, ""
    except (oracle.OracleError, ValueError, ArithmeticError, NoConvergence) as exc:
        return item, "oracle", None, None, f"{type(exc).__name__}: {exc}"


def load_debt(path):
    if not os.path.exists(path):
        return {}, {}
    with open(path, encoding="utf-8") as fh:
        data = json.load(fh)
    if data.get("schema_version") != 1:
        raise AuditError(f"{os.path.basename(path)}: schema_version must be 1")
    blockers = data.get("blockers") or {}
    entries = {}
    for e in data.get("entries") or []:
        key = (e["grid"], e["function"], tuple(e["arguments"]), e["reference"])
        if key in entries:
            raise AuditError(f"duplicate debt entry: {key}")
        if e.get("blocker") not in blockers:
            raise AuditError(f"debt entry {key} names an undefined blocker "
                             f"{e.get('blocker')!r}")
        entries[key] = e
    return entries, blockers


def _debt_key(item):
    grid, _, fn, toks, ref = item
    args = [t for t in toks if t]
    return (grid, fn, tuple(args), ref)


def run(root=HERE, grids=GRIDS, debt_path=None, jobs=None, verbose=False, out=print):
    """Audit the grids under root. Returns the process exit code."""
    problems = []
    debt_path = debt_path or os.path.join(root, DEBT)
    try:
        debt, blockers = load_debt(debt_path)
    except (AuditError, KeyError, ValueError) as exc:
        out(f"FAIL: survival reference debt unreadable: {exc}")
        return 1

    unlisted = sorted(set(discover(root)) - set(grids))
    for rel in unlisted:
        problems.append(f"{rel}: carries survival references but is not in GRIDS")

    items = []
    for rel in grids:
        path = os.path.join(root, rel)
        if not os.path.exists(path):
            problems.append(f"{rel}: listed grid is missing")
            continue
        rows = survival_rows(path, rel)
        if not rows:
            problems.append(f"{rel}: listed grid has no survival rows")
        items.extend(rows)

    unknown = [i for i in items if not oracle.is_survival(i[2])]
    for grid, line, fn, _, _ in unknown:
        problems.append(f"{grid}:{line}: {fn} has no direct survival route")
    items = [i for i in items if oracle.is_survival(i[2])]

    if jobs == 1 or len(items) < 2:
        results = [_evaluate(i) for i in items]
    else:
        with Pool(jobs or os.cpu_count()) as pool:
            results = pool.map(_evaluate, items, chunksize=1)

    used = set()
    per_grid = {}
    for item, status, rel, value, why in results:
        grid, line, fn, toks, ref = item
        stats = per_grid.setdefault(grid, {"rows": 0, "debt": 0, "min_digits": None,
                                           "min_met": None})
        stats["rows"] += 1
        if status == "oracle":
            problems.append(f"{grid}:{line}: {fn}({','.join(t for t in toks if t)}) "
                            f"oracle not certified: {why}")
            continue
        with mp.workdps(30):
            digits = mp.inf if rel == 0 else -mp.log10(rel)
        if stats["min_digits"] is None or digits < stats["min_digits"]:
            stats["min_digits"] = digits
        key = _debt_key(item)
        if status == "below":
            if key in debt:
                used.add(key)
                stats["debt"] += 1
                if verbose:
                    out(f"  debt  {grid}:{line} {fn} {mp.nstr(digits, 4)} digits "
                        f"[{debt[key]['blocker']}]")
                continue
            problems.append(f"{grid}:{line}: {fn}({','.join(t for t in toks if t)}) "
                            f"reference {ref} agrees with the direct value "
                            f"{mp.nstr(value, 34)} to {mp.nstr(digits, 4)} digits, "
                            "below the 25-digit floor")
        else:
            if stats["min_met"] is None or digits < stats["min_met"]:
                stats["min_met"] = digits
            if key in debt:
                used.add(key)
                problems.append(f"{grid}:{line}: {fn} is registered as debt but now "
                                "meets the floor; delete its entry")
            elif verbose:
                out(f"  ok    {grid}:{line} {fn} {mp.nstr(digits, 4)} digits")

    for key in sorted(set(debt) - used):
        problems.append(f"stale debt entry, no such committed row: {key[0]} {key[1]}"
                        f"({','.join(key[2])}) reference {key[3]}")

    total = sum(s["rows"] for s in per_grid.values())
    owed = sum(s["debt"] for s in per_grid.values())
    for grid in grids:
        s = per_grid.get(grid)
        if s:
            md = "-" if s["min_digits"] is None else mp.nstr(s["min_digits"], 4)
            mm = "-" if s["min_met"] is None else mp.nstr(s["min_met"], 4)
            out(f"  {grid:46s} {s['rows']:4d} rows  min {md:>6s} digits, "
                f"{mm:>6s} outside debt  {s['debt']:3d} registered debt")
    if problems:
        out(f"FAIL: survival references ({len(problems)} problem(s)):")
        for p in problems:
            out("  - " + p)
        return 1
    by_blocker = {}
    for key in used:
        b = debt[key]["blocker"]
        by_blocker[b] = by_blocker.get(b, 0) + 1
    detail = ", ".join(f"{n} {b}" for b, n in sorted(by_blocker.items()))
    out(f"PASS: {total} survival references in {len(per_grid)} grids recomputed at "
        f"their binary64 arguments; {total - owed} meet the 25-digit floor, "
        f"{owed} are registered debt ({detail or 'none'})")
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--jobs", type=int, default=None,
                    help="worker processes (default: every CPU; 1 runs serially)")
    ap.add_argument("--verbose", action="store_true", help="print one line per row")
    a = ap.parse_args(argv)
    return run(jobs=a.jobs, verbose=a.verbose)


if __name__ == "__main__":
    sys.exit(main())
