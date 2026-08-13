"""
Verify a promotion + missing-only export, without needing git.

Compares the current grid against a snapshot taken before the promotion and
proves the three properties that matter:

  * no existing observation changed;
  * no existing row moved, and none was added or removed except the appends;
  * every appended row is now filled, and none came back ERROR.

Usage, from benchmark/:

    python check_promotion.py --before <snapshot.csv> --expect-new 101

Take the snapshot BEFORE running promote_grid_rows.py --write:

    copy probability_accuracy_grid.csv grid_before_promotion.csv
"""
import argparse, csv, struct, sys
from collections import Counter


def bits(t):
    t = (t or "").strip()
    if t == "":
        return ""
    try:
        return struct.pack(">d", float(t)).hex()
    except (ValueError, OverflowError):
        return "?" + t


def key(r):
    return (r["function"], bits(r["arg1"]), bits(r["arg2"]), bits(r["arg3"]),
            bits(r["arg4"]), r["regime"], r.get("evidence_set", ""))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--before", default="grid_before_promotion.csv")
    ap.add_argument("--grid", default="probability_accuracy_grid.csv")
    ap.add_argument("--expect-new", type=int, default=101)
    a = ap.parse_args()

    old = list(csv.DictReader(open(a.before)))
    new = list(csv.DictReader(open(a.grid)))
    fails = []

    print(f"rows {len(old)} -> {len(new)}   (expected +{a.expect_new})")
    if len(new) - len(old) != a.expect_new:
        fails.append(f"row delta {len(new) - len(old)}, expected {a.expect_new}")

    head, tail = new[:len(old)], new[len(old):]

    changed = [i for i, (x, y) in enumerate(zip(old, head))
               if x["observed_vba"] != y["observed_vba"]]
    print(f"existing observations changed: {len(changed)}")
    if changed:
        fails.append(f"{len(changed)} existing observation(s) changed")
        for i in changed[:5]:
            print(f"   row {i + 2}: {old[i]['function']} {old[i]['arg1']}")

    moved = [i for i, (x, y) in enumerate(zip(old, head)) if key(x) != key(y)]
    print(f"existing rows moved or re-keyed: {len(moved)}")
    if moved:
        fails.append(f"{len(moved)} existing row(s) moved or re-keyed")

    for col in ("arg1", "arg2", "arg3", "arg4", "reference", "regime",
                "evidence_set", "claim", "metric"):
        d = [i for i, (x, y) in enumerate(zip(old, head)) if x[col] != y[col]]
        if d:
            fails.append(f"{len(d)} existing row(s) changed {col}")

    filled = sum(1 for r in tail if (r["observed_vba"] or "").strip())
    errs = [r for r in tail if (r["observed_vba"] or "").strip().upper() == "ERROR"]
    blank = [r for r in tail if not (r["observed_vba"] or "").strip()]
    print(f"appended rows: {len(tail)}   filled: {filled}   ERROR: {len(errs)}   blank: {len(blank)}")
    if blank:
        fails.append(f"{len(blank)} appended row(s) still unobserved")
    if errs:
        fails.append(f"{len(errs)} appended row(s) returned ERROR")
        for r in errs[:8]:
            print(f"   ERROR: {r['function']} regime={r['regime']} arg1={r['arg1']}")

    print("\nappended rows by function and regime:")
    for k, v in sorted(Counter((r["function"], r["regime"]) for r in tail).items()):
        print(f"   {k[0]}.{k[1]:20s} {v:4d}")

    print()
    if fails:
        for f in fails:
            print("FAIL:", f)
        return 1
    print("PASS: existing evidence untouched; every appended row observed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
