"""
Non-destructive promotion of new evidence rows into probability_accuracy_grid.csv.

The grid holds 2030 Excel observations that cannot be regenerated: the generator
emits observed_vba blank by design, so a full rebuild would discard them. This
tool is the only sanctioned way to add rows, and it is deliberately incapable of
the operations that would lose evidence.

  never deletes a row EXCEPT under --retire, which requires the function,
      the regime, the exact expected row count, and a written reason, and
      which acknowledges the observations it discards
  never reorders rows
  never alters an existing observed_vba
  never alters an existing row at all unless --patch-metadata is given
  appends a row ONLY when --allow-add is given; without it a key the grid does
      not already contain is a hard failure, not a silent insertion

That last rule is the point. "Patch existing" alone is not enough for the
LogGamma and LogGamma1p contracts, because some of their rows must become new
main-grid evidence - and a patcher that quietly adds rows is just another
generator wearing a safety label.

Keys are IEEE-754 bit patterns, not decimal text. The exporter parses each
argument with Val(), so "0.85" and "0.84999999999999998" are the same row.

Report-only unless --write. New rows arrive with observed_vba blank; fill them
with Export_Accuracy_MissingOnly, which leaves every existing cell untouched.
"""
import argparse, csv, importlib.util, os, struct, sys
from collections import Counter

GRID_FIELDS_NOTE = "columns are taken from the existing grid header"


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


def load_generator(path):
    spec = importlib.util.spec_from_file_location("gen", path)
    m = importlib.util.module_from_spec(spec)
    sys.modules["gen"] = m
    spec.loader.exec_module(m)
    import mpmath as mp
    mp.mp.dps = 50
    rows = m.build_rows()
    from _contract_eval import predicted_expected_error
    for r in rows:
        r["expected_error"] = "1" if predicted_expected_error(
            r.get("function", ""), r.get("arg2", ""), r.get("arg3", "")) else ""
    return rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--grid", default="probability_accuracy_grid.csv")
    ap.add_argument("--generator", default="generate_reference_values.py")
    ap.add_argument("--function", action="append", default=[],
                    help="restrict to these grid function names; repeatable")
    ap.add_argument("--allow-add", action="store_true",
                    help="permit appending rows the grid does not contain")
    ap.add_argument("--allow-unclaimed", action="store_true",
                    help="permit rows whose contract does not exist yet, so "
                         "claim/metric are blank. Required because a threshold "
                         "cannot be derived until its evidence is observed; a "
                         "--patch-metadata pass must fill them before the grid "
                         "is committed")
    ap.add_argument("--patch-metadata", action="store_true",
                    help="permit updating reference/claim/metric/expected_error on "
                         "matched rows (regime is part of the key: changing it "
                         "identifies a different row, it does not mutate this one)")
    ap.add_argument("--retire", nargs=3, metavar=("FUNCTION", "REGIME", "COUNT"),
                    help="delete every row with this function and regime. COUNT "
                         "must equal the number found, so a miscount fails "
                         "rather than deleting the wrong set. Requires --reason.")
    ap.add_argument("--reason", default="",
                    help="why the retired rows are being deleted; recorded in "
                         "the output and required by --retire")
    ap.add_argument("--write", action="store_true")
    a = ap.parse_args()

    if not a.function and not a.retire:
        print("REFUSING: --function is required. Promotion is explicit by design;")
        print("  a whole-grid promotion is indistinguishable from a regeneration.")
        return 1


    grid = list(csv.DictReader(open(a.grid)))
    fields = list(grid[0].keys())

    # ---- explicit retirement ------------------------------------------------
    # The only path that deletes. Kept deliberately awkward: the caller must
    # name the function, the regime and the exact count, and give a reason. A
    # superseded contract's evidence should not linger as unclaimed rows, but
    # nor should deletion ever be a side effect of some other operation.
    if a.retire:
        fn_r, reg_r, count_r = a.retire[0], a.retire[1], int(a.retire[2])
        if not a.reason.strip():
            print("REFUSING: --retire requires --reason.")
            return 1
        doomed = [r for r in grid if r["function"] == fn_r and r["regime"] == reg_r]
        observed = sum(1 for r in doomed if (r["observed_vba"] or "").strip())
        print(f"retire {fn_r} / {reg_r}")
        print(f"  rows found          {len(doomed)}   expected {count_r}")
        print(f"  observations lost   {observed}")
        print(f"  reason              {a.reason}")
        if len(doomed) != count_r:
            print(f"\nREFUSING: expected {count_r} rows, found {len(doomed)}.")
            return 1
        if not doomed:
            print("\nnothing to retire.")
            return 0
        left = [r for r in grid if not (r["function"] == fn_r and r["regime"] == reg_r)]
        still = [r for r in left if r["function"] == fn_r and r["regime"] == reg_r]
        assert not still, "retirement left rows behind"
        if not a.write:
            print("\nreport only. Re-run with --write to apply.")
            return 0
        with open(a.grid, "w", newline="") as fh:
            w = csv.DictWriter(fh, fieldnames=fields, lineterminator="\n")
            w.writeheader(); w.writerows(left)
        print(f"\nwrote {a.grid}: {len(doomed)} row(s) retired, "
              f"{len(left)} remain")
        return 0

    gen = load_generator(a.generator)
    want = [r for r in gen if r["function"] in set(a.function)]
    if not want:
        print(f"REFUSING: the generator emits no rows for {a.function}.")
        return 1

    # Duplicate keys are scoped to the functions under promotion. A duplicate
    # elsewhere - the three LogChoose rows from #17 finding D - does not stop
    # this tool from uniquely naming a LogGamma row, and blocking on it would
    # couple every promotion to an unrelated repair. Duplicates that DO affect
    # the promotion are still fatal: a patcher must never guess which of two
    # identical keys it is updating.
    gk = Counter(key(r) for r in grid)
    dup_all = {k for k, v in gk.items() if v > 1}
    targets = set(a.function)
    dup = {k for k in dup_all if k[0] in targets}
    if dup_all - dup:
        print(f"note: {len(dup_all - dup)} duplicate key(s) elsewhere in the grid "
              f"({sorted({k[0] for k in dup_all - dup})}); out of scope here, "
              f"tracked as #17 finding D.")
    if dup:
        print(f"REFUSING: {len(dup)} duplicate key(s) among the promoted "
              f"functions; a patcher cannot uniquely name the row it updates.")
        for k in sorted(dup)[:5]:
            print(f"   {k[0]}  {k[1:5]}")
        return 1
    by = {key(r): r for r in grid}

    # Duplicate keys among the rows selected for promotion are fatal too, not
    # only duplicates already in the grid. If the generator emitted the same new
    # key twice and it is absent from the grid, both copies would otherwise be
    # appended - creating the very ambiguity this tool refuses to work with.
    want_dup = {k for k, c in Counter(key(r) for r in want).items() if c > 1}
    if want_dup:
        print(f"REFUSING: {len(want_dup)} duplicate key(s) among the generator "
              f"rows selected for promotion.")
        for k in sorted(want_dup)[:5]:
            print(f"   {k[0]}  {k[1:5]}  regime={k[5]}")
        return 1

    unclaimed = [r for r in want if not (r.get("claim") or "").strip()]

    add, patch, same = [], [], 0
    for r in want:
        k = key(r)
        cur = by.get(k)
        if cur is None:
            add.append(r)
            continue
        diffs = [c for c in ("reference", "claim", "metric", "expected_error")
                 if cur.get(c, "") != r.get(c, "")]
        if diffs:
            patch.append((k, cur, r, diffs))
        else:
            same += 1

    obs_on_add = 0
    print(f"functions: {', '.join(sorted(set(a.function)))}")
    print(f"  generator rows      {len(want)}")
    print(f"  already identical   {same}")
    print(f"  would APPEND        {len(add)}"
          f"{'' if a.allow_add else '   <- blocked without --allow-add'}")
    print(f"  would PATCH         {len(patch)}"
          f"{'' if a.patch_metadata else '   <- blocked without --patch-metadata'}")
    print(f"  pending-contract    {len(unclaimed)}"
          f"{'' if not unclaimed else ('' if a.allow_unclaimed else '   <- blocked without --allow-unclaimed')}")
    if add:
        print("\n  rows to append (observed_vba blank; fill with "
              "Export_Accuracy_MissingOnly):")
        for r in add[:10]:
            args = ",".join(x for x in (r["arg1"], r["arg2"], r["arg3"], r["arg4"]) if x)
            print(f"     {r['function']:34s} regime={r['regime']:16s} ({args[:44]})")
        if len(add) > 10:
            print(f"     ... {len(add) - 10} more")
    if patch:
        print("\n  rows to patch:")
        for k, cur, new, diffs in patch[:10]:
            print(f"     {k[0]:34s} regime={k[5]:16s} fields={diffs}")
        if len(patch) > 10:
            print(f"     ... {len(patch) - 10} more")

    if add and not a.allow_add:
        print("\nREFUSING: rows would be added. Pass --allow-add to permit it.")
        return 1
    if unclaimed and not a.allow_unclaimed:
        print(f"\nREFUSING: {len(unclaimed)} row(s) have no contract, so claim "
              f"and metric are blank. That is normally a bug - a row generated "
              f"for a claim nobody wrote down. Pass --allow-unclaimed only when "
              f"promoting evidence ahead of its threshold, and follow with a "
              f"--patch-metadata pass before committing.")
        return 1
    if patch and not a.patch_metadata:
        print("\nREFUSING: existing rows would change. Pass --patch-metadata.")
        return 1
    if not a.write:
        print("\nreport only. Re-run with --write to apply.")
        return 0

    before_obs = [r["observed_vba"] for r in grid]
    before_keys = [key(r) for r in grid]
    for k, cur, new, diffs in patch:
        for c in diffs:
            cur[c] = new.get(c, "")
    for r in add:
        row = {f: "" for f in fields}
        for f in fields:
            if f in r:
                row[f] = r[f]
        row["observed_vba"] = ""
        row["evidence_set"] = r.get("evidence_set", "main grid")
        grid.append(row)

    # invariants, asserted rather than trusted
    assert [key(r) for r in grid][:len(before_keys)] == before_keys, \
        "existing row keys or order changed"
    assert [r["observed_vba"] for r in grid][:len(before_obs)] == before_obs, \
        "an existing observation changed"
    assert all(g["observed_vba"] == "" for g in grid[len(before_keys):]), \
        "an appended row arrived with an observation"
    with open(a.grid, "w", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=fields, lineterminator="\n")
        w.writeheader(); w.writerows(grid)
    print(f"\nwrote {a.grid}: {len(add)} appended, {len(patch)} patched, "
          f"{len(before_keys)} existing rows untouched in key, order and observation")
    still = sum(1 for r in grid if not (r.get("claim") or "").strip())
    if still:
        print(f"\n  {still} row(s) in the grid still carry no claim/metric. The "
              f"committed grid must not: run --patch-metadata once the contracts "
              f"are frozen.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
