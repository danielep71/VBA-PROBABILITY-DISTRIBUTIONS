"""
Reference-only migration derived from reference_audit_post_canonicalisation.csv.

Patches the `reference` column of probability_accuracy_grid.csv and NOTHING
else. observed_vba is never touched, no row is added or removed, and no
contract threshold is altered - what the library actually measured against
correct references must be established before deciding what the contracts
should be.

The action per row comes from the audit's replacement_source, not from
re-deriving policy here:

  CANONICAL_GENERATOR  the canonicalised generator agrees with the independent
                       oracle and the committed reference does not: adopt the
                       generator value.
  INDEPENDENT_ORACLE   both candidates sit outside the floor and neither wins:
                       adopt a freshly computed oracle value, serialised to the
                       grid's 25 significant digits. The full-precision value is
                       recorded in the provenance report.
  BLOCKED_GENERATOR    the generator is wrong: patch the generator first.
  BLOCKED_ORACLE       the oracle could not be certified.
  NONE                 leave the committed reference alone.

Refuses to write if any row is blocked, if the audit does not match the current
generator, or if any key is ambiguous. Report-only unless --write is given.
"""
import argparse, csv, importlib.util, os, struct, sys
from collections import Counter

import mpmath as mp

BLOCKING = {"BLOCKED_GENERATOR", "BLOCKED_ORACLE"}


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


def audit_key(r):
    return (r["function"], r["arg1_hex"], r["arg2_hex"], r["arg3_hex"],
            r["arg4_hex"], r["regime"])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--grid", default="probability_accuracy_grid.csv")
    ap.add_argument("--audit", default="reference_audit_post_canonicalisation.csv")
    ap.add_argument("--generator", default="generate_reference_values.py")
    ap.add_argument("--report", default="reference_migration.csv")
    ap.add_argument("--write", action="store_true",
                    help="apply the patch; without it nothing is written to the grid")
    a = ap.parse_args()

    audit = list(csv.DictReader(open(a.audit)))
    grid = list(csv.DictReader(open(a.grid)))

    spec = importlib.util.spec_from_file_location("gen", a.generator)
    g = importlib.util.module_from_spec(spec); sys.modules["gen"] = g
    spec.loader.exec_module(g)
    mp.mp.dps = 50
    gen = {key(r): r for r in g.build_rows()}

    # ---- refuse to proceed on a blocked or unverifiable audit ---------------
    blocked = [r for r in audit if r["replacement_source"] in BLOCKING]
    if blocked:
        print(f"REFUSING: {len(blocked)} row(s) are blocked "
              f"({dict(Counter(r['replacement_source'] for r in blocked))}).")
        print("  Fix the generator or the oracle first; a blocked row must not migrate.")
        return 1

    gk = Counter(key(r) for r in grid)
    dup = {k for k, v in gk.items() if v > 1}
    grid_by = {}
    for r in grid:
        grid_by.setdefault(key(r), r)

    stale, oracle_rows, missing, mismatch = [], [], [], []
    for r in audit:
        ak = audit_key(r)
        hits = [k for k in grid_by if k[:6] == ak]
        if len(hits) != 1:
            missing.append((r["function"], ak, len(hits))); continue
        k = hits[0]
        if k in dup:
            missing.append((r["function"], ak, "ambiguous key")); continue
        grow = grid_by[k]
        if grow["reference"] != r["committed_reference"]:
            mismatch.append(r["function"]); continue
        if r["replacement_source"] == "CANONICAL_GENERATOR":
            grow_gen = gen.get(k)
            if grow_gen is None or grow_gen["reference"] != r["generator_reference"]:
                mismatch.append(r["function"]); continue
            stale.append((k, grow, r, grow_gen["reference"]))
        elif r["replacement_source"] == "INDEPENDENT_ORACLE":
            oracle_rows.append((k, grow, r))

    print(f"audit rows {len(audit)}   grid rows {len(grid)}")
    print(f"  CANONICAL_GENERATOR  {len(stale)}")
    print(f"  INDEPENDENT_ORACLE   {len(oracle_rows)}")
    if missing:
        print(f"  UNMATCHED            {len(missing)}   {missing[:3]}")
    if mismatch:
        print(f"  DRIFTED SINCE AUDIT  {len(mismatch)}   {Counter(mismatch).most_common(4)}")
    if missing or mismatch:
        print("\nREFUSING: the audit does not describe the current grid or generator.")
        print("  Re-run audit_references.py before migrating.")
        return 1

    # ---- build the patch ---------------------------------------------------
    report = []
    for k, grow, ar, newref in stale:
        report.append({
            "function": k[0], "regime": k[5],
            "arguments": ";".join(x for x in (grow["arg1"], grow["arg2"],
                                              grow["arg3"], grow["arg4"]) if x),
            "old_reference": grow["reference"], "new_reference": newref,
            "independent_oracle": ar["independent_oracle"],
            "old_rel_error": ar["committed_rel_error"],
            "new_rel_error": ar["generator_rel_error"],
            "classification": ar["classification"],
            "replacement_source": ar["replacement_source"],
        })
    for k, grow, ar in oracle_rows:
        # Serialise the certified oracle at the grid's 25 significant digits.
        o = mp.mpf(ar["independent_oracle"])
        report.append({
            "function": k[0], "regime": k[5],
            "arguments": ";".join(x for x in (grow["arg1"], grow["arg2"],
                                              grow["arg3"], grow["arg4"]) if x),
            "old_reference": grow["reference"], "new_reference": mp.nstr(o, 25),
            "independent_oracle": ar["independent_oracle"],
            "old_rel_error": ar["committed_rel_error"],
            "new_rel_error": "0",
            "classification": ar["classification"],
            "replacement_source": ar["replacement_source"],
        })

    with open(a.report, "w", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=list(report[0].keys()), lineterminator="\n")
        w.writeheader(); w.writerows(report)
    print(f"\nwrote {a.report}: {len(report)} rows")

    if not a.write:
        print("\nreport only. Re-run with --write to patch the grid.")
        return 0

    patch = {k: nr for k, _, _, nr in stale}
    patch.update({k: mp.nstr(mp.mpf(ar["independent_oracle"]), 25)
                  for k, _, ar in oracle_rows})
    changed = obs_before = obs_after = 0
    for r in grid:
        obs_before += 1 if (r["observed_vba"] or "").strip() else 0
    for r in grid:
        k = key(r)
        if k in patch and r["reference"] != patch[k]:
            r["reference"] = patch[k]; changed += 1
    for r in grid:
        obs_after += 1 if (r["observed_vba"] or "").strip() else 0
    assert obs_before == obs_after, "observations must not change"
    with open(a.grid, "w", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=list(grid[0].keys()), lineterminator="\n")
        w.writeheader(); w.writerows(grid)
    print(f"\npatched {a.grid}: {changed} reference cells; "
          f"observations {obs_before} -> {obs_after}, unchanged")
    return 0


if __name__ == "__main__":
    sys.exit(main())
