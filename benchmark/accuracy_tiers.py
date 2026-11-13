"""
Two-tier accuracy policy (#327): derive each contract's guarantee.

Every contract in accuracy_contracts.csv carries one threshold. Since #327
that threshold is the MEASURED TARGET: reported in every summary, never a
release blocker on its own. The GUARANTEE is what gates the release and the
Accuracy Gate, and it is derived here by a fixed rule, never chosen by hand:

  * forward relative errors (densities, CDFs, survivals, moments, special
    functions):           guarantee = max(target, 1E-13)
  * inverses (quantile_error, or any *Inverse* function scored relative):
                          guarantee = max(target, 1E-12)
  * balanced Beta CDF and survival (shape ratio >= 0.1): guarantee =
    max(target, 2E-13). The route is modelled bit for bit (#174, 368 of 368
    Excel observations) at up to 1.38E-13 for shapes 20 to 100; 2E-13 is the
    recalibration decided in #163. The model covers shapes 0.5 to 100; the
    regime's rows outside that range are measured in the main grid.
  * absolute, scaled, log-absolute and tail-residual metrics: guarantee =
    target, unchanged.

No guarantee is ever stricter than its target. The derived values live in
accuracy_guarantees.csv, a generated file: the registry itself is untouched
because observation_manifest.json binds its hash to the Excel export, and a
policy change must not pretend to be a new measurement.

    python accuracy_tiers.py --write   regenerate accuracy_guarantees.csv
    python accuracy_tiers.py           verify it is current (exit 1 if not)
"""
import argparse
import csv
import io
import os
import sys
from decimal import Decimal

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from _contract_eval import parse_threshold  # noqa: E402

CONTRACTS = os.path.join(HERE, "accuracy_contracts.csv")
GUARANTEES = os.path.join(HERE, "accuracy_guarantees.csv")
FIELDS = ["contract_id", "measured_target", "guarantee", "basis"]

FORWARD_FLOOR = Decimal("1E-13")
INVERSE_FLOOR = Decimal("1E-12")
BALANCED_BETA_FLOOR = Decimal("2E-13")
BALANCED_BETA = frozenset({
    "Beta_Cumulative.balanced.output_rel",
    "Beta_Survival.balanced.output_rel",
})
# Measures whose threshold is not a relative output error of a probability
# or a value; their thresholds stay as they are.
UNCHANGED_MEASURES = frozenset({
    "tail_probability_residual", "scaled_output_error", "log_absolute_error",
})

BASIS_NONE = "no numeric threshold: no guarantee"
BASIS_SAME = "absolute or residual metric: guarantee = target"
BASIS_FORWARD = "forward relative error: max(target, 1E-13)"
BASIS_INVERSE = "inverse: max(target, 1E-12)"
BASIS_BETA = "balanced Beta exception (#163, #174): max(target, 2E-13)"


def is_inverse(contract):
    return (contract["measure"].strip() == "quantile_error"
            or "Inverse" in contract["function"])


def derive(contract):
    """Return (guarantee text, basis) for one registry row."""
    target_text = contract["threshold"].strip()
    target = parse_threshold(target_text)
    if target is None:
        return "", BASIS_NONE
    metric = contract["metric"].strip().lower()
    measure = contract["measure"].strip()
    if metric != "relative" or measure in UNCHANGED_MEASURES:
        return target_text, BASIS_SAME
    if contract["contract_id"] in BALANCED_BETA:
        floor, basis = BALANCED_BETA_FLOOR, BASIS_BETA
    elif is_inverse(contract):
        floor, basis = INVERSE_FLOOR, BASIS_INVERSE
    else:
        floor, basis = FORWARD_FLOOR, BASIS_FORWARD
    if target >= floor:
        return target_text, basis
    return format(floor, "E"), basis


def load_registry(path=CONTRACTS):
    with open(path, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def render(contracts):
    out = io.StringIO()
    w = csv.DictWriter(out, fieldnames=FIELDS, lineterminator="\n")
    w.writeheader()
    for c in sorted(contracts, key=lambda c: c["contract_id"]):
        g, basis = derive(c)
        w.writerow({"contract_id": c["contract_id"],
                    "measured_target": c["threshold"].strip(),
                    "guarantee": g, "basis": basis})
    return out.getvalue()


def problems(contracts, text):
    """Every reason the committed guarantees file is not the generator's output."""
    found = []
    expected = render(contracts)
    if text != expected:
        found.append("accuracy_guarantees.csv differs from the generator output; "
                     "it is never edited by hand - run accuracy_tiers.py --write")
    rows = list(csv.DictReader(io.StringIO(text)))
    for r in rows:
        t = parse_threshold(r.get("measured_target"))
        g = parse_threshold(r.get("guarantee"))
        if t is not None and (g is None or g < t):
            found.append(f"{r.get('contract_id')}: guarantee {r.get('guarantee')!r} "
                         f"is stricter than its measured target {r.get('measured_target')}")
    return found


def load_guarantees(path=GUARANTEES, contracts=None):
    """contract_id -> guarantee text ("" when none). Refuses a stale file.

    The evaluators call this, so a guarantee that was edited by hand or not
    regenerated after a registry change stops the gate instead of quietly
    deciding it.
    """
    if contracts is None:
        contracts = load_registry()
    with open(path, newline="", encoding="utf-8") as f:
        text = f.read()
    bad = problems(contracts, text)
    if bad:
        raise ValueError("; ".join(bad))
    return {r["contract_id"]: r["guarantee"]
            for r in csv.DictReader(io.StringIO(text))}


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[1])
    ap.add_argument("--write", action="store_true")
    a = ap.parse_args(argv)
    contracts = load_registry()
    if a.write:
        with open(GUARANTEES, "w", newline="", encoding="utf-8") as f:
            f.write(render(contracts))
        print(f"wrote {os.path.basename(GUARANTEES)} ({len(contracts)} contracts)")
        return 0
    try:
        with open(GUARANTEES, newline="", encoding="utf-8") as f:
            text = f.read()
    except OSError as exc:
        print(f"FAIL: {exc}")
        return 1
    bad = problems(contracts, text)
    for p in bad:
        print("FAIL: " + p)
    if not bad:
        print(f"accuracy_guarantees.csv is current ({len(contracts)} contracts)")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
