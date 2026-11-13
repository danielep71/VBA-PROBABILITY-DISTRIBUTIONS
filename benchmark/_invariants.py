"""
Always-gating invariants (#327), independent of the accuracy tiers.

The guarantee/target split decides how close a value must be. These checks
decide what may never happen, whatever a row's contract says, and they also
reach rows that no gating contract scores (characterization-only contracts
and rows still awaiting a disposition under #22):

  * SILENT WRONG: a finite result whose relative error exceeds 1E-6 (or the
    row's own documented threshold, when that is looser), or an exact 0 or 1
    where the reference rounds to a different double.
  * MONOTONICITY: along every evaluated sequence of a CDF, survival or
    inverse (rows of one function that differ only in arg1), the observed
    values may not run backwards by more than the two rows' error allowances.
  * P + Q: a CDF row and the survival row at the same arguments must sum to
    one within the sum of their allowances.

The documented error behaviour (#NUM! in envelope-reject regions) is gated
by _contract_eval.dispositions; the #VALUE!/#NUM! input-validation cases are
gated by the Excel regression suite.

An allowance is guarantee * |reference| for a relative row and the guarantee
itself for an absolute one; a row without a gating contract receives the
policy floor for its kind (accuracy_tiers.FORWARD_FLOOR or INVERSE_FLOOR).
This module is stdlib-only, like _contract_eval.
"""
import re
from collections import defaultdict, namedtuple
from decimal import Decimal

from _contract_eval import parse_observed, parse_reference, parse_threshold
from accuracy_tiers import FORWARD_FLOOR, INVERSE_FLOOR

SILENT_WRONG_LIMIT = Decimal("1E-6")
_OUTPUT_MEASURES = ("output_error", "quantile_error")
_SEQUENCE = re.compile(r"_(Cumulative|Survival|InverseCumulative|InverseCumulativeFast|"
                       r"InverseSurvival)$")
_PAIR = re.compile(r"^(.*)_(Cumulative|Survival)$")

Violation = namedtuple("Violation", "check function args detail")
Result = namedtuple("Result", "checked violations reversals")


def _args(row):
    return tuple(row.get(k, "") or "" for k in ("arg1", "arg2", "arg3", "arg4"))


def governing(contracts, guarantees):
    """(function, regime) -> (guarantee, target, metric) for output contracts."""
    out = {}
    for c in contracts:
        if c["measure"].strip() not in _OUTPUT_MEASURES:
            continue
        if c.get("status", "active").strip() != "active":
            continue
        out[(c["function"], c["regime"])] = (
            parse_threshold(guarantees.get(c["contract_id"], "")),
            parse_threshold(c["threshold"]),
            c["metric"].strip().lower())
    return out


def _measured(rows, gov):
    """Yield (row, observed, reference, guarantee, target, relative) for usable rows."""
    for r in rows:
        if (r.get("expected_error") or "").strip():
            continue
        o = parse_observed(r.get("observed_vba"))
        ref = parse_reference(r.get("reference"))
        if o is None or ref is None:
            continue
        key = (r["function"], (r.get("regime") or "all"))
        if key in gov and gov[key][0] is not None:
            g, t, metric = gov[key]
            relative = metric == "relative"
        else:
            relative = (r.get("metric") or "").strip() == "rel"
            if not relative:
                continue
            g = INVERSE_FLOOR if "Inverse" in r["function"] else FORWARD_FLOOR
            t = None
        yield r, o, ref, g, t, relative


def _allowance(ref, g, relative):
    return g * abs(ref) if relative else g


def silent_wrong(rows, gov):
    checked, bad = 0, []
    for r, o, ref, g, t, relative in _measured(rows, gov):
        if not relative:
            continue
        checked += 1
        limit = max(SILENT_WRONG_LIMIT, t or 0, g or 0)
        args = _args(r)
        if ref == 0:
            if o != 0:
                bad.append(Violation("silent_wrong", r["function"], args,
                                     f"{float(o):.17g} where the reference is 0"))
            continue
        err = abs((o - ref) / ref)
        # An exact 0 against a nonzero reference has relative error 1, so the
        # limit test catches it; an exact 1 can hide inside a small error.
        if err > limit:
            bad.append(Violation("silent_wrong", r["function"], args,
                                 f"relative error {float(err):.3g} > {limit}"))
        elif o == 1 and float(ref) != 1.0:
            bad.append(Violation("silent_wrong", r["function"], args,
                                 f"1 where the reference is {float(ref):.17g}"))
    return Result(checked, bad, 0)


def monotonicity(rows, gov):
    seqs = defaultdict(list)
    for r, o, ref, g, t, relative in _measured(rows, gov):
        if not _SEQUENCE.search(r["function"]):
            continue
        x = parse_reference(r.get("arg1"))
        if x is None:
            continue
        a = _args(r)
        seqs[(r["function"],) + a[1:]].append((x, ref, o, _allowance(ref, g, relative)))
    checked, bad, reversals = 0, [], 0
    for key, seq in seqs.items():
        seq.sort(key=lambda s: s[0])
        for (x1, r1, o1, a1), (x2, r2, o2, a2) in zip(seq, seq[1:]):
            if x1 == x2 or r1 == r2:
                continue
            checked += 1
            # Direction of the true function on this step, from the reference.
            back = (o1 - o2) if r2 > r1 else (o2 - o1)
            if back <= 0:
                continue
            reversals += 1
            if back > a1 + a2:
                bad.append(Violation("monotonicity", key[0], (str(x1), str(x2)) + key[1:],
                                     f"runs backwards by {float(back):.3g}, allowance "
                                     f"{float(a1 + a2):.3g}"))
    return Result(checked, bad, reversals)


def p_plus_q(rows, gov):
    sides = defaultdict(dict)
    for r, o, ref, g, t, relative in _measured(rows, gov):
        m = _PAIR.match(r["function"])
        if not m or not relative:
            continue
        sides[(m.group(1),) + _args(r)][m.group(2)] = (o, ref, _allowance(ref, g, relative))
    checked, bad = 0, []
    for key, pair in sides.items():
        if len(pair) != 2:
            continue
        (op, rp, ap), (oq, rq, aq) = pair["Cumulative"], pair["Survival"]
        checked += 1
        # The references themselves sum to one only to their printed digits.
        slack = abs(rp + rq - 1)
        dev = abs(op + oq - 1)
        if dev > ap + aq + slack:
            bad.append(Violation("p_plus_q", key[0], key[1:],
                                 f"|P + Q - 1| = {float(dev):.3g}, allowance "
                                 f"{float(ap + aq + slack):.3g}"))
    return Result(checked, bad, 0)


CHECKS = (("silent wrong", silent_wrong), ("monotonicity", monotonicity),
          ("P + Q", p_plus_q))


def run_all(rows, contracts, guarantees):
    gov = governing(contracts, guarantees)
    return [(name, fn(rows, gov)) for name, fn in CHECKS]


def markdown(results, heading="## Always-gating checks"):
    lines = [heading, "",
             "Independent of the tiers (#327): any violation fails the gate.", "",
             "| Check | Evaluated | Violations | Notes |", "|---|---:|---:|---|"]
    for name, res in results:
        note = (f"{res.reversals} reversal(s) within the error allowance"
                if name == "monotonicity" else "")
        lines.append(f"| {name} | {res.checked} | {len(res.violations)} | {note} |")
    for name, res in results:
        for v in res.violations[:20]:
            lines.append(f"- {name}: `{v.function}` at ({', '.join(a for a in v.args if a)})"
                         f" — {v.detail}")
    return lines


def violation_count(results):
    return sum(len(res.violations) for _, res in results)
