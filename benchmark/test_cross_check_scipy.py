"""
Proofs for the SciPy cross-check classification (#154).

Part 1 needs no SciPy and always runs. On a fixture grid it shows that

  * a function without a SciPy mapping is reported as not mapped, never as
    beyond SciPy's range, even with shape parameters far above the range limit,
    and SciPy is never asked about it;
  * a mapped Beta/F point above the shape limit is beyond range, with the reason;
  * a mapped point where SciPy raises or returns NaN is beyond range, with the
    reason;
  * a mapped point SciPy evaluates is confirmed, with its relative difference;
  * a mapped function with fewer arguments (one-argument moments and normal
    inverses) is classified, not silently dropped;
  * a mapped row with a missing argument or reference fails loudly;
  * every row lands in exactly one class, and the record lists the unmapped
    function in its own table with its row count and reason.

Part 2 needs SciPy. It checks every mapping against known values: hand-computed
closed forms and mpmath, not SciPy against itself. The values are chosen so a
wrong convention changes them: failures versus trials in the geometric and
negative binomial, SciPy's exclusive upper bound in the discrete uniform, and
the hypergeometric argument order. When SciPy is not installed, as in CI, part 2
prints SKIP and the script still exits 0.

Run: python3 test_cross_check_scipy.py   (exit 0 = pass, nonzero = fail)
"""
import csv
import os
import sys
import tempfile

import mpmath as mp

import cross_check_scipy as ccs

FAILS = []
RAN_KNOWN_VALUES = []


def check(cond, msg):
    if not cond:
        FAILS.append(msg)


def _row(fn, reference, *args):
    r = {"function": fn, "reference": reference,
         "arg1": "", "arg2": "", "arg3": "", "arg4": ""}
    for i, a in enumerate(args, start=1):
        r[f"arg{i}"] = a
    return r


def part1_classification():
    asked = []

    def fake(fn, args):
        asked.append(fn)
        if args and args[0] == 0.125:
            raise ValueError("domain")
        if args and args[0] == 0.375:
            return float("nan")
        return 0.5000000000000001

    rows = [
        # unmapped: named reason, shapes far beyond SCIPY_SHAPE_LIMIT
        _row("Lognormal_ParametersFromMeanStdDev", "1.5", "7.5", "1E12", "1E12"),
        # unmapped: default reason, a Beta/F-like name with huge shapes
        _row("F_Imaginary", "0.25", "1.0", "1E9", "1E10"),
        # mapped, above the shape limit
        _row("F_Cumulative", "0.6", "1.0", "2.5", "1E8"),
        # mapped, SciPy raises / returns NaN
        _row("Beta_Cumulative", "0.5", "0.125", "2", "3"),
        _row("Beta_Survival", "0.5", "0.375", "2", "3"),
        # mapped, evaluated
        _row("Beta_Cumulative", "0.5", "0.5", "2", "2"),
        # mapped one-argument functions: must be classified, not dropped
        _row("Poisson_Mean", "0.5", "0.5"),
        _row("NormalStandard_InverseSurvival", "0.5", "0.5"),
    ]
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, "fixture_grid.csv")
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=["function", "arg1", "arg2", "arg3",
                                              "arg4", "reference"])
            w.writeheader()
            for r in rows:
                w.writerow(r)
        result = ccs.cross_check([path], fake)
        md = ccs.render(result, [(path, "0" * 64)], [("SciPy", "fixture")])

    per_fn = result["per_fn"]
    for fn in ("Lognormal_ParametersFromMeanStdDev", "F_Imaginary"):
        d = per_fn.get(fn, {})
        check(d.get(ccs.UNMAPPED) == 1 and d.get(ccs.BEYOND) == 0
              and d.get(ccs.CONFIRMED) == 0,
              f"{fn}: an unmapped function must be counted as not mapped only, got {d}")
        check(fn not in asked, f"{fn}: SciPy must not be asked about an unmapped function")
    check(per_fn["F_Cumulative"][ccs.BEYOND] == 1, "F_Cumulative at df2 1E8 must be beyond range")
    check("F_Cumulative" not in asked, "SciPy must not be asked beyond the shape limit")
    check(per_fn["Beta_Cumulative"][ccs.BEYOND] == 1
          and per_fn["Beta_Cumulative"][ccs.CONFIRMED] == 1,
          f"Beta_Cumulative: expected 1 beyond (raised) and 1 confirmed, got "
          f"{per_fn['Beta_Cumulative']}")
    check(per_fn["Beta_Survival"][ccs.BEYOND] == 1, "a NaN from SciPy must be beyond range")
    check(per_fn["Poisson_Mean"][ccs.CONFIRMED] == 1, "Poisson_Mean row was dropped")
    check(per_fn["NormalStandard_InverseSurvival"][ccs.CONFIRMED] == 1,
          "NormalStandard_InverseSurvival row was dropped")
    counts = result["counts"]
    check(result["total"] == len(rows), f"total {result['total']} != {len(rows)} rows")
    check(counts == {ccs.CONFIRMED: 3, ccs.BEYOND: 3, ccs.UNMAPPED: 2},
          f"unexpected class counts {counts}")
    reasons = [reason for _, _, reason in result["beyond_rows"]]
    check(any("shape parameter" in s for s in reasons), "shape-limit reason missing")
    check(any("raised ValueError" in s for s in reasons), "SciPy-raised reason missing")
    check(any("nan" in s for s in reasons), "NaN reason missing")
    check(abs(per_fn["Beta_Cumulative"]["worst"] - 2.220446049250313e-16) < 1e-30,
          f"relative difference wrong: {per_fn['Beta_Cumulative']['worst']}")

    # Record: unmapped functions sit in their own table, never in the beyond table.
    beyond_part, _, rest = md.partition("Functions not mapped to SciPy")
    check("Lognormal_ParametersFromMeanStdDev" not in beyond_part
          and "F_Imaginary" not in beyond_part,
          "an unmapped function appears in the confirmed or beyond tables")
    check(f"| Lognormal_ParametersFromMeanStdDev | 1 | "
          f"{ccs.UNMAPPED_REASONS['Lognormal_ParametersFromMeanStdDev']} |" in rest,
          "unmapped table lacks the named function, its count or its reason")
    check(f"| F_Imaginary | 1 | {ccs.DEFAULT_UNMAPPED_REASON} |" in rest,
          "unmapped table lacks the default-reason function")
    check("3 beyond SciPy's reliable range" in md and "2 not mapped to SciPy" in md,
          "summary does not separate beyond-range from not-mapped rows")

    # Malformed mapped rows fail loudly instead of being dropped.
    for bad in (_row("Binomial_PMF", "0.375", "2", "", "0.5"),
                _row("Binomial_PMF", "", "2", "4", "0.5")):
        try:
            ccs.classify_row(bad["function"], bad, fake)
        except ValueError:
            pass
        else:
            FAILS.append(f"malformed row was not rejected: {bad}")


# (function, args, exact value as an mpmath expression). Each value is a closed
# form or an mpmath evaluation, independent of SciPy.
def known_values():
    e = mp.e
    z975 = -mp.sqrt(2) * mp.erfinv(2 * mp.mpf("0.025") - 1)   # 1.959963984540054...
    return [
        # Binomial(k, n, p) with n = 4, p = 1/2
        ("Binomial_PMF", (2, 4, 0.5), mp.mpf(6) / 16),
        ("Binomial_LogPMF", (2, 4, 0.5), mp.log(mp.mpf(6) / 16)),
        ("Binomial_Cumulative", (2, 4, 0.5), mp.mpf(11) / 16),
        ("Binomial_Survival", (2, 4, 0.5), mp.mpf(5) / 16),
        ("Binomial_InverseCumulative", (0.5, 4, 0.5), 2),
        ("Binomial_Mean", (4, 0.5), 2),
        ("Binomial_Variance", (4, 0.5), 1),
        ("Binomial_StdDev", (4, 0.5), 1),
        # Poisson(k, mu) with mu = 1
        ("Poisson_PMF", (2, 1), 1 / (2 * e)),
        ("Poisson_LogPMF", (2, 1), -1 - mp.log(2)),
        ("Poisson_Cumulative", (2, 1), mp.mpf(5) / (2 * e)),
        ("Poisson_Survival", (2, 1), 1 - mp.mpf(5) / (2 * e)),
        ("Poisson_InverseCumulative", (0.9, 1), 2),   # CDF(1) = 0.7358, CDF(2) = 0.9197
        ("Poisson_Mean", (3.5,), mp.mpf("3.5")),
        ("Poisson_Variance", (3.5,), mp.mpf("3.5")),
        ("Poisson_StdDev", (4,), 2),
        # Geometric(k failures, p) with p = 1/4: P(k) = p (1-p)^k
        # (the trials convention would give 0.1875 at k = 2)
        ("Geometric_PMF", (2, 0.25), mp.mpf("0.140625")),
        ("Geometric_LogPMF", (2, 0.25), mp.log(mp.mpf("0.140625"))),
        ("Geometric_Cumulative", (2, 0.25), 1 - mp.mpf("0.75") ** 3),
        ("Geometric_Survival", (2, 0.25), mp.mpf("0.75") ** 3),
        ("Geometric_InverseCumulative", (0.5, 0.25), 2),   # CDF(1) = 0.4375, CDF(2) = 0.578
        ("Geometric_InverseCumulative", (0.2, 0.25), 0),
        ("Geometric_Mean", (0.25,), 3),
        ("Geometric_Variance", (0.25,), 12),
        ("Geometric_StdDev", (0.25,), mp.sqrt(12)),
        # NegativeBinomial(k failures, r successes, p) with r = 3, p = 1/2:
        # P(k) = C(k + r - 1, k) p^r (1-p)^k
        ("NegativeBinomial_PMF", (2, 3, 0.5), mp.mpf(6) / 32),
        ("NegativeBinomial_LogPMF", (2, 3, 0.5), mp.log(mp.mpf(6) / 32)),
        ("NegativeBinomial_Cumulative", (2, 3, 0.5), mp.mpf(1) / 2),
        ("NegativeBinomial_Survival", (2, 3, 0.5), mp.mpf(1) / 2),
        ("NegativeBinomial_InverseCumulative", (0.4, 3, 0.5), 2),   # CDF(1) = 0.3125
        ("NegativeBinomial_Mean", (3, 0.25), 9),
        ("NegativeBinomial_Variance", (3, 0.25), 36),
        ("NegativeBinomial_StdDev", (3, 0.25), 6),
        # Hypergeometric(k, n sample, K pop. successes, N pop. size), n = 3, K = 4, N = 10
        ("Hypergeometric_PMF", (1, 3, 4, 10), mp.mpf(1) / 2),
        ("Hypergeometric_LogPMF", (1, 3, 4, 10), -mp.log(2)),
        ("Hypergeometric_Cumulative", (1, 3, 4, 10), mp.mpf(2) / 3),
        ("Hypergeometric_Survival", (1, 3, 4, 10), mp.mpf(1) / 3),
        ("Hypergeometric_InverseCumulative", (0.5, 3, 4, 10), 1),
        # sample of 8 from 10 with K = 4: at least 2 successes, P(2) = C(4,2)C(6,6)/C(10,8)
        ("Hypergeometric_PMF", (2, 8, 4, 10), mp.mpf(6) / 45),
        ("Hypergeometric_Mean", (3, 4, 10), mp.mpf("1.2")),
        ("Hypergeometric_Variance", (3, 4, 10), mp.mpf("0.56")),
        ("Hypergeometric_StdDev", (3, 4, 10), mp.sqrt(mp.mpf("0.56"))),
        # DiscreteUniform(x, a, b), inclusive support a..b
        ("DiscreteUniform_PMF", (6, 1, 6), mp.mpf(1) / 6),
        ("DiscreteUniform_LogPMF", (6, 1, 6), -mp.log(6)),
        ("DiscreteUniform_Cumulative", (3, 1, 6), mp.mpf(1) / 2),
        ("DiscreteUniform_Cumulative", (3.7, 1, 6), mp.mpf(1) / 2),
        ("DiscreteUniform_Survival", (3, 1, 6), mp.mpf(1) / 2),
        ("DiscreteUniform_Cumulative", (-3, -5, -2), mp.mpf(3) / 4),
        ("DiscreteUniform_InverseCumulative", (0.4, 1, 6), 3),
        ("DiscreteUniform_InverseCumulative", (0.99, 1, 6), 6),
        ("DiscreteUniform_Mean", (1, 6), mp.mpf("3.5")),
        ("DiscreteUniform_Variance", (1, 6), mp.mpf(35) / 12),
        ("DiscreteUniform_StdDev", (1, 6), mp.sqrt(mp.mpf(35) / 12)),
        # Continuous
        ("ChiSquare_Density", (2, 2), 1 / (2 * e)),
        ("F_Density", (1, 2, 2), mp.mpf("0.25")),      # 1 / (1 + x)^2 for df 2, 2
        ("NormalStandard_InverseCumulative", (0.975,), z975),
        ("NormalStandard_InverseSurvival", (0.025,), z975),
        ("Normal_InverseSurvival", (0.025, 10, 2), 10 + 2 * z975),
        ("Lognormal_InverseSurvival", (0.5, 1, 0.5), e),
        ("Lognormal_InverseSurvival", (0.025, 0, 1), mp.exp(z975)),
        ("Beta_Density", (0.5, 2, 2), mp.mpf("1.5")),
        ("Beta_Cumulative", (0.25, 2, 2), mp.mpf(5) / 32),   # 3x^2 - 2x^3
        ("Beta_Survival", (0.25, 2, 2), mp.mpf(27) / 32),
        ("Beta_InverseCumulative", (mp.mpf(5) / 32, 2, 2), mp.mpf("0.25")),
        ("F_Cumulative", (1, 2, 2), mp.mpf("0.5")),           # x / (1 + x) for df 2, 2
        ("F_Survival", (3, 2, 2), mp.mpf("0.25")),
        ("F_InverseCumulative", (0.75, 2, 2), 3),
        ("PROB_LogBeta", (2, 3), -mp.log(12)),
    ]


def part2_known_values():
    try:
        evaluate, versions = ccs.scipy_evaluator()
    except ImportError as exc:
        print(f"SKIP: known-value checks need SciPy ({exc})")
        return
    print("known-value checks with " + ", ".join(f"{n} {v}" for n, v in versions))
    RAN_KNOWN_VALUES.append(True)
    covered = set()
    for fn, args, exact in known_values():
        covered.add(fn)
        row = _row(fn, mp.nstr(mp.mpf(exact), 40), *[repr(float(a)) for a in args])
        cls, detail = ccs.classify_row(fn, row, evaluate)
        exact_f = float(exact)
        if cls != ccs.CONFIRMED:
            FAILS.append(f"{fn}{args}: classified {cls} ({detail})")
            continue
        tol = 0.0 if fn.endswith("_InverseCumulative") and fn.split("_")[0] in (
            "Binomial", "Poisson", "Geometric", "NegativeBinomial", "Hypergeometric",
            "DiscreteUniform") else 1e-14
        check(detail <= tol, f"{fn}{args}: SciPy {float(evaluate(fn, [float(a) for a in args]))!r}"
                             f" vs exact {exact_f!r}, relative difference {detail:.2e}")
    missing = sorted(set(ccs.MAPPINGS) - covered)
    check(not missing, f"mappings without a known-value check: {missing}")
    overlap = sorted(set(ccs.MAPPINGS) & set(ccs.UNMAPPED_REASONS))
    check(not overlap, f"functions both mapped and listed as unmapped: {overlap}")


def main():
    part1_classification()
    part2_known_values()
    if FAILS:
        for f in FAILS:
            print("FAIL:", f)
        sys.exit(1)
    print("PASS: SciPy cross-check classification"
          + (" and known-value mappings" if RAN_KNOWN_VALUES else ""))


if __name__ == "__main__":
    main()
