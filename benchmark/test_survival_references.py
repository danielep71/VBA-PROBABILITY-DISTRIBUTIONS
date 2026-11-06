"""Fixtures for the survival-reference audit (#159).

Negative proof that check_survival_references.py fails closed, plus route
cross-checks that keep its oracle honest:

  * a reference stored at the decimal token rather than at the binary64
    argument fails (the #159 defect), and the same row re-referenced passes;
  * a survival formed as 1 - CDF at 53 bits fails (the tchi_large_df defect);
  * a reference one digit short of the floor fails, one at 25-digit
    serialisation passes;
  * registered debt passes only while it matches a failing cell exactly; a
    stale or orphaned entry fails, as does an entry with an undefined blocker;
  * an unknown survival function, an unlisted grid with survival rows, a listed
    grid without one, and an oracle whose two precisions disagree all fail;
  * the pipe-delimited excel_comparison layout is read;
  * quadrature agrees with an independent Lentz continued fraction where both
    beta shapes are huge, with mpmath.betainc where both routes apply, and with
    Temme's uniform expansion where the gamma shape is huge.
"""
import contextlib
import io
import json
import os
import tempfile

import mpmath as mp

import _survival_oracle as O
import check_survival_references as C

fails = []


def check(condition, message):
    if not condition:
        fails.append(message)


FIELDS = ["function", "vba_kernel", "claim", "metric", "arg1", "arg2", "arg3",
          "reference", "observed_vba", "regime", "evidence_set"]


def grid_text(rows):
    lines = [",".join(FIELDS)]
    for fn, args, ref in rows:
        args = list(args) + [""] * (3 - len(args))
        lines.append(",".join([fn, "K_STATS_" + fn, "c", "rel", *args, ref, "", "all", "fixture"]))
    return "\n".join(lines) + "\n"


def run(files, grids, debt=None, patch=None):
    """Run the audit over a temporary tree. Returns (exit code, output)."""
    with tempfile.TemporaryDirectory() as root:
        for rel, text in files.items():
            path = os.path.join(root, rel)
            os.makedirs(os.path.dirname(path), exist_ok=True)
            with open(path, "w", encoding="utf-8", newline="") as fh:
                fh.write(text)
        if debt is not None:
            with open(os.path.join(root, "debt.json"), "w", encoding="utf-8") as fh:
                json.dump(debt, fh)
        buf = io.StringIO()
        saved = O.evaluate
        if patch is not None:
            O.evaluate = patch
        try:
            with contextlib.redirect_stdout(buf):
                code = C.run(root=root, grids=grids, debt_path=os.path.join(root, "debt.json"),
                             jobs=1, out=print)
        finally:
            O.evaluate = saved
        return code, buf.getvalue()


def s(v, digits=34):
    return mp.nstr(v, digits)


mp.mp.dps = 60
Q = O.normal_upper
binary64 = O.exact_args(["1.96"])[0]
at_double = s(Q(binary64))
at_token = s(Q(mp.mpf("1.96")))          # the decimal 1.96, not the Double
check(at_double != at_token, "fixture: the two references differ")

# 1. the #159 defect: a reference at the decimal token fails; at the Double passes
code, text = run({"g.csv": grid_text([("NormalStandard_Survival", ["1.96"], at_token)])}, ["g.csv"])
check(code == 1 and "below the 25-digit floor" in text,
      "a reference at the decimal token fails")
code, text = run({"g.csv": grid_text([("NormalStandard_Survival", ["1.96"], at_double)])}, ["g.csv"])
check(code == 0 and "PASS" in text, "the same row at the binary64 argument passes")

# the same defect on a Beta row: 1 - CDF at the decimal 0.999 is exactly 3.997E-9
code, text = run({"g.csv": grid_text([("Beta_Survival", ["0.999", "2", "3"], "3.997e-9")])},
                 ["g.csv"])
check(code == 1, "Beta survival taken at the decimal 0.999 fails")
beta_ok = s(mp.betainc(3, 2, 0, 1 - O.exact_args(["0.999"])[0], regularized=True))
code, _ = run({"g.csv": grid_text([("Beta_Survival", ["0.999", "2", "3"], beta_ok)])}, ["g.csv"])
check(code == 0, "Beta survival at the Double 0.999 passes")

# 2. the tchi_large_df defect: 1 - CDF formed at 53 bits
with mp.workdps(15):
    sf53 = 1 - mp.ncdf(1)                # 1 - CDF at mpmath's default 53 bits
code, _ = run({"g.csv": grid_text([("NormalStandard_Survival", ["1.0"], s(sf53))])},
              ["g.csv"])
check(code == 1, "a survival accurate only to 53 bits fails")

# 3. the floor: 25-digit serialisation passes, 23 digits fails
v = Q(binary64)
code, _ = run({"g.csv": grid_text([("NormalStandard_Survival", ["1.96"], s(v, 25))])}, ["g.csv"])
check(code == 0, "a reference serialised to 25 digits meets the floor")
code, _ = run({"g.csv": grid_text([("NormalStandard_Survival", ["1.96"],
                                    s(v * (1 + mp.mpf("3e-24"))))])}, ["g.csv"])
check(code == 1, "a reference 3E-24 away fails")

# 4. registered debt
blockers = {"fixture-binding": "a binding that only a fresh export may rebind"}
entry = {"grid": "g.csv", "function": "NormalStandard_Survival", "arguments": ["1.96"],
         "reference": at_token, "blocker": "fixture-binding"}
debt = {"schema_version": 1, "blockers": blockers, "entries": [entry]}
code, text = run({"g.csv": grid_text([("NormalStandard_Survival", ["1.96"], at_token)])},
                 ["g.csv"], debt)
check(code == 0 and "1 are registered debt" in text, "a registered failing cell passes as debt")
code, text = run({"g.csv": grid_text([("NormalStandard_Survival", ["1.96"], at_double)])},
                 ["g.csv"], debt)
check(code == 1 and "stale debt entry" in text,
      "debt left behind after the cell was re-referenced fails")
passing = {"schema_version": 1, "blockers": blockers,
           "entries": [dict(entry, reference=at_double)]}
code, text = run({"g.csv": grid_text([("NormalStandard_Survival", ["1.96"], at_double)])},
                 ["g.csv"], passing)
check(code == 1 and "now meets the floor" in text, "debt for a cell that meets the floor fails")
code, text = run({"g.csv": grid_text([("NormalStandard_Survival", ["2.5"], s(Q(mp.mpf(2.5))))])},
                 ["g.csv"], debt)
check(code == 1 and "stale debt entry" in text, "debt for a removed row fails")
other = dict(entry, reference=s(Q(mp.mpf("1.9600000000000001"))))
code, text = run({"g.csv": grid_text([("NormalStandard_Survival", ["1.96"], at_token)])},
                 ["g.csv"], {"schema_version": 1, "blockers": blockers, "entries": [other]})
check(code == 1 and "below the 25-digit floor" in text,
      "debt covers one exact reference text, not the row")
code, text = run({"g.csv": grid_text([("NormalStandard_Survival", ["1.96"], at_token)])},
                 ["g.csv"], {"schema_version": 1, "blockers": {},
                             "entries": [entry]})
check(code == 1 and "undefined blocker" in text, "a debt entry needs a defined blocker")

# 5. structural failures
code, text = run({"g.csv": grid_text([("Foo_Survival", ["1"], "0.5")])}, ["g.csv"])
check(code == 1 and "no direct survival route" in text, "an unknown survival function fails")
code, text = run({"g.csv": grid_text([("NormalStandard_Survival", ["1.96"], at_double)]),
                  "study/new.csv": grid_text([("NormalStandard_Survival", ["1.96"], at_double)])},
                 ["g.csv"])
check(code == 1 and "not in GRIDS" in text, "an unlisted grid with survival rows fails")
code, text = run({"g.csv": grid_text([("NormalStandard_Cumulative", ["1.96"], "0.975")])},
                 ["g.csv"])
check(code == 1 and "no survival rows" in text, "a listed grid without survival rows fails")
code, text = run({}, ["g.csv"])
check(code == 1 and "missing" in text, "a missing listed grid fails")


def unstable(function, args, dps):
    return mp.mpf(1) / 3 + mp.mpf(10) ** -(dps - 40)


code, text = run({"g.csv": grid_text([("NormalStandard_Survival", ["1.96"], at_double)])},
                 ["g.csv"], patch=unstable)
check(code == 1 and "oracle not certified" in text, "an unstable oracle fails closed")

# 6. the pipe-delimited excel_comparison layout
pipe = ("label|excel_formula|kstats_call|reference|observed_excel|observed_kstats|why\n"
        f"x|-|K_STATS_NormalStandard_Survival(1.96)|{at_token}|0|0|w\n")
code, text = run({"xl.csv": pipe}, ["xl.csv"])
check(code == 1 and "xl.csv:2" in text, "a pipe-delimited grid is read and audited")

# 7. route cross-checks
mp.mp.dps = 60


def lentz_beta(a, b, x):
    """Lentz continued fraction for I_x(a, b) on the side x < (a+1)/(a+b+2)."""
    tiny, eps = mp.mpf(10) ** -150, mp.mpf(10) ** -(mp.mp.dps - 3)
    qab, qap, qam = a + b, a + 1, a - 1
    c, d = mp.mpf(1), 1 - qab * x / qap
    d = 1 / d
    h, m = d, 1
    while True:
        m2 = 2 * m
        for aa in (m * (b - m) * x / ((qam + m2) * (a + m2)),
                   -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))):
            d = 1 + aa * d
            d = 1 / (d if abs(d) > tiny else tiny)
            c = 1 + aa / c
            c = c if abs(c) > tiny else tiny
            h *= d * c
        if abs(d * c - 1) < eps:
            break
        m += 1
    front = mp.e ** (a * mp.log(x) + b * mp.log1p(-x)
                     - (mp.loggamma(a) + mp.loggamma(b) - mp.loggamma(a + b)))
    return front * h / a


def digits(a, b):
    return mp.inf if a == b else -mp.log10(abs(a - b) / abs(b))


with mp.workdps(80):
    for a, b, z in ((mp.mpf(10) ** 8, mp.mpf(10) ** 8, 1), (mp.mpf(10) ** 6, 3 * mp.mpf(10) ** 6, 2)):
        sd = mp.sqrt(a * b / ((a + b) ** 2 * (a + b + 1)))
        x = mp.mpf(float(a / (a + b) + z * sd))
        quad = O.beta_upper(1 - x, a, b)               # quadrature: both shapes huge
        cf = lentz_beta(b, a, 1 - x)                   # I_{1-x}(b, a), the converging side
        check(digits(quad, cf) > 45, f"quadrature agrees with Lentz at a={a}, b={b}")
    for a, b, x in ((mp.mpf(25), mp.mpf(40), mp.mpf("0.45")), (mp.mpf(30), mp.mpf(300), mp.mpf("0.12"))):
        quad = O.beta_lower_quad(1 - x, b, a)
        hyp = mp.betainc(b, a, 0, 1 - x, regularized=True)
        check(digits(quad, hyp) > 45, f"quadrature agrees with betainc at a={a}, b={b}")
    for a in (mp.mpf(5) * 10 ** 15, mp.mpf(5) * 10 ** 29):
        for zz in (-3, 0, 3):
            z = a + zz * mp.sqrt(a)
            quad = O.gamma_upper_quad(a, z)
            lam1 = (z - a) / a
            if lam1 == 0:
                eta, c0, c1 = mp.mpf(0), mp.mpf(-1) / 3, mp.mpf(-1) / 540
            else:
                with mp.workdps(200):
                    eta = mp.sign(lam1) * mp.sqrt(2 * (lam1 - mp.log1p(lam1)))
                    c0 = 1 / lam1 - 1 / eta
                    c1 = 1 / eta ** 3 - 1 / lam1 ** 3 - 1 / lam1 ** 2 - 1 / (12 * lam1)
            temme = (mp.erfc(eta * mp.sqrt(a / 2)) / 2
                     + mp.e ** (-a * eta * eta / 2) / mp.sqrt(2 * mp.pi * a) * (c0 + c1 / a))
            check(digits(quad, temme) > 35, f"gamma quadrature agrees with Temme at a={a}, z={zz}")
    a = mp.mpf(10) ** 8
    z = a + 2 * mp.sqrt(a)
    with mp.workdps(90):
        check(digits(O.gamma_upper_quad(a, z), mp.gammainc(a, z, mp.inf, regularized=True)) > 45,
              "gamma quadrature agrees with mpmath gammainc at a = 1E8")

if fails:
    print("FAIL: survival-reference audit fixtures")
    for failure in fails:
        print("  - " + failure)
    raise SystemExit(1)
print("PASS: survival-reference audit (decimal-token and 53-bit references fail, the floor "
      "boundary, exact-cell debt and stale debt, unknown function, unlisted and empty "
      "grids, unstable oracle, pipe layout, and quadrature cross-checked against Lentz, "
      "betainc and Temme)")
