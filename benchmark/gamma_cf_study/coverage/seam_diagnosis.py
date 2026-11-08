import sys, math, re, json
from pathlib import Path
import mpmath as mp

ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(ROOT / "benchmark/gamma_cf_study"))
from model import X, D, pref, cf, backward

s = (ROOT / "src/M_STATS_PROBDIST_SPECIALFUNCS.bas").read_text()
s = s.split("Private Function PROB_GammaSmallLogGamma1p")[1].split("End Function")[0]
cs = []
for i in range(33):
    t = re.search(r"Const C" + str(i) + r" As Double = (.*)", s)[1]
    cs.append(sum(float(v) for v in re.findall(r"[-+]?\d+\.\d+E[-+]\d+", t)))


def lg(a):
    u = D(X(a) - X(0.5))
    v = cs[-1]
    for c in reversed(cs[:-1]):
        v = D(X(v) * X(u) + X(c))
    return D(X(D(X(a) * (X(a) - 1))) * X(v))


def upper(a, x):
    h, n = cf(a, x, eps=2**-52)
    h = backward(a, x, n)
    e = D(X(a) * X(math.log(x)) - X(lg(a)))
    return D(X(D(X(D(X(a) * X(math.exp(-x)))) * X(math.exp(e)))) * X(h))


def series(a, x, kahan=False):
    ap = a
    v = D(1 / X(a))
    d = v
    c = 0.0
    for i in range(1, 100001):
        ap = D(X(ap) + 1)
        d = D(X(d) * X(x) / X(ap))
        if kahan:
            y = D(X(d) - X(c))
            t = D(X(v) + X(y))
            c = D((X(t) - X(v)) - X(y))
            v = t
        else:
            v = D(X(v) + X(d))
        if abs(d) <= abs(v) * 2**-52:
            return v, i
    raise RuntimeError("cap")


def public(a, x, compensated=False):
    if x < a + 1:
        p = D(X(series(a, x, compensated and a <= 1)[0]) * X(pref(a, x)[0]))
        p = min(1.0, max(0.0, p))
        return p, 1 - p
    q = upper(a, x) if a <= 1 else D(X(cf(a, x)[0]) * X(pref(a, x)[0]))
    q = min(1.0, max(0.0, q))
    return 1 - q, q


def diagnose():
    rows = []
    with mp.workdps(120):
        for x in [math.nextafter(1.5, 0), 1.5, math.nextafter(1.5, math.inf)]:
            a = 0.5
            p = mp.gammainc(a, 0, mp.mpf(x), regularized=True)
            q = 1 - p
            factor = mp.exp(a * mp.log(x) - x - mp.loggamma(a))
            v, n = series(a, x)
            f, _ = pref(a, x)
            variants = {
                "lower": D(X(v) * X(f)),
                "exact_prefactor": D(X(v) * X(float(factor))),
                "exact_sum": D(X(float(p / factor)) * X(f)),
                "kahan": D(X(series(a, x, True)[0]) * X(f)),
            }
            rows.append(
                {
                    "x_hex": x.hex(),
                    "reference_p": mp.nstr(p, 100),
                    "reference_q": mp.nstr(q, 100),
                    "iterations": n,
                    "sum_relative_error": mp.nstr(mp.mpf(v) / (p / factor) - 1, 40),
                    "prefactor_relative_error": mp.nstr(mp.mpf(f) / factor - 1, 40),
                    "variants": {
                        k: {
                            "hex": z.hex(),
                            "value": repr(z),
                            "p_error_ulps": float((mp.mpf(z) - p) / math.ulp(float(p))),
                        }
                        for k, z in variants.items()
                    },
                    "b6": list(public(a, x)),
                    "b7": list(public(a, x, True)),
                }
            )
    return rows


if __name__ == "__main__":
    print(json.dumps(diagnose(), indent=2))
