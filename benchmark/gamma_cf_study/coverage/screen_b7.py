"""Execute the frozen B7 differential screen; no production source is changed."""

import hashlib
import json
import platform
import math
from pathlib import Path
import time
import mpmath as mp
import seam_diagnosis as model

HERE = Path(__file__).resolve().parent


def reference(a, x):
    with mp.workdps(90):
        lo = [
            mp.gammainc(mp.mpf(a), 0, mp.mpf(x), regularized=True),
            mp.gammainc(mp.mpf(a), mp.mpf(x), mp.inf, regularized=True),
        ]
    with mp.workdps(120):
        a, x = mp.mpf(a), mp.mpf(x)
        hi = [
            mp.gammainc(a, 0, x, regularized=True),
            mp.gammainc(a, x, mp.inf, regularized=True),
        ]
        if any(not mp.isfinite(v) or v <= 0 for v in hi):
            raise ValueError("invalid oracle")
        if max(abs(l / h - 1) for l, h in zip(lo, hi)) > mp.mpf("1e-80"):
            raise ValueError("unstable oracle")
        splits = sorted(
            {mp.mpf(0), mp.mpf(1)}
            | {1 - a * k for k in [1, 4, 16, 64, 256] if a * k < 1}
        )
        pquad = (
            x**a
            / mp.gamma(a + 1)
            * mp.quad(lambda u: mp.exp(-x * u ** (1 / a)), splits)
        )
        qquad = mp.quad(lambda t: mp.exp(-t) * t ** (a - 1), [x, mp.inf]) / mp.gamma(a)
        agreement = max(abs(pquad / hi[0] - 1), abs(qquad / hi[1] - 1))
        if not mp.isfinite(agreement) or agreement > mp.mpf("1e-80"):
            raise ValueError("independent quadrature disagreement: " + str(agreement))
        return hi, mp.nstr(agreement, 30)


def main():
    started = time.monotonic()
    bound = [
        HERE / "B7_PROTOCOL.md",
        HERE / "screen_b7.py",
        HERE / "seam_diagnosis.py",
        HERE.parent / "model.py",
        HERE.parents[1] / "beta_balanced_survival/beta_route_model.py",
    ]
    bound += list((model.ROOT / "src").glob("M_STATS_PROBDIST_*.bas"))
    hashes = {
        str(p.relative_to(model.ROOT)): hashlib.sha256(
            p.read_bytes().replace(b"\r\n", b"\n")
        ).hexdigest()
        for p in bound
    }
    design = json.loads((HERE / "protocol.json").read_text())["points"]
    sets = {
        "fitting": [
            (float.fromhex(r["a_hex"]), float.fromhex(r["x_hex"])) for r in design
        ]
    }
    independent = []
    for a in [1e-72, 9e-9, 0.0071, 0.021, 0.083, 0.19, 0.34, 0.61, 0.88, 0.987]:
        s = 1 + a
        xs = [s]
        for _ in range(8):
            xs.insert(0, math.nextafter(xs[0], 0))
            xs.append(math.nextafter(xs[-1], math.inf))
        independent.extend((a, x) for x in [s * 0.17, s * 0.43, s * 0.81] + xs)
    sets["independent"] = independent
    report = {
        "status": "UNCOMPLETE",
        "rows": [],
        "new_accuracy_failures": [],
        "new_monotonicity_failures": [],
        "oracle_failures": [],
    }
    for role, points in sets.items():
        previous = {}
        for i, (a, x) in enumerate(points):
            b6 = model.public(a, x)
            b7 = model.public(a, x, True)
            row = {
                "role": role,
                "index": i,
                "a_hex": a.hex(),
                "x_hex": x.hex(),
                "b6": b6,
                "b7": b7,
            }
            try:
                ref, agreement = reference(a, x)
                with mp.workdps(120):
                    row["reference"] = [mp.nstr(v, 100) for v in ref]
                    row["quadrature_relative_disagreement"] = agreement
                    row["errors"] = {
                        name: [float(abs(mp.mpf(v) / r - 1)) for v, r in zip(vals, ref)]
                        for name, vals in [("b6", b6), ("b7", b7)]
                    }
                for j, tol in enumerate([3e-15, 5e-15]):
                    if row["errors"]["b6"][j] <= tol < row["errors"]["b7"][j]:
                        report["new_accuracy_failures"].append([role, i, j])
            except (ValueError, RuntimeError) as exc:
                row["oracle_error"] = str(exc)
                report["oracle_failures"].append([role, i, str(exc)])
            if a in previous:
                px, old, new = previous[a]
                if px < x and math.nextafter(px, math.inf) == x:
                    for j in [0, 1]:
                        ok6 = (old[j] <= b6[j]) if j == 0 else (old[j] >= b6[j])
                        ok7 = (new[j] <= b7[j]) if j == 0 else (new[j] >= b7[j])
                        if ok6 and not ok7:
                            report["new_monotonicity_failures"].append(
                                [role, i, j, px.hex(), x.hex()]
                            )
            previous[a] = (x, b6, b7)
            report["rows"].append(row)
        print(role, "evaluated", len(points), flush=True)
    left = model.public(0.5, math.nextafter(1.5, 0), True)
    right = model.public(0.5, 1.5, True)
    report["original_new_seam_fixed"] = left[0] <= right[0] and left[1] >= right[1]
    report["status"] = (
        "REJECT"
        if any(
            report[k]
            for k in [
                "new_accuracy_failures",
                "new_monotonicity_failures",
                "oracle_failures",
            ]
        )
        or not report["original_new_seam_fixed"]
        else "DIFFERENTIAL SCREEN PASS; NOT FULL ACCEPTANCE"
    )
    report["runtime_s"] = time.monotonic() - started
    report["provenance"] = {
        "protocol_commit": "02303d694af180354785abbd65ce8fe4f624cf73",
        "python": platform.python_version(),
        "mpmath": mp.__version__,
        "sha256_lf": hashes,
    }
    if any(
        hashlib.sha256(p.read_bytes().replace(b"\r\n", b"\n")).hexdigest()
        != hashes[str(p.relative_to(model.ROOT))]
        for p in bound
    ):
        raise RuntimeError("input changed during screen")
    (HERE / "b7_screen.json").write_text(json.dumps(report, indent=2) + "\n")
    print(
        report["status"],
        {
            k: len(report[k])
            for k in [
                "new_accuracy_failures",
                "new_monotonicity_failures",
                "oracle_failures",
            ]
        },
    )

    return 1 if report["status"] == "REJECT" else 0


if __name__ == "__main__":
    raise SystemExit(main())
