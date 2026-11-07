"""Replay the frozen B1 screen; reproducibility is distinct from acceptance."""

import argparse
import hashlib
import json
import math
import subprocess
import time
from pathlib import Path

import mpmath as mp
from model import HERE, ROOT, X, D, cf, pref, backward

PROTOCOL_COMMIT = "f95d3e04f9ce19749ab479e061966668c3d1901b"
PROTOCOL = HERE / "candidate_protocol.json"
DOCUMENT = HERE / "CANDIDATE_PREREGISTRATION.md"


def digest(path):
    return hashlib.sha256(path.read_bytes().replace(b"\r\n", b"\n")).hexdigest()


def frozen_bytes(path, commit):
    return subprocess.check_output(
        ["git", "show", f"{commit}:{path.relative_to(ROOT).as_posix()}"], cwd=ROOT
    ).replace(b"\r\n", b"\n")


def relative_error(value, reference):
    if not mp.isfinite(reference) or reference <= 0 or not math.isfinite(value):
        raise ValueError("unavailable, nonpositive or non-finite measurement")
    return mp.mpf(value) / reference - 1


def validate_oracles(reference, lower, quadrature, protocol):
    if any(not mp.isfinite(v) or v <= 0 for v in (reference, lower, quadrature)):
        raise ValueError("invalid reference")
    if reference < mp.mpf(2) ** -1022:
        raise ValueError("unexpected subnormal reference: no row may be dropped")
    stability = abs(lower / reference - 1)
    agreement = abs(quadrature / reference - 1)
    if stability > mp.mpf(protocol["reference_stability_limit"]):
        raise ValueError("reference did not stabilize")
    if agreement > mp.mpf(protocol["quadrature_agreement_limit"]):
        raise ValueError("independent quadrature disagrees")
    return stability, agreement


def check_guards(protocol):
    # Missing/unstable independent evidence must not produce a favourable score.
    with mp.workdps(90):
        for args in [
            (mp.nan, 1, 1),
            (1, 1, mp.nan),
            (1, 1, 0),
            (1, 1 + mp.mpf("1e-40"), 1),
            (1, 1, 1 + mp.mpf("1e-60")),
            (mp.mpf(2) ** -1023,) * 3,
        ]:
            try:
                validate_oracles(*args, protocol)
            except ValueError:
                continue
            raise AssertionError("invalid oracle accepted")
        for value, reference in [(math.nan, 1), (math.inf, 1), (1, 0)]:
            try:
                relative_error(value, reference)
            except ValueError:
                continue
            raise AssertionError("invalid measurement scored")
        threshold = mp.mpf(protocol["relative_threshold"])
        if relative_error(1.0, mp.mpf(1)) != 0:
            raise AssertionError("clean fixture failed")
        if abs(relative_error(1.0 + 1e-14, mp.mpf(1))) <= threshold:
            raise AssertionError("degraded fixture was accepted")


def coordinates():
    baseline = json.loads((HERE / "isolation.json").read_text())
    fitting = [(r["point_id"], r["shape"], r["x"]) for r in baseline["rows"]]
    fitting += [
        (f"expanded_{i}", r["shape"], r["x"])
        for i, r in enumerate(baseline["expanded_fitting"])
        if r["shape"] <= 0.5
    ]
    protocol = json.loads(PROTOCOL.read_text())
    holdout = [
        (r["id"], float.fromhex(r["shape_hex"]), float.fromhex(r["x_hex"]))
        for r in protocol["holdout"]
    ]
    identities = lambda rows: {(a.hex(), x.hex()) for _, a, x in rows}
    if (
        len(fitting) != 84
        or len(holdout) != protocol["holdout_count"]
        or len(holdout) != 84
    ):
        raise ValueError("wrong frozen coverage")
    if len(identities(holdout)) != 84 or identities(fitting) & identities(holdout):
        raise ValueError("duplicate or contaminated holdout coordinates")
    return protocol, fitting, holdout


def measure():
    protocol, fitting, holdout = coordinates()
    for path in (PROTOCOL, DOCUMENT):
        if frozen_bytes(path, PROTOCOL_COMMIT) != path.read_bytes().replace(
            b"\r\n", b"\n"
        ):
            raise ValueError("frozen preregistration was modified")
    model = HERE / "model.py"
    if frozen_bytes(model, protocol["model_commit"]) != model.read_bytes().replace(
        b"\r\n", b"\n"
    ):
        raise ValueError(
            "candidate model drifted from its preregistered implementation"
        )
    check_guards(protocol)
    rows = []
    with mp.workdps(protocol["reference_dps"][1]):
        for role, points in (("fitting", fitting), ("independent", holdout)):
            for point_id, a, x in points:
                am, xm = mp.mpf(a), mp.mpf(x)
                reference = mp.gammainc(am, xm, mp.inf, regularized=True)
                with mp.workdps(protocol["reference_dps"][0]):
                    lower = mp.gammainc(mp.mpf(a), mp.mpf(x), mp.inf, regularized=True)
                quadrature = mp.quad(
                    lambda t: mp.exp(-t) * t ** (am - 1), [xm, mp.inf]
                ) / mp.gamma(am)
                stability, agreement = validate_oracles(
                    reference, lower, quadrature, protocol
                )
                factor, _ = pref(a, x)
                old, old_n = cf(a, x, budget=protocol["iteration_cap"])
                tight, n = cf(
                    a,
                    x,
                    eps=float.fromhex(protocol["epsilon_hex"]),
                    budget=protocol["iteration_cap"],
                )
                candidate = backward(a, x, n)
                errors = {
                    name: relative_error(D(X(factor) * X(h)), reference)
                    for name, h in (
                        ("production", old),
                        ("tight_stop_only", tight),
                        ("B1", candidate),
                    )
                }
                accepted = abs(errors["B1"]) <= mp.mpf(protocol["relative_threshold"])
                rows.append(
                    dict(
                        role=role,
                        id=point_id,
                        shape_hex=a.hex(),
                        x_hex=x.hex(),
                        old_iterations=old_n,
                        candidate_iterations=n,
                        reference=mp.nstr(reference, 60),
                        reference_stability=mp.nstr(stability, 15),
                        quadrature_agreement=mp.nstr(agreement, 15),
                        relative_errors={k: mp.nstr(v, 30) for k, v in errors.items()},
                        candidate_pass=accepted,
                    )
                )
    failed = [f"{r['role']}:{r['id']}" for r in rows if not r["candidate_pass"]]
    maxima = {
        role: max(
            (r for r in rows if r["role"] == role),
            key=lambda r: abs(mp.mpf(r["relative_errors"]["B1"])),
        )
        for role in ("fitting", "independent")
    }
    paths = [
        PROTOCOL,
        DOCUMENT,
        model,
        Path(__file__),
        HERE / "isolation.json",
        HERE.parent / "beta_balanced_survival/beta_route_model.py",
        ROOT / "src/M_STATS_PROBDIST_SPECIALFUNCS.bas",
        ROOT / "src/M_STATS_PROBDIST_CORE.bas",
    ]
    return dict(
        protocol_commit=PROTOCOL_COMMIT,
        status="REJECTED" if failed else "MODEL SCREEN PASSED - EXCEL NOT RUN",
        scope="B1 feasibility only; not production acceptance or issue closure",
        bindings={str(p.relative_to(ROOT)): digest(p) for p in paths},
        fitting_count=len(fitting),
        independent_count=len(holdout),
        failed_ids=failed,
        worst={
            role: dict(id=r["id"], relative_error=r["relative_errors"]["B1"])
            for role, r in maxima.items()
        },
        rows=rows,
    )


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--write", action="store_true")
    args = parser.parse_args()
    started = time.monotonic()
    result = measure()
    expected = json.dumps(result, indent=2) + "\n"
    path = HERE / "candidate_results.json"
    if args.write:
        path.write_text(expected)
    elif not path.exists() or path.read_text() != expected:
        raise SystemExit("FAIL: candidate report missing or stale")
    print(
        f"PASS: report reproduced; numerical status={result['status']}; "
        f"failures={len(result['failed_ids'])}; elapsed={time.monotonic() - started:.1f}s"
    )
    print(json.dumps(result["worst"]))


if __name__ == "__main__":
    main()
