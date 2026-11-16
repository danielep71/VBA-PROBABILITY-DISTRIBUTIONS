"""
#348: tables of the Gamma call profile, from gamma_call_profile.csv.

Nothing here is a verdict: the study measures and the README decides. The
program fails only when the file is not a complete export.

  layers    for each of the eight points, the time of each layer of a
            public cumulative and of a public survival call, with the cost
            of the timing loop (the part empty) subtracted, and what each
            layer adds to the one below it;
  parts     the parts that could be timed alone;
  inverse   for each shape and probability, the solver's iterations, which
            routine served each tail evaluation, and where the time of a
            public inverse goes.

A part reached through Application.Run includes what that call costs; the
part run_pow2 measures it and is subtracted. A part with ok=0 was refused or
raised and is shown as "-".

Usage: python3 analyze.py [--check]
--check prints nothing but the outcome.
"""
import argparse
import csv
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
FILE = os.path.join(HERE, "gamma_call_profile.csv")
RUN_PARTS = ("run_factor", "run_small_lg1p", "run_validate")


def read(path=FILE):
    with open(path, encoding="utf-8", newline="") as handle:
        return list(csv.DictReader(handle))


def note_fields(note):
    out = {}
    for item in note.split(";"):
        key, _, value = item.partition("=")
        out[key] = value
    return out


def problems(rows):
    """Why the file is not a complete export; empty when it is."""
    out = []
    for row in rows:
        where = f"{row['point_id']} ({row['part']})"
        if not row["note"].strip():
            out.append(f"{where}: not exported")
            continue
        try:
            int(row["calls"])
            float(row["us_per_call"])
        except ValueError:
            out.append(f"{where}: calls or us_per_call is not a number")
            continue
        if row["part"] == "solver":
            fields = note_fields(row["note"])
            if fields.get("same") != "1":
                out.append(f"{where}: the replica did not return the Double of the solver")
        elif row["part"] in ("empty", "public_cdf", "public_sf", "routed_p", "routed_q",
                             "public_inv", "inv_p") and row["note"] != "ok=1":
            out.append(f"{where}: a part every table needs returned no value")
    return out


def measured(rows):
    """{(kind, a, x, part): microseconds, or None where ok=0}."""
    out = {}
    for row in rows:
        ok = row["part"] == "solver" or row["note"] == "ok=1"
        out[(row["kind"], row["a"], row["x"], row["part"])] = \
            float(row["us_per_call"]) if ok else None
    return out


def net(values, key, empty, run=None):
    """A time with the loop cost, and for a Run part the Run cost, removed."""
    value = values.get(key)
    if value is None:
        return None
    if key[3] in RUN_PARTS:
        return None if run is None else value - empty - run
    return value - empty


def cell(value, digits=2):
    return "-" if value is None else f"{value:.{digits}f}"


def minus(a, b):
    return None if a is None or b is None else a - b


def tables(rows):
    values = measured(rows)
    empty = values[("fixed", "1", "1", "empty")]
    run = values.get(("fixed", "1", "1", "run_pow2"))
    run = None if run is None else run - empty
    points = []
    for row in rows:
        if row["kind"] == "call" and (row["a"], row["x"]) not in points:
            points.append((row["a"], row["x"]))

    def at(a, x, part):
        return net(values, ("call", a, x, part), empty, run)

    out = ["### Layers of a public call, microseconds", "",
           "| Shape | X | Public CDF | Public SF | Routed P | Routed Q | Binary64 kernel "
           "| Loop before #242 | Double-double P | Double-double Q |",
           "| ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |"]
    for a, x in points:
        out.append(f"| {a} | {x} | " + " | ".join(cell(at(a, x, part)) for part in (
            "public_cdf", "public_sf", "routed_p", "routed_q", "b64_kernel", "legacy_loop",
            "kernel_p", "kernel_q")) + " |")

    out += ["", "### What each layer adds, microseconds", "",
            "| Shape | X | Wrapper, CDF | Wrapper, SF | Route test | Kernel over the old loop "
            "| Prefactor alone | Validation alone | Stirling or small log-gamma |",
            "| ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |"]
    for a, x in points:
        routed = at(a, x, "routed_p" if float(x) < float(a) + 1.0 else "routed_q")
        small = at(a, x, "stirling") if float(a) > 1.0 else at(a, x, "run_small_lg1p")
        out.append(f"| {a} | {x} | " + " | ".join(cell(v) for v in (
            minus(at(a, x, "public_cdf"), at(a, x, "routed_p")),
            minus(at(a, x, "public_sf"), at(a, x, "routed_q")),
            minus(routed, at(a, x, "b64_kernel")),
            minus(at(a, x, "b64_kernel"), at(a, x, "legacy_loop")),
            at(a, x, "run_factor"), at(a, x, "run_validate"), small)) + " |")

    out += ["", "### Parts timed once, microseconds", "",
            "| Part | Time | Loop cost removed |", "| --- | ---: | ---: |"]
    for part in ("empty", "set_status", "try_divide", "classify", "is_finite", "run_pow2"):
        value = values.get(("fixed", "1", "1", part))
        out.append(f"| `{part}` | {cell(value, 3)} | "
                   f"{cell(None if value is None or part == 'empty' else value - empty, 3)} |")

    out += ["", "### The inverse solver", "",
            "| Shape | p | Public inverse, us | Solver alone, us | Iterations | Series "
            "| Fraction | Double-double | Tail evaluations, us | Density evaluations, us "
            "| Rest, us |",
            "| ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |"]
    for row in rows:
        if row["part"] != "solver":
            continue
        a, p = row["a"], row["x"]
        fields = note_fields(row["note"])
        public = net(values, ("inverse", a, p, "public_inv"), empty)
        solver = net(values, ("inverse", a, p, "inv_p"), empty)
        tails = float(row["us_per_call"])
        density = float(fields["pdf_us"])
        out.append(f"| {a} | {p} | {cell(public, 1)} | {cell(solver, 1)} | {row['calls']} | "
                   f"{fields['series']} | {fields['fraction']} | {fields['kernel']} | "
                   f"{tails:.1f} | {density:.1f} | {cell(minus(solver, tails + density), 1)} |")
    return out


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--check", action="store_true")
    a = ap.parse_args(argv)
    rows = read()
    found = problems(rows)
    for text in found:
        print(f"FAIL: {text}")
    if found:
        return 1
    if not a.check:
        print("\n".join(tables(rows)))
    print(f"PASS: {len(rows)} rows exported")
    return 0


if __name__ == "__main__":
    sys.exit(main())
