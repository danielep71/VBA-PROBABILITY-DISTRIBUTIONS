"""
Fixtures for the #17 grid-regeneration safeguards.

Each safeguard is proved against temporary grids, never against the committed
files, with two deliberate exceptions that exercise the real default command
against the real committed grid and verify it is left byte-identical:

  G0  the documented default `python generate_reference_values.py` writes
      nothing and leaves the committed grid unchanged;
  G3  `--out` aimed at the committed grid is refused before anything is built.

Both snapshot the committed grid first and restore it if a regression ever
does write to it, so a failing run cannot leave the evidence damaged.

Run: python3 test_grid_regeneration.py   (exit 0 = pass, nonzero = fail)
"""
import contextlib
import copy
import csv
import importlib.util
import io
import json
import os
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import _grid_safety as gs  # noqa: E402

fails = []


def check(cond, label):
    if not cond:
        fails.append(label)


def load(name, filename):
    spec = importlib.util.spec_from_file_location(name, os.path.join(HERE, filename))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def row(fn, arg1, ref, obs="", claim="rel<=1E-15", metric="rel", regime="all",
        evidence_set="main grid"):
    r = {f: "" for f in gs.GRID_FIELDS}
    r.update(function=fn, vba_kernel=fn, claim=claim, metric=metric, arg1=arg1,
             reference=ref, observed_vba=obs, regime=regime,
             evidence_set=evidence_set)
    return r


def read_bytes(path):
    with open(path, "rb") as fh:
        return fh.read()


@contextlib.contextmanager
def guard_committed_grid(label, path=gs.COMMITTED_GRID):
    """Fail and restore if a committed grid changes inside the block."""
    before = read_bytes(path)
    try:
        yield
    finally:
        if read_bytes(path) != before:
            with open(path, "wb") as fh:
                fh.write(before)
            fails.append(f"{label}: {os.path.relpath(path, HERE)} was modified (restored)")


def quiet(fn, *args):
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        rc = fn(*args)
    return rc, buf.getvalue()


# ---------------------------------------------------------------------------
# K. Canonical-key parity. Four tools keep their own key(); all must agree with
# _grid_safety.canonical_key, including across decimal spellings of one Double.
# ---------------------------------------------------------------------------
samples = [
    row("Fixture_A", "0.85", "1"),
    row("Fixture_A", "0.84999999999999998", "1"),   # same binary64 as 0.85
    dict(row("Fixture_B", "1E-300", "2"), arg2="5", arg3="", arg4="7.5"),
    row("Fixture_C", "3", "4", regime="deep_tail", evidence_set="holdout"),
]
check(gs.canonical_key(samples[0]) == gs.canonical_key(samples[1]),
      "K: decimal spellings of one Double share a canonical key")
cov = load("cgc_fixture", "check_grid_coverage.py")
for mod_name, filename in (("recon_fixture", "reconcile_grid.py"),
                           ("promote_fixture", "promote_grid_rows.py"),
                           ("migrate_fixture", "migrate_references.py"),
                           ("checkpromo_fixture", "check_promotion.py")):
    mod = load(mod_name, filename)
    check(all(mod.key(s) == gs.canonical_key(s) for s in samples),
          f"K: {filename} key() agrees with canonical_key")
check(all(cov.identity_key(cov.row_identity(s)) == gs.canonical_key(s)
          for s in samples),
      "K: check_grid_coverage identity agrees with canonical_key")


# ---------------------------------------------------------------------------
# G. generate_reference_values.py
# ---------------------------------------------------------------------------
gen = load("grv_fixture", "generate_reference_values.py")
real_build_rows = gen.build_rows

with tempfile.TemporaryDirectory() as tmp:
    grid = os.path.join(tmp, "grid.csv")
    grid_rows = [row("Fixture_A", "0.5", "1", obs="1.0E+000;0"),
                 row("Fixture_A", "0.25", "2", obs="2.0E+000;0"),
                 row("Fixture_B", "1", "3", obs="3.0E+000;0")]
    gs.write_grid(grid, grid_rows)
    grid_before = read_bytes(grid)

    # G1. --out naming the grid under comparison is refused before any build.
    built = []
    gen.build_rows = lambda: built.append(1) or []
    rc, out = quiet(gen.main, ["--grid", grid, "--out", grid])
    check(rc == 2 and "REFUSING" in out, "G1: --out equal to --grid is refused")
    check(not built, "G1: refusal happens before the reference build")
    check(read_bytes(grid) == grid_before, "G1: the grid is byte-identical")

    # G2. --out naming any other file that carries observations is refused.
    copy_path = os.path.join(tmp, "renamed_copy.csv")
    gs.write_grid(copy_path, grid_rows)
    copy_before = read_bytes(copy_path)
    rc, out = quiet(gen.main, ["--grid", grid, "--out", copy_path])
    check(rc == 2 and "3 observation(s)" in out,
          "G2: an observation-bearing copy under another name is refused")
    check(read_bytes(copy_path) == copy_before, "G2: the copy is byte-identical")

    # G3. --out aimed at the committed grid itself, as the old default did.
    with guard_committed_grid("G3"):
        rc, out = quiet(gen.main, ["--out", gs.COMMITTED_GRID])
    check(rc == 2 and "authoritative grid" in out,
          "G3: --out naming the committed grid is refused")
    with guard_committed_grid("G3h", gs.COMMITTED_HOLDOUT):
        rc, out = quiet(gen.main, ["--out", gs.COMMITTED_HOLDOUT])
    check(rc == 2, "G3h: --out naming the committed holdout grid is refused")

    # G4. a fresh non-authoritative --out is written: blank observations,
    # canonical header, LF endings; grid-only rows are reported as retained.
    fresh_rows = [row("Fixture_A", "0.5", "1.5", obs="leak"),   # reference differs
                  row("Fixture_A", "0.25", "2"),                # unchanged
                  row("Fixture_D", "9", "9")]                   # generator only
    gen.build_rows = lambda: copy.deepcopy(fresh_rows)
    fresh = os.path.join(tmp, "reference_rows.csv")
    rc, out = quiet(gen.main, ["--grid", grid, "--out", fresh])
    check(rc == 0, "G4: writing a fresh non-authoritative file succeeds")
    check(read_bytes(grid) == grid_before, "G4: the compared grid is untouched")
    if os.path.exists(fresh):
        raw = read_bytes(fresh)
        written = list(csv.DictReader(io.StringIO(raw.decode("utf-8"))))
        check(b"\r" not in raw, "G4: output uses LF line endings")
        check(raw.decode("utf-8").splitlines()[0].split(",") == gs.GRID_FIELDS,
              "G4: output carries the canonical header")
        check(len(written) == 3 and all(r["observed_vba"] == "" for r in written),
              "G4: every observed_vba in reference output is blank")
    else:
        check(False, "G4: the fresh output file was written")
    check("reference differs            1" in out and
          "generator only               1" in out and
          "grid only, retained          1" in out,
          "G4: report counts reference change, generator-only and retained rows")

    # G5. duplicate canonical keys from the generator fail hard and write nothing.
    dup_rows = [row("Fixture_A", "0.85", "1"),
                row("Fixture_A", "0.84999999999999998", "1")]
    gen.build_rows = lambda: copy.deepcopy(dup_rows)
    dup_out = os.path.join(tmp, "dup_rows.csv")
    rc, out = quiet(gen.main, ["--grid", grid, "--out", dup_out])
    check(rc == 1 and "duplicate canonical key" in out,
          "G5: duplicate generator keys fail hard")
    check(not os.path.exists(dup_out), "G5: nothing is written on a duplicate")

    # G6. a duplicate in the compared grid fails hard too.
    dup_grid = os.path.join(tmp, "dup_grid.csv")
    gs.write_grid(dup_grid, grid_rows + [copy.deepcopy(grid_rows[2])])
    gen.build_rows = lambda: copy.deepcopy(fresh_rows)
    rc, out = quiet(gen.main, ["--grid", dup_grid])
    check(rc == 1 and "duplicate canonical key" in out,
          "G6: a duplicate key in the grid fails hard")

gen.build_rows = real_build_rows

# G0. The documented default, run for real as a subprocess from benchmark/.
listing_before = sorted(os.listdir(HERE))
with guard_committed_grid("G0"):
    proc = subprocess.run([sys.executable, "generate_reference_values.py"],
                          cwd=HERE, capture_output=True, text=True)
check(proc.returncode == 0, f"G0: default run exits 0 (got {proc.returncode})")
check("report only: nothing written." in proc.stdout,
      "G0: default run reports that it wrote nothing")
check("grid only, retained" in proc.stdout,
      "G0: default run reports retained grid-only rows")
check(sorted(os.listdir(HERE)) == listing_before,
      "G0: default run creates no file in benchmark/")


# ---------------------------------------------------------------------------
# P. promote_grid_rows.py, driven as a subprocess with a fixture generator.
# ---------------------------------------------------------------------------
PROMOTE = os.path.join(HERE, "promote_grid_rows.py")


def fake_generator(path, rows):
    with open(path, "w", encoding="utf-8") as fh:
        fh.write("import copy, json\n")
        fh.write(f"ROWS = json.loads({json.dumps(json.dumps(rows))})\n")
        fh.write("def build_rows():\n    return copy.deepcopy(ROWS)\n")


def promote(tmp, grid, generator, *args):
    proc = subprocess.run([sys.executable, PROMOTE, "--grid", grid,
                           "--generator", generator] + list(args),
                          cwd=tmp, capture_output=True, text=True)
    return proc.returncode, proc.stdout + proc.stderr


with tempfile.TemporaryDirectory() as tmp:
    base = [row("Fixture_A", "0.5", "1", obs="1.0E+000;0"),
            row("Fixture_A", "0.25", "2", obs="2.0E+000;0"),
            row("Fixture_B", "1", "3", obs="3.0E+000;0")]
    grid = os.path.join(tmp, "grid.csv")
    generator = os.path.join(tmp, "fixture_gen.py")

    def reset(rows=base):
        gs.write_grid(grid, rows)
        return read_bytes(grid)

    # The generator changes one reference, one claim (spelled differently but
    # the same Double), adds one row, and omits Fixture_B entirely.
    gen_rows = [dict(base[0], reference="1.5", observed_vba=""),
                dict(base[1], arg1="0.250000000000000000", claim="rel<=2E-15",
                     observed_vba=""),
                row("Fixture_A", "0.125", "4")]
    fake_generator(generator, gen_rows)
    common = ["--function", "Fixture_A", "--allow-add", "--patch-metadata", "--write"]

    before = reset()
    rc, out = promote(tmp, grid, generator, *common)
    check(rc == 1 and "--accept-reference-changes 1" in out,
          "P1: a reference change without acknowledgement is refused")
    check(read_bytes(grid) == before, "P1: grid unchanged")

    rc, out = promote(tmp, grid, generator, *common,
                      "--accept-reference-changes", "2", "--reason", "fixture")
    check(rc == 1 and "but 1 reference(s) would change" in out,
          "P2: a miscounted acknowledgement is refused")
    check(read_bytes(grid) == before, "P2: grid unchanged")

    rc, out = promote(tmp, grid, generator, *common,
                      "--accept-reference-changes", "1")
    check(rc == 1 and "requires --reason" in out,
          "P3: an acknowledgement without a reason is refused")
    check(read_bytes(grid) == before, "P3: grid unchanged")

    rc, out = promote(tmp, grid, generator, "--function", "Fixture_A",
                      "--patch-metadata", "--write",
                      "--accept-reference-changes", "1", "--reason", "fixture")
    check(rc == 1 and "Pass --allow-add" in out,
          "P4: an append without --allow-add is refused")
    check(read_bytes(grid) == before, "P4: grid unchanged")

    rc, out = promote(tmp, grid, generator, *common,
                      "--accept-reference-changes", "1", "--reason", "fixture")
    check(rc == 0, f"P5: acknowledged promotion succeeds (rc={rc})")
    after = list(csv.DictReader(io.StringIO(read_bytes(grid).decode("utf-8"))))
    check([r["observed_vba"] for r in after[:3]] ==
          [r["observed_vba"] for r in base],
          "P5: every existing observation is preserved")
    check([gs.canonical_key(r) for r in after[:3]] ==
          [gs.canonical_key(r) for r in base],
          "P5: existing row keys and order are preserved")
    check(after[0]["reference"] == "1.5", "P5: the acknowledged reference changed")
    check(after[1]["claim"] == "rel<=2E-15" and after[1]["arg1"] == "0.25",
          "P5: metadata patched without rewriting the argument spelling")
    check(after[2]["function"] == "Fixture_B",
          "P5: a row absent from the generator is retained, never inferred deleted")
    check(len(after) == 4 and after[3]["observed_vba"] == "",
          "P5: the new row is appended with a blank observation")
    check(b"\r" not in read_bytes(grid), "P5: grid written with LF endings")

    # P6. a metadata-only change needs no reference acknowledgement.
    fake_generator(generator, [dict(base[1], claim="rel<=3E-15", observed_vba="")])
    before = reset()
    rc, out = promote(tmp, grid, generator, "--function", "Fixture_A",
                      "--patch-metadata", "--write")
    check(rc == 0, "P6: metadata-only patch succeeds without acknowledgement")

    # P7. an acknowledgement with no reference change to match is stale.
    before = reset()
    rc, out = promote(tmp, grid, generator, "--function", "Fixture_A",
                      "--patch-metadata", "--write",
                      "--accept-reference-changes", "1", "--reason", "stale")
    check(rc == 1 and "but 0 reference(s) would change" in out,
          "P7: a stale acknowledgement is refused")
    check(read_bytes(grid) == before, "P7: grid unchanged")

    # P8. a duplicate outside the promoted function now blocks too.
    before = reset(base + [copy.deepcopy(base[2])])
    rc, out = promote(tmp, grid, generator, "--function", "Fixture_A",
                      "--patch-metadata", "--write")
    check(rc == 1 and "duplicate canonical key" in out,
          "P8: a duplicate anywhere in the grid is refused")
    check(read_bytes(grid) == before, "P8: grid unchanged")

    # P9. retirement is explicit: reason and exact count, nothing else lost.
    before = reset()
    rc, out = promote(tmp, grid, generator, "--retire", "Fixture_B", "all", "1",
                      "--write")
    check(rc == 1 and "requires --reason" in out, "P9: retire without reason refused")
    rc, out = promote(tmp, grid, generator, "--retire", "Fixture_B", "all", "2",
                      "--reason", "fixture", "--write")
    check(rc == 1 and "expected 2 rows, found 1" in out,
          "P9: retire with a wrong count refused")
    check(read_bytes(grid) == before, "P9: grid unchanged by refused retirements")
    rc, out = promote(tmp, grid, generator, "--retire", "Fixture_B", "all", "1",
                      "--reason", "fixture", "--write")
    after = list(csv.DictReader(io.StringIO(read_bytes(grid).decode("utf-8"))))
    check(rc == 0 and [r["function"] for r in after] == ["Fixture_A", "Fixture_A"],
          "P9: explicit retirement removes exactly the named rows")
    check([r["observed_vba"] for r in after] ==
          [r["observed_vba"] for r in base[:2]],
          "P9: retirement preserves every other observation")


# ---------------------------------------------------------------------------
# M. migrate_references.py: the audited reference path changes references only.
# ---------------------------------------------------------------------------
MIGRATE = os.path.join(HERE, "migrate_references.py")
AUDIT_FIELDS = ["function", "arg1_hex", "arg2_hex", "arg3_hex", "arg4_hex",
                "regime", "replacement_source", "committed_reference",
                "generator_reference", "independent_oracle",
                "committed_rel_error", "generator_rel_error", "classification"]


def audit_row(r, source, gen_ref):
    return {"function": r["function"], "arg1_hex": gs.f64_bits(r["arg1"]),
            "arg2_hex": gs.f64_bits(r["arg2"]), "arg3_hex": gs.f64_bits(r["arg3"]),
            "arg4_hex": gs.f64_bits(r["arg4"]), "regime": r["regime"],
            "replacement_source": source, "committed_reference": r["reference"],
            "generator_reference": gen_ref, "independent_oracle": gen_ref,
            "committed_rel_error": "1E-10", "generator_rel_error": "0",
            "classification": "fixture"}


with tempfile.TemporaryDirectory() as tmp:
    base = [row("Fixture_A", "0.5", "1", obs="1.0E+000;0"),
            row("Fixture_A", "0.25", "2", obs="2.0E+000;0"),
            row("Fixture_B", "1", "3", obs="3.0E+000;0")]
    grid = os.path.join(tmp, "grid.csv")
    gs.write_grid(grid, base)
    generator = os.path.join(tmp, "fixture_gen.py")
    fake_generator(generator, [dict(base[0], reference="1.25", observed_vba="")])
    audit = os.path.join(tmp, "audit.csv")
    with open(audit, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=AUDIT_FIELDS, lineterminator="\n")
        w.writeheader()
        w.writerow(audit_row(base[0], "CANONICAL_GENERATOR", "1.25"))
        w.writerow(audit_row(base[1], "NONE", ""))

    def migrate(*args):
        proc = subprocess.run([sys.executable, MIGRATE, "--grid", grid, "--audit",
                               audit, "--generator", generator, "--report",
                               os.path.join(tmp, "migration.csv")] + list(args),
                              cwd=tmp, capture_output=True, text=True)
        return proc.returncode, proc.stdout + proc.stderr

    before = read_bytes(grid)
    rc, out = migrate()
    check(rc == 0 and read_bytes(grid) == before,
          "M1: migration is report-only without --write")
    rc, out = migrate("--write")
    after = list(csv.DictReader(io.StringIO(read_bytes(grid).decode("utf-8"))))
    check(rc == 0 and after[0]["reference"] == "1.25",
          f"M2: the audited reference is patched (rc={rc})")
    check([r["observed_vba"] for r in after] == [r["observed_vba"] for r in base],
          "M2: every observation is preserved")
    check([gs.canonical_key(r) for r in after] == [gs.canonical_key(r) for r in base],
          "M2: row keys and order are preserved; nothing added or removed")
    check(all(a[f] == b[f] for a, b in zip(after, base)
              for f in gs.GRID_FIELDS if f != "reference"),
          "M2: no field other than reference changes")


# ---------------------------------------------------------------------------
# C. check_grid_keys.py on fixture grids.
# ---------------------------------------------------------------------------
ck = load("ckeys_fixture", "check_grid_keys.py")
with tempfile.TemporaryDirectory() as tmp:
    clean = os.path.join(tmp, "clean.csv")
    gs.write_grid(clean, [row("Fixture_A", "0.5", "1"), row("Fixture_A", "0.25", "2")])
    rc, _ = quiet(ck.main, [clean])
    check(rc == 0, "C1: a duplicate-free grid passes")
    dup = os.path.join(tmp, "dup.csv")
    gs.write_grid(dup, [row("Fixture_A", "0.85", "1"),
                        row("Fixture_A", "0.84999999999999998", "1")])
    rc, out = quiet(ck.main, [dup])
    check(rc == 1 and "x2" in out, "C2: two spellings of one Double are a duplicate")
    bad = os.path.join(tmp, "schema.csv")
    gs.write_grid(bad, [row("Fixture_A", "0.5", "1")],
                  fields=[f for f in gs.GRID_FIELDS if f != "evidence_set"])
    rc, out = quiet(ck.main, [bad])
    check(rc == 1 and "canonical" in out, "C3: a non-canonical schema fails")


if fails:
    print("FAIL: grid-regeneration safeguards (#17)")
    for f in fails:
        print("  - " + f)
    raise SystemExit(1)
print("PASS: grid-regeneration safeguards (#17): key parity; generator refusal of "
      "grid, copy, committed and holdout targets; fresh non-authoritative output; "
      "duplicate keys; real default run; promotion acknowledgement, metadata, "
      "stale ack, duplicates, retention, retirement; audited reference "
      "migration; duplicate guard")
