"""
Fixtures for check_study_grid_binding.py and _row_binding.py (#302).

  1. The repository passes, and every declaration carries a reason.
  2. NEGATIVE: a study file that names a live grid and hashes whole files is
     reported until it is declared; a declaration for a file that is no longer
     flagged, or one with no reason, is reported too.
  3. A file that only names the grid, or only hashes, is not flagged; files
     directly in benchmark/ are outside the scan.
  4. row_digest: an appended row of another function, a different row order
     and a text spelling of the same binary64 argument leave the digest
     unchanged; a changed observation, reference or claim of a bound row, an
     added or removed bound row change it; a duplicate key raises.

Run: python3 test_study_grid_binding.py   (exit 0 = pass)
"""
import csv
import os
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import check_study_grid_binding as C  # noqa: E402
from _row_binding import row_digest  # noqa: E402

fails = []


def check(cond, msg):
    if not cond:
        fails.append(msg)


# 1.
check(C.problems() == [], f"1: the repository passes: {C.problems()}")
check(all(r.strip() for r in C.DECLARED.values()), "1: every declaration has a reason")

# 2 and 3.
with tempfile.TemporaryDirectory(prefix="grid-binding-") as d:
    os.makedirs(os.path.join(d, "study"))

    def put(rel, text):
        with open(os.path.join(d, rel), "w", encoding="utf-8") as f:
            f.write(text)

    put("study/binds.py", 'import hashlib\nGRID = "probability_accuracy_grid.csv"\n')
    put("study/reads.py", 'GRID = "probability_accuracy_grid.csv"\n')
    put("study/hashes.py", "import hashlib\n")
    put("top_level.py", 'import hashlib\nGRID = "holdout_grid.csv"\n')
    check(C.flagged(d) == ["study/binds.py"], f"3: only the binding file is flagged: {C.flagged(d)}")
    check(len(C.problems(d, {})) == 1 and "study/binds.py" in C.problems(d, {})[0],
          "2: an undeclared flagged file is reported")
    check(C.problems(d, {"study/binds.py": "hashes its own output"}) == [],
          "2: a declared file passes")
    check(any("no reason" in p for p in C.problems(d, {"study/binds.py": " "})),
          "2: a declaration without a reason is reported")
    check(any("no longer flagged" in p
              for p in C.problems(d, {"study/binds.py": "x", "study/gone.py": "x"})),
          "2: a stale declaration is reported")

# 4.
FIELDS = ["function", "vba_kernel", "claim", "metric", "arg1", "arg2", "arg3", "arg4",
          "reference", "observed_vba", "regime", "evidence_set", "expected_error"]


def row(function, x, observed="0.5;0E+000", reference="0.5", claim="rel<=5E-15"):
    return {"function": function, "claim": claim, "metric": "rel", "arg1": x, "arg2": "1.0",
            "reference": reference, "observed_vba": observed, "regime": "all",
            "evidence_set": "main grid"}


def digest(rows, functions=("Foo_Survival",)):
    with tempfile.NamedTemporaryFile("w", suffix=".csv", delete=False, newline="",
                                     encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS)
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in FIELDS})
    try:
        return row_digest(f.name, functions)
    finally:
        os.unlink(f.name)


base = [row("Foo_Survival", "0.85"), row("Foo_Survival", "2.0"), row("Bar_Density", "1.0")]
n, h = digest(base)
check(n == 2, f"4: two bound rows, got {n}")
check(digest(base + [row("Bar_Density", "3.0")]) == (n, h), "4: an unrelated appended row")
check(digest(base[::-1]) == (n, h), "4: another row order")
check(digest([row("Foo_Survival", "0.84999999999999998")] + base[1:]) == (n, h),
      "4: another spelling of the same binary64 argument")
for label, changed in (
        ("observation", [row("Foo_Survival", "0.85", observed="0.6;0E+000")] + base[1:]),
        ("reference", [row("Foo_Survival", "0.85", reference="0.51")] + base[1:]),
        ("claim", [row("Foo_Survival", "0.85", claim="rel<=1E-13")] + base[1:]),
        ("added bound row", base + [row("Foo_Survival", "3.0")]),
        ("removed bound row", base[1:])):
    check(digest(changed)[1] != h, f"4: a changed {label} changes the digest")
try:
    digest(base + [row("Foo_Survival", "0.85")])
    check(False, "4: a duplicate canonical key must raise")
except ValueError:
    pass

if fails:
    print(f"FAIL: {len(fails)} study grid-binding check(s)")
    for m in fails:
        print("  - " + m)
    raise SystemExit(1)
print("PASS: no study binds a whole live grid, undeclared or stale declarations are refused, and "
      "row digests ignore unrelated rows and order but follow every bound cell")
