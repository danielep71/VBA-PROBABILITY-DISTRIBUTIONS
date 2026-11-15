"""
Fixtures for check_generator_grid.py (#349), on synthetic rows: the comparison
is proved without building the generator.

Run: python3 test_check_generator_grid.py   (exit 0 = pass)
"""
import copy
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import check_generator_grid as C  # noqa: E402

fails = []


def check(cond, label):
    print(("PASS  " if cond else "FAIL  ") + label)
    if not cond:
        fails.append(label)


def row(function="F", arg1="0.5", regime="all", reference="1.0",
        claim="rel<=1E-13", metric="rel", evidence_set="main grid",
        expected_error=""):
    return {"function": function, "vba_kernel": "K", "claim": claim,
            "metric": metric, "arg1": arg1, "arg2": "", "arg3": "", "arg4": "",
            "reference": reference, "observed_vba": "1;0", "regime": regime,
            "evidence_set": evidence_set, "expected_error": expected_error}


gen = [row(), row(arg1="0.25"), row(function="G", regime="tail")]
grid = copy.deepcopy(gen) + [row(function="H", evidence_set="a_study")]

missing, different, grid_only = C.compare(gen, grid)
check(not missing and not different, "identical rows agree")
check(grid_only == 1, "a grid row the generator does not emit is counted, not failed")

# The argument is compared as a Double, not as text.
respelled = copy.deepcopy(grid)
respelled[0]["arg1"] = "0.50000000000000000"
missing, different, _ = C.compare(gen, respelled)
check(not missing and not different, "another spelling of the same Double agrees")

# An observation is not compared: it is Excel's, never the generator's.
observed = copy.deepcopy(grid)
observed[0]["observed_vba"] = "2;0"
missing, different, _ = C.compare(gen, observed)
check(not missing and not different, "the observation is not compared")

for field, value in (("reference", "1.0000000000000001"), ("claim", "rel<=1E-12"),
                     ("metric", "abs"), ("expected_error", "1")):
    bad = copy.deepcopy(grid)
    bad[1][field] = value
    missing, different, _ = C.compare(gen, bad)
    check(not missing and len(different) == 1 and different[0][1] == [field],
          f"a changed {field} is reported on its row")

one_ulp = copy.deepcopy(grid)
one_ulp[1]["arg1"] = "0.25000000000000006"
missing, different, grid_only = C.compare(gen, one_ulp)
check(len(missing) == 1 and grid_only == 2,
      "an argument one ulp away is a different row: absent on one side, extra on the other")

for field, value in (("regime", "other"), ("evidence_set", "a_study")):
    moved = copy.deepcopy(grid)
    moved[2][field] = value
    missing, different, _ = C.compare(gen, moved)
    check(len(missing) == 1 and not different,
          f"a generator row committed under another {field} is absent")

missing, different, _ = C.compare(gen, grid[1:])
check(len(missing) == 1, "a generator row removed from the grid is absent")

missing, different, grid_only = C.compare([], grid)
check(not missing and not different and grid_only == len(grid),
      "an empty generator disagrees with nothing")

print()
if fails:
    print(f"{len(fails)} check(s) failed")
    sys.exit(1)
print("all checks passed")
