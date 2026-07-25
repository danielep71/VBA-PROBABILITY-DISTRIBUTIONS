"""
Regression tests for the observation-provenance manifest (_manifest.py).

Proves the source binding actually catches drift: a matching tree verifies
clean, and a changed / added / removed module, or a schema change, is reported
so the gate can refuse to certify stale evidence.
Run: python3 test_manifest.py   (exit 0 = pass, nonzero = fail)
"""
import os
import tempfile
import _manifest as M

fails = []
def check(cond, msg):
    if not cond:
        fails.append(msg)


def _tree():
    """A tiny fake repo: root/src/a.bas, root/benchmark/{grid,contracts}."""
    d = tempfile.mkdtemp()
    os.makedirs(os.path.join(d, "src"))
    os.makedirs(os.path.join(d, "benchmark"))
    open(os.path.join(d, "src", "a.bas"), "wb").write(b"Attribute VB_Name = \"a\"\r\nSub X()\r\nEnd Sub\r\n")
    grid = os.path.join(d, "benchmark", "probability_accuracy_grid.csv")
    open(grid, "w", newline="\n").write("function,arg1,reference,observed_vba,regime,expected_error\n")
    contracts = os.path.join(d, "benchmark", "accuracy_contracts.csv")
    open(contracts, "w", newline="\n").write("contract_id,threshold\n")
    return d, grid, contracts


d, grid, contracts = _tree()
man = M.build_manifest(d, grid, contracts, generated_utc="t0")

# 1. clean tree verifies with no problems
check(M.verify_source_binding(d, man, grid, contracts) == [], "clean tree verifies")

# 2. CRLF vs LF must hash identically (content binding, not line endings)
p = os.path.join(d, "src", "a.bas")
raw = open(p, "rb").read()
open(p, "wb").write(raw.replace(b"\r\n", b"\n"))            # rewrite as LF
check(M.verify_source_binding(d, man, grid, contracts) == [], "LF vs CRLF hashes identically")

# 3. a real content change is caught
open(p, "wb").write(raw.replace(b"Sub X()", b"Sub X() 'edited"))
probs = M.verify_source_binding(d, man, grid, contracts)
check(any("a.bas" in x and "changed" in x for x in probs), "content change detected")
open(p, "wb").write(raw)                                    # restore

# 4. an added module (not in manifest) is caught
extra = os.path.join(d, "src", "b.bas")
open(extra, "wb").write(b"Attribute VB_Name = \"b\"\r\n")
probs = M.verify_source_binding(d, man, grid, contracts)
check(any("b.bas" in x and "absent from the manifest" in x for x in probs), "added module detected")
os.remove(extra)

# 5. a removed module (in manifest, missing from tree) is caught
os.remove(p)
probs = M.verify_source_binding(d, man, grid, contracts)
check(any("a.bas" in x and "missing from the tree" in x for x in probs), "removed module detected")
open(p, "wb").write(raw)                                    # restore

# 6. a grid schema (column) change is caught
grid_raw = open(grid).read()
open(grid, "w", newline="\n").write("function,arg1,reference,observed_vba,regime\n")  # dropped expected_error
probs = M.verify_source_binding(d, man, grid, contracts)
check(any("schema columns" in x for x in probs), "schema change detected")
open(grid, "w", newline="\n").write(grid_raw)               # restore

# 7. an OBSERVATION-CONTENT change with UNCHANGED columns is caught (P1-03 gap)
open(grid, "a", newline="\n").write("F,0.5,1.0,2.0,all,\n")  # append a data row, same columns
probs = M.verify_source_binding(d, man, grid, contracts)
check(any("observation grid contents changed" in x for x in probs), "observation-content change detected")
check(not any("schema columns" in x for x in probs), "content change is not misreported as a schema change")
open(grid, "w", newline="\n").write(grid_raw)               # restore

# 8. a contract-registry change is caught
open(contracts, "w", newline="\n").write("contract_id,threshold\nX,9E-9\n")
probs = M.verify_source_binding(d, man, grid, contracts)
check(any("contract registry changed" in x for x in probs), "contract registry change detected")

# 9. a manifest predating grid-content binding is rejected
old_style = dict(man); old_style.pop("observation_grid_sha256")
probs = M.verify_source_binding(d, old_style, grid, contracts)
check(any("predates observation-content binding" in x for x in probs), "pre-binding manifest rejected")

if fails:
    print("FAIL: manifest verification")
    for f in fails:
        print("  - " + f)
    raise SystemExit(1)
print("PASS: manifest binds evidence to source (clean, LF/CRLF, module change/add/remove, "
      "schema, observation-content, contract registry, pre-binding rejection)")
