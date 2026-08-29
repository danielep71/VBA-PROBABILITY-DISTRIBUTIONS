"""
Regression tests for main-grid and holdout provenance (_manifest.py).

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


def _holdout_tree():
    """A tiny fake repo with the exact source classes a holdout depends on."""
    d = tempfile.mkdtemp()
    os.makedirs(os.path.join(d, "src"))
    os.makedirs(os.path.join(d, "benchmark", "holdout"))
    src = os.path.join(d, "src", "a.bas")
    exporter = os.path.join(d, "benchmark", "holdout", "M_STATS_PROBDIST_HOLDOUT.bas")
    open(src, "wb").write(b"Attribute VB_Name = \"a\"\r\nSub X()\r\nEnd Sub\r\n")
    open(exporter, "wb").write(b"Attribute VB_Name = \"h\"\r\nSub Export()\r\nEnd Sub\r\n")
    grid = os.path.join(d, "benchmark", "holdout", "holdout_grid.csv")
    open(grid, "w", newline="\n").write(
        "function,arg1,reference,observed_vba,regime,expected_error\n"
        "F,0.5,1.0,1.0,all,\n")
    contracts = os.path.join(d, "benchmark", "accuracy_contracts.csv")
    open(contracts, "w", newline="\n").write("contract_id,threshold\nF.all.output,1E-9\n")
    return d, src, exporter, grid, contracts


# The holdout binding is deliberately tested as its own evidence record.  The
# committed historical holdout has no manifest yet; these fixtures prove the
# mechanism before the first truthful post-#23 binding is written.
hd, hsrc, hexporter, hgrid, hcontracts = _holdout_tree()
hman = M.build_holdout_manifest(hd, hgrid, hcontracts, generated_utc="t0")

# 10. a complete binding verifies and records the actual row count
check(M.verify_holdout_binding(hd, hman, hgrid, hcontracts) == [],
      "clean holdout binding verifies")
check(hman["observation_row_count"] == 1, "holdout row count recorded")

# 11. missing provenance is fail-closed rather than treated as legacy evidence
probs = M.verify_holdout_binding(hd, None, hgrid, hcontracts)
check(any(M.HOLDOUT_MANIFEST_NAME in x and "unbound" in x for x in probs),
      "missing holdout manifest rejected")

# 12. CRLF/LF presentation changes do not invalidate source content
hsrc_raw = open(hsrc, "rb").read()
hexporter_raw = open(hexporter, "rb").read()
open(hsrc, "wb").write(hsrc_raw.replace(b"\r\n", b"\n"))
open(hexporter, "wb").write(hexporter_raw.replace(b"\r\n", b"\n"))
check(M.verify_holdout_binding(hd, hman, hgrid, hcontracts) == [],
      "holdout LF vs CRLF hashes identically")
open(hsrc, "wb").write(hsrc_raw)
open(hexporter, "wb").write(hexporter_raw)

# 13. a production-source or exporter change invalidates the binding
open(hsrc, "wb").write(hsrc_raw.replace(b"Sub X()", b"Sub X() 'changed"))
probs = M.verify_holdout_binding(hd, hman, hgrid, hcontracts)
check(any("src/a.bas" in x and "source changed" in x for x in probs),
      "holdout production-source change detected")
open(hsrc, "wb").write(hsrc_raw)
open(hexporter, "wb").write(hexporter_raw.replace(b"Sub Export()", b"Sub Export() 'changed"))
probs = M.verify_holdout_binding(hd, hman, hgrid, hcontracts)
check(any("M_STATS_PROBDIST_HOLDOUT.bas" in x and "source changed" in x for x in probs),
      "holdout exporter change detected")
open(hexporter, "wb").write(hexporter_raw)

# 14. changed observation bytes and row count are detected
hgrid_raw = open(hgrid).read()
open(hgrid, "a", newline="\n").write("F,0.9,2.0,2.0,all,\n")
probs = M.verify_holdout_binding(hd, hman, hgrid, hcontracts)
check(any("grid contents changed" in x for x in probs), "holdout grid change detected")
check(any("row count changed" in x for x in probs), "holdout row-count change detected")
open(hgrid, "w", newline="\n").write(hgrid_raw)

# 15. changed registry and schema are each detected
hcontracts_raw = open(hcontracts).read()
open(hcontracts, "w", newline="\n").write("contract_id,threshold\nF.all.output,2E-9\n")
probs = M.verify_holdout_binding(hd, hman, hgrid, hcontracts)
check(any("contract registry changed" in x for x in probs),
      "holdout contract-registry change detected")
open(hcontracts, "w", newline="\n").write(hcontracts_raw)
open(hgrid, "w", newline="\n").write(
    "function,arg1,reference,observed_vba,regime\nF,0.5,1.0,1.0,all\n")
probs = M.verify_holdout_binding(hd, hman, hgrid, hcontracts)
check(any("schema columns changed" in x for x in probs), "holdout schema change detected")
open(hgrid, "w", newline="\n").write(hgrid_raw)

# 16. malformed JSON is converted into an explicit fail-closed mismatch
hmanifest_path = os.path.join(hd, "benchmark", "holdout", M.HOLDOUT_MANIFEST_NAME)
open(hmanifest_path, "w", newline="\n").write("{not json}\n")
loaded = M.load_holdout_manifest(hd)
probs = M.verify_holdout_binding(hd, loaded, hgrid, hcontracts)
check(any("malformed" in x for x in probs), "malformed holdout manifest rejected")

if fails:
    print("FAIL: manifest verification")
    for f in fails:
        print("  - " + f)
    raise SystemExit(1)
print("PASS: manifests bind main and holdout evidence to source (clean, LF/CRLF, "
      "module/exporter change/add/remove, schema, observation content/row count, "
      "contract registry, missing/malformed/pre-binding rejection)")
