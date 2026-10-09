"""
Negative and positive controls for render_readme_assurance.py (#28).

Every case runs against a private copy of the committed inputs, so the real
README and evidence are never touched. The cases prove that the block is
current, that it cannot drift by hand, that stale evidence is rendered visibly
and blocks, and that missing, malformed or contradictory inputs fail without
writing anything - never falling back to an earlier figure or to zero.

Run: python3 test_render_readme_assurance.py   (exit 0 = pass, nonzero = fail)
"""
import contextlib
import glob
import io
import json
import os
import shutil
import subprocess
import sys
import tempfile

import render_readme_assurance as R
from _manifest import normalized_hash

fails = []
REAL_ROOT = R.ROOT
BAS_PATTERNS = ("src/**/*.bas", "tests/**/*.bas", "benchmark/**/*.bas")


def check(cond, msg):
    if not cond:
        fails.append(msg)


def fixture():
    root = tempfile.mkdtemp(prefix="readme-assurance-")
    # Retain real candidate objects without copying/mutating the source repo.
    # No checkout: the private working inputs are copied explicitly below.
    subprocess.run(["git", "clone", "--shared", "--no-checkout", REAL_ROOT, root],
                   check=True, capture_output=True)
    paths = list(R.INPUTS)
    for pattern in BAS_PATTERNS:
        paths += [os.path.relpath(p, REAL_ROOT).replace(os.sep, "/")
                  for p in glob.glob(os.path.join(REAL_ROOT, pattern), recursive=True)]
    for rel in paths:
        target = os.path.join(root, *rel.split("/"))
        os.makedirs(os.path.dirname(target), exist_ok=True)
        shutil.copyfile(os.path.join(REAL_ROOT, *rel.split("/")), target)
    return root


def path(root, rel):
    return os.path.join(root, *rel.split("/"))


def read(root, rel):
    with open(path(root, rel), encoding="utf-8", newline="") as f:
        return f.read()


def write(root, rel, text):
    with open(path(root, rel), "w", encoding="utf-8", newline="\n") as f:
        f.write(text)


def edit(root, rel, old, new):
    text = read(root, rel)
    if old not in text:
        raise AssertionError(f"fixture text not found in {rel}: {old[:60]!r}")
    write(root, rel, text.replace(old, new, 1))


def edit_json(root, rel, change):
    data = json.loads(read(root, rel))
    change(data)
    write(root, rel, json.dumps(data, indent=2) + "\n")


def run(root, *argv):
    out = io.StringIO()
    with contextlib.redirect_stdout(out):
        code = R.main(list(argv), root=root)
    return code, out.getvalue()


def expect_failure(label, mutate, needle):
    """A broken input must fail --write with no README change, and --check too."""
    root = fixture()
    try:
        mutate(root)
        before = read(root, R.README)
        code, out = run(root, "--write")
        check(code == 1, f"{label}: --write exits 1 (got {code})")
        check(needle in out, f"{label}: diagnostic names the problem ({needle!r} not in {out!r})")
        check(read(root, R.README) == before, f"{label}: README left untouched")
        code, _ = run(root, "--check")
        check(code == 1, f"{label}: --check exits 1")
    finally:
        shutil.rmtree(root)


# 1. The committed README is current, and rendering is deterministic.
root = fixture()
try:
    code, out = run(root, "--check")
    check(code == 0, f"committed README renders current: {out}")
    before = read(root, R.README)
    code, _ = run(root, "--write")
    check(code == 0 and read(root, R.README) == before, "--write on a current README is a no-op")
    model = R.build_model(root)
    check(R.render(model) == R.render(R.build_model(root)), "two renders are identical")
    rendered = R.render(model)
    check("2,088" in rendered["assurance table"], "thousands use one comma format")
    for region in R.REGIONS:
        check(rendered[region] in before, f"{region} region is present verbatim")

    # 2. Manual drift inside a generated region fails --check; --write repairs it.
    table = rendered["assurance table"]
    first_figure = table.split("| **", 2)[1].split("**", 1)[0]
    drifted = before.replace(f"| **{first_figure}** |", "| **999** |", 1)
    write(root, R.README, drifted)
    code, out = run(root, "--check")
    check(code == 1 and "differ from a fresh render" in out, "hand-edited figure fails --check")
    code, _ = run(root, "--write")
    check(code == 0 and read(root, R.README) == before, "--write restores the rendered figure")

    # 3. Prose outside the generated regions is not the renderer's business.
    write(root, R.README, before.replace("A transparent, tail-aware", "A transparent and tail-aware", 1))
    code, _ = run(root, "--check")
    check(code == 0, "edits outside generated regions do not fail --check")
finally:
    shutil.rmtree(root)

# 4. Stale source: rendered visibly, and blocking in both modes.
root = fixture()
try:
    edit(root, "src/M_STATS_PROBDIST_CORE.bas", "Option Explicit", "Option Explicit\n' changed")
    code, out = run(root, "--write")
    check(code == 1 and "BLOCKING: stale evidence" in out, "stale source blocks --write")
    text = read(root, R.README)
    check("Excel_Regression-STALE" in text, "stale Excel evidence shows a STALE badge")
    check("Accuracy_Contracts-STALE" in text and "Holdout-STALE" in text,
          "stale main-grid and holdout bindings show STALE badges")
    check("| **STALE** |" in text and "src/M_STATS_PROBDIST_CORE.bas" in text,
          "the table and state list name the stale evidence")
    check("909_of_909" not in text, "stale evidence never keeps the previous figure")
    code, out = run(root, "--check")
    check(code == 1 and "stale evidence" in out, "stale source fails --check after --write")
finally:
    shutil.rmtree(root)

# 5. Missing and malformed inputs fail closed.
def retarget_candidate(root, candidate):
    def change(record):
        record["candidate_sha"] = candidate
        record["runner"]["workflow"]["sha"] = candidate
    edit_json(root, R.EXCEL_RECORD, change)


expect_failure("nonexistent candidate with matching workflow SHA",
               lambda r: retarget_candidate(r, "0" * 40), "excel_regression_record.json")
expect_failure("missing candidate history",
               lambda r: shutil.rmtree(path(r, ".git")), "excel_regression_record.json")
expect_failure("source digest inconsistent with candidate",
               lambda r: edit_json(r, R.EXCEL_RECORD,
                                   lambda d: d["sources"][0].update(sha256="sha256:" + "0" * 64)),
               "source digest")


def unrelated_candidate(root):
    subprocess.run(["git", "read-tree", "HEAD"], cwd=root, check=True, capture_output=True)
    source = "src/M_STATS_PROBDIST_CORE.bas"
    edit(root, source, "Option Explicit", "Option Explicit\n' unrelated candidate")
    subprocess.run(["git", "add", source], cwd=root, check=True, capture_output=True)
    subprocess.run(["git", "-c", "user.name=Fixture", "-c", "user.email=fixture@example.invalid",
                    "commit", "-qm", "Unrelated candidate"],
                   cwd=root, check=True, capture_output=True)
    candidate = subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=root,
                                        text=True).strip()
    retarget_candidate(root, candidate)


expect_failure("existing but unrelated candidate", unrelated_candidate, "source digest")
for field in ("excel_version", "excel_build", "office_bitness"):
    expect_failure(f"PASS stages with unavailable {field}",
                   lambda r, field=field: edit_json(r, R.EXCEL_RECORD,
                       lambda d: d["environment"].update({field: "unavailable"})),
                   "green evidence requires observed")

expect_failure("uncommitted policy drift",
               lambda r: edit_json(r, R.POLICY, lambda d: d.update(expected_assertions=910)),
               "policy differs from committed HEAD")

# Failed sessions may truthfully retain unavailable environment fields, but
# they must never become green evidence. The canonical schema still applies.
root = fixture()
try:
    edit_json(root, R.EXCEL_RECORD,
              lambda d: d["stages"]["cleanup"].update(status="FAIL"))
    edit_json(root, R.EXCEL_RECORD,
              lambda d: d["environment"].update(office_bitness="unavailable"))
    check(R.excel_facts(root)["green"] is False,
          "failed session with unavailable environment stays non-green")
finally:
    shutil.rmtree(root)

expect_failure("missing Excel record",
               lambda r: os.remove(path(r, R.EXCEL_RECORD)), "required input missing")

# A supported prerequisite failure must be a clean error before any mutation,
# including the binding modes. Use real entry points with an empty PATH rather
# than mocking away the executable lookup that caused the review finding.
before_inputs = {rel: read(REAL_ROOT, rel) for rel in R.INPUTS}
with tempfile.TemporaryDirectory(prefix="no-git-path-") as empty_path:
    env = dict(os.environ, PATH=empty_path)
    for script, args in (
        ("render_readme_assurance.py", ["--write"]),
        ("render_readme_assurance.py", ["--check"]),
        ("refresh_evidence.py", []),
        ("refresh_evidence.py", ["--check"]),
        ("refresh_evidence.py", ["--bind-exported-main", "--bind-exported-holdout"]),
    ):
        proc = subprocess.run([sys.executable, os.path.join(R.HERE, script)] + args,
                              cwd=REAL_ROOT, env=env, capture_output=True, text=True)
        output = proc.stdout + proc.stderr
        check(proc.returncode == 1, f"{script} {args}: missing Git exits 1")
        check("requires Git on PATH" in output, f"{script} {args}: actionable prerequisite")
        check("Traceback" not in output, f"{script} {args}: no uncaught exception")
        check("Regenerating" not in output, f"{script} {args}: no regeneration attempted")
        check(all(read(REAL_ROOT, rel) == text for rel, text in before_inputs.items()),
              f"{script} {args}: all evidence inputs remain unchanged")
expect_failure("missing readiness registry",
               lambda r: os.remove(path(r, R.READINESS)), "required input missing")
expect_failure("malformed Excel record",
               lambda r: write(r, R.EXCEL_RECORD, "{not json"), "not valid JSON")
expect_failure("Excel counters inconsistent",
               lambda r: edit_json(r, R.EXCEL_RECORD,
                                   lambda d: d["harness"].update(passed=908)), "inconsistent")
expect_failure("unknown contract status",
               lambda r: edit(r, R.CONTRACTS, ",active,", ",retired,"), "unknown status")
expect_failure("public API line malformed",
               lambda r: edit(r, R.PUBLIC_API, "\tFunction\tK_STATS_", "\tFunction\tX_"),
               "not a K_STATS_ declaration")
expect_failure("readiness schema",
               lambda r: edit_json(r, R.READINESS, lambda d: d.update(schema_version="v0")),
               "unsupported schema_version")
expect_failure("resolved blocker without evidence",
               lambda r: edit_json(r, R.READINESS,
                                   lambda d: d["blockers"][0].update(status="resolved")),
               "blocker keys differ")
expect_failure("README markers missing",
               lambda r: edit(r, R.README, R._markers("evidence state")[0], ""),
               "expected exactly one 'evidence state' region")

# 6. Contradictory inputs fail closed.
expect_failure("verdict tally contradicts table",
               lambda r: edit(r, R.SUMMARY, "FAIL: 0, KNOWN", "FAIL: 1, KNOWN"),
               "tally states FAIL")
expect_failure("summary grid rows contradict grid",
               lambda r: edit(r, R.SUMMARY, "- total grid rows: 2,088", "- total grid rows: 2,089"),
               "grid rows")
expect_failure("summary source commit contradicts manifest",
               lambda r: edit(r, R.SUMMARY, "- source commit: `74041b3`",
                              "- source commit: `0000000`"),
               "reports source commit")
expect_failure("holdout tally contradicts table",
               lambda r: edit(r, R.HOLDOUT_SUMMARY, "> 80 pass,", "> 81 pass,"),
               "contradicts its table")
expect_failure("registry drops the silent-wrong blocker",
               lambda r: edit_json(r, R.READINESS, lambda d: d.update(
                   blockers=[b for b in d["blockers"] if b["issue"] != 34])),
               "#34 must stay registered")
expect_failure("registry claims certification with open blockers",
               lambda r: edit_json(r, R.READINESS, lambda d: d.update(certified=True)),
               "claims v1.0.0 is certified")
expect_failure("incomplete gate without an owner",
               lambda r: edit_json(r, R.READINESS, lambda d: d.update(
                   blockers=[b for b in d["blockers"] if b["issue"] != 47])),
               "main_grid_export is incomplete")


def certify_main_grid(r):
    digest = normalized_hash(path(r, R.GRID))
    edit_json(r, R.EXCEL_RECORD, lambda d: d["grids"][R.GRID].update(sha256=digest))


expect_failure("satisfied gate still owned by an open blocker", certify_main_grid,
               "still lists #47 open")

if fails:
    print("FAIL: README assurance renderer")
    for f in fails:
        print("  - " + f)
    raise SystemExit(1)
print("PASS: README assurance renderer (current, deterministic, manual drift, stale source, "
      "missing/malformed inputs, contradictory totals and readiness-registry controls)")
