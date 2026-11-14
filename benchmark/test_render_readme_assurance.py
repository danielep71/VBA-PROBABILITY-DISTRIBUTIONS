"""
Negative and positive controls for render_readme_assurance.py (#28).

Every case runs against a private, pinned coherent evidence snapshot, so the real
README and evidence are never touched. The cases prove that the block is
current, that it cannot drift by hand, that stale evidence is rendered visibly
and blocks, and that missing, malformed or contradictory inputs fail without
writing anything - never falling back to an earlier figure or to zero.

Run: python3 test_render_readme_assurance.py   (exit 0 = pass, nonzero = fail)
"""
import csv
import contextlib
import io
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile

import render_readme_assurance as R

fails = []
REAL_ROOT = R.ROOT


def check(cond, msg):
    if not cond:
        fails.append(msg)


def fixture():
    root = tempfile.mkdtemp(prefix="readme-assurance-")
    # Retain real candidate objects without copying/mutating the source repo.
    # Check out only the private historical data fixture below.
    subprocess.run(["git", "clone", "--shared", "--no-checkout", REAL_ROOT, root],
                   check=True, capture_output=True)
    # Unit controls need coherent evidence even while live source is unexported.
    # Use a real, immutable evidence snapshot; R itself is imported from the
    # current working tree. Live freshness is checked separately by CI.
    subprocess.run(["git", "checkout", "--quiet", "8fd0af5b8fa3754d112c06f4845e2a896933c7b6"],
                   cwd=root, check=True, capture_output=True)
    # The snapshot's README predates the study row: render it as it was.
    # The row has its own controls at the end of this file.
    study = os.path.join(root, ".github", "excel-study-policy.json")
    if os.path.exists(study):
        os.remove(study)
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
    with open(path(root, "benchmark/probability_accuracy_grid.csv"), newline="") as grid:
        row_count = sum(1 for _ in csv.DictReader(grid))
    check(f"{row_count:,}" in rendered["assurance table"], "grid count uses comma formatting")
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

# 4a (#301). A comment-only source change is carried, visibly, and is never a
# certified release.
root = fixture()
try:
    edit(root, "src/M_STATS_PROBDIST_CORE.bas", "Option Explicit",
         "Option Explicit\n' a comment added after the Excel run")
    code, out = run(root, "--write")
    text = read(root, R.README)
    check(code == 0 and "STALE" not in text, f"a comment-only change is not stale: {out}")
    check("**Carried from `" in text and "`src/M_STATS_PROBDIST_CORE.bas` changed since that run "
          "in comments or whitespace only" in text and "| **carried**:" not in text.split(
              "<!-- BEGIN")[0], "the carried certification is labelled in the evidence state")
    check(text.count("**Carried**: 1 module(s)") == 2 and "**carried**: 1 regression module(s)" in text,
          "the table and both grid bindings name the carried module")
    code, _ = run(root, "--check")
    check(code == 0, "a carried render is current after --write")
    model = R.build_model(root)
    check(model["carried"] == ["src/M_STATS_PROBDIST_CORE.bas"], "the model lists the carried module")

    certified = {"certified": True, "release": "v9.9.9"}
    check(R.certification_problem(certified, True, []) is None,
          "a clean, uncarried release may be certified")
    check("carried" in (R.certification_problem(certified, True, model["carried"]) or ""),
          "a certified release on a carried binding is refused")
    check(R.certification_problem(certified, False, []) is not None,
          "a certified release with open evidence is refused")
    check(R.certification_problem({"certified": False, "release": "v9.9.9"}, False,
                                  model["carried"]) is None,
          "an uncertified release may carry")
finally:
    shutil.rmtree(root)

# 4. Stale source: rendered visibly, and blocking in both modes.
root = fixture()
try:
    edit(root, "src/M_STATS_PROBDIST_CORE.bas", "Option Explicit",
         "Option Explicit\nPublic Const PROB_FIXTURE_CHANGED As Long = 1")
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

# 5b. refresh_evidence.py preflight: complete history and a valid candidate
# before any write.

# HEAD existing is not sufficient. A real shallow checkout must be rejected
# before refresh rewrites even one artifact, in ordinary and binding modes.
with tempfile.TemporaryDirectory(prefix="shallow-refresh-") as shallow:
    subprocess.run(["git", "clone", "--depth=1", Path(REAL_ROOT).as_uri(), shallow],
                   check=True, capture_output=True)
    for script in ("refresh_evidence.py", "render_readme_assurance.py", "excel_certification.py",
                   "_manifest.py", "vba_code_identity.py"):
        shutil.copyfile(os.path.join(R.HERE, script), path(shallow, "benchmark/" + script))
    before = {rel: read(shallow, rel) for rel in R.INPUTS}
    for args in ([], ["--check"], ["--bind-exported-main", "--bind-exported-holdout"]):
        proc = subprocess.run([sys.executable, path(shallow, "benchmark/refresh_evidence.py")] + args,
                              cwd=shallow, capture_output=True, text=True)
        output = proc.stdout + proc.stderr
        check(proc.returncode == 1 and "complete Git history is required" in output,
              f"shallow refresh {args}: fails its history preflight")
        check("Traceback" not in output and "Regenerating" not in output,
              f"shallow refresh {args}: stops cleanly before writes")
        check(all(read(shallow, rel) == text for rel, text in before.items()),
              f"shallow refresh {args}: all evidence inputs unchanged")

# Complete history alone is not sufficient either: validate the retained
# candidate before regeneration, independently of the shallow-repository guard.
with tempfile.TemporaryDirectory(prefix="invalid-candidate-refresh-") as invalid:
    subprocess.run(["git", "clone", "--shared", REAL_ROOT, invalid],
                   check=True, capture_output=True)
    for script in ("refresh_evidence.py", "render_readme_assurance.py", "excel_certification.py",
                   "_manifest.py", "vba_code_identity.py"):
        shutil.copyfile(os.path.join(R.HERE, script), path(invalid, "benchmark/" + script))
    retarget_candidate(invalid, "0" * 40)
    before = {rel: read(invalid, rel) for rel in R.INPUTS}
    for args in ([], ["--check"], ["--bind-exported-main", "--bind-exported-holdout"]):
        proc = subprocess.run([sys.executable, path(invalid, "benchmark/refresh_evidence.py")] + args,
                              cwd=invalid, capture_output=True, text=True)
        output = proc.stdout + proc.stderr
        check(proc.returncode == 1 and "evidence refresh prerequisite" in output,
              f"invalid candidate refresh {args}: fails candidate preflight")
        check("Traceback" not in output and "Regenerating" not in output,
              f"invalid candidate refresh {args}: stops cleanly before writes")
        check(all(read(invalid, rel) == text for rel, text in before.items()),
              f"invalid candidate refresh {args}: all evidence inputs unchanged")

# 6. Contradictory inputs fail closed.
expect_failure("verdict tally contradicts table",
               lambda r: edit(r, R.SUMMARY, "FAIL: 0, KNOWN", "FAIL: 1, KNOWN"),
               "tally states FAIL")
def bump_main_grid_count(root):
    """Contradict the current total without pinning the live grid's size."""
    text = read(root, R.SUMMARY)
    match = re.search(r"^- total grid rows: ([\d,]+)$", text, re.M)
    if match is None:
        raise AssertionError("summary grid-row count is missing")
    count = int(match.group(1).replace(",", ""))
    edit(root, R.SUMMARY, match.group(0), f"- total grid rows: {count + 1:,}")


expect_failure("summary grid rows contradict grid", bump_main_grid_count, "grid rows")
# The manifest's commit changes with every main-grid binding, so read it rather
# than pin it here (#181).
expect_failure("summary source commit contradicts manifest",
               lambda r: edit(r, R.SUMMARY, "- source commit: `%s`"
                              % json.loads(read(r, R.MAIN_MANIFEST))["commit_sha"],
                              "- source commit: `0000000`"),
               "reports source commit")
def bump_holdout_tally(root):
    """Raise the holdout pass tally by one, whatever it is today."""
    text = read(root, R.HOLDOUT_SUMMARY)
    m = re.search(r"^> (\d+) pass,", text, re.M)
    if not m:
        raise AssertionError(f"fixture text not found in {R.HOLDOUT_SUMMARY}: '> N pass,'")
    edit(root, R.HOLDOUT_SUMMARY, m.group(0), f"> {int(m.group(1)) + 1} pass,")


expect_failure("holdout tally contradicts table", bump_holdout_tally,
               "contradicts its table")
expect_failure("registry drops the silent-wrong blocker",
               lambda r: edit_json(r, R.READINESS, lambda d: d.update(
                   blockers=[b for b in d["blockers"] if b["issue"] != 34])),
               "#34 must stay registered")
# The balanced-Beta contract breaches are measured on holdout_older and by the
# binary64 route model, neither of which the hosted gate scores (#163, #173).
for _issue in (163, 173):
    expect_failure(f"registry drops balanced-Beta blocker #{_issue}",
                   lambda r, i=_issue: edit_json(r, R.READINESS, lambda d: d.update(
                       blockers=[b for b in d["blockers"] if b["issue"] != i])),
                   f"#{_issue} must stay registered")
expect_failure("registry claims certification with open blockers",
               lambda r: edit_json(r, R.READINESS, lambda d: d.update(certified=True)),
               "claims v1.0.0 is certified")
# Which gates are incomplete changes with the evidence: a certified holdout
# export satisfies holdout_export (#181). Drop the owner of whichever gate is
# open in the committed registry.
OPEN_GATE = next(b["gate"] for b in json.loads(read(REAL_ROOT, R.READINESS))["blockers"]
                 if b["gate"] and b["status"] == "open")
expect_failure("incomplete gate without an owner",
               lambda r: edit_json(r, R.READINESS, lambda d: d.update(
                   blockers=[b for b in d["blockers"] if b["gate"] != OPEN_GATE])),
               f"{OPEN_GATE} is incomplete")


def reopen_47(d):
    for blocker in d["blockers"]:
        if blocker["issue"] == 47:
            blocker["status"] = "open"
            del blocker["resolution"]


expect_failure("satisfied gate still owned by an open blocker",
               lambda r: edit_json(r, R.READINESS, reopen_47), "still lists #47 open")


def restore_pre_47_record(r):
    """The record as it was before #47: the raw CRLF export digest stored as
    the grid digest, with no correction trail. A record finalized from a later
    export has no corrections to drop (#181)."""
    def change(d):
        entry = d["grids"][R.GRID]
        entry["sha256"] = entry["raw_sha256"]
        d.pop("corrections", None)
    edit_json(r, R.EXCEL_RECORD, change)


expect_failure("uncertified main grid with #47 marked resolved", restore_pre_47_record,
               "main_grid_export is incomplete")

# The study batteries: a declared count, shown only where the policy exists,
# and refused when its total is not the sum of its batteries.
check(R.study_rows(None) == [], "no study row without a study policy")
root = fixture()
try:
    os.makedirs(path(root, ".github"), exist_ok=True)
    write(root, R.STUDY_POLICY, json.dumps(
        {"expected_assertions": 30, "batteries": {"A": 10, "B": 15},
         "production_batteries": {"C": 5}}))
    facts = R.study_facts(root)
    check(facts == {"assertions": 30, "batteries": 3}, f"study policy is summed: {facts}")
    row = R.study_rows(facts)
    check(len(row) == 1 and "**30**" in row[0] and "3 study and slow production" in row[0]
          and "not in the retained record" in row[0], f"study row states what it is: {row}")
    write(root, R.STUDY_POLICY, json.dumps({"expected_assertions": 31, "batteries": {"A": 10}}))
    try:
        R.study_facts(root)
        check(False, "a study total that is not the sum of its batteries is refused")
    except R.AssuranceError:
        pass
    os.remove(path(root, R.STUDY_POLICY))
    check(R.study_facts(root) is None, "a checkout without the policy has no study facts")
finally:
    shutil.rmtree(root)

if fails:
    print("FAIL: README assurance renderer")
    for f in fails:
        print("  - " + f)
    raise SystemExit(1)
print("PASS: README assurance renderer (current, deterministic, manual drift, stale source, "
      "missing/malformed inputs, contradictory totals and readiness-registry controls)")
