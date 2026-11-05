"""
Fixture proof that `refresh_evidence.py --check` changes nothing (#88), and
that the default refresh mode still rewrites the committed summaries.

Every case runs the real refresh_evidence.main() in a private clone of this
repository, never in the real checkout. The clone carries the working-tree
copy of the benchmark tooling and commits it, so it starts as a clean tree. Its
accuracy and holdout summaries are committed stale on purpose: a verify step
that writes to its committed path then changes their bytes, which an
up-to-date summary would hide. Before #88 the strict gate step wrote
accuracy_summary.md that way, and the first case below failed.

The VERIFY step that runs test_evidence_tools.py is dropped in the clone,
because that shim runs this file and the run would recurse. Every other step
runs as committed.

Run: python3 test_refresh_evidence.py   (exit 0 = pass, nonzero = fail)
"""
import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
GUARD = "REFRESH_EVIDENCE_FIXTURE"
SUMMARIES = ("benchmark/accuracy_summary.md", "benchmark/holdout/holdout_summary.md",
             "benchmark/holdout_older/holdout_older_summary.md")
STALE = "\n<!-- fixture: stale line that only a refresh may remove (#88) -->\n"
CHECK_LINE = "--check left the working tree unchanged"

# Runs inside the clone's benchmark/ directory, so `import refresh_evidence`
# loads the clone's copy and its HERE and ROOT point into the clone.
DRIVER = r'''
import json
import sys
import refresh_evidence as R
config = json.loads(sys.argv[1])
kept = [step for step in R.VERIFY if step[2][0] != "test_evidence_tools.py"]
if len(kept) != len(R.VERIFY) - 1:
    raise SystemExit("fixture: VERIFY must run test_evidence_tools.py exactly once")
R.VERIFY = kept + [(label, R.HERE, argv) for label, argv in config["extra"]]
sys.argv = ["refresh_evidence.py"] + config["args"]
R.main()
'''

fails = []


def check(cond, msg):
    if not cond:
        fails.append(msg)


def git(root, *args):
    return subprocess.run(["git", "-C", root] + list(args), check=True,
                          capture_output=True).stdout


def status(root):
    return git(root, "status", "--porcelain=v1", "--untracked-files=all").decode()


def snapshot(root):
    """SHA-256 of every tracked file's bytes, keyed by path."""
    paths = git(root, "ls-files", "-z").decode().split("\0")
    out = {}
    for rel in filter(None, paths):
        full = os.path.join(root, *rel.split("/"))
        if os.path.isfile(full):
            with open(full, "rb") as f:
                out[rel] = hashlib.sha256(f.read()).hexdigest()
    return out


def read(root, rel):
    with open(os.path.join(root, *rel.split("/")), encoding="utf-8", newline="") as f:
        return f.read()


def fixture(root):
    """A clean clone whose committed summaries are stale."""
    subprocess.run(["git", "clone", "--quiet", "--shared", ROOT, root],
                   check=True, capture_output=True)
    for rel in filter(None, git(ROOT, "ls-files", "-z", "--", "benchmark").decode().split("\0")):
        source = os.path.join(ROOT, *rel.split("/"))
        if rel.endswith(".py") and os.path.isfile(source):
            shutil.copyfile(source, os.path.join(root, *rel.split("/")))
    for rel in SUMMARIES:
        with open(os.path.join(root, *rel.split("/")), "a", encoding="utf-8", newline="\n") as f:
            f.write(STALE)
    git(root, "add", "-A")
    git(root, "-c", "user.name=fixture", "-c", "user.email=fixture@example.invalid",
        "-c", "commit.gpgsign=false", "commit", "--quiet", "--no-verify",
        "-m", "fixture: working-tree tooling, stale summaries")


def refresh(root, args, extra=()):
    config = json.dumps({"args": args, "extra": [list(e) for e in extra]})
    env = dict(os.environ, **{GUARD: "1"})
    proc = subprocess.run([sys.executable, "-c", DRIVER, config],
                          cwd=os.path.join(root, "benchmark"), env=env,
                          capture_output=True, text=True, timeout=900)
    return proc.returncode, proc.stdout + proc.stderr


def ran_verify(output):
    return ("Verifying" in output and "strict accuracy gate" in output
            and "independent holdout" in output and "Traceback" not in output)


if os.environ.get(GUARD):
    print("FAIL: test_refresh_evidence.py ran inside its own fixture (recursion)")
    raise SystemExit(1)

# 1. --check on a clean tree leaves every tracked file and the status unchanged.
with tempfile.TemporaryDirectory(prefix="refresh-check-") as root:
    fixture(root)
    clean, before = status(root), snapshot(root)
    check(clean == "", f"fixture must start clean, got {clean!r}")
    _code, output = refresh(root, ["--check"])
    check(ran_verify(output), "--check: the verify steps ran\n" + output)
    check(f"PASS  {CHECK_LINE}" in output, "--check: reports the tree unchanged\n" + output)
    check(status(root) == clean, f"--check: git status unchanged, got {status(root)!r}")
    after = snapshot(root)
    changed = sorted(rel for rel in before if after.get(rel) != before[rel])
    check(changed == [], f"--check: no tracked file changed, changed {changed}")
    check(all(read(root, rel).endswith(STALE) for rel in SUMMARIES),
          "--check: the stale summaries are left as committed")

# 2. Negative control: a verify step that writes a committed file (the
# pre-#88 strict gate step, compute_errors.py with its default --out) fails
# --check instead of passing silently.
with tempfile.TemporaryDirectory(prefix="refresh-check-rogue-") as root:
    fixture(root)
    code, output = refresh(root, ["--check"],
                           [("pre-#88 strict gate", ["compute_errors.py"])])
    check(ran_verify(output), "rogue --check: the verify steps ran\n" + output)
    check(code == 1 and f"FAIL  {CHECK_LINE}" in output,
          "rogue --check: a write to the tree fails the run\n" + output)
    check("benchmark/accuracy_summary.md" in status(root),
          "rogue --check: the control did write the tracked summary")

# 3. The default refresh mode still rewrites the committed summaries.
with tempfile.TemporaryDirectory(prefix="refresh-default-") as root:
    fixture(root)
    _code, output = refresh(root, [])
    check(ran_verify(output) and "Regenerating" in output,
          "refresh: regenerated, then verified\n" + output)
    check(CHECK_LINE not in output, "refresh: the --check guard does not run")
    for rel in SUMMARIES:
        text = read(root, rel)
        check(STALE not in text and text.startswith("# "),
              f"refresh: {rel} regenerated in place")
        check(rel in status(root), f"refresh: {rel} reported as changed")

if fails:
    print(f"FAIL: {len(fails)} refresh_evidence.py --check fixture check(s)")
    for msg in fails:
        print("  - " + msg)
    raise SystemExit(1)
print("PASS: refresh_evidence.py --check writes nothing to the tree (outputs go to a "
      "temporary directory and a guard fails any write); the default refresh still "
      "rewrites the committed summaries")
