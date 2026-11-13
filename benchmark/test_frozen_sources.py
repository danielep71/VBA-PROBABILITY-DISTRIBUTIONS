"""
The closed Gamma studies do not read the working tree's production sources
(#338).

  1. _frozen_sources.text() and tree() return the pinned blobs: equal to
     `git show PIN:path`, with LF line ends, including a path that
     `git archive` would drop (.github/).
  2. In a throwaway clone carrying the working tree's benchmark/ and tools/,
     every production module and the test module are changed, in a comment
     and in code. The replay fixtures and preparation checks of the Gamma
     continued-fraction, cache and recurrence studies still pass there.
  3. NEGATIVE: in the same clone, the working-tree check
     test_production_follows_study.py fails when the generated public Gamma
     block changes in code, and passes when only a comment in it changes. The
     clone's copy of the check always compares, even after the committed one
     declares that production has moved on.

Run: python3 test_frozen_sources.py   (exit 0 = pass)
"""
import os
import re
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import _frozen_sources as F  # noqa: E402

# The fixtures that used to compare the live sources with history.
FIXTURES = (
    "gamma_cf_study/coverage/test_scaled_vba_port.py",
    "gamma_cf_study/coverage/test_scaled_dd_vba_port.py",
    "gamma_cf_study/coverage/test_scaled_gamma_vba_port.py",
    "gamma_cf_study/coverage/test_extended_gamma.py",
    "gamma_cf_study/coverage/test_wide_gamma.py",
    "gamma_cf_study/coverage/test_public_integration.py",
    "gamma_cache_study/test_cache_vba.py",
    "gamma_recurrence_study/test_prepare.py",
)
LIVE_CHECK = "gamma_cf_study/coverage/test_production_follows_study.py"
BLOCK = b"' BEGIN generated public Gamma integration"

fails = []


def check(cond, msg):
    if not cond:
        fails.append(msg)


def git(root, *args):
    return subprocess.run(["git", "-C", root] + list(args), check=True,
                          capture_output=True).stdout


# 1.
for rel in ("src/M_STATS_PROBDIST_CORE.bas", ".github/excel-evidence-policy.json"):
    shown = git(ROOT, "show", f"{F.GAMMA_STUDY_PIN}:{rel}")
    check(F.blob(rel) == shown, f"1: blob({rel}) is the pinned blob")
    check(F.text(rel) == shown.decode().replace("\r\n", "\n"), f"1: text({rel}) has LF line ends")
frozen = F.tree(("src", ".github/excel-evidence-policy.json"))
check((frozen / ".github/excel-evidence-policy.json").is_file()
      and len(list((frozen / "src").glob("*.bas"))) == 6,
      "1: tree() holds the six production modules and the export-ignored policy")


def run(root, rel):
    proc = subprocess.run([sys.executable, os.path.basename(rel)],
                          cwd=os.path.join(root, "benchmark", os.path.dirname(rel)),
                          capture_output=True, text=True)
    return proc.returncode, (proc.stdout + proc.stderr)[-600:]


def mutate(path, old, new):
    with open(path, "rb") as f:
        data = f.read()
    if old not in data:
        raise AssertionError(f"mutation anchor missing in {path}")
    with open(path, "wb") as f:
        f.write(data.replace(old, new, 1))


with tempfile.TemporaryDirectory(prefix="frozen-sources-clone-") as clone:
    subprocess.run(["git", "clone", "--quiet", "--shared", ROOT, clone],
                   check=True, capture_output=True)
    listed = git(ROOT, "ls-files", "-z", "--", "benchmark", "tools", "src", "tests", ".github")
    for rel in filter(None, listed.decode().split("\0")):
        source = os.path.join(ROOT, *rel.split("/"))
        if os.path.isfile(source):
            target = os.path.join(clone, *rel.split("/"))
            os.makedirs(os.path.dirname(target), exist_ok=True)
            shutil.copyfile(source, target)

    # 2. Change every production module and the test module, comment and code.
    modules = sorted(os.listdir(os.path.join(clone, "src")))
    for index, name in enumerate(modules + ["../tests/M_STATS_PROBDIST_TEST.bas"]):
        path = os.path.join(clone, "src", name)
        mutate(path, b"Option Explicit",
               b"Option Explicit\r\n' fixture comment (#338)\r\n"
               + f"Private Const FIXTURE_338_{index} As Long = 1".encode())
    for rel in FIXTURES:
        code, tail = run(clone, rel)
        check(code == 0, f"2: {rel} must pass whatever the working-tree sources hold\n{tail}")

    # 3. The narrow working-tree check. When production has moved on, the
    # committed check skips the comparison on purpose; the clone's copy is
    # switched back on, so the detection itself stays proved.
    live_check = os.path.join(clone, "benchmark", *LIVE_CHECK.split("/"))
    with open(live_check, encoding="utf-8") as f:
        check_text = f.read()
    check_text = re.sub(r"^PRODUCTION_FOLLOWS_STUDY = .*$", "PRODUCTION_FOLLOWS_STUDY = True",
                        check_text, count=1, flags=re.M)
    check_text, replaced = re.subn(r'^MOVED_ON = (""".*?"""|"[^"\n]*"|\(.*?\))$', 'MOVED_ON = ""',
                                   check_text, count=1, flags=re.M | re.S)
    check(replaced == 1 and "PRODUCTION_FOLLOWS_STUDY = True" in check_text,
          "3: the live check declares PRODUCTION_FOLLOWS_STUDY and MOVED_ON as single assignments")
    with open(live_check, "w", encoding="utf-8") as f:
        f.write(check_text)
    special = os.path.join(clone, "src", "M_STATS_PROBDIST_SPECIALFUNCS.bas")
    code, tail = run(clone, LIVE_CHECK)
    check(code == 0, f"3: changes outside the generated block leave the live check green\n{tail}")
    mutate(special, BLOCK, BLOCK + b"\r\n' a comment inside the block")
    code, tail = run(clone, LIVE_CHECK)
    check(code == 0, f"3: a comment inside the generated block is not a code change\n{tail}")
    mutate(special, BLOCK, BLOCK + b"\r\nPrivate Const FIXTURE_338_BLOCK As Long = 1")
    code, tail = run(clone, LIVE_CHECK)
    check(code != 0, "3: a code change inside the generated block must fail the live check")

if fails:
    print(f"FAIL: {len(fails)} frozen-source check(s)")
    for m in fails:
        print("  - " + m)
    raise SystemExit(1)
print(f"PASS: the pinned sources are the Git blobs; {len(FIXTURES)} closed-study fixtures pass "
      "with every production and test module changed; the generated public Gamma block is "
      "still checked against the working tree")
