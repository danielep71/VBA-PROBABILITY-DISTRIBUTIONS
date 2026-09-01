"""
Fixtures for the root README structural guard (check_root_readme.py).

The guard exists because the root README was replaced by an unrelated study
README and nothing caught it. These fixtures prove it catches that case and
the adjacent ones, and - just as important - that it does NOT fire on the
legitimate edits the file is expected to receive.

Run: python3 test_root_readme.py   (exit 0 = pass, nonzero = fail)
"""
import os

import check_root_readme as G
from _manifest import repo_root

fails = []


def check(cond, msg):
    if not cond:
        fails.append(msg)


def real_readme():
    with open(os.path.join(repo_root(), "README.md"), encoding="utf-8") as f:
        return f.read()


REAL = real_readme()
REAL_LINES = REAL.count("\n") + 1

# 1. the committed root README passes
check(G.check(REAL, REAL_LINES) == [], "real root README passes the guard")

# 2. THE FAILURE THAT ACTUALLY HAPPENED: replaced by the study README.
STUDY = (
    "# Chi-square inverse reference study (#22, v1.0.0 plan Track A2 item 6)\n"
    "\n**Status: PROVISIONAL.**\n\n## Frozen design\n\n69 references.\n"
)
probs = G.check(STUDY, STUDY.count("\n") + 1)
check(any("overwritten" in p for p in probs),
      "study README at the root path is detected as an overwrite")
check(any("line floor" in p or "below the" in p for p in probs),
      "study README at the root path also trips the length floor")

# 3. truncation: the real document cut to its first 50 lines
TRUNC = "\n".join(REAL.splitlines()[:50])
probs = G.check(TRUNC, 50)
check(any("below the" in p for p in probs), "truncated root README detected")

# 4. a major section dropped, everything else intact
for anchor in ("# 🧩 Distribution catalogue", "# 📄 License"):
    missing = REAL.replace(anchor, "# Something else entirely")
    probs = G.check(missing, missing.count("\n") + 1)
    check(any("required section missing" in p and anchor in p for p in probs),
          f"dropped section detected: {anchor}")

# 5. sections reordered - right headings, wrong document shape
lines = REAL.splitlines()
i_cat = next(i for i, l in enumerate(lines) if l.startswith("# 🧩 Distribution catalogue"))
i_lic = next(i for i, l in enumerate(lines) if l.startswith("# 📄 License"))
swapped = lines[:]
swapped[i_cat], swapped[i_lic] = swapped[i_lic], swapped[i_cat]
swapped_text = "\n".join(swapped)
probs = G.check(swapped_text, len(swapped))
check(any("out of order" in p for p in probs), "reordered sections detected")

# 6. LEGITIMATE EDITS MUST PASS. A guard that fires on ordinary maintenance
# gets deleted, so these matter as much as the negative cases.

# 6a. #28 regenerating the assurance figures.
regenerated = REAL
for old, new in (("**112**", "**118**"), ("**835**", "**909**"),
                 ("**161**", "**166**"), ("**1 905**", "**2 088**")):
    regenerated = regenerated.replace(old, new)
check(G.check(regenerated, regenerated.count("\n") + 1) == [],
      "regenerated assurance figures still pass (#28 must not trip the guard)")

# 6b. ordinary prose editing.
reworded = REAL.replace(
    "A transparent, tail-aware numerical probability library for pure Excel VBA",
    "A transparent, tail-aware probability library for Excel VBA")
check(G.check(reworded, reworded.count("\n") + 1) == [],
      "reworded prose still passes")

# 6c. adding a new section.
extended = REAL + "\n\n# 🔬 New section\n\nAdded later.\n"
check(G.check(extended, extended.count("\n") + 1) == [],
      "an added section still passes")

if fails:
    print("FAIL: root README guard")
    for f in fails:
        print("  - " + f)
    raise SystemExit(1)
print("PASS: root README guard (real file, study-README overwrite, truncation, "
      "dropped section, reordering, and three legitimate-edit cases)")
