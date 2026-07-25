"""
Guard against measured accuracy thresholds being duplicated in the public VBA
source headers, where they drift out of step with the authoritative registry
(accuracy_contracts.csv). This is the P2-01 failsafe: the F-quantile header claim
once read 5.9E-13 while the frozen contract was 2E-10, because the number was
restated in source instead of referenced.

Public source headers should point to accuracy_contracts.csv rather than restate
a threshold. This check scans src/*.bas for the standard threshold-claim form
"<= <number> relative|absolute" and exits non-zero if any remain, listing them.

Run: python3 check_source_thresholds.py   (exit 0 = clean, 1 = drift risk found)
"""
import glob
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)

# The "<= N.NeM relative/absolute" contract-claim form. Deliberately narrow so it
# flags restated measured thresholds without tripping on incidental numbers
# (e.g. "delta is 1.67E-04 near N = 501", which is not a threshold claim).
CLAIM = re.compile(r"<=\s*[0-9.]+\s*[Ee]-?[0-9]+\s*(relative|absolute)", re.IGNORECASE)


def main():
    offenders = []
    for path in sorted(glob.glob(os.path.join(ROOT, "src", "*.bas"))):
        with open(path, encoding="utf-8", errors="replace") as f:
            for lineno, line in enumerate(f, 1):
                if CLAIM.search(line):
                    offenders.append((os.path.relpath(path, ROOT).replace(os.sep, "/"),
                                      lineno, line.strip()))
    if offenders:
        print("Measured-threshold claims found in source headers. Move these to "
              "accuracy_contracts.csv and reference it instead:")
        for rel, lineno, text in offenders:
            print(f"  {rel}:{lineno}: {text}")
        sys.exit(1)
    print("OK: no measured thresholds duplicated in src/*.bas "
          "(accuracy_contracts.csv is authoritative).")


if __name__ == "__main__":
    main()
