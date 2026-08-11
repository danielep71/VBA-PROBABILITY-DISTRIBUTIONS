"""
Guard against measured accuracy thresholds being duplicated in the public VBA
source headers, where they drift out of step with the authoritative registry
(accuracy_contracts.csv). This is the P2-01 failsafe: the F-quantile header claim
once read 5.9E-13 while the frozen contract was 2E-10, because the number was
restated in source instead of referenced.

Public source headers should point to accuracy_contracts.csv rather than restate
a threshold. This check scans src/*.bas for threshold and CAP claims and exits
non-zero if any remain, listing them.

WHY THERE IS MORE THAN ONE PATTERN
    The original check recognised only "<= <number> relative|absolute". An
    external review found five live contradictions that phrasing did not catch,
    every one of which had drifted from the code beside it:

      "df above 1E5 are REJECTED"          while PROB_F_MAX_DF was 1E10
      "validated to roughly 1E9"           while the cap was 1E8
      "up to 1E100 is accepted"            while the wrapper rejected past 1E8
      "density and survival to 5E-15"      while the frozen contract was 1E-14
      "is closed-form and unrestricted"    while the density cap was 1E20

    A checker that reports clean while the source contradicts the registry is
    worse than no checker, because it converts an open question into a false
    assurance. The patterns below therefore cover caps and accuracy claims in
    the phrasings that actually occurred, not just the canonical one.

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
CLAIMS = [
    # The canonical "<= N.NeM relative/absolute" contract-claim form.
    ("threshold claim",
     re.compile(r"<=\s*[0-9.]+\s*[Ee]-?[0-9]+\s*(relative|absolute)", re.IGNORECASE)),
    # "... to 5E-15" / "... to 2E-14" as an accuracy statement. Requires an
    # accuracy noun on the line so ordinary prose containing a number is not
    # flagged.
    ("accuracy claim",
     re.compile(r"\b(density|survival|cdf|quantile|residual|accuracy|contract)\b"
                r"[^\n]{0,60}\bto\s+[0-9.]+\s*[Ee]-[0-9]+", re.IGNORECASE)),
    # A restated public cap: "df above 1E5 are REJECTED", "up to 1E100 is accepted",
    # "validated to roughly 1E9". These belong to the PROB_*_MAX_DF constants and
    # numerical_limitations.csv.
    ("cap claim",
     re.compile(r"\b(?:df|shape|degrees of freedom)\b[^\n]{0,40}"
                r"\b(?:above|beyond|up to|over)\s+[0-9.]+\s*[Ee]\+?[0-9]+", re.IGNORECASE)),
    ("cap claim",
     re.compile(r"\bvalidated to (?:roughly |about |approximately )?[0-9.]+\s*[Ee]\+?[0-9]+",
                re.IGNORECASE)),
    # "unrestricted" / "NOT enveloped" assertions about a public surface. Every
    # density now carries PROB_DENSITY_SHAPE_MAX, so these claims went stale
    # silently.
    # "unrestricted" / "NOT enveloped" said ABOUT A DENSITY OR ENVELOPE. Every
    # density now carries PROB_DENSITY_SHAPE_MAX, so these claims went stale
    # silently. Requires the subject nearby so "an unrestricted scale parameter"
    # - a statement about an argument, not an envelope - is not flagged.
    ("envelope claim",
     # (?:\b|_) not \b: "F_Density" is one word to the regex engine because the
     # underscore is a word character, so \bdensity\b never matches inside it.
     re.compile(r"(?:\b|_)(?:density|envelope|enveloped)\b[^\n]{0,60}"
                r"\b(?:unrestricted|not enveloped)\b"
                r"|\b(?:unrestricted|not enveloped)\b[^\n]{0,60}"
                r"(?:\b|_)(?:density|envelope|enveloped)\b", re.IGNORECASE)),
]

# A constant declaration IS the authoritative value, not a restatement of one.
# Flagging it would tell the maintainer to delete the very thing the prose is
# supposed to point at.
AUTHORITATIVE = re.compile(r"^\s*(?:Private|Public)\s+Const\b", re.IGNORECASE)


def main():
    offenders = []
    for path in sorted(glob.glob(os.path.join(ROOT, "src", "*.bas"))):
        with open(path, encoding="utf-8", errors="replace") as f:
            for lineno, line in enumerate(f, 1):
                if AUTHORITATIVE.match(line):
                    continue
                for kind, pattern in CLAIMS:
                    if pattern.search(line):
                        offenders.append((os.path.relpath(path, ROOT).replace(os.sep, "/"),
                                          lineno, kind, line.strip()))
                        break
    if offenders:
        print("Measured-threshold claims found in source headers. Move these to "
              "accuracy_contracts.csv and reference it instead:")
        for rel, lineno, kind, text in offenders:
            print(f"  {rel}:{lineno} [{kind}]: {text}")
        sys.exit(1)
    print("OK: no measured thresholds or public caps duplicated in src/*.bas "
          "(accuracy_contracts.csv and numerical_limitations.csv are "
          "authoritative).")


if __name__ == "__main__":
    main()
