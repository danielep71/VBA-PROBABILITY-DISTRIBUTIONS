"""Proposed bounded dispatch only; never imported by production or its oracle."""
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))
import scaled_gamma
import seam_diagnosis


def proposed(a, x):
    if 0 < a <= 1 and 0 < x <= 32:
        return scaled_gamma.public(a, x)
    return seam_diagnosis.public(a, x)
