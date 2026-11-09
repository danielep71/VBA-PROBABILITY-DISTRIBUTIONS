"""Proposed bounded dispatch only; never imported by production or its oracle."""
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))
import extended_gamma_v2
import seam_diagnosis


def proposed(a, x):
    if 0 < a <= 128 and 0 < x <= 1024:
        return extended_gamma_v2.public(a, x)
    return seam_diagnosis.public(a, x)
