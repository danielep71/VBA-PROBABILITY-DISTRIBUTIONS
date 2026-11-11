"""
The binary64 route block that #312 appended to src/M_STATS_PROBDIST_SPECIALFUNCS.bas
and that left production after #307 closed, as frozen text. Standard library only.

Studies that measured the module while it carried the block (the LogGamma
cache and recurrence studies) bind a SHA-256 of the whole module. They compare
as_measured(current), the current module with this block appended, so any other
drift is still detected. b64_route_vba.py --check proves that this file is
byte for byte the block it generates and that the fit candidate 88ff6ba carried.
"""
from pathlib import Path

BLOCK_FILE = Path(__file__).with_name("b64_route_block.txt")
BEGIN = "' BEGIN generated binary64 Gamma route (#307)"
ANCHOR = "' END generated public Gamma integration\n\n"


def block():
    """The block as appended: the file stores it without its final blank line,
    which Git's whitespace check refuses at the end of a file."""
    return BLOCK_FILE.read_text(encoding="ascii") + "\n"


def as_measured(text):
    """The module text (LF) as the #307 studies measured it."""
    if BEGIN in text:
        return text
    if not text.endswith(ANCHOR):
        raise ValueError("the module does not end with the public Gamma block")
    return text + block()
