"""Production sources at a pinned commit, for closed studies (#338).

A closed study proves facts about the sources it measured. Its replay
fixtures and its preparation tooling must not read src/ or tests/ from the
working tree: every later production change, a comment included, then breaks
them, and the only repair is another inverse in the study's own tooling.

GAMMA_STUDY_PIN is release/1.0.0 as it stood when the Gamma studies closed
(#307) and before #13 changed the production modules again. The Gamma
continued-fraction, cache and recurrence studies read their sources from it.

    text(rel)      the file at the pin, as text with LF line ends
    tree(paths)    a temporary directory holding those paths at the pin

Files are read as Git blobs. `git archive` is not used: it honours
export-ignore, which drops .github/. Standard library and git only.
"""
import atexit
import shutil
import subprocess
import tempfile
from functools import lru_cache
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
GAMMA_STUDY_PIN = "6c6a2d74daba677f3511c4230d2628e63c443004"


def blob(rel, pin=GAMMA_STUDY_PIN):
    return subprocess.run(["git", "-C", str(ROOT), "show", f"{pin}:{rel}"],
                          check=True, capture_output=True).stdout


def text(rel, pin=GAMMA_STUDY_PIN):
    """The file at the pin, decoded, with LF line ends, as Path.read_text gives."""
    return blob(rel, pin).decode("utf-8").replace("\r\n", "\n")


@lru_cache(maxsize=None)
def tree(paths, pin=GAMMA_STUDY_PIN):
    """A temporary directory holding `paths` (a tuple) as committed at the pin."""
    target = Path(tempfile.mkdtemp(prefix="frozen-sources-"))
    atexit.register(shutil.rmtree, target, ignore_errors=True)
    listed = subprocess.run(["git", "-C", str(ROOT), "ls-tree", "-r", "--name-only", pin, "--", *paths],
                            check=True, capture_output=True, text=True).stdout.split("\n")
    for rel in filter(None, listed):
        path = target / rel
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(blob(rel, pin))
    return target
