"""Replay historical #26 evidence against its pinned source, never current VBA.

Current study inputs are copied so their existing drift checks remain active.
Only VBA source is restored from the immutable, preregistered baseline. No
observations or reports are regenerated, and no Git reference is moved.
"""
import argparse
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
BASELINE = "23a56f2ea60ab9391a93367431dabeefc51e5d37"


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("study", choices=("isolation", "candidate"))
    args = parser.parse_args()
    git_dir = subprocess.check_output(
        ["git", "rev-parse", "--absolute-git-dir"], cwd=ROOT, text=True
    ).strip()
    with tempfile.TemporaryDirectory(prefix="gamma-historical-") as folder:
        root = Path(folder)
        bench = root / "benchmark"
        bench.mkdir()
        shutil.copytree(HERE, bench / HERE.name, ignore=shutil.ignore_patterns("__pycache__"))
        helper = bench / "beta_balanced_survival"
        helper.mkdir()
        shutil.copy2(ROOT / "benchmark/beta_balanced_survival/beta_route_model.py", helper)
        grid = bench / "positive_ratio_subnormal"
        grid.mkdir()
        shutil.copy2(ROOT / "benchmark/positive_ratio_subnormal/gamma_q_complement.csv", grid)
        (root / "src").mkdir()
        for name in ("CORE", "SPECIALFUNCS", "CONTINUOUS"):
            path = f"src/M_STATS_PROBDIST_{name}.bas"
            data = subprocess.check_output(["git", "show", f"{BASELINE}:{path}"], cwd=ROOT)
            (root / path).write_bytes(data)
        env = dict(os.environ, GIT_DIR=git_dir, GIT_WORK_TREE=str(root))
        script = "run_isolation.py" if args.study == "isolation" else "check_candidate.py"
        result = subprocess.run([sys.executable, str(bench / HERE.name / script)],
                                cwd=root, env=env, check=False)
        raise SystemExit(result.returncode)


if __name__ == "__main__":
    main()
