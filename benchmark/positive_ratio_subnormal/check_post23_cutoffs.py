"""Replay the post-#23 fitting crossovers and frozen holdout claims."""
import hashlib
import json
from pathlib import Path
import analyze_positive_ratio_subnormal as study

HERE = Path(__file__).resolve().parent


def require(condition, detail):
    if not condition:
        raise RuntimeError(str(detail))


def main():
    record = json.loads((HERE / "post23_cutoffs.json").read_text())
    for name, expected in record["fitting"].items():
        path = HERE / (name + ".csv")
        require(hashlib.sha256(path.read_bytes()).hexdigest() == expected["sha256"], name)
        points = study.load(path)
        require(not study.verify_constructions(points), name)
        require(len(points) == expected["rows"], name)
        for surface, bits in expected["crossovers"].items():
            actual = study.crossover_of(study.envelope(
                [p for p in points if p.surface == surface], True))[0]
            require(actual == bits, (name, surface, actual, bits))
    for name, expected in record["holdout"].items():
        path = HERE / (name + ".csv")
        require(hashlib.sha256(path.read_bytes()).hexdigest() == expected["sha256"], name)
        points = study.load(path)
        require(len(points) == expected["rows"], name)
        require(not study.verify_constructions(points), name)
        require(study.test_claims(points), name)
    print("PASS: post-#23 fitting crossovers and frozen holdout claims reproduced")


if __name__ == "__main__":
    main()
