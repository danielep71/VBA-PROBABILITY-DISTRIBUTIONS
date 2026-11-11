"""
Summarize the TIMING and COST lines of an Excel regression log (#298).

TIMING lines give the seconds each test procedure of the suite took, grouped
by the suite driver that called it. COST lines give the microseconds per call
of the double-double Gamma P and Q (dd_p, dd_q) and of the binary64 routine of
the same region (binary64) at fixed points.

Run: python3 summarize_timing.py [LOG]   (default: the retained run 42766552574)
"""
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT = os.path.join(HERE, "run_42766552574_log.txt")
TIMING = re.compile(r" TIMING (\S+) (\S+)$")
COST = re.compile(r" COST (\w+) A=(\S+) X=(\S+) ok=(\d) calls=(\d+) us_per_call=(\S+)$")


def parse(path):
    suites, procs, cost = {}, [], {}
    suite = None
    with open(path, encoding="utf-8-sig") as f:
        for line in f:
            line = line.rstrip()
            m = TIMING.search(line)
            if m:
                label, value = m.groups()
                if label == "suite":
                    suite = value
                    suites.setdefault(suite, 0.0)
                else:
                    seconds = float(value)
                    suites[suite] += seconds
                    procs.append((seconds, suite, label))
                continue
            m = COST.search(line)
            if m:
                route, a, x, ok, calls, us = m.groups()
                cost[(float(a), float(x), route)] = (float(us), int(calls), ok == "1")
    return suites, procs, cost


def main(argv=None):
    path = (argv or sys.argv[1:] or [DEFAULT])[0]
    suites, procs, cost = parse(path)
    total = sum(suites.values())
    print(f"Timed test procedures: {len(procs)}, {total:.1f} s in total")
    for name, seconds in suites.items():
        print(f"  {name:22s} {seconds:8.1f} s {100 * seconds / total:6.1f}%")
    print("Slowest procedures:")
    for seconds, suite, label in sorted(procs, reverse=True)[:5]:
        print(f"  {seconds:8.1f} s {100 * seconds / total:6.1f}%  {label} ({suite})")
    print("Per-call cost (microseconds): A, X, double-double P, Q, binary64, ratio")
    for a, x in sorted({(a, x) for a, x, _ in cost}):
        p, q, b = (cost[(a, x, r)][0] for r in ("dd_p", "dd_q", "binary64"))
        print(f"  {a:>8g} {x:>8g} {p:12.1f} {q:12.1f} {b:9.3f} {max(p, q) / b:9.0f}x")
    return 0


if __name__ == "__main__":
    sys.exit(main())
