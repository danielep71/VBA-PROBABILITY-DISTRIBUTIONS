"""Validate complete four-way differential evidence and report every case."""
import json
import statistics
import sys
from pathlib import Path

MODES = ('baseline', 'lower', 'upper', 'combined')
POINTS = ((.0001, .5), (.5, .25), (2.5, 1.), (2.5, 6.), (100., 90.),
          (100., 110.), (10000., 9900.), (10000., 10100.))
CASES = [(a, x, r) for a, x in POINTS for r in ('p', 'q')] + [(2., .5, 'inverse'), (100., .9, 'inverse')]


def summarize(path, emit=True):
    data = json.loads(Path(path).read_text(encoding='utf-8-sig'))
    assert data['status'] == 'PASS' and not data['error']
    assert len(data['candidate_sha']) == 40
    rows = data['samples']
    expected = {(a, x, r, m, n) for a, x, r in CASES for m in MODES for n in (1, 2, 3)}
    actual = {(r['a'], r['x'], r['route'], r['mode'], r['sample']) for r in rows}
    assert len(rows) == len(actual) == 216 and actual == expected
    for row in rows:
        assert row['cold'] is False
        assert row['calls'] == (1 if row['route'] == 'inverse' else 4)
        assert row['elapsed_us'] > 0
        assert len(row['bits']) == 16 and int(row['bits'], 16) >= 0
    if emit:
        print('| Shape | x / p | Route | Baseline ms | Lower speedup | Upper speedup | Combined speedup |')
        print('|---:|---:|:---|---:|---:|---:|---:|')
    for a, x, route in CASES:
        case = [r for r in rows if (r['a'], r['x'], r['route']) == (a, x, route)]
        assert len({r['bits'] for r in case}) == 1, (a, x, route, 'different bits')
        medians = {m: statistics.median(r['elapsed_us'] / r['calls'] for r in case if r['mode'] == m) for m in MODES}
        if emit:
            ratios = ' | '.join(f"{medians['baseline'] / medians[m]:.3f}x" for m in MODES[1:])
            print(f"| {a:g} | {x:g} | {route} | {medians['baseline'] / 1000:.3f} | {ratios} |")
    assert sum(r['calls'] + 1 for r in rows) == 1008
    return data


if __name__ == '__main__':
    summarize(sys.argv[1])
