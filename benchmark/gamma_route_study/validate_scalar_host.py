"""Reject incomplete B5 evidence and compare contemporaneous B4/B5 warm costs."""
import hashlib
import json
import os
import statistics
import struct
import sys
from pathlib import Path

from validate_pair_host import validate as validate_pair


def validate(report, design):
    assert design['study'] == 'B5_scalar' and design['baseline'] == 'B4'
    assert report['candidate_sha'] == design['candidate_sha']
    assert report['arithmetic_check'] == 'PASS checks=576'
    points = validate_pair(report, design)
    expected = {r['id']: r['expected'] for r in design['cases']}
    # Both modes implement the frozen B4 graph; baseline mismatch is also fatal.
    for row in report['timings']:
        key = ':'.join(struct.pack('>d', row[v]).hex().upper() for v in ('a', 'x'))
        assert row['p'] + ',' + row['q'] == expected[key]
    return points


def costs(report, points):
    result = []
    for a, x in points:
        med = {b: statistics.median(r['elapsed_us'] / r['calls'] for r in report['timings']
                                   if (r['a'], r['x'], r['baseline']) == (a, x, b))
               for b in (True, False)}
        result.append(dict(a=a, x=x, b4_us=med[True], b5_us=med[False],
                           reduction=1 - med[False] / med[True]))
    return result


def main():
    path = Path(sys.argv[1])
    report = json.loads((path / 'scalar-results.json').read_text(encoding='utf-8-sig'))
    design = json.loads((path / 'inputs.json').read_text())
    points = validate(report, design)
    assert len(design['cases']) == 1371
    assert sum(r['expected'] == 'DECLINED' for r in design['cases']) == 106
    if os.environ.get('GITHUB_SHA'):
        assert report['candidate_sha'] == os.environ['GITHUB_SHA']
    for name, key in (('candidate.txt', 'candidate_sha256'), ('probe.txt', 'probe_sha256')):
        assert hashlib.sha256((path / name).read_bytes()).hexdigest() == design[key]
    here = Path(__file__).resolve().parent
    for name, expected in design['sha256'].items():
        assert hashlib.sha256((here / name).read_text().encode()).hexdigest() == expected
    rows = costs(report, points)
    useful = all(r['reduction'] >= .05 for r in rows if r['a'] == 10000)
    no_regression = all(r['reduction'] >= -.05 for r in rows)
    within_budget = all(r['b5_us'] <= 100 for r in rows)
    print('PASS: 576 scalar arithmetic checks, 11 normalization checks, 1371 exact inputs, 48 timing rows')
    print('| A | X | B4 pair us | B5 scalar pair us | Cost reduction |')
    print('|---:|---:|---:|---:|---:|')
    for r in rows:
        print(f"| {r['a']:g} | {r['x']:g} | {r['b4_us']:.3f} | {r['b5_us']:.3f} | {r['reduction']:.1%} |")
    print('Development signal (>=5% reduction at both A=10000 points):', useful)
    print('No material regression (<=5% extra cost at all eight points):', no_regression)
    print('Conservative 100-us paired budget at all measured points:', within_budget)
    print('Warm paired kernels only. No cold/public/inverse claim, holdout acceptance or production promotion.')
    (path / 'scalar-summary.json').write_text(json.dumps(dict(
        candidate_sha=report['candidate_sha'], rows=rows, development_signal=useful,
        no_material_regression=no_regression, paired_budget_met=within_budget), indent=2) + '\n')


if __name__ == '__main__':
    main()
