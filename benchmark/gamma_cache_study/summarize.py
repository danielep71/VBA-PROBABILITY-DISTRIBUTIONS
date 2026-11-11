"""Validate cache differential measurements and print paired timing medians."""
import json
import statistics
import sys
from pathlib import Path


def summarize(path, emit=True):
    data = json.loads(Path(path).read_text(encoding='utf-8-sig'))
    assert data['status'] == 'PASS', data.get('error')
    assert data['private_checks'] == {
        'uncached': 'PASS checks=40 preserved_error=5',
        'cached': 'PASS checks=40 preserved_error=5',
    }
    points = ((.0001, .5), (.5, .25), (2.5, 1), (2.5, 6),
              (100, 90), (100, 110), (10000, 9900), (10000, 10100))
    expected = {(a, x, r) for a, x in points for r in ('p', 'q')}
    expected |= {(2, .5, 'inverse'), (100, .9, 'inverse')}
    groups = {}
    for row in data['samples']:
        key = row['a'], row['x'], row['route']
        groups.setdefault(key, []).append(row)
    assert set(groups) == expected
    assert len(data['samples']) == 216
    for key, rows in groups.items():
        assert len(rows) == 12
        assert len({r['bits'] for r in rows}) == 1, key
        assert {(r['sample'], r['mode'], r['cold']) for r in rows} == {
            (n, mode, cold) for n in (1, 2, 3)
            for mode in ('uncached', 'cached') for cold in (True, False)}
        assert all(r['calls'] == (1 if key[2] == 'inverse' else 4) for r in rows)
    assert sum(r['calls'] + 1 for r in data['samples']) == 1008
    if not emit:
        return data
    print('PASS: 216 samples, 1008 target evaluations, 40 private checks per path')
    print('| Shape | x/probability | Function | Cache | Original ms | Cached ms | Ratio |')
    print('|---:|---:|:---|:---|---:|---:|---:|')
    for key, rows in groups.items():
        for cold in (True, False):
            values = {mode: statistics.median(r['elapsed_us'] / r['calls'] / 1000
                      for r in rows if r['cold'] == cold and r['mode'] == mode)
                      for mode in ('uncached', 'cached')}
            a, x, route = key
            old, new = values['uncached'], values['cached']
            state = 'cold' if cold else 'warm'
            print(f'| {a:g} | {x:g} | {route} | {state} | {old:.3f} | {new:.3f} | {old/new:.2f}x |')
    return data


if __name__ == '__main__':
    summarize(sys.argv[1])
