"""Replay revealed fitting sets; never call them independent validation."""
import json
import math
from pathlib import Path
import mpmath as mp
import precision_v2 as candidate

HERE = Path(__file__).resolve().parent


def main():
    count = 0
    failures = []
    previous = {}
    with mp.workdps(120):
        for filename in ['b7_screen.json', 'precision_screen.json']:
            for row in json.loads((HERE / filename).read_text())['rows']:
                a, x = float.fromhex(row['a_hex']), float.fromhex(row['x_hex'])
                if a > 1:
                    continue
                values = candidate.public(a, x)
                errors = [abs(mp.mpf(v) / mp.mpf(r) - 1)
                          for v, r in zip(values, row['reference'])]
                count += 1
                if errors[0] > mp.mpf('3e-15') or errors[1] > mp.mpf('5e-15'):
                    failures.append([filename, row['index'], 'accuracy'])
                key = (filename, row.get('role', ''), a)
                if key in previous:
                    px, pv = previous[key]
                    if math.nextafter(px, math.inf) == x:
                        if pv[0] > values[0] or pv[1] < values[1]:
                            failures.append([filename, row['index'], 'monotonicity'])
                previous[key] = (x, values)
    if count != 489:
        raise RuntimeError('unexpected fitting coverage')
    print('REVEALED FITTING ONLY', count, 'failures', failures)
    return 1 if failures else 0


if __name__ == '__main__':
    raise SystemExit(main())
