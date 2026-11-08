"""Generate exact-binary64 Gamma P/Q references; never evaluate VBA outputs."""
import json
from pathlib import Path
import mpmath as mp

HERE = Path(__file__).resolve().parent


def main():
    protocol = json.loads((HERE / 'protocol.json').read_text())
    output = []
    for point in protocol['points']:
        a, x = (float.fromhex(point[k]) for k in ('a_hex', 'x_hex'))
        values = []
        for dps in protocol['reference_dps']:
            with mp.workdps(dps):
                values.append((mp.gammainc(mp.mpf(a), 0, mp.mpf(x), regularized=True),
                               mp.gammainc(mp.mpf(a), mp.mpf(x), mp.inf, regularized=True)))
        with mp.workdps(protocol['reference_dps'][-1]):
            for low, high in zip(*values):
                if not mp.isfinite(high) or high <= 0:
                    raise ValueError('unavailable or nonpositive reference')
                if abs(low / high - 1) > mp.mpf(protocol['stability_limit']):
                    raise ValueError('unstable reference')
            output.append(dict(point, p=mp.nstr(values[-1][0], 100),
                               q=mp.nstr(values[-1][1], 100)))
    (HERE / 'references.json').write_text(json.dumps(output, indent=2) + '\n')
    print(f'{len(output)} reference pairs stable at both precisions; no VBA observations')


if __name__ == '__main__':
    main()
