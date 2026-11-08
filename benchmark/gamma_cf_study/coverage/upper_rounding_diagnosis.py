"""Post-hoc decomposition of the revealed A=.39 pair; not a candidate fit."""
import hashlib
import json
import math
from pathlib import Path
import mpmath as mp
import seam_diagnosis as model

HERE = Path(__file__).resolve().parent


def main():
    report = {'status': 'POST-HOC MODEL DIAGNOSIS ONLY', 'rows': []}
    with mp.workdps(120):
        a = .39
        for raw in ['0x1.63d70a3d70a41p+0', '0x1.63d70a3d70a42p+0']:
            x = float.fromhex(raw)
            _, depth = model.cf(a, x, eps=2**-52)
            h = model.backward(a, x, depth)
            X, D = model.X, model.D
            e = D(X(a) * X(math.log(x)) - X(model.lg(a)))
            f = D(X(D(X(a) * X(math.exp(-x)))) * X(math.exp(e)))
            exact_f = mp.exp(mp.mpf(a) * mp.log(mp.mpf(x)) - mp.mpf(x)
                             - mp.loggamma(mp.mpf(a)))
            exact_h, _ = model.cf(a, x, steps=depth, exact=True)
            ref = mp.gammainc(mp.mpf(a), mp.mpf(x), mp.inf, regularized=True)
            report['rows'].append({'a_hex': a.hex(), 'x_hex': raw, 'depth': depth,
                'prefactor_hex': f.hex(), 'backward_hex': h.hex(),
                'q_hex': model.upper(a, x).hex(),
                'exact_prefactor_only_hex': D(X(float(exact_f)) * X(h)).hex(),
                'exact_fraction_only_hex': D(X(f) * X(float(exact_h))).hex(),
                'both_exact_hex': float(exact_f * exact_h).hex(),
                'reference_q': mp.nstr(ref, 100)})
    paths = [Path(__file__), HERE / 'seam_diagnosis.py', HERE.parent / 'model.py',
             model.ROOT / 'src/M_STATS_PROBDIST_SPECIALFUNCS.bas']
    report['sha256_lf'] = {str(p.relative_to(model.ROOT)): hashlib.sha256(
        p.read_bytes().replace(b'\r\n', b'\n')).hexdigest() for p in paths}
    (HERE / 'upper_rounding_diagnosis.json').write_text(json.dumps(report, indent=2) + '\n')
    print('Post-hoc diagnosis retained; no acceptance verdict')


if __name__ == '__main__':
    main()
