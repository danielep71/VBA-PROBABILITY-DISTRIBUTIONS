"""Retained-row rejection gate; fitting/holdout are never read by this runner."""
import hashlib
import json
from pathlib import Path
import mpmath as mp
import run_baseline as baseline
import vba_order_model as model

HERE = Path(__file__).resolve().parent


def main():
    original = json.loads((HERE/'baseline_result.json').read_text())
    if original['status'] != 'PASS':
        raise RuntimeError('Baseline faithfulness failed')
    for path, digest in original['bindings'].items():
        if hashlib.sha256((model.ROOT/path).read_bytes()).hexdigest() != digest:
            raise RuntimeError('Stale baseline: '+path)
    results = []
    for candidate in ['F-A','F-B']:
        rows = []
        for old in original['rows']:
            p,d1,d2 = map(float.fromhex,old['inputs'])
            row = dict(inputs=old['inputs'],kind=old['kind'])
            try:
                value,x,y,trace = model.f_inverse(p,d1,d2,candidate)
                r50 = baseline.reference(p,d1,d2,50)
                r80 = baseline.reference(p,d1,d2,80)
                with mp.workdps(80):
                    stable = all(abs(a-b)<=mp.mpf('1e-40')*abs(b) for a,b in zip(r50,r80))
                    error = float(abs(mp.mpf(value)-r80[0])/r80[0])
                passed = stable and error < 2**-48
                if old['kind']=='control':
                    passed = passed and error <= 4*old['error']
                row.update(error=error,passes=passed,oracle_stable=stable,
                           pair=[x.hex(),y.hex()],result=value.hex(),trace=trace)
            except (ValueError,ArithmeticError) as exc:
                row.update(passes=False,failure=str(exc))
            rows.append(row)
            print(candidate,old['kind'],p,d1,d2,row.get('error',row.get('failure')),row['passes'],flush=True)
        results.append(dict(candidate=candidate,retained_pass=all(r['passes'] for r in rows),rows=rows))
    (HERE/'retained_candidates.json').write_text(json.dumps(dict(
        stage='retained rows only',fitting_evaluated=False,holdout_evaluated=False,
        baseline_sha256=hashlib.sha256((HERE/'baseline_result.json').read_bytes()).hexdigest(),
        runner_sha256=hashlib.sha256(Path(__file__).read_bytes()).hexdigest(),results=results),indent=2)+'\n')


if __name__ == '__main__':
    main()
