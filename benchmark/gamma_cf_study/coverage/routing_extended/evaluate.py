"""Evaluate the frozen bounded routing preflight, retaining every result."""
import hashlib
import json
import math
from pathlib import Path
import subprocess
import sys

import mpmath as mp
from candidate import proposed, seam_diagnosis

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]


def digest(path):
    return hashlib.sha256(path.read_bytes().replace(b'\r\n', b'\n')).hexdigest()


def check_identity(protocol):
    for name, expected in protocol['sha256_lf'].items():
        if digest(ROOT/name) != expected:
            raise ValueError('frozen source changed: '+name)


def reference(a, x):
    refs = []
    for precision in (160, 320):
        with mp.workdps(precision):
            aa, xx = mp.mpf(a), mp.mpf(x)
            refs.append((mp.gammainc(aa, 0, xx, regularized=True),
                         mp.gammainc(aa, xx, mp.inf, regularized=True)))
    if [float(r).hex() for r in refs[0]] != [float(r).hex() for r in refs[1]]:
        raise ValueError('oracle binary64 rounding is unstable')
    return refs[1]


def score(values, refs):
    failures = []
    errors = []
    with mp.workdps(320):
        for j, (value, ref) in enumerate(zip(values, refs)):
            if not math.isfinite(value) or not 0 <= value <= 1:
                raise ValueError('invalid probability')
            error = abs(mp.mpf(value)/ref-1)
            errors.append(mp.nstr(error, 30))
            if float(ref) < sys.float_info.min:
                bad = value.hex() != float(ref).hex()
            else:
                bad = error > mp.mpf(('3e-15', '5e-15')[j])
            if bad:
                failures.append(j)
    return errors, failures


def run():
    protocol = json.loads((HERE/'protocol.json').read_text())
    check_identity(protocol)
    report = dict(status='INCOMPLETE', protocol_commit=subprocess.check_output(
        ['git','log','-1','--format=%H','--',str(HERE/'protocol.json')],cwd=ROOT).decode().strip(),
        protocol_sha256=digest(HERE/'protocol.json'), mpmath=mp.__version__, groups=[],
        failures=[], new_monotonicity_failures=[], retained_monotonicity_failures=[])
    for group in protocol['groups']:
        out = dict(name=group['name'], axis=group['axis'], rows=[])
        for point in group['points']:
            a,x = [float.fromhex(point[k]) for k in ('a_hex','x_hex')]
            row = dict(point, route='scaled' if a<=128 and x<=1024 else 'legacy')
            try:
                refs = reference(a,x)
                row['reference'] = [mp.nstr(r,330) for r in refs]
                row['reference_hex'] = [float(r).hex() for r in refs]
                for name,fn in [('baseline',seam_diagnosis.public),('proposed',proposed)]:
                    try:
                        values = fn(a,x)
                        row[name] = [v.hex() for v in values]
                        errors,bad = score(values,refs)
                        row[name+'_relative_errors'] = errors
                        row[name+'_accuracy_failures'] = bad
                        if name=='proposed' and bad:
                            report['failures'].append(dict(group=out['name'],point=point,reason='accuracy',tails=bad))
                    except (ValueError,RuntimeError,OverflowError,ZeroDivisionError) as exc:
                        row[name+'_error'] = str(exc)
                        if name=='proposed':
                            report['failures'].append(dict(group=out['name'],point=point,reason=str(exc)))
            except (ValueError,RuntimeError,OverflowError,ZeroDivisionError) as exc:
                row['oracle_error'] = str(exc)
                report['failures'].append(dict(group=out['name'],point=point,reason='oracle: '+str(exc)))
            out['rows'].append(row)
        for i,(left,right) in enumerate(zip(out['rows'],out['rows'][1:])):
            for tail in (0,1):
                increasing = (out['axis']=='x') == (tail==0)
                def monotone(name):
                    if name not in left or name not in right:
                        return None
                    l,r = float.fromhex(left[name][tail]),float.fromhex(right[name][tail])
                    return l<=r if increasing else l>=r
                if monotone('proposed') is False:
                    key = 'new_monotonicity_failures' if monotone('baseline') is True else 'retained_monotonicity_failures'
                    report[key].append(dict(group=out['name'],interval=i,tail=tail,baseline_monotone=monotone('baseline')))
        report['groups'].append(out)
        print('evaluated',out['name'],flush=True)
    check_identity(protocol)
    rejected = any(report[k] for k in ('failures','new_monotonicity_failures','retained_monotonicity_failures'))
    report['status'] = 'REJECT' if rejected else 'PREFLIGHT PASS; NOT PUBLIC ACCEPTANCE'
    (HERE/'result.json').write_text(json.dumps(report,indent=2)+'\n')
    print(report['status'],{k:len(report[k]) for k in ('failures','new_monotonicity_failures','retained_monotonicity_failures')})
    return int(rejected)


if __name__ == '__main__':
    raise SystemExit(run())
