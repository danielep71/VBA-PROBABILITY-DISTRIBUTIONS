"""Evaluate the frozen shape-disjoint wide Gamma model screen."""
import hashlib
import json
import math
from pathlib import Path
import subprocess
import sys

import mpmath as mp
from mpmath.libmp.libhyper import NoConvergence
from wide_gamma import public

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[2]


def digest(path):
    return hashlib.sha256(path.read_bytes().replace(b'\r\n',b'\n')).hexdigest()


def identity(protocol):
    for name, expected in protocol['sha256_lf'].items():
        if digest(ROOT/name)!=expected:
            raise ValueError('frozen candidate mismatch: '+name)


def direct_lower(a,x):
    try:
        return mp.gammainc(a,0,x,regularized=True),'mpmath.gammainc'
    except NoConvergence:
        if not (a>128 and x>a):
            raise
        sigma=mp.sqrt(a)
        end=(x-a)/sigma
        lg=mp.loggamma(a)
        def density(t):
            y=a+t*sigma
            if y<=0:return mp.mpf(0)
            return mp.exp((a-1)*mp.log(y)-y-lg)*sigma
        cuts=sorted(set([-sigma,end]+[mp.mpf(t) for t in (-20,-10,-5,0,5,10,20) if -sigma<t<end]))
        return mp.quad(density,cuts),'direct standardized-density quadrature'


def main():
    protocol=json.loads((HERE/'wide_protocol.json').read_text())
    identity(protocol)
    result=dict(status='INCOMPLETE',protocol_commit=subprocess.check_output(
        ['git','log','-1','--format=%H','--',str(HERE/'wide_protocol.json')],cwd=ROOT).decode().strip(),
        protocol_sha256=digest(HERE/'wide_protocol.json'),screen_sha256=digest(Path(__file__)),
        mpmath=mp.__version__,initial_attempt_sha256=digest(HERE/'wide_initial_attempt.txt'),rows=[],failures=[],monotonicity_failures=[])
    previous={}
    for i,point in enumerate(protocol['points']):
        a,x=[float.fromhex(point[k]) for k in ('a_hex','x_hex')]
        row=dict(point,index=i)
        try:
            trace={}
            values=public(a,x,trace)
            row['trace']=trace
            row['values_hex']=[v.hex() for v in values]
            refs=[]
            row['oracle_methods']=[]
            for precision in protocol['oracle_dps']:
                with mp.workdps(precision):
                    aa,xx=mp.mpf(a),mp.mpf(x)
                    lower,method=direct_lower(aa,xx)
                    refs.append((lower,mp.gammainc(aa,xx,mp.inf,regularized=True)))
                    row['oracle_methods'].append(dict(dps=precision,P=method,Q='mpmath.gammainc'))
            if [float(v).hex() for v in refs[0]]!=[float(v).hex() for v in refs[1]]:
                raise ValueError('unstable oracle rounding')
            row['reference']=[mp.nstr(v,330) for v in refs[1]]
            row['reference_hex']=[float(v).hex() for v in refs[1]]
            row['relative_errors']=[]
            with mp.workdps(320):
                for j,(v,ref) in enumerate(zip(values,refs[1])):
                    if not math.isfinite(v) or not 0<=v<=1 or not mp.isfinite(ref) or ref<=0:
                        raise ValueError('invalid probability/reference')
                    error=abs(mp.mpf(v)/ref-1)
                    row['relative_errors'].append(mp.nstr(error,35))
                    bad = v.hex()!=float(ref).hex() if float(ref)<sys.float_info.min else error>mp.mpf(('3e-15','5e-15')[j])
                    if bad:result['failures'].append(dict(index=i,tail=j,reason='accuracy'))
            if a in previous:
                px,pv=previous[a]
                if not px<x:raise ValueError('unordered coordinates')
                for j in (0,1):
                    if (pv[j]>values[j] if j==0 else pv[j]<values[j]):
                        result['monotonicity_failures'].append(dict(index=i,tail=j,previous_x=px.hex()))
            previous[a]=(x,values)
        except Exception as exc:
            row['error']=str(exc)
            result['failures'].append(dict(index=i,reason=str(exc)))
            previous.pop(a,None)
        result['rows'].append(row)
        print('evaluated',i+1,'/',len(protocol['points']),flush=True)
    identity(protocol)
    rejected=bool(result['failures'] or result['monotonicity_failures'])
    result['status']='REJECT' if rejected else 'BOUNDED MODEL PASS; NOT PUBLIC ACCEPTANCE'
    (HERE/'wide_complete_screen.json').write_text(json.dumps(result,indent=2)+'\n')
    print(result['status'],'failures',len(result['failures']),'monotonicity',len(result['monotonicity_failures']))
    return int(rejected)


if __name__=='__main__':
    raise SystemExit(main())
