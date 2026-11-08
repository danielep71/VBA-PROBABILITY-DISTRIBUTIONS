"""Evaluate the separately frozen public candidate cohort without tuning."""
import hashlib
import json
import math
from pathlib import Path
import sys
import mpmath as mp
from scipy.special import gammaincinv
from wide_gamma import public
from reference_binary64 import rounded as reference_float
from screen_wide_complete_v2 import direct_lower,direct_upper

HERE=Path(__file__).resolve().parent
ROOT=HERE.parents[2]


def digest(path):return hashlib.sha256(path.read_bytes().replace(b'\r\n',b'\n')).hexdigest()


def identity(protocol):
    for path,sha in protocol['sha256_lf'].items():
        if digest(ROOT/path)!=sha:raise ValueError('frozen source changed: '+path)


def inverse(a,p,q,dps):
    with mp.workdps(dps):
        aa=mp.mpf(a);target=mp.mpf(p if p<=q else q)
        seed=float(gammaincinv(a,p))
        if not seed>0 or not math.isfinite(seed):raise ValueError('invalid inverse seed')
        def objective(z):
            x=mp.exp(z)
            return (direct_lower(aa,x)[0] if p<=q else direct_upper(aa,x)[0])-target
        def derivative(z):return (1 if p<=q else -1)*mp.exp(aa*z-mp.exp(z)-mp.loggamma(aa))
        z=mp.findroot(objective,mp.log(mp.mpf(seed)),df=derivative,solver='newton',tol=mp.eps*32,maxsteps=30,verify=True)
        value=mp.exp(z)
        residual=abs(objective(z)/target)
        if residual>mp.sqrt(mp.eps):raise ValueError('inverse oracle residual')
        return value


def main():
    protocol=json.loads((HERE/'public_fresh_protocol.json').read_text());identity(protocol)
    out=dict(status='INCOMPLETE',candidate_commit=protocol['candidate_commit'],protocol_sha256=digest(HERE/'public_fresh_protocol.json'),original_screen_sha256=digest(HERE/'public_fresh_screen.json'),conversion_sha256=digest(HERE/'reference_binary64.py'),completion_sha256=digest(Path(__file__)),rows=[],inverses=[],failures=[],monotonicity_failures=[])
    previous={}
    for i,point in enumerate(protocol['points']):
        a,x=[float.fromhex(point[k]) for k in ('a_hex','x_hex')];row=dict(point,index=i)
        try:
            values=public(a,x);row['values_hex']=[v.hex() for v in values];refs=[];row['oracle_methods']=[]
            for dps in protocol['oracle_dps']:
                with mp.workdps(dps):
                    p,pm=direct_lower(mp.mpf(a),mp.mpf(x));q,qm=direct_upper(mp.mpf(a),mp.mpf(x));refs.append((p,q));row['oracle_methods'].append(dict(dps=dps,P=pm,Q=qm))
            if [reference_float(v).hex() for v in refs[0]]!=[reference_float(v).hex() for v in refs[1]]:raise ValueError('unstable rounded reference')
            row['reference']=[mp.nstr(v,330) for v in refs[1]];row['reference_hex']=[reference_float(v).hex() for v in refs[1]];row['relative_errors']=[]
            with mp.workdps(320):
                for j,(v,ref) in enumerate(zip(values,refs[1])):
                    if not math.isfinite(v) or not 0<=v<=1 or not mp.isfinite(ref) or ref<=0:raise ValueError('invalid result/reference')
                    err=abs(mp.mpf(v)/ref-1);row['relative_errors'].append(mp.nstr(err,35))
                    bad=v.hex()!=reference_float(ref).hex() if reference_float(ref)<sys.float_info.min else err>mp.mpf(('3e-15','5e-15')[j])
                    if bad:out['failures'].append(dict(index=i,tail=j,reason='accuracy'))
            if a in previous:
                oldx,old=previous[a]
                if oldx>=x:raise ValueError('unordered cohort')
                for j in (0,1):
                    if (old[j]>values[j] if j==0 else old[j]<values[j]):out['monotonicity_failures'].append(dict(index=i,tail=j))
            previous[a]=(x,values)
        except Exception as e:row['error']=str(e);out['failures'].append(dict(index=i,reason=str(e)));previous.pop(a,None)
        out['rows'].append(row);print('forward',i+1,'/',len(protocol['points']),flush=True)
    for i,point in enumerate(protocol['inverses']):
        row=dict(point,index=i)
        try:
            a,p,q=[float.fromhex(point[k]) for k in ('a_hex','p_hex','q_hex')]
            refs=[inverse(a,p,q,dps) for dps in protocol['oracle_dps']]
            if reference_float(refs[0]).hex()!=reference_float(refs[1]).hex():raise ValueError('unstable inverse reference')
            row['reference']=mp.nstr(refs[1],330);row['reference_hex']=reference_float(refs[1]).hex()
        except Exception as e:row['error']=str(e);out['failures'].append(dict(inverse_index=i,reason=str(e)))
        out['inverses'].append(row);print('inverse references',i+1,'/',len(protocol['inverses']),flush=True)
    identity(protocol)
    rejected=bool(out['failures'] or out['monotonicity_failures'])
    out['status']='REJECT' if rejected else 'FORWARD MODEL PASS; INVERSE REFERENCES READY; VBA PUBLIC ACCEPTANCE PENDING'
    (HERE/'public_fresh_complete_screen.json').write_text(json.dumps(out,indent=2)+'\n');print(out['status'])
    return int(rejected)


if __name__=='__main__':raise SystemExit(main())
