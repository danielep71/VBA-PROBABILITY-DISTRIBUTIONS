"""B3 development: smaller-tail routing on revealed fitting shapes only.

Radius 1 is retained alongside the wider radius-16 rejection. No holdout input
is read and no production code is modified.
"""
import csv
import hashlib
import json
import math
import struct
from pathlib import Path
import preflight_phase_b as old
HERE = Path(__file__).resolve().parent

def route(a,x):
 r,v,_=old.rm.route(a,x)
 if a>10000 or v is None or not 0<=v<=.5:return old.public(a,x)
 return (v,old.rm.sub(1.,v)) if r=='series' else (old.rm.sub(1.,v),v)
def bits(x):return struct.unpack('>Q',struct.pack('>d',x))[0]
def value(n):return struct.unpack('>d',struct.pack('>Q',n))[0]

def evaluate(radius):
    fail=[];rows=[]
    shapes=sorted({float(r['shape']) for r in csv.DictReader((HERE/'b64_fit_grid.csv').open()) if 0<float(r['shape'])<=10000})
    for a in shapes:
     centers=[('seam',a+1)]
     lo=1;hi=bits(math.nextafter(a+1,0))
     def above(n):
      _,p,_=old.rm.route(a,value(n));return p is not None and p>.5
     if not above(lo) and above(hi):
      while hi-lo>1:
       mid=(lo+hi)//2
       if above(mid):hi=mid
       else:lo=mid
      centers.append(('half',value(hi)))
     for label,c in centers:
      xs=[value(n) for n in range(max(1,bits(c)-radius),bits(c)+radius+1)]
      vs=[route(a,x) for x in xs]
      rows.append(dict(a=a,label=label,x=[x.hex() for x in xs],v=[[v.hex() for v in pair] for pair in vs]))
      for i in range(len(xs)-1):
       if vs[i+1][0]<vs[i][0] or vs[i+1][1]>vs[i][1]:fail.append(dict(a=a,label=label,left=xs[i].hex(),right=xs[i+1].hex(),left_v=[v.hex() for v in vs[i]],right_v=[v.hex() for v in vs[i+1]],baseline_left=[v.hex() for v in old.public(a,xs[i])],baseline_right=[v.hex() for v in old.public(a,xs[i+1])]))
    return dict(status='REJECT' if fail else 'DEVELOPMENT PASS ONLY', failures=fail, triples=len(rows), evaluations=sum(len(row['x']) for row in rows), rows_sha256=hashlib.sha256(json.dumps(rows,sort_keys=True).encode()).hexdigest(), radius=radius)

def result():
    return dict(kind='development only; no holdout',candidate='B3: fast only when the directly computed tail is <=0.5; otherwise existing DD; A2=10000', narrow=evaluate(1),wide=evaluate(16),sha256={p:hashlib.sha256((HERE/p).read_bytes()).hexdigest() for p in ('smaller_tail_development.py','preflight_phase_b.py','route_model.py','b64_fit_grid.csv')})

if __name__=='__main__':
    out=result()
    (HERE/'smaller_tail_development.json').write_text(json.dumps(out,indent=2)+'\n')
    print(out['wide']['status'],len(out['wide']['failures']))
    raise SystemExit(out['wide']['status']=='REJECT')
