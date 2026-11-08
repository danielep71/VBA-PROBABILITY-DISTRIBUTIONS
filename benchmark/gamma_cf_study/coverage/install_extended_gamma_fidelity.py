"""Generate the separate private extended-Gamma port and host fidelity cases."""
import argparse
from functools import lru_cache
import json
import math
import random
from install_scaled_fidelity import HERE, ROOT, documented, literal
from extended_gamma_v2 import public

BEGIN="' BEGIN generated extended-Gamma study"
END="' END generated extended-Gamma study"
BASELINE=8190


def strip(text):
    if BEGIN not in text:
        return text
    before,rest=text.split(BEGIN,1)
    _,after=rest.split(END+'\n\n',1)
    return (before+after).replace('    Test_ExtendedGamma_Fidelity\n','')


@lru_cache(maxsize=1)
def dataset():
    seen={}
    def add(a,x,kind):
        key=(a,x)
        seen.setdefault(key,[]).append(kind)
    for r in json.loads((HERE/'references.json').read_text()):
        add(r['a_hex'],r['x_hex'],'revealed B6')
    for r in json.loads((HERE/'scaled_gamma_fidelity.json').read_text())['rows']:
        add(r['a'],r['x'],'revealed bounded Gamma')
    screen=json.loads((HERE/'extended_screen.json').read_text())
    assert screen['status']=='BOUNDED MODEL PASS; NOT PUBLIC ACCEPTANCE'
    for r in screen['rows']:
        add(r['a_hex'],r['x_hex'],'newly revealed extended screen')
    for g in json.loads((HERE/'routing/result.json').read_text())['groups']:
        for r in g['rows']:
            add(r['a_hex'],r['x_hex'],'revealed bounded-routing preflight')
    for a in (1.,2.,10.,100.,128.):
        for x in (2.**-1074,2.**-1022,1e-100,a+1,1024.):
            add(a.hex(),x.hex(),'extended domain edges')
    for x in (math.nextafter(2.**-537,0.),2.**-537,math.nextafter(2.**-537,math.inf)):
        add((2.).hex(),x.hex(),'final lower subnormal seam')
    rng=random.Random(202611080326)
    for _ in range(128):
        x=math.ldexp(rng.uniform(.5,1.),rng.randrange(-537,-511))
        add((2.).hex(),x.hex(),'shape-two rational interval')
    for n in range(1,32,2):
        center=n*2.**-537
        for x in (math.nextafter(center,0.),center,math.nextafter(center,math.inf)):
            add((2.).hex(),x.hex(),'shape-two rational interval')
    rows=[]
    for (a,x),kinds in seen.items():
        p,q=public(float.fromhex(a),float.fromhex(x))
        rows.append(dict(a=a,x=x,p=p.hex(),q=q.hex(),kinds=sorted(set(kinds))))
    invalid=[(0.,1.),(-1.,1.),(math.nextafter(128.,math.inf),1.),
             (1.,0.),(1.,math.nextafter(1024.,math.inf)),(math.inf,1.),(1.,math.nan),(math.nan,1.)]
    return dict(schema=1,scope='private Python/VBA fidelity, no public integration',rows=rows,
                invalid=[[a.hex(),x.hex()] for a,x in invalid],baseline_assertions=BASELINE,
                additional_assertions=len(rows)*3+len(invalid)+3)


def render(data):
    lines=[(HERE/'extended_gamma_vba.template').read_text(),'']
    names=[]
    for start in range(0,len(data['rows']),20):
        name=f'Test_ExtendedGamma_Chunk{start//20}'
        names.append(name)
        lines.append(f'Private Sub {name}()')
        for i,r in enumerate(data['rows'][start:start+20],start):
            lines += [f'    EG_Case "EG {i:03d}", {literal(r["a"])}, {literal(r["x"])}, _',
                      f'        {literal(r["p"])}, {literal(r["q"])}']
        lines+=['End Sub','']
    lines+=['Private Sub Test_ExtendedGamma_Fidelity()']
    lines+=['    '+n for n in names]
    for a,x in data['invalid']:
        lines.append(f'    EG_Invalid {literal(a)}, {literal(x)}')
    lines+=['    EG_RoutingControl','End Sub','']
    return BEGIN+'\n'+documented('\n'.join(lines))+'\n'+END+'\n\n'


def install():
    data=dataset()
    target=ROOT/'tests/M_STATS_PROBDIST_TEST.bas'
    text=strip(target.read_text())
    text=text.replace('Private Function PV2_H(',render(data)+'Private Function PV2_H(',1)
    text=text.replace('    Test_ScaledGamma_Fidelity\n','    Test_ScaledGamma_Fidelity\n    Test_ExtendedGamma_Fidelity\n',1)
    policy=ROOT/'.github/excel-evidence-policy.json'
    current=json.loads(policy.read_text())
    total=BASELINE+data['additional_assertions']
    assert current['expected_assertions'] in (BASELINE,total)
    current['expected_assertions']=total
    target.write_bytes(text.replace('\n','\r\n').encode('ascii'))
    policy.write_text(json.dumps(current,indent=2)+'\n')
    (HERE/'extended_gamma_fidelity.json').write_text(json.dumps(data,indent=2)+'\n')
    print(len(data['rows']),'pairs;',data['additional_assertions'],'new assertions;',total,'total')


if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--write',action='store_true',required=True)
    parser.parse_args()
    install()
