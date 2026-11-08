"""Install the separately frozen wide Gamma model as a private VBA study."""
import argparse
from functools import lru_cache
import json
import math
from install_scaled_fidelity import HERE, ROOT, documented, literal

BEGIN="' BEGIN generated wide-Gamma study"
END="' END generated wide-Gamma study"
BASELINE=9962


def strip(text):
    from public_integration import restore
    text=restore(text)
    if BEGIN not in text:return text
    before,rest=text.split(BEGIN,1)
    _,after=rest.split(END+'\n\n',1)
    return (before+after).replace('    Test_WideGamma_Fidelity\n','')


@lru_cache(maxsize=1)
def dataset():
    seen={}
    def add(row,kind):
        key=(row['a_hex'],row['x_hex'])
        if key in seen:
            assert seen[key]['values_hex']==row['values_hex']
            seen[key]['kinds'].append(kind)
        else:seen[key]=dict(a=key[0],x=key[1],values_hex=row['values_hex'],kinds=[kind])
    screen=json.loads((HERE/'wide_complete_v2_screen.json').read_text())
    assert screen['status']=='BOUNDED MODEL PASS; NOT PUBLIC ACCEPTANCE'
    for row in screen['rows']:add(row,'newly revealed wide screen')
    replay=json.loads((HERE/'wide_outer_replay.json').read_text())
    assert replay['status']=='PASS'
    for group in replay['groups']:
        for row in group['rows']:add(row,'revealed outer boundary')
    for a in (1.,129.):
        add(dict(a_hex=a.hex(),x_hex=(0.).hex(),values_hex=[(0.).hex(),(1.).hex()]),'endpoint')
    invalid=[(0.,1.),(-1.,1.),(1.,-1.),(math.inf,1.),(1.,math.inf),(math.nan,1.),(1.,math.nan)]
    rows=list(seen.values())
    return dict(schema=1,scope='private Python/VBA fidelity; no public acceptance',
                rows=rows,invalid=[[a.hex(),x.hex()] for a,x in invalid],baseline_assertions=BASELINE,
                additional_assertions=3*len(rows)+len(invalid)+4+5)


def render(data):
    lines=[(HERE/'wide_gamma_vba.template').read_text(),''];names=[]
    for start in range(0,len(data['rows']),20):
        name=f'Test_WideGamma_Chunk{start//20}';names.append(name)
        lines.append(f'Private Sub {name}()')
        for i,r in enumerate(data['rows'][start:start+20],start):
            lines += [f'    WG_Case "WG {i:03d}", {literal(r["a"])}, {literal(r["x"])}, _',
                      f'        {literal(r["values_hex"][0])}, {literal(r["values_hex"][1])}']
        lines+=['End Sub','']
    lines+=['Private Sub Test_WideGamma_Fidelity()']+['    '+n for n in names]
    for a,x in data['invalid']:lines.append(f'    WG_Invalid {literal(a)}, {literal(x)}')
    lines+=['    WG_RoutingControl','    WG_CapControls','End Sub','']
    return BEGIN+'\n'+documented('\n'.join(lines))+'\n'+END+'\n\n'


def install():
    data=dataset();target=ROOT/'tests/M_STATS_PROBDIST_TEST.bas'
    text=strip(target.read_text())
    text=text.replace("' BEGIN generated extended-Gamma study",render(data)+"' BEGIN generated extended-Gamma study",1)
    text=text.replace('    Test_ExtendedGamma_Fidelity\n','    Test_ExtendedGamma_Fidelity\n    Test_WideGamma_Fidelity\n',1)
    policy=ROOT/'.github/excel-evidence-policy.json';current=json.loads(policy.read_text())
    total=BASELINE+data['additional_assertions']
    assert current['expected_assertions'] in (BASELINE,total)
    current['expected_assertions']=total
    target.write_bytes(text.replace('\n','\r\n').encode('ascii'))
    policy.write_text(json.dumps(current,indent=2)+'\n')
    (HERE/'wide_gamma_fidelity.json').write_text(json.dumps(data,indent=2)+'\n')
    print(len(data['rows']),'pairs;',data['additional_assertions'],'new assertions;',total,'total')


if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--write',action='store_true',required=True)
    parser.parse_args();install()
