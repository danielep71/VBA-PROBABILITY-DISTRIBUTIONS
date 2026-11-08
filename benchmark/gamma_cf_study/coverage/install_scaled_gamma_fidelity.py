"""Generate private Gamma fidelity checks from revealed development points."""
import argparse
from functools import lru_cache
import json
import math
from install_scaled_fidelity import HERE, ROOT, documented, literal
from precision_probe import LN2
from precision_v2 import COEFFICIENTS
from scaled_dd import ScaledDD as S
from scaled_gamma import public

BEGIN = "' BEGIN generated scaled-Gamma study"
END = "' END generated scaled-Gamma study"
BASELINE = 7715


def scaled_literal(value):
    return f'Array({literal(value.hi.hex())}, {literal(value.lo.hex())}, {value.exponent})'


@lru_cache(maxsize=1)
def dataset():
    pairs = []
    for r in json.loads((HERE/'references.json').read_text()):
        a, x = float.fromhex(r['a_hex']), float.fromhex(r['x_hex'])
        if 0 < a <= 1 and 0 < x <= 32:
            pairs.append((a, x))
    for a in [2.**-1074, 2.**-1022, 2.**-400, .01, .5, 1.]:
        for x in [2.**-1074, 2.**-1022, .25, 1., a+1, 32.]:
            pairs.append((a, x))
    seams = []
    for a in [2.**-1074, 1e-100, 1e-6, .0001, .01, .25, .5, 1.]:
        c = a+1
        xs = [math.nextafter(c, 0.), c, math.nextafter(c, math.inf)]
        seams.append([a.hex()]+[x.hex() for x in xs])
        pairs.extend((a,x) for x in xs)
    rows = []
    for a,x in pairs:
        p,q = public(a,x)
        rows.append(dict(a=a.hex(), x=x.hex(), p=p.hex(), q=q.hex()))
    invalid = [[float(a).hex(), float(x).hex()] for a,x in
               [(0,1),(-1,1),(2,1),(1,0),(1,33),(math.inf,1),(1,math.nan),(math.nan,1)]]
    return dict(schema=1, scope='revealed model fidelity, not independent Gamma acceptance',
                rows=rows, seams=seams, invalid=invalid, baseline_assertions=BASELINE,
                additional_assertions=len(rows)*3+len(invalid)+len(seams)*4)


def render(data):
    template = (HERE/'scaled_gamma_vba.template').read_text()
    template = template.replace('@LN2@', scaled_literal(S._normalize(LN2, 0)))
    template = template.replace('@LN2HI@', literal(LN2.hi.hex()))
    coefficients = []
    for i,c in enumerate(reversed(COEFFICIENTS)):
        value = scaled_literal(S._normalize(c, 0))
        coefficients.append('    V = '+(value if i == 0 else f'SG_Add(SG_Mul(V, U), {value})'))
    template = template.replace('@COEFFICIENTS@', '\n'.join(coefficients))
    lines = [template, '']
    names = []
    for start in range(0, len(data['rows']), 20):
        name = f'Test_ScaledGamma_Chunk{start//20}'
        names.append(name)
        lines.append(f'Private Sub {name}()')
        for i,r in enumerate(data['rows'][start:start+20], start):
            lines += [f'    SG_Case "SG {i:03d}", {literal(r["a"])}, {literal(r["x"])}, _',
                      f'        {literal(r["p"])}, {literal(r["q"])}']
        lines += ['End Sub', '']
    lines += ['Private Sub Test_ScaledGamma_Fidelity()']
    lines += ['    '+name for name in names]
    for a,x in data['invalid']:
        lines.append(f'    SG_Invalid {literal(a)}, {literal(x)}')
    for a,l,c,r in data['seams']:
        lines += [f'    SG_Seam {literal(a)}, {literal(l)}, _',
                  f'        {literal(c)}, {literal(r)}']
    lines += ['End Sub', '']
    return BEGIN+'\n'+documented('\n'.join(lines))+'\n'+END+'\n\n'


def install(data):
    target = ROOT/'tests/M_STATS_PROBDIST_TEST.bas'
    text = target.read_text()
    if BEGIN in text:
        before,rest = text.split(BEGIN,1)
        _,after = rest.split(END+'\n\n',1)
        text = before+render(data)+after
    else:
        text = text.replace('Private Function PV2_H(',render(data)+'Private Function PV2_H(',1)
        text = text.replace('    Test_ScaledDD_Fidelity\n',
                            '    Test_ScaledDD_Fidelity\n    Test_ScaledGamma_Fidelity\n',1)
    policy = ROOT/'.github/excel-evidence-policy.json'
    current = json.loads(policy.read_text())
    total = BASELINE+data['additional_assertions']
    assert current['expected_assertions'] in (BASELINE,total)
    current['expected_assertions'] = total
    target.write_bytes(text.replace('\n','\r\n').encode('ascii'))
    policy.write_text(json.dumps(current,indent=2)+'\n')
    (HERE/'scaled_gamma_fidelity.json').write_text(json.dumps(data,indent=2)+'\n')
    print(f'{len(data["rows"])} pairs; {data["additional_assertions"]} new assertions; {total} total')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--write',action='store_true',required=True)
    parser.parse_args()
    install(dataset())
