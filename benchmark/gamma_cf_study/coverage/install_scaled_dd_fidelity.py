"""Private scaled-DD VBA generation; model fidelity is not independent acceptance."""
import argparse
import json
import math
import random
from install_scaled_fidelity import HERE, ROOT, documented, literal
from scaled_dd import ScaledDD as S

BEGIN = "' BEGIN generated scaled-DD study"
END = "' END generated scaled-DD study"
BASELINE = 7104


def calculate(op, a, b):
    if op == 1:
        return a+b
    if op == 2:
        return a-b
    if op == 3:
        return a*b
    if op == 4:
        return a/b
    raise ValueError('unknown operation')


def encode(a):
    return [a.hi.hex(), a.lo.hex(), a.exponent]


def dataset():
    pairs = []
    rng = random.Random(24232768)
    for i in range(64):
        a = S(rng.uniform(.5, 1.), rng.uniform(-.49, .49)*2.**-53, rng.randrange(-2100, 2101))
        b = S(rng.uniform(.5, 1.), rng.uniform(-.49, .49)*2.**-53,
              a.exponent+rng.randrange(-120, 121) if i % 2 else rng.randrange(-2100, 2101))
        pairs.append((a, -b if i % 3 else b))
    for e in [-2100, -1022, 0, 1024, 2100]:
        for k in [54, 106, 500, 1074]:
            for sign in [-1, 1]:
                pairs.append((S(.75, sign*2.**-k, e), S(.75, 0., e)))
        pairs.append((S(math.nextafter(1., 0.), 2.**-55, e), S(.5, -2.**-55, e+1)))
    for gap in [0, 1, 53, 54, 106, 119, 120, 121, 122, 4200]:
        for sign in [-1, 1]:
            pairs.append((S(.5, 2.**-55, 2100), S(sign*.75, sign*2.**-55, 2100-gap)))
    # Single-round alignment of subnormal low components (double-rounding traps).
    for gap in [1, 2, 53, 120]:
        pairs.append((S(.75), S(.625, 3*2.**-1074, -gap)))
    for a in [0., -0., 2., -2.]:
        for b in [0., -0., 2., -2.]:
            pairs.append((S.from_float(a), S.from_float(b)))
    rows = []
    def row(op, a, b):
        try:
            c = calculate(op, S(*a), S(*b))
            ok = True
        except (ValueError, ZeroDivisionError):
            c, ok = S(0.), False
        return dict(operation=op, a=[a[0].hex(), a[1].hex(), a[2]],
                    b=[b[0].hex(), b[1].hex(), b[2]], ok=ok, result=encode(c))
    for a, b in pairs:
        for op in range(1, 5):
            rows.append(row(op, (a.hi, a.lo, a.exponent), (b.hi, b.lo, b.exponent)))
    for bad in [(math.inf, 0., 0), (.5, math.nan, 0), (1., 0., 0),
                (.5, .1, 0), (0., 0., 1), (.5, 0., 32768)]:
        rows.append(row(1, bad, (.5, 0., 0)))
        rows.append(row(1, (.5, 0., 0), bad))
    rows.extend([row(0, (.5, 0., 0), (.5, 0., 0)),
                 row(3, (.75, 0., 32767), (.5, 0., 3)),
                 row(4, (.75, 0., -32767), (.5, 0., 3))])
    return dict(schema=1, scope='revealed arithmetic model fidelity, not a Gamma holdout',
                baseline_assertions=BASELINE, additional_assertions=len(rows), rows=rows)


def render(data):
    lines = [(HERE/'scaled_dd_vba.template').read_text(), '']
    names = []
    for start in range(0, len(data['rows']), 24):
        name = f'Test_ScaledDD_Chunk{start//24}'
        names.append(name)
        lines.append(f'Private Sub {name}()')
        for i, r in enumerate(data['rows'][start:start+24], start):
            a, b, c = r['a'], r['b'], r['result']
            lines += [f'    SD_Case "SD {i:04d}", {r["operation"]}, _',
                      f'        {literal(a[0])}, {literal(a[1])}, {a[2]}, _',
                      f'        {literal(b[0])}, {literal(b[1])}, {b[2]}, _',
                      f'        {r["ok"]}, {literal(c[0])}, {literal(c[1])}, {c[2]}']
        lines += ['End Sub', '']
    lines += ['Private Sub Test_ScaledDD_Fidelity()']
    lines += ['    '+name for name in names]
    lines += ['End Sub', '']
    return BEGIN+'\n'+documented('\n'.join(lines))+'\n'+END+'\n\n'


def install(data):
    target = ROOT/'tests/M_STATS_PROBDIST_TEST.bas'
    text = target.read_text()
    if BEGIN in text:
        before, rest = text.split(BEGIN, 1)
        _, after = rest.split(END+'\n\n', 1)
        text = before+render(data)+after
    else:
        text = text.replace('Private Function PV2_H(', render(data)+'Private Function PV2_H(', 1)
        text = text.replace('    Test_ScaledProduct_Fidelity\n',
                            '    Test_ScaledProduct_Fidelity\n    Test_ScaledDD_Fidelity\n', 1)
    policy = ROOT/'.github/excel-evidence-policy.json'
    current = json.loads(policy.read_text())
    total = BASELINE+data['additional_assertions']
    assert current['expected_assertions'] in (BASELINE, total)
    current['expected_assertions'] = total
    target.write_bytes(text.replace('\n', '\r\n').encode('ascii'))
    policy.write_text(json.dumps(current, indent=2)+'\n')
    (HERE/'scaled_dd_fidelity.json').write_text(json.dumps(data, indent=2)+'\n')
    print(f'{data["additional_assertions"]} new assertions; {total} total')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--write', action='store_true', required=True)
    parser.parse_args()
    install(dataset())
