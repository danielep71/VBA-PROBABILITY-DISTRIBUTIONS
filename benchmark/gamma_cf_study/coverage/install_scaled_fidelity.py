"""Generate a private VBA primitive port with independent exact-rational cases.

No probability oracle or new Gamma acceptance set is evaluated here.
"""
import argparse
import json
import math
import random
import re
import struct
import sys
from fractions import Fraction as F
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[2]
BEGIN = "' BEGIN generated scaled-product study"
END = "' END generated scaled-product study"
TYPES = """' BEGIN scaled-product bit fixtures (little-endian Windows Excel)
Private Type SP_Number
    Value As Double
End Type
Private Type SP_Words
    Low As Long
    High As Long
End Type
' END scaled-product bit fixtures

"""


def power(e):
    return F(2**e) if e >= 0 else F(1, 2**-e)


def reference(a, b):
    """Derive the pair from exact rationals, never from scaled_product.py."""
    value = F(a) * F(b)
    sign = math.copysign(1., a) * math.copysign(1., b)
    if not value:
        return math.copysign(0., sign), 0., 0, True, math.copysign(0., sign)
    m = abs(value)
    e = m.numerator.bit_length() - m.denominator.bit_length() + 1
    m /= power(e)
    if m < F(1, 2):
        m *= 2
        e -= 1
    h = float(m)
    if h == 1:
        m /= 2
        h = .5
        e += 1
    l = float(m - F(h))
    assert F(h) + F(l) == m
    try:
        result = float(value)
    except OverflowError:
        return sign*h, sign*l, e, False, 0.
    if result == 0:
        result = math.copysign(0., sign)
    return sign*h, sign*l, e, True, result


def dataset():
    pairs = [(2.**1000, 2.**-1000),
             (2.**-1022, math.nextafter(.5, 1.)),
             (2.**-1074, 2.**-1074),
             (sys.float_info.max, sys.float_info.max)]
    for a in [sys.float_info.max, math.nextafter(sys.float_info.max, 0.),
              2.**-1022, math.nextafter(2.**-1022, 0.)]:
        for b in [math.nextafter(1., 0.), 1., math.nextafter(1., 2.), 2.]:
            pairs.append((a, b))
            pairs.append((-a, b))
    for a in [0., -0., 2.**-1074, -2.**-1074]:
        for b in [0., -0., .25, -.25, sys.float_info.max]:
            pairs.append((a, b))
    rng = random.Random(24220261108)
    def operand():
        while True:
            v = struct.unpack('>d', rng.getrandbits(64).to_bytes(8, 'big'))[0]
            if math.isfinite(v):
                return v
    pairs.extend((operand(), operand()) for _ in range(160))
    rows = []
    for a, b in pairs:
        h, l, e, ok, result = reference(a, b)
        rows.append(dict(a=a.hex(), b=b.hex(), hi=h.hex(), lo=l.hex(),
                         exponent=e, rounded_ok=ok, rounded=result.hex()))
    rounds = []
    for h, e in [(.625, -1072), (.875, -1072), (.5, -1074)]:
        for l in [0., 2.**-56, -2.**-56]:
            for sign in [-1., 1.]:
                value = (F(sign*h)+F(sign*l))*power(e)
                result = float(value)
                if result == 0:
                    result = math.copysign(0., sign)
                rounds.append(dict(hi=(sign*h).hex(), lo=(sign*l).hex(),
                                   exponent=e, rounded=result.hex()))
    return dict(schema=1, reference='exact Fraction of binary64 inputs',
                scope='primitive regression, not a Gamma holdout',
                products=rows, rounding=rounds, baseline_assertions=2447,
                additional_assertions=len(rows)*2+len(rounds)+2098*2+11)


def literal(hex_value):
    high, low = struct.unpack('>II', struct.pack('>d', float.fromhex(hex_value)))
    return f'SP_FromWords(&H{high:08X}&, &H{low:08X}&)'


def documented(text):
    lines = []
    pending = None
    for line in text.splitlines():
        lines.append(line)
        match = re.match(r'Private (?:Function|Sub) (\w+)', line)
        if match:
            pending = match[1]
        if pending and not line.endswith('_'):
            lines += ["'==============================================================================",
                      "' PURPOSE", f"'   {pending}: private exponent-separated arithmetic fixture.",
                      "' NOTES", "'   Explicit Double stores; Windows little-endian word fixtures.",
                      "'   Study-only: no production routing or probability acceptance change.",
                      "' ERROR POLICY", "'   Try functions return False; test failures remain visible.",
                      "' UPDATED", "'   2026-11-08",
                      "'=============================================================================="]
            pending = None
    return '\n'.join(lines)


def render(data):
    template = (HERE / 'scaled_product_vba.template').read_text()
    lines = [template, '']
    chunk_names = []
    rows = data['products']
    for start in range(0, len(rows), 32):
        name = f'Test_ScaledProduct_Chunk{start//32}'
        chunk_names.append(name)
        lines.append(f'Private Sub {name}()')
        for i in range(start, min(start+32, len(rows))):
            r = rows[i]
            lines += [f'    SP_ProductCase "SP rational {i:03d}", _',
                      f'        {literal(r["a"])}, {literal(r["b"])}, _',
                      f'        {literal(r["hi"])}, {literal(r["lo"])}, {r["exponent"]}, _',
                      f'        {r["rounded_ok"]}, {literal(r["rounded"])}']
        lines += ['End Sub', '']
    lines += ['Private Sub Test_ScaledProduct_Fidelity()',
              '    Test_ScaledProduct_Exponents', '    Test_ScaledProduct_Invalid']
    lines += ['    '+name for name in chunk_names]
    for i, r in enumerate(data['rounding']):
        lines += [f'    SP_RoundCase "SP halfway {i:02d}", _',
                  f'        {literal(r["hi"])}, {literal(r["lo"])}, _',
                  f'        {r["exponent"]}, {literal(r["rounded"])}']
    lines += ['End Sub', '']
    return BEGIN+'\n'+documented('\n'.join(lines))+'\n'+END+'\n\n'


def install(data):
    target = ROOT / 'tests/M_STATS_PROBDIST_TEST.bas'
    text = target.read_text()
    block = render(data)
    if BEGIN in text:
        before, rest = text.split(BEGIN, 1)
        _, after = rest.split(END+'\n\n', 1)
        text = before+block+after
    else:
        text = text.replace('Private Function PV2_H(', block+'Private Function PV2_H(', 1)
        text = text.replace('    Test_PrecisionV2_Fidelity\n',
                            '    Test_PrecisionV2_Fidelity\n    Test_ScaledProduct_Fidelity\n', 1)
    if TYPES not in text:
        marker = 'Private Const TOL_ABS_TIGHT'
        assert marker in text
        text = text.replace(marker, TYPES+marker, 1)
    total = data['baseline_assertions']+data['additional_assertions']
    policy = ROOT / '.github/excel-evidence-policy.json'
    old = json.loads(policy.read_text())
    assert old['expected_assertions'] in (data['baseline_assertions'], total)
    old['expected_assertions'] = total
    target.write_bytes(text.replace('\n', '\r\n').encode('ascii'))
    policy.write_text(json.dumps(old, indent=2)+'\n')
    (HERE / 'scaled_product_fidelity.json').write_text(json.dumps(data, indent=2)+'\n')
    print(f'{len(data["products"])} rational pairs; {data["additional_assertions"]} new assertions; total {total}')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--write', action='store_true', required=True)
    parser.parse_args()
    install(dataset())
