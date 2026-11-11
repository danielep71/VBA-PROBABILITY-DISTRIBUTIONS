"""Prepare exact revealed input words and predicted pairs for the Excel study."""
import csv
import hashlib
import json
import struct
from pathlib import Path

import pair_candidate as candidate
from screen_pair_candidate import bits, value

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent.parent
OUT = ROOT/'artifacts/gamma-pair'
SOURCE_SHA256 = 'f568e3a3d3e3517d5c66294d980f411e711188a2df40f2b608ace267b041fb6e'


def hex_bits(x):
    return struct.pack('>d', x).hex().upper()


def inputs():
    with (HERE/'b64_fit_grid.csv').open() as stream:
        fit = list(csv.DictReader(stream))
    retained = json.loads((HERE/'arithmetic_diagnosis.json').read_text())['rows']
    points = {(float(r['shape']), float(r['x'])) for r in fit}
    centers = {(a, a+1) for a, _ in points if 1e-8 <= a <= 10000}
    for r in retained:
        a = float.fromhex(r['a'])
        points.add((a, float.fromhex(r['left'])))
        points.add((a, float.fromhex(r['right'])))
        centers.add((a, float.fromhex(r['left'])))
    for a, x in centers:
        points.update((a, value(n)) for n in range(bits(x)-16, bits(x)+17))
    points.update(((.0001,.5),(.5,.25),(2.5,1.),(2.5,6.),(100.,90.),(100.,110.),(10000.,9900.),(10000.,10100.)))
    out = []
    for a, x in sorted(points):
        try:
            p, q = candidate.public(a, x)
            expected = hex_bits(p)+','+hex_bits(q)
        except candidate.Declined:
            expected = 'DECLINED'
        out.append(dict(id=hex_bits(a)+':'+hex_bits(x), a=a, x=x, expected=expected))
    return out


if __name__ == '__main__':
    source = (ROOT/'src/M_STATS_PROBDIST_SPECIALFUNCS.bas').read_text()
    if hashlib.sha256(source.encode()).hexdigest() != SOURCE_SHA256:
        raise RuntimeError('production source drift: review before regenerating')
    OUT.mkdir(parents=True, exist_ok=True)
    (OUT/'M_STATS_PROBDIST_B4.bas').write_text((HERE/'pair_probe_vba.txt').read_text())
    (OUT/'candidate.txt').write_text((HERE/'pair_candidate_vba.txt').read_text())
    rows = inputs()
    (OUT/'inputs.json').write_text(json.dumps(dict(source_sha256=SOURCE_SHA256, cases=rows,
        sha256={p:hashlib.sha256((HERE/p).read_bytes()).hexdigest() for p in
            ('pair_candidate.py','pair_candidate_vba.txt','pair_probe_vba.txt','prepare_pair_host.py')}),indent=2)+'\n')
    print('prepared',len(rows),'unique revealed inputs')
