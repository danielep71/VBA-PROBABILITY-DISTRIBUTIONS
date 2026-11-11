"""Generate isolated recurrence candidates without changing production source."""
import hashlib
import importlib.util
import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]


def _route_block():
    """b64_route_vba.py, loaded by path: the measured module carried its block."""
    path = ROOT / 'benchmark/gamma_route_study/b64_route_vba.py'
    import sys
    sys.path.insert(0, str(path.parent))
    spec = importlib.util.spec_from_file_location('b64_route_vba', path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def measured_source():
    """The module as this study measured it (BASELINE): the current source with
    the binary64 route block of #312, which left production after #307 closed."""
    return _route_block().as_measured(SOURCE.read_text())

SOURCE = ROOT / 'src/M_STATS_PROBDIST_SPECIALFUNCS.bas'
BASELINE = 'f568e3a3d3e3517d5c66294d980f411e711188a2df40f2b608ace267b041fb6e'
LOWER = [
    ('    Stage = "lower geometric remainder"\n',
     '    Dim Denominator As Variant\n    Stage = "lower geometric remainder"\n'),
    ('        Term = PROB_DD_SG_Div(PROB_DD_SG_Mul(Term, XX), PROB_DD_SG_Add(AA, N))\n',
     '        Denominator = PROB_DD_SG_Add(AA, N)\n        Term = PROB_DD_SG_Div(PROB_DD_SG_Mul(Term, XX), Denominator)\n'),
    ('        Ratio = PROB_DD_SG_Div(XX, PROB_DD_SG_Add(PROB_DD_SG_Add(AA, N), 1#))\n',
     '        Ratio = PROB_DD_SG_Div(XX, PROB_DD_SG_Add(Denominator, 1#))\n'),
]
UPPER = [
    ('    Dim N As Long, K As Long, Converged As Boolean\n',
     '    Dim N As Long, K As Long, Converged As Boolean, BaseB As Variant\n'),
    ('    B = PROB_DD_SG_Sub(PROB_DD_SG_Add(XX, 1#), AA)\n',
     '    B = PROB_DD_SG_Sub(PROB_DD_SG_Add(XX, 1#), AA)\n    BaseB = B\n'),
    ('            PROB_DD_SG_Add(PROB_DD_SG_Add(PROB_DD_SG_Sub(PROB_DD_SG_Add(XX, 1#), AA), 2# * K), Tail))\n',
     '            PROB_DD_SG_Add(PROB_DD_SG_Add(BaseB, 2# * K), Tail))\n'),
    ('    H = PROB_DD_SG_Div(1#, PROB_DD_SG_Add(PROB_DD_SG_Sub(PROB_DD_SG_Add(XX, 1#), AA), Tail))\n',
     '    H = PROB_DD_SG_Div(1#, PROB_DD_SG_Add(BaseB, Tail))\n'),
]


def transform(source, mode, reverse=False):
    if mode not in ('baseline', 'lower', 'upper', 'combined'):
        raise ValueError('unknown mode')
    groups = []
    if mode in ('lower', 'combined'):
        groups.append(('Function', 'PROB_DD_WG_Pair', LOWER))
    if mode in ('upper', 'combined'):
        groups.append(('Sub', 'PROB_DD_WG_Upper', UPPER))
    for kind, name, pairs in groups:
        match = re.search(r'^Private ' + kind + ' ' + name + r'\(.*?^End ' + kind + r'\n', source, re.M | re.S)
        if match is None:
            raise ValueError('missing procedure ' + name)
        original = match.group(0)
        code = original
        if reverse:
            pairs = [(new, old) for old, new in reversed(pairs)]
        for old, new in pairs:
            if code.count(old) != 1:
                raise ValueError('replacement site drift: ' + old)
            code = code.replace(old, new, 1)
        source = source.replace(original, code, 1)
    return source


def digest(code):
    return hashlib.sha256(code.encode()).hexdigest()


def main():
    source = measured_source()
    if digest(source) != BASELINE:
        raise ValueError('baseline source drift')
    out = ROOT / 'artifacts/gamma-recurrence'
    out.mkdir(parents=True, exist_ok=True)
    metadata = {'source_sha256_lf': digest(source), 'generated_source_sha256': {}}
    for mode in ('baseline', 'lower', 'upper', 'combined'):
        code = transform(source, mode)
        if transform(code, mode, reverse=True) != source:
            raise ValueError('reversal failed')
        code = re.sub(r'^Attribute .*\n', '', code, flags=re.M)
        (out / (mode + '.txt')).write_bytes(code.encode())
        metadata['generated_source_sha256'][mode] = digest(code)
    (out / 'inputs.json').write_text(json.dumps(metadata, indent=2) + '\n')
    print('PASS: four candidates generated; all reverse exactly to pinned production')


if __name__ == '__main__':
    main()
