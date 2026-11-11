"""Generate disposable, reversible DD instrumentation; never edit production VBA."""
import hashlib
import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / 'artifacts/kernel-profile'
GROUPS = {
    'coarse': ('WG_Pair', 'WG_Upper', 'WG_Deviance', 'WG_Stirling',
               'EG_Pair', 'EG1_Pair', 'EG_LogGamma', 'EG_LogPair', 'EG_LowerPair',
               'SG_LogGamma', 'SG_Log', 'SG_Log1p', 'SG_Exp', 'SG_Expm1'),
    'arithmetic': ('SG_Op', 'SG_Coerce', 'SGD_Try', 'SGD_Valid',
                   'SGD_Canonicalize', 'SP_Frexp', 'SP_Shift',
                   'PV2_Add', 'PV2_Mul', 'PV2_Div', 'PV2_Sum'),
}
PROC = re.compile(r'^Private (Function|Sub) (PROB_DD_\w+)\b.*?^End \1$', re.M | re.S)


def instrument(source, names):
    """Insert entry/exit statements only; fail closed on an unexpected layout."""
    seen = []

    def replace(match):
        kind, name = match.group(1, 2)
        if name not in names:
            return match.group(0)
        seen.append(name)
        index = names.index(name) + 1
        lines = match.group(0).splitlines()
        # Find the first executable line after the multiline signature, comments
        # and declarations. Target procedures have no continued Dim statements.
        pos = 0
        while lines[pos].endswith(' _'):
            pos += 1
        pos += 1
        while (not lines[pos].strip() or lines[pos].lstrip().startswith("'")
               or lines[pos].lstrip().startswith('Dim ')):
            assert not lines[pos].endswith(' _'), name
            pos += 1
        for i in range(pos, len(lines)):
            if not lines[i].lstrip().startswith("'"):
                lines[i] = lines[i].replace('Exit ' + kind,
                    f'KP_Leave {index}: Exit {kind}')
        lines.insert(pos, f'    KP_Enter {index}')
        lines.insert(len(lines) - 1, f'    KP_Leave {index}')
        return '\n'.join(lines)

    result = PROC.sub(replace, source)
    assert set(seen) == set(names), (seen, names)
    # Removing our exact statements must recover every source byte (LF form).
    restored = re.sub(r'^    KP_(?:Enter|Leave) \d+\n', '', result, flags=re.M)
    restored = re.sub(r'KP_Leave \d+: ', '', restored)
    assert restored == source
    return result


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    path = ROOT / 'src/M_STATS_PROBDIST_SPECIALFUNCS.bas'
    source = path.read_text(encoding='utf-8')
    inventory = {}
    for group, suffixes in GROUPS.items():
        names = ['PROB_DD_' + s for s in suffixes]
        result = instrument(source, names)
        # CodeModule.AddFromString accepts code, not exported Attribute lines.
        result = re.sub(r'^Attribute .*\n', '', result, flags=re.M)
        (OUT / (group + '.txt')).write_text(result, encoding='utf-8')
        inventory[group] = dict(enumerate(names, 1))
    inventory['source_sha256_lf'] = hashlib.sha256(source.encode()).hexdigest()
    (OUT / 'instrumentation.json').write_text(json.dumps(inventory, indent=2) + '\n')
    print('PASS: two reversible instrumentation passes; production files unchanged')


if __name__ == '__main__':
    main()
