"""Freeze the B5 scalar representation experiment before any Excel measurement."""
import hashlib
import json
import subprocess
from pathlib import Path

from prepare_pair_host import SOURCE_SHA256, inputs

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent.parent
OUT = ROOT / 'artifacts/gamma-scalar'
FILES = (
    'pair_candidate.py', 'pair_candidate_vba.txt', 'prepare_pair_host.py',
    'scalar_candidate_vba.txt', 'scalar_probe_vba.txt', 'prepare_scalar_host.py',
    'Run-ScalarStudy.ps1', 'validate_scalar_host.py', 'SCALAR_CANDIDATE.md',
)


def digest(text):
    return hashlib.sha256(text.encode('utf-8')).hexdigest()


def prepare():
    source = (ROOT / 'src/M_STATS_PROBDIST_SPECIALFUNCS.bas').read_text()
    if digest(source) != SOURCE_SHA256:
        raise RuntimeError('production source drift: review before regenerating')
    rows = inputs()  # Unchanged B4 Python graph; never use B5 to generate expectations.
    assert len(rows) == 1371
    assert sum(r['expected'] == 'DECLINED' for r in rows) == 106
    candidate = (HERE / 'pair_candidate_vba.txt').read_text() + '\n' + (HERE / 'scalar_candidate_vba.txt').read_text()
    probe = (HERE / 'scalar_probe_vba.txt').read_text()
    sha = subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=ROOT, text=True).strip()
    design = dict(study='B5_scalar', baseline='B4', candidate_sha=sha,
                  source_sha256=SOURCE_SHA256, candidate_sha256=digest(candidate),
                  probe_sha256=digest(probe), cases=rows,
                  sha256={p: digest((HERE / p).read_text()) for p in FILES})
    OUT.mkdir(parents=True, exist_ok=True)
    (OUT / 'M_STATS_PROBDIST_B5.bas').write_text(probe)
    (OUT / 'probe.txt').write_text(probe)
    (OUT / 'candidate.txt').write_text(candidate)
    (OUT / 'inputs.json').write_text(json.dumps(design, indent=2) + '\n')
    print('prepared', len(rows), 'unique revealed inputs; B4 expectations unchanged')


if __name__ == '__main__':
    prepare()
