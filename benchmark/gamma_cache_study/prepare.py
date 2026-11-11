"""Generate uncached reference and test-only cache diagnostics for Excel."""
import hashlib
import json
import re
from pathlib import Path
import cache_vba as C

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / 'artifacts/gamma-cache'


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    source = C.SRC.read_text()
    original = C.strip(source)
    assert C.apply(original) == source
    prior = json.loads((ROOT / 'benchmark/harness_timing/run_43017348872_profile.json').read_text())
    assert hashlib.sha256(original.encode()).hexdigest() == prior['instrumentation']['source_sha256_lf']
    reference = original + '\n' + C.procedure(source, C.SLOW)
    # Reference contains the same slow function under both names so the private
    # diagnostic macro can compile in either pass. Original public calls stay
    # on the original body, without a cache lookup.
    for name, code in (('uncached', reference), ('cached', source)):
        code = re.sub(r'^Attribute .*\n', '', code, flags=re.M)
        (OUT / (name + '.txt')).write_text(code + '\n' + Path(__file__).with_name('private_checks.txt').read_text())
    metadata = {'source_sha256_lf': hashlib.sha256(source.encode()).hexdigest(),
                'uncached_source_sha256_lf': hashlib.sha256(original.encode()).hexdigest()}
    (OUT / 'inputs.json').write_text(json.dumps(metadata, indent=2) + '\n')
    print('PASS: cache reverses exactly; reference preserves the original arithmetic')


if __name__ == '__main__':
    main()
