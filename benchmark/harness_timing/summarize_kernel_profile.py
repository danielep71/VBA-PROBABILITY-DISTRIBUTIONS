"""Validate and summarize the retained Excel DD profiling evidence (#307)."""
import json, statistics, sys
from pathlib import Path
p = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).with_name('run_43017348872_profile.json')
d = json.loads(p.read_text(encoding='utf-8-sig'))
m = d['instrumentation']
assert d['status']=='PASS', d.get('error')
groups={}
for r in d['samples']:
    key=(r['a'],r['x'],r['route'])
    groups.setdefault(key,{}).setdefault(r['mode'],[]).append(r)
expected = {(a, x, r) for a, x in ((.0001, .5), (.5, .25), (2.5, 1), (2.5, 6),
                                     (100, 90), (100, 110), (10000, 9900), (10000, 10100))
            for r in ('p', 'q')}
expected |= {(2, .5, 'inverse'), (100, .9, 'inverse')}
assert set(groups) == expected
assert len(d['samples']) == 216
assert sum(r['calls'] + 1 for r in d['samples']) == 1008
print('Candidate',d['candidate_sha'],'Excel',d['excel_version'],d['excel_build'])
print('Empty entry/exit pair us:',d['calibration_us_per_enter_leave'])
for key,modes in groups.items():
    assert set(modes)=={'baseline','coarse','arithmetic','restored'}
    assert all(len(rows)==3 for rows in modes.values())
    assert len({r['bits'] for rows in modes.values() for r in rows})==1
    med={mode:statistics.median(r['elapsed_us']/r['calls'] for r in rows) for mode,rows in modes.items()}
    print('\n',key,'us/call', {k:round(v,2) for k,v in med.items()})
    for mode in ('coarse','arithmetic'):
        rows=modes[mode]
        ids={s['id'] for r in rows for s in r['scopes']}
        scopes=[]
        for id in ids:
            values=[(s,r) for r in rows for s in r['scopes'] if s['id']==id]
            assert len(values)==3
            calls=statistics.median(s['calls']/r['calls'] for s,r in values)
            incl=statistics.median(s['inclusive_us']/r['calls'] for s,r in values)
            excl=statistics.median(s['exclusive_us']/r['calls'] for s,r in values)
            scopes.append((excl,m[mode][str(id)],calls,incl))
        print(mode, sorted(scopes,reverse=True)[:6])
