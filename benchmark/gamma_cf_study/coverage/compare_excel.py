"""Compare retained Excel failure excerpts without treating old failures as PASS."""
import argparse
import json
from pathlib import Path
import re

HERE = Path(__file__).resolve().parent


def read(name):
    text = (HERE / name).read_text()
    match = re.search(r'TOTAL=(\d+);PASS=(\d+);FAIL=(\d+)', text)
    if not match:
        raise ValueError('missing Excel summary')
    total, passed, failed = map(int, match.groups())
    failures = {}
    for line in text.splitlines():
        if ']   ' not in line:
            continue
        item = line.split(']   ', 1)[1]
        parts = item.split(' -> got ', 1)
        label, detail = parts[0], parts[1] if len(parts) == 2 else None
        if label in failures:
            raise ValueError('duplicate assertion label')
        failures[label] = detail
    if passed + failed != total or len(failures) != failed:
        raise ValueError('incomplete failure excerpt')
    return {'total': total, 'passed': passed, 'failed': failed,
            'candidate': re.search(r'Candidate SHA: ([0-9a-f]{40})', text)[1],
            'failures': failures}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--baseline', default='baseline_excel_run_42467939162.txt')
    parser.add_argument('--candidate', default='b6_excel_run_42467893663.txt')
    parser.add_argument('--output', default='paired_excel_result.json')
    args = parser.parse_args()
    baseline = read(args.baseline)
    candidate = read(args.candidate)
    if baseline['total'] != candidate['total']:
        raise ValueError('unequal assertion totals')
    old, new = baseline['failures'], candidate['failures']
    report = {
        'status': 'BLOCKED: residual contract failures remain',
        'baseline': baseline, 'b6': candidate,
        'new_failure_labels': sorted(set(new) - set(old)),
        'resolved_failure_labels': sorted(set(old) - set(new)),
        'common_changed_display_values': sorted(k for k in new if k in old and new[k] != old[k]),
        'comparison_precision': 'Excel diagnostic text, not bitwise output comparison',
    }
    (HERE / args.output).write_text(json.dumps(report, indent=2) + '\n')
    print(f"Baseline {baseline['passed']}/{baseline['total']}; B6 {candidate['passed']}/{candidate['total']}")
    print(f"Resolved {len(report['resolved_failure_labels'])}; new {len(report['new_failure_labels'])}; residual {len(new)}")


if __name__ == '__main__':
    main()
