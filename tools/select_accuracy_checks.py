#!/usr/bin/env python3
"""Skip expensive proofs only for a positively identified metadata-only change."""
from __future__ import annotations

import argparse
import json
import re
import subprocess
from pathlib import Path

# These files are not inputs to either numerical proof suite. Do not broaden
# this to all Markdown, all workflows, or all tools: those include evidence.
SAFE_FILES = {
    '.github/labels.json', '.github/dependabot.yml',
    '.github/PULL_REQUEST_TEMPLATE.md',
    '.github/workflows/labels-sync.yml', '.github/workflows/labels-drift.yml',
    '.github/workflows/issue-triage.yml',
    '.github/scripts/labels-sync.mjs', '.github/scripts/labels-drift.mjs',
    'tools/check_issues.py', 'tests/test_check_issues.py',
    'docs/repo_governance/ISSUE_LABELS.md',
}


def metadata_only(paths: list[str]) -> bool:
    return bool(paths) and all(
        p in SAFE_FILES or (p.startswith('.github/ISSUE_TEMPLATE/')
                            and p.endswith(('.yml', '.yaml')))
        for p in paths
    )


def git(*args: str) -> str:
    return subprocess.check_output(['git', *args], text=True, stderr=subprocess.PIPE)


def select(event: str, base: str, head: str = 'HEAD') -> tuple[bool, str]:
    if event not in ('pull_request', 'push'):
        return True, 'Manual, scheduled or unknown event: full proofs'
    if not re.fullmatch(r'[0-9a-fA-F]{40}', base) or set(base) == {'0'}:
        return True, 'Missing comparison base: full proofs'
    try:
        git('cat-file', '-e', base + '^{commit}')
        git('merge-base', '--is-ancestor', base, head)
        # The checked-out PR head is the synthetic merge. Comparing its base
        # to that commit covers the complete PR, including base integration.
        # --no-renames exposes both paths of a rename, including deletions.
        paths = git('diff', '--no-renames', '--name-only', '-z', base, head).split('\0')
        paths = [p for p in paths if p]
        # Also inspect intermediate commits: an earlier relevant edit followed
        # by a restoration must not hide behind an innocuous final commit.
        history = git('log', '--format=', '--name-only', '--no-renames', '-z',
                      base + '..' + head)
        paths += [p.strip('\n') for p in history.split('\0') if p.strip('\n')]
    except (OSError, subprocess.CalledProcessError):
        return True, 'Comparison unavailable or non-ancestral: full proofs'
    if metadata_only(paths):
        return False, 'Only recognized issue/label metadata changed'
    return True, 'Relevant, unknown or empty change set: full proofs'


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--event', required=True)
    parser.add_argument('--base', default='')
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--summary', type=Path, required=True)
    args = parser.parse_args()
    full, reason = select(args.event, args.base)
    with args.output.open('a') as handle:
        handle.write(f'full={str(full).lower()}\n')
    with args.summary.open('a') as handle:
        handle.write(f'## Accuracy check selection\n\nFull regression proofs: **{full}**\n\n{reason}. '
                     'Live evidence checks always run.\n')
    print(json.dumps({'full': full, 'reason': reason}))


if __name__ == '__main__':
    main()
