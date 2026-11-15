#!/usr/bin/env python3
"""Select the depth of the Accuracy Gate proofs.

The full proofs run unless one of two things is positively established:

- the change touches only recognized issue and label metadata; or
- the exact tree under test was already verified by a full, successful
  Accuracy Gate run of this repository (#386). Such a run leaves an artifact
  named after the tree it checked out; see ``verified_tree``.

Every doubt selects the full proofs: an unknown event, a missing base, an API
error, an expired or foreign attestation. Live evidence checks always run.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any, Callable

WORKFLOW_PATH = '.github/workflows/accuracy-gate.yml'
ATTESTATION_PREFIX = 'accuracy-verified-'
# A run of a pull request from a fork executes the fork's copy of the workflow
# and can upload an artifact of any name, so only these events of this
# repository's own branches attest a tree.
ATTESTING_EVENTS = ('push', 'pull_request')

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


def api_get(url: str, token: str) -> Any:
    headers = {'Accept': 'application/vnd.github+json',
               'X-GitHub-Api-Version': '2022-11-28'}
    if token:
        headers['Authorization'] = 'Bearer ' + token
    request = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(request, timeout=30) as response:
        return json.load(response)


def verified_tree(tree: str, repo: str, token: str = '', api: str = 'https://api.github.com',
                  fetch: Callable[[str, str], Any] = api_get) -> tuple[bool, str]:
    """True only when a full, successful run of this workflow attested ``tree``.

    The attestation is an artifact named ``accuracy-verified-<tree>``. The final
    job of the Accuracy Gate uploads it after every job has succeeded with the
    full proofs selected, and names it after the tree the run checked out. An
    equal tree holds equal sources, tests, tools, pinned dependencies and
    workflow, so the proofs that do not read the live evidence would repeat.

    An attestation counts only if its run belongs to this repository and not
    to a fork, ran this workflow file on a push or pull-request event, and
    completed with success. Anything else, and any error, is "not verified".
    """
    if not re.fullmatch(r'[0-9a-f]{40}([0-9a-f]{24})?', tree):
        return False, 'No tree to look up'
    if not re.fullmatch(r'[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+', repo):
        return False, 'No repository to look up'
    try:
        listing = fetch(f'{api}/repos/{repo}/actions/artifacts'
                        f'?name={ATTESTATION_PREFIX}{tree}&per_page=100', token)
        for artifact in listing['artifacts']:
            if artifact.get('expired') or artifact.get('name') != ATTESTATION_PREFIX + tree:
                continue
            origin = artifact.get('workflow_run') or {}
            run_id = origin.get('id')
            if not isinstance(run_id, int):
                continue
            if origin.get('repository_id') != origin.get('head_repository_id'):
                continue
            run = fetch(f'{api}/repos/{repo}/actions/runs/{run_id}', token)
            if (run.get('id') == run_id
                    and run.get('path') == WORKFLOW_PATH
                    and run.get('event') in ATTESTING_EVENTS
                    and run.get('status') == 'completed'
                    and run.get('conclusion') == 'success'
                    and (run.get('repository') or {}).get('full_name') == repo
                    and (run.get('head_repository') or {}).get('full_name') == repo):
                return True, f'Tree {tree[:12]} was verified in full by run {run_id}'
    except (OSError, ValueError, KeyError, TypeError, AttributeError,
            urllib.error.URLError):
        return False, 'Attestation lookup failed'
    return False, 'No valid attestation for this tree'


def head_tree() -> str:
    try:
        return git('rev-parse', 'HEAD^{tree}').strip()
    except (OSError, subprocess.CalledProcessError):
        return ''


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--event', required=True)
    parser.add_argument('--base', default='')
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--summary', type=Path, required=True)
    args = parser.parse_args()
    full, reason = select(args.event, args.base)
    tree = head_tree()
    # Manual and scheduled runs stay full whatever was verified before: they
    # are how the proofs are repeated on purpose.
    if full and args.event in ATTESTING_EVENTS:
        verified, why = verified_tree(tree, os.environ.get('GITHUB_REPOSITORY', ''),
                                      os.environ.get('GH_TOKEN', ''),
                                      os.environ.get('GITHUB_API_URL', 'https://api.github.com'))
        if verified:
            full, reason = False, why
        else:
            reason += f' ({why[0].lower()}{why[1:]})'
    with args.output.open('a') as handle:
        handle.write(f'full={str(full).lower()}\ntree={tree}\n')
    with args.summary.open('a') as handle:
        handle.write(f'## Accuracy check selection\n\nFull regression proofs: **{full}**\n\n{reason}. '
                     'Live evidence checks always run.\n')
    print(json.dumps({'full': full, 'reason': reason}))


if __name__ == '__main__':
    main()
