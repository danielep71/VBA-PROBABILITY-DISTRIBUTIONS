#!/usr/bin/env python3
"""Preview, apply or verify the repository settings baseline through the GitHub CLI.

The baseline is docs/repo_governance/repository-settings.json: visibility,
merge methods and repository features, the security features, the GitHub
Actions settings, and the repository rulesets with their required checks.
Preview (the default) and verify only read; apply writes and then reads
everything back. Actions settings are read and reported, never written.
"""
# Postponed annotations keep the version check in main() reachable on old Pythons.
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import subprocess
import sys
from collections.abc import Callable
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import quote

ROOT = Path(__file__).resolve().parents[1]
BASELINE = ROOT / 'docs' / 'repo_governance' / 'repository-settings.json'
SCHEMA_VERSION = 2
EVIDENCE_SCHEMA_VERSION = 2
TOP_KEYS = {'schema_version', 'repository', 'visibility', 'settings', 'security', 'actions', 'rulesets'}
RULE_KEYS = ('name', 'target', 'enforcement', 'conditions', 'rules', 'bypass_actors')
SETTING_KEYS = (
    'allow_merge_commit', 'allow_squash_merge', 'allow_rebase_merge', 'allow_auto_merge',
    'allow_update_branch', 'delete_branch_on_merge', 'web_commit_signoff_required',
    'has_issues', 'has_projects', 'has_wiki', 'has_discussions', 'has_downloads',
)
# Write order: Dependabot alerts before the security updates that depend on them.
SECURITY_KEYS = (
    'private_vulnerability_reporting', 'vulnerability_alerts', 'automated_security_fixes',
    'secret_scanning', 'secret_scanning_push_protection',
)
ANALYSIS_KEYS = ('secret_scanning', 'secret_scanning_push_protection')
# Free only on public repositories; a private one needs a paid licence for secret scanning.
PUBLIC_ONLY_KEYS = ('private_vulnerability_reporting', 'secret_scanning', 'secret_scanning_push_protection')
# Dependabot security updates that are enabled but paused open no pull requests.
PAUSED = 'on (paused)'
PAUSE_NOTE = ('NOTE: Dependabot security updates are enabled but paused, so they open no pull '
              'requests. GitHub pauses them when Dependabot pull requests go unanswered; '
              'answering or closing the open Dependabot pull requests resumes them. Until then '
              'they are drift, and --apply cannot resume them.')
# GitHub Actions settings: baseline key -> (endpoint under repos/{r}/, response field, type, allowed).
ACTIONS_SETTINGS: dict[str, tuple[str, str, type, tuple[Any, ...]]] = {
    'actions_enabled': ('actions/permissions', 'enabled', bool, ()),
    'sha_pinning_required': ('actions/permissions', 'sha_pinning_required', bool, ()),
    'default_workflow_permissions': ('actions/permissions/workflow', 'default_workflow_permissions',
                                     str, ('read', 'write')),
    'can_approve_pull_request_reviews': ('actions/permissions/workflow',
                                         'can_approve_pull_request_reviews', bool, ()),
    'fork_pr_approval_policy': ('actions/permissions/fork-pr-contributor-approval', 'approval_policy',
                                str, ('first_time_contributors_new_to_github',
                                      'first_time_contributors', 'all_external_contributors')),
    'artifact_and_log_retention_days': ('actions/permissions/artifact-and-log-retention', 'days',
                                        int, ()),
}
MERGE_METHOD_SETTINGS = {
    'merge': 'allow_merge_commit', 'squash': 'allow_squash_merge', 'rebase': 'allow_rebase_merge',
}
REPOSITORY_PATTERN = re.compile(r'[A-Za-z0-9][A-Za-z0-9-]*/[A-Za-z0-9_.-]+')

# GitHub API values are JSON; a call takes a method, an endpoint and an optional payload.
Json = Any
Call = Callable[..., Json]
Step = tuple[str, str, Json]


class SetupError(Exception):
    """A refusal or failure; the command exits with status 2."""


class ApiError(SetupError):
    """A GitHub API call that failed, with its HTTP status when gh reported one."""

    def __init__(self, method: str, endpoint: str, status: str) -> None:
        super().__init__(f'{method} {endpoint} failed (HTTP {status}). Check gh authentication, '
                         'repository access and Administration permission. No automatic rollback '
                         'is performed; rerun the preview to inspect partial changes.')
        self.status = status


def api(method: str, endpoint: str, payload: Json = None) -> Json:
    """Call the GitHub REST API through an authenticated gh CLI."""
    command = ['gh', 'api', '--hostname', 'github.com', '--method', method,
               '-H', 'Accept: application/vnd.github+json',
               '-H', 'X-GitHub-Api-Version: 2022-11-28', endpoint]
    if payload is not None:
        command += ['--input', '-']
    # Never inherit HTTP debug logging: it can print authentication headers.
    env = dict(os.environ)
    env.pop('GH_DEBUG', None)
    # gh writes UTF-8 whatever the locale code page is.
    result = subprocess.run(command, input=json.dumps(payload) if payload is not None else None,
                            text=True, encoding='utf-8', capture_output=True, env=env, check=False)
    if result.returncode:
        # Do not repeat gh's stderr: it can hold server or credential detail.
        match = re.search(r'HTTP (\d{3})', result.stderr)
        raise ApiError(method, endpoint, match.group(1) if match else 'unknown')
    return json.loads(result.stdout) if result.stdout.strip() else None


def pages(call: Call, endpoint: str, key: str | None = None) -> list[Json]:
    """Every item of a paginated GET endpoint."""
    result: list[Json] = []
    page = 1
    while True:
        separator = '&' if '?' in endpoint else '?'
        value = call('GET', f'{endpoint}{separator}per_page=100&page={page}')
        batch = value[key] if key else value
        result.extend(batch)
        if len(batch) < 100:
            return result
        page += 1


def canonical(value: Json) -> Json:
    """Ruleset arrays are sets for this baseline: ignore API ordering, never values."""
    if isinstance(value, dict):
        return {k: canonical(v) for k, v in sorted(value.items())}
    if isinstance(value, list):
        return sorted((canonical(x) for x in value), key=lambda x: json.dumps(x, sort_keys=True))
    return value


def rule_spec(ruleset: dict[str, Json]) -> Json:
    """The managed part of a ruleset; response metadata is ignored, unknown rules are kept."""
    return canonical({k: ruleset.get(k, [] if k == 'bypass_actors' else None) for k in RULE_KEYS})


def _bools(value: Json, keys: tuple[str, ...], what: str) -> None:
    if not isinstance(value, dict) or set(value) != set(keys) or any(type(v) is not bool for v in value.values()):
        raise SetupError(f'The baseline must give exactly the documented boolean {what}: {", ".join(keys)}.')


def load_baseline(path: Path) -> dict[str, Json]:
    """Read and validate the baseline; refuse anything the tool cannot apply as written."""
    value = json.loads(path.read_text(encoding='utf-8'))
    if not isinstance(value, dict) or set(value) != TOP_KEYS or value['schema_version'] != SCHEMA_VERSION:
        raise SetupError('Unsupported baseline schema.')
    if not isinstance(value['repository'], str) or not REPOSITORY_PATTERN.fullmatch(value['repository']):
        raise SetupError('The baseline repository must be an owner/name on github.com.')
    if value['visibility'] not in ('public', 'private'):
        raise SetupError('The baseline visibility must be public or private.')
    _bools(value['settings'], SETTING_KEYS, 'repository settings')
    _bools(value['security'], SECURITY_KEYS, 'security features')
    security = value['security']
    public_only = [key for key in PUBLIC_ONLY_KEYS if security[key]]
    if value['visibility'] == 'private' and public_only:
        raise SetupError('A private baseline cannot turn on ' + ', '.join(public_only) + ': GitHub '
                         'offers them free only on public repositories, and a private repository '
                         'needs a paid licence this tool does not manage.')
    actions = value['actions']
    if not isinstance(actions, dict) or set(actions) != set(ACTIONS_SETTINGS):
        raise SetupError('The baseline must give exactly the documented Actions settings: '
                         + ', '.join(ACTIONS_SETTINGS) + '.')
    for key, (_, _, kind, allowed) in ACTIONS_SETTINGS.items():
        if type(actions[key]) is not kind or (allowed and actions[key] not in allowed):
            raise SetupError(f'The baseline Actions setting {key} must be a {kind.__name__}'
                             + (f' among {", ".join(allowed)}.' if allowed else '.'))
    if not 1 <= actions['artifact_and_log_retention_days'] <= 400:
        raise SetupError('The baseline artifact and log retention must be 1 to 400 days.')
    if security['secret_scanning_push_protection'] and not security['secret_scanning']:
        raise SetupError('Push protection requires secret scanning.')
    if security['automated_security_fixes'] and not security['vulnerability_alerts']:
        raise SetupError('Dependabot security updates require Dependabot alerts.')
    names: list[str] = []
    for ruleset in value['rulesets']:
        if not isinstance(ruleset, dict) or set(ruleset) != set(RULE_KEYS):
            raise SetupError('Each baseline ruleset must have exactly: ' + ', '.join(RULE_KEYS) + '.')
        if ruleset['target'] not in ('branch', 'tag') or ruleset['enforcement'] != 'active':
            raise SetupError(f'Ruleset "{ruleset["name"]}" must be an active branch or tag ruleset.')
        if ruleset['bypass_actors'] or ruleset['conditions']['ref_name']['exclude']:
            raise SetupError(f'Ruleset "{ruleset["name"]}" must not have bypass actors or excluded refs.')
        types = [rule['type'] for rule in ruleset['rules']]
        if len(types) != len(set(types)):
            raise SetupError(f'Ruleset "{ruleset["name"]}" repeats a rule type.')
        for rule in ruleset['rules']:
            parameters = rule.get('parameters') or {}
            for method in parameters.get('allowed_merge_methods', []):
                setting = MERGE_METHOD_SETTINGS.get(method)
                if setting is None or not value['settings'][setting]:
                    raise SetupError(f'Ruleset "{ruleset["name"]}" allows {method} merges, which the '
                                     'baseline repository settings disable.')
            for check in parameters.get('required_status_checks', []):
                if not isinstance(check.get('context'), str) or type(check.get('integration_id')) is not int:
                    raise SetupError(f'Ruleset "{ruleset["name"]}" has a required check without a '
                                     'context and an integration ID.')
        names.append(ruleset['name'])
    if not names or len(names) != len(set(names)):
        raise SetupError('The baseline needs at least one ruleset, each with a unique name.')
    return dict(value)


def read(call: Call, endpoint: str) -> tuple[Json, str | None]:
    """GET an endpoint; return its value, or None and the HTTP status of a failure."""
    try:
        return call('GET', endpoint), None
    except ApiError as error:
        return None, error.status


def security_states(call: Call, root: str, metadata: dict[str, Json]) -> dict[str, dict[str, Json]]:
    """Each security feature as observed: True, False or None when GitHub does not say."""
    analysis = metadata.get('security_and_analysis') or {}

    def status(key: str) -> dict[str, Json]:
        entry = analysis.get(key)
        if not entry:
            return {'observed': None, 'source': f'security_and_analysis.{key} not reported (needs admin access)'}
        return {'observed': entry.get('status') == 'enabled', 'source': f'security_and_analysis.{key}'}

    states: dict[str, dict[str, Json]] = {}
    value, failure = read(call, root + '/private-vulnerability-reporting')
    states['private_vulnerability_reporting'] = (
        {'observed': bool((value or {}).get('enabled')), 'source': 'GET private-vulnerability-reporting'}
        if failure is None else
        {'observed': None, 'source': f'GET private-vulnerability-reporting unreadable (HTTP {failure})'})

    # Only the endpoint reports a pause, so it is always read; security_and_analysis
    # says enabled for paused updates too, and stands in only when the endpoint is unreadable.
    value, failure = read(call, root + '/automated-security-fixes')
    if failure is None:
        enabled, paused = bool((value or {}).get('enabled')), bool((value or {}).get('paused'))
        fixes: dict[str, Json] = {'observed': (PAUSED if paused else True) if enabled else False,
                                  'paused': paused and enabled, 'source': 'GET automated-security-fixes'}
    elif failure == '404':
        fixes = {'observed': False, 'paused': False, 'source': 'GET automated-security-fixes (404: off)'}
    else:
        fixes = dict(status('dependabot_security_updates'), paused=None)
        fixes['source'] += (f'; paused state unknown (GET automated-security-fixes unreadable: '
                            f'HTTP {failure})')

    # 204 means on and 404 off. When the endpoint cannot be read, enabled
    # security updates still prove the alerts on: GitHub requires them.
    _, failure = read(call, root + '/vulnerability-alerts')
    if failure is None:
        alerts: dict[str, Json] = {'observed': True, 'source': 'GET vulnerability-alerts (204: on)'}
    elif failure == '404':
        alerts = {'observed': False, 'source': 'GET vulnerability-alerts (404: off)'}
    elif fixes['observed'] in (True, PAUSED):
        alerts = {'observed': True, 'source': f'inferred from Dependabot security updates being on '
                                              f'(GET vulnerability-alerts unreadable: HTTP {failure})'}
    else:
        alerts = {'observed': None, 'source': f'GET vulnerability-alerts unreadable (HTTP {failure})'}

    states['vulnerability_alerts'] = alerts
    states['automated_security_fixes'] = fixes
    for key in ANALYSIS_KEYS:
        states[key] = status(key)
    return {key: states[key] for key in SECURITY_KEYS}


def actions_states(call: Call, root: str) -> dict[str, dict[str, Json]]:
    """Each Actions setting as observed, or None with the reason it could not be read."""
    states: dict[str, dict[str, Json]] = {}
    bodies: dict[str, tuple[Json, str | None]] = {}
    for key, (endpoint, field, _, _) in ACTIONS_SETTINGS.items():
        if endpoint not in bodies:
            bodies[endpoint] = read(call, f'{root}/{endpoint}')
        value, failure = bodies[endpoint]
        if failure is not None:
            states[key] = {'observed': None, 'source': f'GET {endpoint} unreadable (HTTP {failure})'}
        elif not isinstance(value, dict) or field not in value:
            states[key] = {'observed': None, 'source': f'GET {endpoint} does not report {field}'}
        else:
            states[key] = {'observed': value[field], 'source': f'GET {endpoint}'}
    return states


def observe(call: Call, repo: str) -> dict[str, Json]:
    """Read every setting the baseline manages. Read-only."""
    root = f'repos/{repo}'
    metadata = call('GET', root)
    listed = pages(call, root + '/rulesets?includes_parents=true')
    own = [r for r in listed if r.get('source_type', 'Repository') == 'Repository']
    names = [r['name'] for r in own]
    duplicates = sorted({name for name in names if names.count(name) > 1})
    if duplicates:
        raise SetupError('Ambiguous duplicate ruleset names: ' + ', '.join(duplicates))
    return {
        'metadata': metadata,
        'rulesets': {r['name']: call('GET', f'{root}/rulesets/{r["id"]}') for r in own},
        'inherited': [r['name'] for r in listed if r.get('source_type', 'Repository') != 'Repository'],
        'security': security_states(call, root, metadata),
        'actions': actions_states(call, root),
    }


_ABSENT = '<absent>'


def _keyed(spec: Json) -> Json:
    """A ruleset spec with its rules keyed by type, for readable difference paths."""
    keyed = dict(spec)
    rules = spec.get('rules') or []
    types = [rule.get('type') for rule in rules]
    if len(types) == len(set(types)):
        keyed['rules'] = {rule.get('type'): rule.get('parameters', {}) for rule in rules}
    return keyed


def differences(observed: Json, baseline: Json, path: str = '') -> list[str]:
    """Human-readable paths where an observed value departs from the baseline."""
    if isinstance(observed, dict) and isinstance(baseline, dict):
        found: list[str] = []
        for key in sorted(set(observed) | set(baseline), key=str):
            found += differences(observed.get(key, _ABSENT), baseline.get(key, _ABSENT),
                                 f'{path}.{key}' if path else str(key))
        return found
    if canonical(observed) == canonical(baseline):
        return []
    return [f'{path}: observed {json.dumps(canonical(observed))}, baseline {json.dumps(canonical(baseline))}']


def assess(baseline: dict[str, Json], state: dict[str, Json]) -> dict[str, Json]:
    """Compare every managed setting with what GitHub reports."""
    metadata = state['metadata']
    settings = [{'setting': key, 'baseline': value, 'observed': metadata.get(key),
                 'matches': metadata.get(key) == value}
                for key, value in baseline['settings'].items()]
    security = [{'feature': key, 'baseline': value, **state['security'][key],
                 'matches': state['security'][key]['observed'] == value}
                for key, value in baseline['security'].items()]
    rulesets: list[dict[str, Json]] = []
    for desired in baseline['rulesets']:
        actual = state['rulesets'].get(desired['name'])
        if actual is None:
            rulesets.append({'name': desired['name'], 'id': None, 'status': 'missing', 'differences': []})
            continue
        found = differences(_keyed(rule_spec(actual)), _keyed(rule_spec(desired)))
        if rule_spec(actual) != rule_spec(desired) and not found:
            found = ['rules: observed and baseline differ']
        rulesets.append({'name': desired['name'], 'id': actual.get('id'),
                         'status': 'differs' if found else 'matches', 'differences': found})
    # Actions settings are never written: an unreadable one is unverified (None), not drift.
    actions = [{'setting': key, 'baseline': value, **state['actions'][key],
                'matches': None if state['actions'][key]['observed'] is None
                else state['actions'][key]['observed'] == value}
               for key, value in baseline['actions'].items()]
    managed = {r['name'] for r in baseline['rulesets']}
    visibility = {'baseline': baseline['visibility'], 'observed': metadata.get('visibility'),
                  'matches': metadata.get('visibility') == baseline['visibility']}
    drift = (not visibility['matches'] or not all(row['matches'] for row in settings + security)
             or any(row['matches'] is False for row in actions)
             or any(row['status'] != 'matches' for row in rulesets))
    return {'result': 'DRIFT' if drift else 'VERIFIED', 'visibility': visibility,
            'settings': settings, 'security': security, 'actions': actions, 'rulesets': rulesets,
            'unmanaged_rulesets': sorted(name for name in state['rulesets'] if name not in managed),
            'inherited_rulesets': sorted(state['inherited'])}


def plan(baseline: dict[str, Json], state: dict[str, Json], report: dict[str, Json], repo: str) -> list[Step]:
    """The writes that would make the repository match the baseline. Visibility is never written."""
    root = f'repos/{repo}'
    steps: list[Step] = []
    desired = {r['name']: r for r in baseline['rulesets']}
    for row in report['rulesets']:
        if row['status'] == 'missing':
            steps.append(('POST', root + '/rulesets', desired[row['name']]))
        elif row['status'] == 'differs':
            steps.append(('PUT', f'{root}/rulesets/{row["id"]}', desired[row['name']]))
    analysis: dict[str, dict[str, str]] = {}
    for row in report['security']:
        if row['matches']:
            continue
        if row['feature'] in ANALYSIS_KEYS:
            analysis[row['feature']] = {'status': 'enabled' if row['baseline'] else 'disabled'}
        else:
            endpoint = root + '/' + row['feature'].replace('_', '-')
            steps.append(('PUT' if row['baseline'] else 'DELETE', endpoint, None))
    if analysis:
        steps.append(('PATCH', root, {'security_and_analysis': analysis}))
    changes = {row['setting']: row['baseline'] for row in report['settings'] if not row['matches']}
    if changes:
        steps.append(('PATCH', root, changes))
    return steps


def print_report(report: dict[str, Json]) -> None:
    visibility = report['visibility']
    if not visibility['matches']:
        print(f'DIFF visibility: observed {visibility["observed"]}, baseline {visibility["baseline"]} '
              '(never changed by this tool; follow docs/repo_governance/DISTRIBUTION.md)')
    for row in report['settings']:
        if not row['matches']:
            print(f'DIFF setting {row["setting"]}: observed {json.dumps(row["observed"])}, '
                  f'baseline {json.dumps(row["baseline"])}')
    for row in report['security']:
        if not row['matches']:
            print(f'DIFF security {row["feature"]}: observed {json.dumps(row["observed"])}, '
                  f'baseline {json.dumps(row["baseline"])} [{row["source"]}]')
        if row.get('paused'):
            print(PAUSE_NOTE)
    for row in report['actions']:
        if row['matches'] is False:
            print(f'DIFF actions {row["setting"]}: observed {json.dumps(row["observed"])}, '
                  f'baseline {json.dumps(row["baseline"])} (never changed by this tool; change it '
                  'in the repository Actions settings)')
        elif row['matches'] is None:
            print(f'NOTE: actions {row["setting"]} not verified: {row["source"]}.')
    for row in report['rulesets']:
        if row['status'] == 'missing':
            print(f'DIFF ruleset "{row["name"]}": missing')
        for line in row['differences']:
            print(f'DIFF ruleset "{row["name"]}" ({row["id"]}) {line}')
    for name in report['unmanaged_rulesets']:
        print(f'NOTE: unmanaged repository ruleset "{name}" stays active; this tool never changes it.')
    for name in report['inherited_rulesets']:
        print(f'NOTE: inherited ruleset "{name}" also applies; this tool never changes it.')


def evidence(baseline_path: Path, repo: str, state: dict[str, Json], report: dict[str, Json],
             steps: list[Step], mode: str, now: datetime) -> dict[str, Json]:
    """A dated JSON record of the managed settings as GitHub reported them."""
    try:
        shown = baseline_path.resolve().relative_to(ROOT).as_posix()
    except ValueError:
        shown = baseline_path.name
    metadata = state['metadata']
    return {
        'schema_version': EVIDENCE_SCHEMA_VERSION,
        'record': 'repository-settings-evidence',
        'checked_at': now.strftime('%Y-%m-%dT%H:%M:%SZ'),
        'tool': 'tools/setup_repository.py',
        'mode': mode,
        'repository': metadata.get('full_name', repo),
        'default_branch': metadata.get('default_branch'),
        'baseline': {'path': shown, 'sha256': hashlib.sha256(baseline_path.read_bytes()).hexdigest()},
        'result': report['result'],
        'visibility': report['visibility'],
        'settings': report['settings'],
        'security': report['security'],
        'actions': report['actions'],
        'rulesets': report['rulesets'],
        'unmanaged_rulesets': report['unmanaged_rulesets'],
        'inherited_rulesets': report['inherited_rulesets'],
        'planned_changes': [{'method': method, 'endpoint': endpoint} for method, endpoint, _ in steps],
        'scope': 'Only the settings in the baseline are recorded; everything else is outside this record.',
    }


def write_evidence(path: Path, record: dict[str, Json], now: datetime) -> None:
    """Write the record; a {date} token in the path becomes the UTC date of the check."""
    target = Path(str(path).replace('{date}', now.strftime('%Y-%m-%d')))
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(json.dumps(record, indent=2) + '\n', encoding='utf-8', newline='\n')
    print(f'Evidence written to {target}')


def check_prerequisites(call: Call, repo: str, baseline: dict[str, Json], state: dict[str, Json],
                        steps: list[Step]) -> None:
    """Refuse to apply anything unless every write can succeed and stay safe."""
    metadata = state['metadata']
    if metadata.get('archived') or metadata.get('disabled'):
        raise SetupError('The repository is archived or disabled.')
    if not (metadata.get('permissions') or {}).get('admin'):
        raise SetupError('Applying the baseline needs administrative access to the repository.')
    if metadata.get('visibility') != baseline['visibility']:
        raise SetupError(f'The repository is {metadata.get("visibility")}, the baseline says '
                         f'{baseline["visibility"]}. This tool never changes visibility; complete '
                         'the distribution review first.')
    if not any('/rulesets' in endpoint for _, endpoint, _ in steps):
        return
    root = f'repos/{repo}'
    branch = call('GET', root + '/branches/' + quote(metadata['default_branch'], safe=''))
    runs = pages(call, f'{root}/commits/{branch["commit"]["sha"]}/check-runs?filter=latest', 'check_runs')
    required = [check for ruleset in baseline['rulesets'] for rule in ruleset['rules']
                if rule['type'] == 'required_status_checks'
                for check in rule['parameters']['required_status_checks']]
    for check in required:
        if not any(run.get('name') == check['context']
                   and (run.get('app') or {}).get('id') == check['integration_id']
                   and run.get('status') == 'completed' and run.get('conclusion') == 'success'
                   for run in runs):
            raise SetupError(f'Required check "{check["context"]}" has not succeeded on the default '
                             'branch HEAD. Make it green before applying protection that requires it.')


def run(args: argparse.Namespace, call: Call = api, now: datetime | None = None) -> int:
    baseline = load_baseline(args.baseline)
    repo = args.repo or baseline['repository']
    if not REPOSITORY_PATTERN.fullmatch(repo) or repo.split('/')[1] in ('.', '..'):
        raise SetupError('The repository must be an explicit owner/name on github.com.')
    mode = 'apply' if args.apply else 'verify' if args.verify else 'preview'
    state = observe(call, repo)
    report = assess(baseline, state)
    steps = plan(baseline, state, report, repo)
    print(f'Repository: {state["metadata"].get("full_name", repo)}; mode: {mode}; '
          f'baseline: {args.baseline.name}')
    print_report(report)

    if args.apply and report['result'] == 'DRIFT':
        check_prerequisites(call, repo, baseline, state, steps)
        if any(method == 'PUT' and '/rulesets/' in endpoint for method, endpoint, _ in steps) \
                and not args.update_existing:
            raise SetupError('A named ruleset differs. Review the preview, then add --update-existing '
                             'to replace that whole ruleset; rules outside the baseline are removed.')
        for method, endpoint, payload in steps:
            call(method, endpoint, payload)
            print(f'Applied {method} {endpoint}')
        state = observe(call, repo)
        report = assess(baseline, state)
        steps = plan(baseline, state, report, repo)

    stamp = now or datetime.now(timezone.utc)
    if args.evidence:
        write_evidence(args.evidence, evidence(args.baseline, repo, state, report, steps, mode, stamp), stamp)
    if report['result'] == 'VERIFIED':
        unread = [row['setting'] for row in report['actions'] if row['matches'] is None]
        print('VERIFIED: every setting in the baseline matches'
              + (f', except {len(unread)} Actions setting(s) that could not be read (see NOTE)'
                 if unread else '') + '. Settings outside it are not checked.')
        return 0
    if args.apply:
        print_report(report)
        raise SetupError('Read-back verification failed; the differences above remain.')
    if args.verify:
        print('DRIFT: the repository does not match the baseline.')
        return 1
    for method, endpoint, payload in steps:
        print(f'PLAN {method} {endpoint}')
        if payload is not None:
            print(json.dumps(payload, indent=2))
    print('PREVIEW ONLY: nothing was changed. Use --apply to write this plan.')
    return 0


def main() -> int:
    if sys.version_info < (3, 10):
        print('ERROR: Python 3.10 or later is required.', file=sys.stderr)
        return 2
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument('--repo', help='owner/name to inspect (default: the baseline repository)')
    parser.add_argument('--baseline', type=Path, default=BASELINE, help='settings baseline (JSON)')
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument('--preview', action='store_true', help='show differences and the plan; read-only (default)')
    mode.add_argument('--apply', action='store_true', help='write the plan, then read everything back')
    mode.add_argument('--verify', action='store_true', help='read-only; exit 1 on drift')
    parser.add_argument('--update-existing', action='store_true',
                        help='with --apply: allow replacing a differing named ruleset')
    parser.add_argument('--evidence', type=Path,
                        help='write a dated JSON evidence record here ({date} becomes YYYY-MM-DD)')
    args = parser.parse_args()
    if args.update_existing and not args.apply:
        parser.error('--update-existing requires --apply')
    try:
        return run(args)
    except (SetupError, OSError, ValueError, KeyError, TypeError) as exc:
        print(f'ERROR: {exc}', file=sys.stderr)
        return 2


if __name__ == '__main__':
    sys.exit(main())
