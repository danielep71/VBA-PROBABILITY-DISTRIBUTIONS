# Security and dependency audit — 2026-11-15

Baseline: `074b800b9f8413a3fa3f57ccdc971927044b367c` on `release/1.0.0`, the
v1.0.0 candidate after the last numerical change. This repeats the audit of
[2026-09-30](SECURITY_DEPENDENCY_AUDIT_2026-09-30.md), whose baseline
`3c205bf` predates eleven of the eighteen workflows.

It is a point-in-time review of the repository's controls, not a penetration
test and not release certification. A change to a workflow, a requirements
file or the settings baseline after this baseline is outside it and has to be
read again before the tag.

## What was read

- All 18 files of `.github/workflows/`, `.github/dependabot.yml`, the four
  files of `.github/scripts/`, and the two requirements files.
- The live repository settings, through `tools/setup_repository.py --verify`.
- PyPI's record of each pinned package.

Eleven workflows are new since `3c205bf`: `daily-traffic`,
`excel-runner-watch`, the five `gamma-*` study workflows, `issue-triage`,
`link-check`, `pr-milestones` and `release`. Seven existed and have changed:
`accuracy-gate`, `codeql`, `excel-vba-regression`, `labels-drift`,
`labels-sync`, `scorecard` and `static-checks`.

## Verified by the maintainer

The token this audit ran with could not read the following. The maintainer
read each of them in the repository's settings on 2026-11-15, after the
integration of `release/1.0.0` into `main`.

| Item | Found |
| --- | --- |
| Code-scanning alerts on `main` | none open; ten closed. The seven Scorecard alerts the 2026-09-30 audit left open are among the closed |
| Dependabot alerts | none open, none closed, for vulnerabilities and for malware |
| Secret-scanning alerts | none open, none closed |
| Actions enabled | yes; every action and reusable workflow is allowed, which the baseline permits. What restrains an action is the pin below |
| Actions pinned to a full-length commit SHA | required |
| Default permissions of the workflow token | read |
| Actions may create or approve pull requests | no |
| Approval of workflows from forks | required for all external contributors |
| Retention of logs and artifacts | 90 days |
| The `analytics` environment | deployment restricted to selected branches: `main` only. The App's private key is a secret of the environment |
| Collaborators | none: only the owner can contribute |

The six Actions settings match `docs/repo_governance/repository-settings.json`.
They were read in the settings page, not by `setup_repository.py`, whose
record of the same day covers the settings its token could read.

The runner host changed on 2026-11-15. The reference Excel runner is
`excel-runner-b`, a desktop owned and used by the maintainer alone, set up as
`docs/EXCEL_VBA_CI.md` describes; `excel-runner-a` stays registered without the
`excel` label. The host's own configuration beyond that guide was not
inspected.

### Changes after the baseline

Between `074b800` and `302ea78`, the paths this audit covers changed in three
files: the regression workflow gained the Dependabot guard of finding 1, and
the two Excel policy files carry a new assertion count. No action pin, no
requirements file and no settings baseline changed.

## Workflows

What holds in every workflow, read from the files and kept by
`tests/test_workflow_controls.py`:

- every action is pinned to a 40-hex commit; there is no tag, branch, local
  action or reusable workflow;
- top-level permissions grant no write. Ten jobs write, each to the scopes
  listed in that test and no other;
- every checkout sets `persist-credentials: false`, except the `traffic` job,
  which pushes the daily snapshot with the job's own token;
- every `pip install` uses `--require-hashes` and `--only-binary=:all:`;
- every job has a timeout and every workflow a concurrency group.

What holds by reading, without a test:

- no event field, head ref or free-text input is interpolated into a script.
  Values a pull request supplies reach scripts through `env:`. The one
  free-text dispatch input, `export_studies`, is checked against a list in
  `ci/ExcelGridExport.psm1` before use;
- no workflow checks out pull-request code in a privileged context.
  `pr-milestones` is the only `pull_request_target` workflow: it checks out
  nothing, reads the event file, and can only set or clear a milestone;
- `excel-runner-watch` is the only `workflow_run` workflow. It checks out the
  default branch and re-runs a failed run at most once;
- no secret is exposed to a pull-request job. The traffic App's key is used
  by one job, on schedule and dispatch, behind the `analytics` environment;
- pull-request CodeQL analysis runs without `security-events: write` and
  uploads nothing.

### Findings

| # | Severity | Finding | Disposition |
| ---: | --- | --- | --- |
| 1 | Medium | Dependabot's pull requests reached the self-hosted Excel runner. The job refused forks, but Dependabot's branches live in the repository, and it opens action and package bumps against `release/1.0.0`. A compromised upstream release would have run on the Excel host before anyone read the bump. | **Fixed with this audit.** The job also refuses pull requests opened by Dependabot; a rule in `tests/test_workflow_controls.py` requires both guards on every job of that runner a pull request can reach. The bump first runs on the host with the push that follows its merge, after review. |
| 2 | Low | An account with write access can run arbitrary code on the Excel runner: by dispatching any of the six workflows that use it from a branch it edited, or by a pull request from such a branch. No condition written in a workflow can prevent this, because the dispatched ref supplies the workflow. | **Accepted.** The control is write access to the repository, which the owner alone holds (no collaborator, confirmed above), and a runner that is not shared. Recorded in `docs/EXCEL_VBA_CI.md`. To be read again if a second account gains write access. |
| 3 | Low | The `traffic` job combines dispatch, `contents: write`, persisted credentials and the App's private key, with no condition on the ref. As in finding 2, a condition in the file would not bind a dispatched branch. | **Accepted.** The binding control is the `analytics` environment's deployment-branch rule, confirmed above: only `main` may deploy to it, so a run dispatched from another branch does not receive the key. |
| 4 | Low | A rehearsal of `release.yml` checks out the ref it is given and runs that ref's two release scripts with a token. | **Accepted.** The token is `contents: read` and `actions: read`; the signing and draft jobs are skipped in a rehearsal by the event name, not by an input. |
| 5 | Informational | The signing and draft jobs of `release.yml` have no environment approval. A tag pushed by an account with write access reaches signing if the checks pass. | **Accepted.** The workflow requires an annotated tag equal to `VERSION`, a commit that is an ancestor of `main`, the Excel certification of that exact commit, and publishes only a draft. The tag ruleset forbids moving or deleting a `v*` tag; it does not restrict who creates one. |
| 6 | Informational | The five `gamma-*` study workflows belong to a closed study. Their push triggers name branches that no longer exist, so only dispatch can start them. | **Left as they are.** They keep the closed study reproducible. Finding 2 covers their exposure. |

Findings 2 to 5 share one fact: every account with write access is trusted
with the runner and with the release path. That is true of this repository
today and is the assumption to revisit first if it stops being true.

## Dependencies

The VBA library has no downloaded dependency. CI installs two hash-locked
sets, binary wheels only:

| File | Packages |
| --- | --- |
| `benchmark/requirements-ci.txt` | mpmath 1.3.0, sympy 1.14.0 |
| `tools/requirements-lint-ci.txt` | ruff 0.17.4, mypy 2.4.1, pyyaml 6.0.3, and five transitive dependencies: ast-serialize 0.13.1, librt 0.16.0, mypy-extensions 1.2.0, pathspec 1.1.1, typing-extensions 4.16.0 |

On 2026-11-15 PyPI listed no known vulnerability for any of the ten pinned
versions. Eight are the latest release. mpmath 1.4.1 exists and is held back
on purpose: Dependabot ignores mpmath 1.4 and later, because the references
are computed with 1.3.0. ruff 0.17.5 exists; the weekly group will propose it.

The actionlint binary is downloaded and checked against a pinned SHA-256.

Dependabot watches three sets weekly, all against `release/1.0.0`: the
actions, and the two requirements directories. After the release its target
branch has to follow the branch that is developed next.

This says nothing about vulnerabilities not yet known.

## Settings

`python tools/setup_repository.py --verify` on 2026-11-15: every setting of
`docs/repo_governance/repository-settings.json` that the token could read
matches, including visibility, merge methods, the five security features and
the five rulesets with their required checks. The six Actions settings could
not be read by the tool; the maintainer read them in the settings page, as
recorded above. The record is
`docs/repo_governance/evidence/settings-verify-2026-11-15.json`.

## Verification boundary

The workflow statements above come from reading the files at the baseline;
the five tested rules are also enforced on every pull request. The findings
table was checked line by line against the files. Nothing here certifies an
Excel run, the accuracy evidence or the release: issue 31 owns those.
