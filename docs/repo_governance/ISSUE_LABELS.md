# Issue labels and triage

![Labels](https://img.shields.io/badge/labels-28-1D76DB)

The canonical catalogue is `.github/labels.json`: 28 issue labels, made of 20
common labels with standardized colors and descriptions, and 8
repository-specific labels that name the area of the library a change touches.
The [catalogue](#catalogue) below lists each one. I change the manifest through a
pull request, and `tests/test_issue_labels.py` fails if this guide, the
manifest, the issue forms and Dependabot disagree.

The manifest sets `"prune": true`. Synchronization therefore deletes live labels
outside the catalogue, but only unused ones. Deleting a label on GitHub removes
it from every issue and pull request, closed ones included, and cannot be undone.

Before applying any change, synchronization checks every label it would delete
against all issues and pull requests, open and closed. If one still carries a
label, it stops with no change applied and names the labels. A name containing a
comma cannot be passed to the API's label filter. For such a name it reads the
labels of every issue and pull request instead, up to 10,000, and stops if a
response is malformed or the limit is reached. When it stops:
- remove the label from its issues and pull requests, or add it to the
  manifest;
- then rerun **Sync issue labels**.

## Automation

- `labels-sync.yml` validates pull requests without write permission. On the
  default branch, trusted pushes and manual dispatches reconcile the live
  catalogue and verify an exact match, including descriptions and colors. The
  reconcile job has `issues: write` and `pull-requests: read`: without the
  latter, the issues API omits pull requests, and the check before a deletion
  would miss a label only pull requests carry.
- `labels-drift.yml` runs offline fixtures on pull requests and checks the live
  catalogue daily or on manual dispatch. It is read-only and retains evidence.
- Both job summaries render every name, colour, path and API message as code
  that the value cannot break: a backtick, a pipe or a line break in a live
  value stays inside its table row.
- `issue-triage.yml` audits all open issues on issue and milestone events or
  manual dispatch. It is read-only: a failed report identifies metadata for me
  to resolve.
- `pr-milestones.yml` (*Keep milestones issue-only*) removes any milestone set
  on a pull request:
  - when the pull request is opened, reopened, edited or given a milestone;
  - in a daily, manual or post-change scan of every pull request.

  It runs on `pull_request_target` from the default branch, checks out nothing
  and runs no pull-request code. Its only write permission is
  `pull-requests: write`. The job summary lists each pull request and the
  milestone removed.

## Issue rules

I assign every open issue to the configured maintainer, choose exactly one of
`P1`, `P2`, `P3`, and attach an open milestone. P1 is reserved for confirmed
calculation bugs producing wrong results for valid inputs. P2 covers material
correctness, robustness or maintainability defects; P3 covers non-blocking work.

Each issue has:
- a target end date (YYYY-MM-DD, Europe/Rome) strictly before its milestone's
  due date;
- a Dependencies and blocking status section.

I use `blocked` exactly while a named prerequisite issue or external condition
is unresolved. Write an external blocker on its own line, like this one (an
optional list bullet is accepted):

```text
External condition: waiting for access to an Excel host
```

Use `None` when there are no prerequisites. Pull request references do not
count as prerequisite issues.

Forms provide default priorities and collect dates and dependencies; triage
confirms their meaning. Existing issues are audited under the same rules, with
findings reported rather than silently rewritten. Dependabot uses `ci` and `P3`.

## Using the labels

Every issue has exactly one priority, the type labels that fit the work, and the
area labels of the library it touches. Pull requests carry the labels of the
work they deliver.

### Priorities

| Label | Use |
| --- | --- |
| `P1` | A confirmed calculation bug: wrong results for valid inputs. Only together with `bug`; never for tooling, documentation or evidence work. |
| `P2` | A material correctness, robustness or maintainability defect, including a defect in the evidence or CI that could let a wrong result through. Default of the bug and security contact forms. |
| `P3` | Non-blocking quality, testing, documentation, setup or technical-debt work. Default of the feature and documentation forms and of Dependabot. |

A form sets the default priority; triage confirms or changes it. Raise a bug to
`P1` only once the calculation defect is confirmed against an independent
reference.

### Types

| Label | Use |
| --- | --- |
| `bug` | A confirmed defect. The bug form sets it. |
| `enhancement` | A new capability or a meaningful improvement. The feature form sets it. |
| `documentation` | Documentation work. The documentation form sets it. |
| `behavior-change` | Add it to work that changes externally observable behavior: a computed value, an error, a parameterization or a worksheet signature. |
| `refactor` | Restructuring that preserves supported behavior. If supported behavior changes, use `behavior-change` instead. |
| `tests` | Regression assertions, test harnesses, fixtures, grids used as tests, or manual validation. |
| `ci` | Workflows, automated checks, the self-hosted Excel runner, and dependency updates, including every Dependabot pull request. |
| `repository` | Repository structure, governance, metadata or tooling that is not a workflow. |
| `release` | Release preparation, certification, evidence packaging or publication. |
| `security` | Hardening, supply-chain work and private security contact requests, which never contain vulnerability details. Never use it for an undisclosed vulnerability: report it as `SECURITY.md` describes. |

### Status and triage

| Label | Use |
| --- | --- |
| `blocked` | A prerequisite named in the dependency section, as an issue link or an `External condition:` line, is unresolved. Remove it once all are resolved; in the same edit, delete each resolved `External condition:` line, because triage treats every such line as unresolved. |
| `question` | I need information or a decision before the work can proceed. |
| `help wanted` | I welcome outside investigation, implementation, validation or review. |
| `good first issue` | A well-scoped task suitable for a first contribution. |
| `duplicate` | Close the issue and link the issue or pull request that already covers the work. |
| `invalid` | Close the issue: it is incomplete, unsupported or not actionable as submitted. |
| `wontfix` | Close the issue: I have decided not to do this work. |

### Areas

The 8 repository-specific labels name the part of the library a change touches.
Add every area that applies; work on CI, governance or documentation alone
needs none. The discrete distributions have no area label.

| Label | Use |
| --- | --- |
| `accuracy` | Accuracy contracts, reference values, grids and holdouts, numerical envelopes and limitations, Excel evidence, and the benchmark studies. |
| `api` | The worksheet and VBA surface: function names, parameters and parameterization, return values, error values, and compatibility for callers. |
| `continuous` | Gamma, Beta, Exponential, Weibull, Uniform and the other continuous distributions of `M_STATS_PROBDIST_CONTINUOUS`. |
| `core` | The shared floating-point primitives, constants, guards and arithmetic helpers of `M_STATS_PROBDIST_CORE`. |
| `normal-family` | Standard Normal, Normal and Lognormal, their tails, intervals and inverses, and the normal kernel of `M_STATS_PROBDIST_NORMALFAMILY`. |
| `performance` | Runtime, algorithmic efficiency or scalability, and speed benchmarks; not accuracy, which is `accuracy`. |
| `special-functions` | Gamma, beta, log-gamma and log-beta, the incomplete functions, and the combinatorial and inverse kernels of `M_STATS_PROBDIST_SPECIALFUNCS`. |
| `t-family` | Student t, Chi-square and F, and their shared numerical paths in `M_STATS_PROBDIST_TFAMILY`. |

## Pull request rules

Milestones belong to issues only. A pull request carries the labels of the
work it delivers and no milestone; I plan through the issue it closes or
relates to. Do not set a milestone on a pull request: *Keep milestones
issue-only* removes it and records the removal in its job summary.

Pull requests opened by Dependabot follow the same rule. During triage, link
each one to an existing issue or a new tracking issue, and give that issue the
milestone. Their workflow runs target `main` and so receive the write
permission. If a run ever gets a read-only token, it warns instead of failing,
and the next daily scan removes the milestone.

## Closing issues

I close an issue as completed only when linked evidence meets its acceptance
criteria. I tick an acceptance box only when that evidence demonstrates it.
When I close the issue, I add a **Resolution** section to its body with:

- the change: the pull request and the merge commit;
- the checks actually run, with their results, including the Excel run and its
  assertion count when VBA changed;
- what was not validated, any deviation from the proposed fix, and any accepted
  limitation.

I close for another reason (duplicate, not planned or superseded) only with an
explicit explanation; the `duplicate`, `invalid` and `wontfix` rows above say
which label applies.

## Catalogue

`.github/labels.json` is the source of truth; this table restates it, in its
order.

| Label | Colour | Description |
| --- | --- | --- |
| `accuracy` | `#E99695` | Accuracy contracts, numerical envelopes, references, holdouts, benchmarks, and limitations |
| `api` | `#0366D6` | Worksheet/VBA API, parameterization, compatibility, return values, and caller contracts |
| `behavior-change` | `#C24E00` | Intentional externally observable behavior change |
| `blocked` | `#B60205` | Cannot proceed until a named dependency is complete |
| `bug` | `#D73A4A` | Confirmed defect in behavior, correctness, error handling, or a documented contract |
| `ci` | `#24292F` | Automated quality checks, workflow maintenance, or dependency updates |
| `continuous` | `#BFE5BF` | Gamma, Beta, Exponential, Weibull, Uniform, and other continuous-distribution surfaces |
| `core` | `#BFDADC` | Core floating-point primitives, constants, guards, arithmetic helpers, and shared numerics |
| `documentation` | `#0075CA` | Documentation for users or for my repository maintenance |
| `duplicate` | `#CFD3D7` | An existing issue or pull request already covers this work |
| `enhancement` | `#A2EEEF` | New capability or meaningful improvement to supported behavior |
| `good first issue` | `#7057FF` | Well-scoped task suitable for a first contribution |
| `help wanted` | `#008672` | I welcome external investigation, implementation, validation, or review |
| `invalid` | `#E4E669` | Report is incomplete, unsupported, or not actionable as submitted |
| `normal-family` | `#C5DEF5` | Standard Normal, Normal, Lognormal, Gaussian tails, intervals, and inverse-normal behavior |
| `P1` | `#B60205` | Confirmed calculation bug: wrong results for valid inputs (reserved for this case) |
| `P2` | `#D93F0B` | Material correctness, robustness, or maintainability defect |
| `P3` | `#FBCA04` | Non-blocking quality, testing, documentation, setup or technical-debt work |
| `performance` | `#0E8A16` | Runtime performance, algorithmic efficiency, scalability, or performance benchmarking |
| `question` | `#D876E3` | I need further information or need to make a decision |
| `refactor` | `#6F42C1` | Internal restructuring intended to preserve supported behavior |
| `release` | `#B60276` | Release preparation, certification, evidence, packaging, or publication |
| `repository` | `#1D76DB` | Repository structure, governance, metadata, tooling, or maintenance |
| `security` | `#8B0000` | Security boundary, vulnerability handling, supply chain, or hardening |
| `special-functions` | `#5319E7` | Gamma, beta, log-gamma/log-beta, incomplete functions, combinatorial and inverse kernels |
| `t-family` | `#0052CC` | Student t, Chi-square, F distributions, and their shared numerical paths |
| `tests` | `#006B75` | Regression coverage, test harnesses, fixtures, or manual validation |
| `wontfix` | `#6E7781` | I have decided not to implement or continue this work |

## Local validation

```sh
node .github/scripts/labels-sync.mjs --self-test
node .github/scripts/labels-sync.mjs --mode validate
node .github/scripts/labels-drift.mjs --self-test
python tools/check_issues.py --check-config
python -m unittest discover -s tests -p 'test_check_issues.py'
python -m unittest discover -s tests -p 'test_pr_milestones_workflow.py'
python -m unittest discover -s tests -p 'test_labels_sync.py'
python -m unittest discover -s tests -p 'test_issue_labels.py'
```
