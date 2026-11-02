# Issue labels and triage

The canonical catalogue is `.github/labels.json`: 20 common labels plus eight
repository-specific labels already assigned to issues (open or closed), for a
total of 28. I preserve their existing colors, descriptions and assignments:
`accuracy`, `api`, `continuous`, `core`, `normal-family`, `performance`,
`special-functions`, and `t-family`. Common labels use standardized colors
and descriptions. I change the manifest through a pull request.
Synchronization removes only labels outside this complete catalogue.

## Automation

- `labels-sync.yml` validates pull requests without write permission. On the default
  branch, trusted pushes and manual dispatches reconcile the live catalogue and
  verify an exact match, including descriptions and colors.
- `labels-drift.yml` runs offline fixtures on pull requests and checks the live
  catalogue daily or on manual dispatch. It is read-only and retains evidence.
- `issue-triage.yml` audits all open issues on issue and milestone events or manual
  dispatch. It is read-only: a failed report identifies metadata for me to resolve.

## Issue rules

I assign every open issue to the configured maintainer, choose exactly one of
`P1`, `P2`, `P3`, and attach an open milestone. P1 is reserved for confirmed
calculation bugs producing wrong results for valid inputs. P2 covers material
correctness, robustness or maintainability defects; P3 covers non-blocking work.

Each issue has a target end date (YYYY-MM-DD, Europe/Rome) strictly before its
milestone due date and a Dependencies and blocking status section. I use
`blocked` exactly while a named prerequisite issue or external condition is
unresolved. Pull request references do not count as prerequisite issues.

Forms provide default priorities and collect dates and dependencies; triage
confirms their meaning. Existing issues are audited under the same rules, with
findings reported rather than silently rewritten. Dependabot uses `ci` and `P3`.

## Local validation

```sh
node .github/scripts/labels-sync.mjs --self-test
node .github/scripts/labels-sync.mjs --mode validate
node .github/scripts/labels-drift.mjs --self-test
python tools/check_issues.py --check-config
python -m unittest discover -s tests -p 'test_check_issues.py'
```
