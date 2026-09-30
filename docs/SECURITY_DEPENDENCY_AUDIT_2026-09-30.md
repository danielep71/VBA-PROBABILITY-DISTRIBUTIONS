# Security and dependency audit — 2026-09-30

Baseline: `3c205bf6a14735edcad16f52c25f6b50cc574cd6`, after CI PR #50.
This is a point-in-time repository-controls audit, not a penetration test or
release certification. No VBA implementation or numerical evidence is changed.

## Repository controls

Applied and verified in GitHub settings during this audit:

- Private vulnerability reporting enabled, alongside the email fallback in SECURITY.md.
- Secret scanning enabled, with push protection for supported secret patterns.
- Full-length commit SHA pins required for Actions.
- Workflow approval required for all external contributors, rather than only first-time contributors.

Already enabled: dependency graph, Dependabot vulnerability/malware alerts,
security updates and grouped security updates. The default workflow token is
read-only; Actions cannot create/approve PRs by default. Workflow-specific
write grants are limited to label synchronization and security-result
publication on trusted events. PR CodeQL jobs have read permission only and
retain SARIF artifacts. There is no `pull_request_target` execution path.

The persistent Windows/Excel runner excludes fork PRs in its workflow and was
offline in the preceding CI audit. Approval is not a sandbox: inspect workflow
and executable changes before approving external runs. Do not expose this host
to untrusted candidate code. Runner host configuration was not inspected.

## Dependencies

The production VBA library has no downloaded package dependencies. CI uses
mpmath 1.3.0, SymPy 1.14.0, Ruff 0.16.6, checksum-pinned actionlint 1.7.12 and
commit-pinned Actions. Optional gmpy2/MPFR and R/Rmpfr research environments are
outside the CI lockfiles and must retain their own evidence/environment records.

The audit patch adds wheel hashes for the existing mpmath/SymPy versions,
requires binary-only hash-checked installation, and extends Dependabot to both
Python requirements directories. It preserves the CodeQL Action grouping.
Updates are assigned to the maintainer with P3 for initial triage; raise the
priority for security advisories and set a milestone before integration.

Both wheel hashes match PyPI release metadata and independently hashed downloads.
A clean install using only those wheels passed; deliberately substituting an
incorrect hash failed before installation. The PyPI metadata for all three
pinned Python packages listed no known vulnerabilities at inspection time.
GitHub showed zero open Dependabot vulnerability, malware and secret alerts.
These signals do not establish the absence of undiscovered vulnerabilities.

## Scanning-alert triage

All eight initially open findings were read individually in GitHub. Alert
numbers below are code-scanning alert IDs, not repository issue numbers.

| Alert | Finding | Disposition |
| --- | --- | --- |
| [#1](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/security/code-scanning/1) | Branch-Protection, 3/10 | Valid: no required PR/check enforcement. Remains open; issue #31 owns configuration and proof. |
| [#3](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/security/code-scanning/3) | SAST, 7/10 | Tool detected, but sampled history not fully checked. CodeQL is newly configured and PR analysis is artifact-only. Keep open and re-evaluate after subsequent scans; do not claim every historical commit was scanned. |
| [#4](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/security/code-scanning/4) | Fuzzing, 0/10 | No recognized fuzzer integration. Existing numerical and mutation fixtures are valuable but do not meet this control. Remains open as an improvement. |
| [#5](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/security/code-scanning/5) | Code-Review, 0/10 | No approved changesets in the sample. Bot review completion and maintainer self-review do not imply independent approval. Remains open; review policy is under #31. |
| [#6](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/security/code-scanning/6) | Maintained, 0/10 | Repository is younger than 90 days. This is a limited-history warning, not evidence of abandonment. Keep open and re-evaluate after the age window. |
| [#7](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/security/code-scanning/7) | CII-Best-Practices, 0/10 | No OpenSSF badge registration. Optional assurance improvement; no badge or compliance is claimed. Remains open. |
| [#8](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/security/code-scanning/8) | CodeQL network data written to file | Dismissed as false positive with rationale: GitHub label metadata is intentionally appended to the runner-provided `GITHUB_STEP_SUMMARY`; network data selects neither the destination nor executable code. Reassess if the sink or trust boundary changes. |
| [#9](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/security/code-scanning/9) | Pinned-Dependencies, 8/10 | Genuine missing pip hash enforcement. Addressed by this patch; leave open until a post-merge Scorecard scan confirms resolution. |

CodeQL covers Python, JavaScript and Actions, not VBA. Its successful workflow
status proves analysis execution, not a finding-free result. PR SARIF is retained
without publication, so reviewers must inspect it; branch enforcement remains
separate. Seven Scorecard alerts remain open at the audit checkpoint.

## Verification boundary

Local checks: clean hash-locked install, deliberate bad-hash rejection,
Dependabot scope/label/assignment validation, actionlint, Ruff, release semantics
and whitespace. Hosted results belong to the actual PR SHA and must be reviewed
before merge. These changes do not certify a live Excel run, fresh exports,
zero coverage debt, or v1.0.0 readiness. Issues #29, #31 and #47 retain those gates.

References: [pip secure installs](https://pip.pypa.io/en/stable/topics/secure-installs/)
and [Dependabot configuration](https://docs.github.com/en/code-security/concepts/supply-chain-security/about-the-dependabot-yml-file).
