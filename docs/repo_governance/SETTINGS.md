# ⚙️ Repository settings · Non-default configuration

[![Scope](https://img.shields.io/badge/Scope-Non--default_settings-00338D?style=flat-square)](#scope)
[![Format](https://img.shields.io/badge/Format-Portable-005EB8?style=flat-square)](#reuse)
[![Comparison](https://img.shields.io/badge/Comparison-Documented_defaults-483698?style=flat-square)](#sources)

**Only deliberate differences · No repository identifiers · Reusable configuration**

<a id="scope"></a>

## 🧭 Comparison basis

This document records only settings that differ from the comparison baseline:
a new standalone GitHub.com repository, without inherited organization rules,
an applied security configuration or copied template customizations. General
options use GitHub's documented REST parameter defaults; those defaults are
not a guarantee of every account's repository-creation UI behavior.

Custom active rulesets are compared with an otherwise unprotected repository,
not with individual checkboxes preselected inside the ruleset editor. In
particular, deletion protection may be preselected when creating a ruleset;
creating and activating that targeted ruleset is still a customization.

Reference documentation checked: **2026-10-29**. I verified the repository settings on **2026-11-02**.
Copying this file does not apply settings or verify another repository.

[General differences](#general) · [Security features](#security) · [Custom protection](#rulesets) · [Reuse](#reuse) · [Sources](#sources)

<a id="general"></a>

## 🏷️ General settings — changed values only

| Setting | Documented baseline | Configured value | Effect / qualification |
| --- | --- | --- | --- |
| Automatically delete merged head branches | Off (`delete_branch_on_merge: false`) | **On** | Cleans up eligible task branches; protected retained branches remain subject to their deletion rules. |
| Always suggest updating PR branches | Off (`allow_update_branch: false`) | **On** | Offers updates even without a freshness requirement; local governance still governs use on shared integration branches. |
| Repository Projects feature | On (`has_projects: true`) | **Off** | Disables the repository feature; this does not describe account-level project boards. Organization defaults may differ. |
| Legacy downloads flag | On (`has_downloads: true`) | **Off** | API metadata difference; not a prohibition on release assets or source downloads. |

Source: GitHub's [repository REST documentation][repo-api]. These values are
configuration differences, not blanket recommendations for every project.

<a id="security"></a>

## 🔐 Security features — always enabled

Their initial state depends on the account, the repository's visibility and
any applied security configuration, so they are not classified as differences
from a default. I verify these features in the repository security settings.

| Feature | Configured value | Availability / effect |
| --- | --- | --- |
| Private vulnerability reporting | **On** | Public repositories only; a private baseline that turns it on is refused. Adds **Security → Report a vulnerability**, the private channel that `SECURITY.md` names. |
| Dependabot alerts | **On** | Every repository. Alerts on vulnerable dependencies. |
| Dependabot security updates | **On** | Every repository; requires Dependabot alerts. Opens pull requests that fix vulnerable dependencies, separately from the scheduled version updates in `.github/dependabot.yml`. GitHub pauses them when Dependabot pull requests go unanswered; paused updates open nothing, so they count as drift until the open Dependabot pull requests are answered or closed. |
| Secret scanning | **On** | Free on public repositories; a private repository needs a paid licence, so a private baseline that turns it on is refused. |
| Push protection | **On** | Same availability as secret scanning, which it requires. Blocks pushes that contain a detected secret. |

Sources: [private vulnerability reporting][pvr], [Dependabot alerts][dependabot-alerts],
[Dependabot security updates][dependabot-updates] and [secret scanning][secret-scanning].

<a id="rulesets"></a>

## 🛡️ Custom active protection

The baseline has no corresponding custom active rulesets. The following five
rulesets are explicit additions. All are **Active**, have **no excluded refs**
and **no bypass actors**. Default-valued supporting parameters below are
included only to describe each custom rule completely; they are not counted
as separate deviations from GitHub defaults.

| Custom ruleset | Target | Added restrictions |
| --- | --- | --- |
| Protect main | `~DEFAULT_BRANCH` | Require PR, resolved conversations and strict up-to-date `Repository integrity (static, no Excel)` and `Strict accuracy gate (pure Python, no Excel)` from GitHub Actions; allow merge commits only; block deletion and force pushes. |
| Protect release branches | `refs/heads/release/**` | Require PR, resolved conversations and the same strict required check; allow squash or merge commits; block deletion and force pushes. |
| Prevent release branch deletion | `refs/heads/release/**` | Independent deletion restriction. |
| Protect release tags | `refs/tags/v*` | Block updates, deletion and force pushes. |
| Protect traffic history | `refs/heads/traffic-history` | Block deletion and force pushes. Direct pushes stay allowed for the optional Traffic history workflow (the repository's daily traffic workflow). |

### 🔬 Supporting parameters for the two PR rulesets

| Parameter | Default-branch rule | Release-branch rule |
| --- | --- | --- |
| Required status context | `Repository integrity (static, no Excel)` and `Strict accuracy gate (pure Python, no Excel)` | `Repository integrity (static, no Excel)` |
| Expected check integration | GitHub Actions (`15368`) | GitHub Actions (`15368`) |
| Up-to-date requirement | Strict | Strict |
| Required check enforced on creation | Yes | No — branch bootstrap exception |
| Required approving reviews | 0 | 0 |
| Extra approval when Copilot opens a PR without a human collaborator | Yes | Yes |
| Required review-thread resolution | Yes | Yes |
| Permitted PR merge methods | Merge | Squash, merge |
| Administrator/app bypass | None | None |

**Operational implications:** zero approvals does not establish independent
review. The extra-approval parameter adds one required review only when the
required count is above zero, so it has no effect here until that count is
raised; GitHub's documentation does not state its default. The release
bootstrap exception does not exempt subsequent changes. Synchronize
integration branches through the approved PR process; do not rely on an
administrator bypass. Tag protection does not itself restrict tag creation or
authorize publishing a release.

The check name and branch patterns are reusable conventions, not built-in
GitHub jobs or automatically existing branches. Confirm that the required job
reports on every relevant PR before adopting these rules.

<a id="reuse"></a>

## ♻️ Reuse and maintenance

- Keep repository names, usernames, URLs, issue numbers and rule IDs out of this file.
- Verify the destination's actual default and inherited policies before copying a value.
- Include a setting only when both its configured value and comparison default are known.
- Remove rows that match the destination's default; retain only its actual differences.
- Keep unverified settings and historical evidence in the setup record.
- Recheck this comparison after platform-default changes or an approved configuration change.

> [!NOTE]
> This is intentionally not a complete settings inventory. Security features
> whose initial state depends on account, visibility or an applied security
> configuration are not classified as non-default without a verified comparator.
> Omission does not mean a feature is disabled or unimportant.

<a id="sources"></a>

## 📚 Default and ruleset references

| Reference | Supports |
| --- | --- |
| [Repository REST API][repo-api] | Declared option defaults and parameter meanings |
| [Private vulnerability reporting][pvr] | Availability on public repositories and how the setting is enabled |
| [Dependabot alerts][dependabot-alerts] and [security updates][dependabot-updates] | What the two Dependabot switches do and how they depend on each other |
| [Secret scanning][secret-scanning] | Availability by visibility and the push protection it enables |
| [Creating rulesets][ruleset-create] | Custom targets, bypass configuration and activation; the editor initially uses Disabled enforcement |
| [Available rules][ruleset-rules] | Meaning of selected protection rules, including editor defaults |
| [Security configuration defaults][security-defaults] | Why organization-applied defaults must be distinguished from a standalone baseline |

[repo-api]: https://docs.github.com/en/rest/repos/repos
[pvr]: https://docs.github.com/en/code-security/how-tos/report-and-fix-vulnerabilities/configure-vulnerability-reporting/configure-for-a-repository
[dependabot-alerts]: https://docs.github.com/en/code-security/concepts/supply-chain-security/dependabot-alerts
[dependabot-updates]: https://docs.github.com/en/code-security/concepts/supply-chain-security/dependabot-security-updates
[secret-scanning]: https://docs.github.com/en/code-security/concepts/secret-security/secret-scanning
[ruleset-create]: https://docs.github.com/en/repositories/configuring-branches-and-merges-in-your-repository/managing-rulesets/creating-rulesets-for-a-repository
[ruleset-rules]: https://docs.github.com/en/repositories/configuring-branches-and-merges-in-your-repository/managing-rulesets/available-rules-for-rulesets
[security-defaults]: https://docs.github.com/en/enterprise-cloud@latest/code-security/how-tos/secure-at-scale/configure-enterprise-security/establish-complete-coverage/create-custom-configuration

---

[↑ Back to comparison basis](#scope)

## Repository-specific protections

I retain the required numerical accuracy gate on the default branch. I also
require workflow approval for all external contributors and keep malware alerts
enabled. These protections are intentional additions.

Actions require full commit SHA pinning, use read-only default token permissions,
and cannot create or approve pull requests. Log and artifact retention is 90 days.
The `actions` block of `repository-settings.json` records these values, with
Actions enabled and the approval policy for all external contributors.
`tools/setup_repository.py` reads them and reports drift, but never writes them:
change one in the repository's Actions settings. An endpoint it cannot read is
reported as not verified.
CodeQL uses the versioned workflow with advanced setup.

This document records configuration; copying it does not apply settings.
