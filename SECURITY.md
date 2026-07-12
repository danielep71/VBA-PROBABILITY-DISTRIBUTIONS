# 🔒 Security Policy

<p align="left">
  <img alt="Reporting" src="https://img.shields.io/badge/Reporting-Private-orange">
  <img alt="Scope" src="https://img.shields.io/badge/Scope-VBA_Source-6f42c1">
  <img alt="Stable version" src="https://img.shields.io/badge/Stable_Tag-v1.0.0-217346">
  <img alt="Development branch" src="https://img.shields.io/badge/Main-Unreleased-lightgrey">
</p>

This project distributes plain-text Excel VBA source modules. There is no installer,
external DLL, compiled executable, or background service.

The attack surface is therefore limited, but responsible disclosure still matters.
This policy explains which versions are supported, what should be reported privately,
and which numerical issues should instead be submitted as ordinary bug reports.

---

## 📦 Supported versions

| Version or branch | Support status |
|---|---|
| `v1.0.0` — latest tagged stable version | ✅ Supported |
| `main` — unreleased development code | ⚠️ Best effort |
| Older tags, snapshots, and copied third-party versions | ❌ Not supported |

Security fixes are normally developed on `main` and included in the next tagged
version.

When reporting an issue, identify the exact source state you used:

- a release tag, such as `v1.0.0`; or
- the full Git commit SHA when testing `main` or another snapshot.

Do not report only “latest,” because the default branch can change after the issue
is observed.

---

## 📣 Reporting a vulnerability

**Do not open a public GitHub issue for a suspected security vulnerability.**

Use one of these private channels:

1. **GitHub private vulnerability reporting**
   - Open the repository’s **Security** tab.
   - Choose **Report a vulnerability**.

2. **Email the maintainer**

```text
danielep71@gmail.com
```

Include:

- the affected tag or full commit SHA;
- the affected module and procedure;
- the exact `K_STATS_*` or `PROB_*` call, where applicable;
- Excel version, Office bitness, and operating system;
- minimal reproduction steps;
- observed behavior;
- expected behavior;
- practical security impact;
- any proposed remediation.

Please do not attach workbooks containing confidential, personal, client, or
production data.

---

## 🎯 What qualifies as a security issue

Examples that should be reported privately include:

- execution of unintended code;
- unintended modification of workbook data or VBA projects;
- disclosure of information outside the documented calculation result;
- uncontrolled or effectively non-terminating resource consumption;
- a crafted input that causes Excel to hang persistently or exhaust resources;
- a numerical defect that creates a concrete security, control, or integrity impact.

### Ordinary numerical bugs

A wrong numerical result is important, but it is not automatically a security
vulnerability.

Use the public **Bug report** template when the issue is:

- an incorrect density, CDF, survival probability, quantile, or moment;
- an accuracy problem within an otherwise terminating calculation;
- a wrong `#NUM!` versus `#VALUE!` classification;
- a documentation or parameterization inconsistency;
- a discrepancy with an independent reference implementation;
- a performance problem that does not create a security or availability impact.

When uncertain, report privately. The maintainer can reclassify the report safely.

---

## 🧭 Scope

### In scope

- VBA source under `src/`;
- the regression harness under `tests/`;
- repository-supplied example or release artifacts;
- numerical kernels that fail to terminate or consume uncontrolled resources;
- behavior that violates the documented security or integrity boundary.

### Out of scope

- vulnerabilities in Microsoft Excel, Office, Windows, macOS, or the VBA runtime;
- macro-security configuration controlled by the user or organization;
- unrelated VBA code in the host workbook;
- modified copies obtained from third parties;
- unsupported historical snapshots;
- ordinary numerical differences within documented tolerances.

---

## ⏱️ Disclosure process

This is a solo-maintained open-source project, so response times are best effort.

The expected process is:

1. The report is acknowledged.
2. The issue is assessed and reproduced where possible.
3. A remediation and release approach is agreed.
4. A fix is developed on a private or controlled branch when necessary.
5. A corrected tagged version is published.
6. Public disclosure follows after users have had reasonable time to update.

Please allow reasonable time for investigation and remediation before public
disclosure.

Credit will be included in release notes when requested.

---

## 🧰 Safe-use guidance

- Obtain the source only from the official repository or a tagged release.
- Review the plain-text `.bas` files before importing them.
- Keep Excel macro security enabled at the organization’s approved level.
- Do not lower macro-security settings for this library.
- Compile the project after importing the modules.
- Run `Test_STATS_PROBDIST_RunAll` before production use.
- Record the tag or commit SHA used in controlled workbooks and validation evidence.
- Do not treat an unreviewed `main` snapshot as a stable release.

If Windows marks downloaded source files as blocked, review their origin before
using **Properties → Unblock**.

---

## 👤 Maintainer

Maintained by **Daniele Penza**.
