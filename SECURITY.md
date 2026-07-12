# 🔒 Security Policy

<p align="left">
  <img alt="Reporting" src="https://img.shields.io/badge/reporting-private-orange">
  <img alt="Scope" src="https://img.shields.io/badge/scope-VBA_source-6f42c1">
  <img alt="Platform" src="https://img.shields.io/badge/platform-Excel_VBA-blue">
</p>

This project ships **VBA source only** — a set of `.bas` modules that are
imported into an Excel workbook and run with the user's Excel and Windows
privileges. There is no add-in, no installer, no external DLL and no worksheet-
function marshalling. The attack surface is therefore small, but security reports
are still taken seriously. This page explains what is supported, how to report an
issue privately, and what is in scope.

---

## 📦 Supported versions

<p align="left">
  <img alt="Support" src="https://img.shields.io/badge/Support-Latest_release-217346">
</p>

Security fixes are applied to the latest release. Older tags are not patched —
please upgrade before reporting.

| Version | Supported |
| --- | --- |
| `v1.0.0` (latest) | ✅ |
| earlier / unreleased | ❌ |

---

## 📣 Reporting a vulnerability

<p align="left">
  <img alt="Channel" src="https://img.shields.io/badge/Channel-Private_disclosure-orange">
  <img alt="Public_issues" src="https://img.shields.io/badge/Public_issues-Do_not_use-red">
</p>

**Please do not open a public issue for a security problem**, and do not post
proof-of-concept exploit code in a public thread.

Report privately through either:

- **GitHub private vulnerability reporting** — on the repository, go to the
  **Security** tab → **Report a vulnerability** (enable it under
  *Settings → Security* if it is not already on), or
- email the maintainer at:

```text
danielep71@gmail.com
```

Helpful details to include:

- Affected version (`v1.0.0`, etc.) and Excel version / bitness / OS
- The specific module and, if applicable, the `K_STATS_*` call and arguments
- A clear description of the issue and its impact
- Minimal reproduction steps
- Any suggested remediation, if you have one

---

## ⏱️ What to expect

<p align="left">
  <img alt="Response" src="https://img.shields.io/badge/response-best_effort-blue">
</p>

This is a solo-maintained project, so responses are best-effort rather than
guaranteed within a fixed window. You can expect:

- Acknowledgement of your report,
- An assessment of validity and severity,
- A fix in a new release when a valid issue is confirmed,
- Credit in the release notes if you would like it.

Please allow reasonable time for a fix before any public disclosure.

---

## 🎯 Scope

<p align="left">
  <img alt="In_scope" src="https://img.shields.io/badge/In_scope-This_project-217346">
  <img alt="Out_of_scope" src="https://img.shields.io/badge/Out_of_scope-Host_environment-orange">
</p>

**In scope**

- The VBA source under `src/` (`CORE`, `SPECIALFUNCS`, and the distribution-
  family modules)
- The regression harness under `tests/`
- Behavior that could cause the library to return a **silently wrong numerical
  result** where the error contract says it should fail (this is treated as a
  correctness-security issue, not just a bug)
- Any input that causes uncontrolled resource consumption (e.g. a non-terminating
  iteration in an inverse/root-finding routine)

**Out of scope**

- Microsoft Excel, Office, Windows, or the VBA runtime themselves
- Issues that require the user to disable Excel macro security or to import
  untrusted code unrelated to this project
- Copies of the modules obtained from anywhere other than this repository
- General numerical inaccuracy within documented tolerances — report these as
  ordinary issues, not security reports

---

## 🧰 Safe use guidance

<p align="left">
  <img alt="Trust" src="https://img.shields.io/badge/Import-Official_source_only-217346">
</p>

- Obtain the modules only from this repository. All code is published in plain
  text under `src/`, so you can review it before importing.
- If a `.bas` file reaches you as a download, right-click the file →
  **Properties** → **Unblock** before importing it into the VBE.
- Keep Excel macro security at its default protected level; nothing in this
  library requires lowering it.

---

## 👤 Maintainer

Maintained by **Daniele Penza**.
