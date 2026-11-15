# 2. Supported environment

## 2.1 What you need

| Item | Requirement |
|---|---|
| Host application | Desktop Microsoft Excel with VBA |
| Macros | Enabled for the workbook |
| Workbook format | `.xlsm`, `.xlsb` or `.xlam` |
| Add-in, installer, DLL, COM component | None |
| Non-standard VBA references | None |

The library consists of standard VBA modules (`.bas` files). It has no class modules, no user forms and no Windows API declarations.

## 2.2 What is tested

The retained runtime evidence covers **64-bit Office on Windows**. The recorded Excel regression run used Excel 16.0 (build 20430), 64-bit. This is the only environment for which the project holds test evidence.

## 2.3 What is not tested

- **32-bit Office.** The modules declare no Windows API calls, so there is no bitness-specific code path. The project nevertheless holds no retained evidence for 32-bit Office. Do not infer a certified 32-bit environment from the use of pure VBA.
- **Anything else.** The project documentation makes no statement about other hosts or operating systems. Treat them as untested.

Compatibility claims apply only to environments certified for the release you selected. In any other environment, validate the library yourself first (section 4.5).

## 2.4 Certification status

The README assurance block states the current test evidence and whether a release is certified. Read the [README](../../../README.md), the changelog and the release notes of the release you install. Do not assume a certification state from this manual.

## 2.5 Security

VBA code runs with your Office permissions. Before you enable macros:

- obtain the files from the official repository or a verified release;
- compare the tag, the `VERSION` file and any published hash;
- follow your organisation's macro, trusted-location and signing policy.

Report a suspected vulnerability privately, as described in `SECURITY.md`.
