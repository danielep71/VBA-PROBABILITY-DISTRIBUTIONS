# Excel/VBA Regression CI

This repository executes the complete VBA regression harness through Microsoft
Excel on a self-hosted Windows GitHub Actions runner. The workflow is
`.github/workflows/excel-vba-regression.yml`; the host adapter is
`ci/Run-ExcelVbaTests.ps1`.

For the exact-SHA evidence contract and fresh-export procedure, see
[EXCEL_CERTIFICATION.md](EXCEL_CERTIFICATION.md).

## Requirements

- A dedicated Windows machine or VM.
- Desktop Microsoft Excel installed and activated.
- A self-hosted GitHub Actions runner labeled `excel`.
- Excel Trust Center **Trust access to the VBA project object model** enabled
  before the job starts.
- An interactive logged-in Windows session for reliable Office COM automation.

GitHub-hosted `windows-latest` runners do not include desktop Excel, so this
workflow cannot use a standard hosted runner.

## Execution model

For each run, the PowerShell adapter:

1. resolves the full checked-out Git SHA and rejects disagreement with
   `GITHUB_SHA`;
2. reads `.github/excel-evidence-policy.json` and hashes the canonical Git bytes
   of every policy-declared VBA source;
3. creates an isolated Excel COM instance and temporary macro-enabled workbook;
4. records Excel version/build and reads Office bitness from `EXCEL.EXE` rather
   than inferring it from the Windows runner architecture;
5. imports the exact candidate modules and test module;
6. injects a CI-only bridge into `M_STATS_PROBDIST_TEST`;
7. executes all suites in dependency order and reads the private assertion
   counters;
8. requires the exact policy count (currently 909), zero failures, and
   consistent TOTAL/PASS/FAIL counters;
9. closes its workbook and Excel instance and records cleanup separately;
10. writes `test-result.txt` and `excel-certification.json`.

The CI bridge is not committed to the production test module.

The automated compile stage is recorded as PASS only after Excel successfully
executes the imported project through `Application.Run`. It is execution-backed
compile evidence, not a fabricated claim that **Debug > Compile VBAProject** was
separately clicked.

## Artifact contract

The workflow uploads exactly these evidence files:

```text
artifacts/excel-vba-ci/test-result.txt
artifacts/excel-vba-ci/excel-certification.json
```

The JSON record binds the run to the exact candidate SHA and canonical source
hashes. The log digest in the JSON binds the human-readable output to the same
record.

The normal regression job does not export numerical benchmark grids; both grid
entries therefore remain `exported: false`. A green regression artifact alone
cannot refresh `observation_manifest.json` or `holdout_manifest.json`.

## Runner setup

1. In GitHub, open **Settings > Actions > Runners > New self-hosted runner**.
2. Install the runner under a dedicated Windows account.
3. Add the custom label `excel`.
4. Open Excel once under that account and complete activation, first-run,
   privacy, and update prompts.
5. Enable:

```text
File
  > Options
  > Trust Center
  > Trust Center Settings
  > Macro Settings
  > Trust access to the VBA project object model
```

6. Run the GitHub runner interactively:

```powershell
.\run.cmd
```

Microsoft does not recommend unattended Office automation from a
non-interactive Windows service. An interactive dedicated session is materially
more reliable for Excel COM execution.

## Security model

Self-hosted runners execute repository code with the runner account's
permissions. The workflow therefore does not execute pull requests from forks:

```yaml
if: >-
  github.event_name != 'pull_request' ||
  github.event.pull_request.head.repo.full_name == github.repository
```

For an external contribution, review the change and copy or cherry-pick it to a
maintainer-controlled branch before running the Excel workflow. Do not switch
this workflow to `pull_request_target` for untrusted code.

The adapter sets `AutomationSecurity = 1` only on the isolated Excel process it
creates. It does not alter persistent Trust Center configuration; VBProject
access must already be configured for the runner account.

## Local execution

From a configured Windows/Excel machine:

```powershell
Set-Location C:\path\to\VBA-PROBABILITY-DISTRIBUTIONS
.\ci\Run-ExcelVbaTests.ps1
```

A local run records `execution = manual` and no GitHub workflow identity. It is
still bound to the exact local `HEAD` and source hashes.

## Result contract

The injected VBA function returns:

```text
TOTAL=<count>;PASS=<count>;FAIL=<count>
```

The PowerShell process exits with code `0` only when all policy-declared
assertions pass and cleanup succeeds. Missing modules, source/checkout identity
problems, Excel automation failures, invalid counters, assertion-count drift,
test failures, and cleanup failures all produce a nonzero result.

## Building the workbook under test from source

The adapter above needs **Trust access to the VBA project object model**
because it imports the modules and injects the CI bridge through the VBE, and
it assembles the tested workbook at run time. An offline builder produces the
workbook under test instead, as a reproducible file bound to one commit
(#110). The Excel job does not use it yet: the adapter and the workflow are
unchanged, and they stay the path of record until a built workbook has been
certified on the runner.

Build and check it on any machine with Python 3.10+ and Git; no Excel is
needed:

```bash
python3 tools/devbuild/build_workbook.py
python3 tools/devbuild/verify_workbook.py \
  test-results/workbook-build/probdist-regression.manifest.json --require-verified
```

The builder writes `test-results/workbook-build/probdist-regression.xlsm` and,
beside it, `probdist-regression.manifest.json`:

- **Modules.** The VBA project holds exactly the modules listed in
  `regression_sources` of `.github/excel-evidence-policy.json`, in that order,
  plus empty `ThisWorkbook` and `shBuild` document modules. A listed module that
  is missing, a `.bas` under `src/` or `tests/` that is not listed, or a module
  whose `VB_Name` differs from its file name stops the build. Every module is
  embedded as it is, with CRLF line endings; no tracked file is changed.
- **Compiled on open.** The source is stored compressed as [MS-OVBA]
  specifies, and `_VBA_PROJECT` carries no p-code cache, so Excel compiles the
  project from source when it opens the workbook. Writing the project needs no
  VBE access.
- **Commit stamp.** The `Build` sheet holds the commit and its committer time
  under the workbook-level defined names `build_commit` and
  `build_commit_time`. A build from a working tree with tracked changes appends
  `(working tree with changes: NOT evidence-grade)` to the commit.
- **Reproducible.** A rebuild of the same commit and working tree is
  byte-identical: the stamp uses the committer time, not the clock, the VBA
  project ID derives from the commit, and the package parts are stored in a
  fixed order with a fixed timestamp.
- **Manifest.** The manifest records the commit, whether the build was clean,
  the stamp, and the SHA-256 of the workbook and of every module, both of its
  Git blob (the same `sha256:` digest as the `sources` of a regression record)
  and of the text embedded in the workbook.

`verify_workbook.py` reports the binding of a workbook to its sources:

| Binding | Meaning |
| --- | --- |
| `VERIFIED` | The build was clean and every check passes. The workbook's SHA-256 and stamp match the manifest. Its package parts are the builder's. Its VBA project holds exactly the builder's streams: `_VBA_PROJECT` without p-code; `dir`, `PROJECT` and `PROJECTwm` equal to what the builder writes for the recorded modules; the generated empty document modules; and each module with the recorded digest. The commit's policy and Git blobs agree with the manifest. Finally, the workbook is byte-identical to a rebuild from the commit |
| `FAIL` | At least one of those checks found a mismatch, or `--commit` names another commit |
| `UNVERIFIED` | Nothing to check against (no manifest, no workbook, or a commit this clone does not have), or a build from a working tree with changes |

It exits 1 on `FAIL`, and on `UNVERIFIED` with `--require-verified`. The
binding proves which sources the workbook holds; it does not prove that the
project compiles or that any test passes. Only an Excel run proves that.

The built workbook holds no CI bridge. The bridge reads the private assertion
counters of `M_STATS_PROBDIST_TEST`, so it can only live inside that module,
and committing it there changes an evidence-bound source. Until that is
decided, a built workbook can run `Test_STATS_PROBDIST_RunAll` but cannot
return the counters that the result contract above requires.

What remains before the Excel job uses a built workbook:

1. a CI entry point that the built workbook can return counters through,
   committed and listed in the policy;
2. an Excel job that builds the workbook, requires `VERIFIED` against
   `GITHUB_SHA`, opens it without VBE access, runs the entry point and records
   the workbook's SHA-256 in the regression record;
3. a first certified run of that job on the runner, recorded before the COM
   path is retired.

## Troubleshooting

### Programmatic access denied

Enable **Trust access to the VBA project object model** under the exact Windows
account running the self-hosted runner.

### Excel COM class not registered

Confirm that desktop Excel is installed and activated. Repair Office if
`New-Object -ComObject Excel.Application` fails.

### Workflow remains queued

Confirm that the runner is online and has all labels:

```text
self-hosted
Windows
X64
excel
```

### Orphaned Excel process

Check for modal Excel dialogs, first-run prompts, add-ins, Protected View
prompts, or Office update dialogs. The adapter closes only its own temporary
workbook and Excel instance; it never kills unrelated Excel processes.

## Operational controls

- Use a dedicated machine or VM.
- Limit the runner account's permissions.
- Keep Windows and Office patched.
- Do not store unrelated credentials in the account profile.
- Review workflow and evidence-policy changes carefully.
- Keep the Excel regression job as a required release check once the runner is
  stable.
