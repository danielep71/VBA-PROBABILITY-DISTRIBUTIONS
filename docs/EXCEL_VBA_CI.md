# Excel/VBA Regression CI

This repository executes the complete VBA regression harness through Microsoft
Excel on a self-hosted Windows GitHub Actions runner. The workflow is
`.github/workflows/excel-vba-regression.yml`; the host adapter is
`ci/Run-ExcelVbaTests.ps1`. A manual dispatch of the same workflow can also
export the accuracy grids unattended through `ci/Export-ExcelGrids.ps1`; see
[Unattended grid export](#unattended-grid-export).

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
8. requires the exact policy count (`expected_assertions` in
   [`.github/excel-evidence-policy.json`](../.github/excel-evidence-policy.json);
   the study batteries, run apart since #300, and the slow production
   batteries, run apart since #13, have their own count in
   [`.github/excel-study-policy.json`](../.github/excel-study-policy.json)),
   zero failures, and
   consistent TOTAL/PASS/FAIL counters;
9. closes its workbook and Excel instance and records cleanup separately;
10. writes `test-result.txt` and `excel-certification.json`.

The CI bridge is not committed to the production test module.

The automated compile stage is recorded as PASS only after Excel successfully
executes the imported project through `Application.Run`. It is execution-backed
compile evidence, not a fabricated claim that **Debug > Compile VBAProject** was
separately clicked.

## Suite levels

The regression has two levels (#299).

- **Full.** Every suite, held to `expected_assertions` in
  `.github/excel-evidence-policy.json`. Only a full run is certification
  evidence. It runs on every push to `main` and `release/**` that changes VBA,
  the runner or a policy, on every dispatch unless the fast level is asked
  for, on every export, and on every pull request that could change what
  Excel executes.
- **Fast.** The same run without the test procedures listed in
  `.github/excel-fast-policy.json`, held to that file's own
  `expected_assertions`. Those procedures drive the double-double Gamma kernel
  and take almost all of the full suite's time. The adapter comments out their
  calls in the temporary workbook only; the committed test module is not
  changed. Each name must be called exactly once, or the run fails. A fast run
  is a smoke test of the runner and harness. Its record states "FAST level,
  not certification evidence" and carries a count the evidence policy rejects.

`tools/select_excel_suite.py` chooses the level in the `select` job. A pull
request runs the fast level only when no changed path can alter the Excel run:
no VBA module anywhere, nothing under `ci/`, no `.github/excel-*.json`, not
the workflow and not the selection tool. Anything else is full, and so is any
event, base or path the tool cannot classify; if the selection job itself
fails, the regression job runs the full suite. A pull request that changes
VBA or the runner therefore runs the full suite on every push, its final head
included. The study batteries and the slow production batteries never run at
the fast level, nor on any pull-request run.

A dispatch can ask for the fast level with `suite: fast`; the input is ignored
when the dispatch also exports a grid.

## Artifact contract

Every run uploads exactly these evidence files as
`excel-vba-regression-<run id>`:

```text
artifacts/excel-vba-ci/test-result.txt
artifacts/excel-vba-ci/excel-certification.json
```

The JSON record binds the run to the exact candidate SHA and canonical source
hashes. The log digest in the JSON binds the human-readable output to the same
record.

The regression itself does not export numerical benchmark grids; both grid
entries therefore remain `exported: false`. A green regression artifact alone
cannot refresh `observation_manifest.json` or `holdout_manifest.json`. A run
dispatched with export inputs uploads a second artifact,
`excel-grid-export-<run id>`, described below.

## Unattended grid export

A `workflow_dispatch` run of the Excel VBA Regression workflow takes seven
inputs. They default to a regression-only run, so push and pull request runs,
and a dispatch without inputs, behave as before.

| Input | Values | Effect |
| --- | --- | --- |
| `export_main` | `true`/`false` | Export `benchmark/probability_accuracy_grid.csv` |
| `export_holdout` | `true`/`false` | Export `benchmark/holdout/holdout_grid.csv` |
| `export_holdout_older` | `true`/`false` | Export `benchmark/holdout_older/holdout_older_grid.csv` |
| `export_survival_boundary` | `true`/`false` | Export `benchmark/survival_boundary/survival_boundary_grid.csv` |
| `export_tchi_large_df` | `true`/`false` | Export `benchmark/tchi_large_df/tchi_large_df_grid.csv` |
| `export_studies` | comma-separated names, default empty | Export these study grids (#225): `prs_gamma_probe`, `prs_chisquare_probe`, `prs_gamma_holdout`, `prs_chisquare_holdout`, `prs_gamma_q_complement` (the `benchmark/positive_ratio_subnormal/` CSVs), `gss_small_shape` (the #23 study grid), `glq_lower_q` (the #24 holdout) and `gbr_binary64_map`, `prs_gamma_holdout2`, `prs_chisquare_holdout2` (the #307 binary64 route map). The #307 fit set `gbf_b64_fit` is no longer exportable: the routines it calls left production after #307 closed |
| `run_studies` | boolean, default `true` | Also run the study batteries of #300 and the slow production batteries of #13 (`Test_STATS_PROBDIST_RunStudies`, against `.github/excel-study-policy.json`) after a green regression. Always on when any grid is exported, and a failure stops the job before the export; pull-request runs never run them |
| `main_mode` | `full`, `missing-only` | Main grid only: `Export_Accuracy_Observations` or `Export_Accuracy_MissingOnly`. The holdout exporters always fill every row |

`export_studies` is one text input rather than one boolean per grid, so a new
study grid needs only an export spec in `ci/ExcelGridExport.psm1` and
`benchmark/apply_excel_export.py`. The workflow passes it to the adapter
through the environment, never on the `cmd` command line, and the adapter
refuses a name it does not know or a name given twice.

When any export input is set, the job runs the regression as usual and, only
if that step succeeded, runs `ci/Export-ExcelGrids.ps1` on the same runner and
checkout, with a job timeout of 120 minutes instead of 20. The job keeps
`contents: read`, receives no secrets and never commits or pushes.

### Why the inputs live in this workflow

GitHub dispatches a workflow only if its file exists on the default branch,
but a dispatched run uses the file as it exists at the dispatched ref. I put
the export inputs in the existing workflow, which is already on `main`, rather
than in a new workflow file. A dispatch on a branch therefore runs the
branch's copy of the export, so a change to it can be tested on the runner
before it merges.

### What the export adapter does

`ci/Export-ExcelGrids.ps1` keeps its pure checks in
`ci/ExcelGridExport.psm1`, which touches neither Excel nor files. It:

1. resolves the full checked-out SHA, rejects disagreement with `GITHUB_SHA`
   and refuses a working tree that differs from it;
2. refuses unless `artifacts/excel-vba-ci/excel-certification.json` names this
   commit and this workflow run, every stage passed, every policy assertion
   passed, and the retained log matches its digest;
3. copies the selected grids, and always the main grid, into a private work
   folder (`artifacts/excel-grid-export-work`) and saves a new macro-enabled
   workbook there. The exporters look for their grid in the workbook's
   folder, so they find it without a dialog and without a change to any
   exporter module;
4. starts its own Excel instance and refuses unless its version, build and
   bitness equal the regression record's. It notes the process ID and checks
   that no Excel process had that ID before it started;
5. imports the `src/` modules, the exporter modules and a CI-only bridge
   module, `M_CI_EXPORT_DIALOG_BRIDGE`, which is never committed. The bridge
   defines a public `MsgBox`, which takes precedence over `VBA.MsgBox` for
   unqualified calls in the project, so each exporter's closing message is
   recorded instead of blocking. A "Could not locate" prompt makes the bridge
   raise an error, so an exporter fails before it can open a file picker;
6. runs `Export_ExcelEnvironment`, then each selected exporter, under a
   watchdog. The watchdog polls for a visible dialog owned by that Excel
   process and enforces a timeout (20 minutes per macro by default, so the
   watchdog acts before the job timeout, which would leave Excel running). On
   either, it records the dialog text or the timeout and kills that one
   process, never another Excel process. Closing Excel runs under the same
   watchdog, with a two-minute limit;
7. requires exactly one exporter message per macro: the success message,
   naming the staged file and, for the holdout exporters, the number of rows
   in scope;
8. checks each exported grid against the checked-out grid. The bytes must be
   the CRLF lines VBA's `Print #` writes; every in-scope row (main grid:
   `evidence_set` is `main grid`, and in missing-only mode `observed_vba` was
   blank) must have a filled `observed_vba`; no other cell and no out-of-scope
   row may change. `ERROR` tokens are counted, not rejected. A missing-only
   run with no blank cell fails, because it would export nothing. A study
   probe CSV is written whole by the probe's `Export_` entry point, so it is
   checked as a *generated* grid instead: the same header and line count,
   every row with the header's field count, and each row's leading key fields
   (the point's identity and echoed arguments) unchanged. Only the
   measurement fields after them may differ;
9. completes `excel_environment.json`: the three fields Excel wrote, unchanged,
   plus `commit_sha` and `commit_sha_recorded_by`
   (`ci/Export-ExcelGrids.ps1`), because `Export_ExcelEnvironment` cannot know
   the commit;
10. writes the artifact and exits 0 only if every step passed.

### Export artifact

`excel-grid-export-<run id>` is uploaded whether the export passed or failed:

```text
export-manifest.json          status, candidate SHA, run, Excel, sources,
                              per-grid row counts, ERROR counts, raw SHA-256,
                              timings, and the failure reason if any
export-log.txt                the adapter's log
excel-certification.json      the regression record of the same run (exact bytes)
test-result.txt               its retained log (exact bytes)
excel_environment.json        the completed environment file
excel_environment.excel.json  the bytes Export_ExcelEnvironment wrote
grids/benchmark/...           each exported grid, exactly as Excel wrote it
diagnostics/                  on failure only: what the exporters left behind
```

A failed export has `"status": "FAIL"`, and `benchmark/apply_excel_export.py`
refuses it whatever files it holds. To apply a passing export, see
[Automated fresh export](EXCEL_CERTIFICATION.md#automated-fresh-export).

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

## Hardened runner set-up

One machine runs every Excel regression and every export. It has failed in
four ways: the runner lost communication with GitHub, an action download timed
out, the checkout could not resolve `github.com`, and the machine went to
standby (#295). The set-up below removes the causes that are on the machine.
It keeps the interactive session that Excel COM needs.

**Status: written on 2026-11-13, not yet applied or tested on the runner
machine.** Apply it there once and record the date here.

Run the commands in an elevated PowerShell on the runner machine.

1. **Sign in automatically.** Use Sysinternals Autologon (it stores the
   password as an LSA secret, not in clear text) for the dedicated runner
   account. The account should have no rights beyond the runner's needs.
2. **Start the runner at logon and restart it when it exits.** Replace the
   path with the runner directory:

   ```powershell
   $action = New-ScheduledTaskAction -Execute "cmd.exe" `
       -Argument "/c run.cmd" -WorkingDirectory "C:\actions-runner"
   $trigger = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
   $settings = New-ScheduledTaskSettingsSet -RestartCount 999 `
       -RestartInterval (New-TimeSpan -Minutes 1) -ExecutionTimeLimit ([TimeSpan]::Zero) `
       -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
   Register-ScheduledTask -TaskName "GitHub Actions runner (Excel)" -Action $action `
       -Trigger $trigger -Settings $settings -RunLevel Limited
   ```

   The task must run in the interactive session: leave "Run only when user is
   logged on" selected. Do not install the runner as a Windows service.
3. **Never sleep.**

   ```powershell
   powercfg /change standby-timeout-ac 0
   powercfg /change hibernate-timeout-ac 0
   powercfg /hibernate off
   ```

   Keep the machine on mains power, and do not close a laptop lid unless the
   lid action is "Do nothing".
4. **Keep Windows Update out of the waves.** Set active hours to the hours
   when waves run (**Settings > Windows Update > Advanced options > Active
   hours**), and restart the machine deliberately after updates.
5. **Let the runner update itself.** Do not pass `--disableupdate` to
   `config.cmd`. After a GitHub runner release, the first job updates the
   runner and restarts it; the scheduled task of step 2 brings it back.

To test the set-up: restart the machine, do not sign in by hand, and dispatch
the Excel workflow with `suite: fast`. It must start within a few minutes.

### What the repository does about the rest

Network failures cannot be fixed on the machine. The workflow
`.github/workflows/excel-runner-watch.yml` handles them, with the rules in
`tools/excel_runner_watch.py`:

- **One automatic re-run of an infrastructure failure.** When an Excel run
  fails in job set-up or checkout, or GitHub reports a lost runner and no step
  failed, its failed jobs are re-run once. A failure of the regression or
  export step is never re-run automatically, and no run is re-run twice.
- **An alert issue.** Twice an hour a check reads the queue. It opens or
  updates one issue, "[CI]: The Excel runner needs attention", when:
  - a run of a workflow that uses this runner has been queued for more than
    30 minutes with no job created;
  - a job has waited for the runner for more than 60 minutes while nothing
    runs on it, which means the runner is probably offline, or for more than
    180 minutes while another job holds it.

  It closes the issue when nothing is waiting any more. Runs older than 48
  hours are left out.
- **A run queued with no job** cannot be cancelled through the API. Close and
  reopen its pull request, or push a new commit to the branch; cancel the old
  run from the Actions page if it remains listed.

Both jobs run from the default branch, so they act only once this file is on
`main`. The workflow token cannot list runners, so "offline" is inferred from
a job that waits while nothing runs.

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

The export runs only on `workflow_dispatch`, which only accounts with write
access can trigger, and the fork guard above still applies to the job. Export
runs have the same `contents: read` token, no secrets and no step that writes
to the repository; the evidence reaches the repository only through a commit
a maintainer makes after applying the artifact.

## Local execution

From a configured Windows/Excel machine:

```powershell
Set-Location C:\path\to\VBA-PROBABILITY-DISTRIBUTIONS
.\ci\Run-ExcelVbaTests.ps1
```

A local run records `execution = manual` and no GitHub workflow identity. It is
still bound to the exact local `HEAD` and source hashes.

The export adapter runs locally after a green local regression of the same
commit:

```powershell
.\ci\Run-ExcelVbaTests.ps1
.\ci\Export-ExcelGrids.ps1 -Main true -Holdout true
```

`-HoldoutOlder true` adds the older-families holdout, `-SurvivalBoundary true`
the survival-boundary sweep, `-TChiLargeDf true` the large-df t/Chi-square study, `-MainMode missing-only`
selects the missing-only main export and `-MacroTimeoutMinutes` changes the
watchdog's limit. The artifact lands in `artifacts\excel-grid-export`. Its
manifest records no workflow run.

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

Build and check it on any machine with Python 3.11+ and Git; no Excel is
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

### Export stopped by the watchdog

`export-log.txt` and the manifest's `failure` name the macro, the elapsed time
and either the timeout or the dialog's class, title and text. A VBA compile
error dialog means the imported project does not compile; any other dialog
means something outside the exporters prompted (an add-in, an Office notice).
Dismiss its cause on the runner under the runner account and dispatch again.
The watchdog killed only the adapter's own Excel process.

### Export refused or failed verification

The first `EXPORT REFUSED` line in `export-log.txt` gives the reason. The usual
ones are a regression that was not green, an Excel version or build that
differs from the regression record's, a missing-only main export with no blank
cell to fill, and an exporter message other than its success message (for
example, `Accuracy export failed: ...`). A verification failure lists the
first rows that changed outside `observed_vba` or were left blank;
`diagnostics/` holds the files as the exporters left them.

## Operational controls

- Use a dedicated machine or VM.
- Limit the runner account's permissions.
- Keep Windows and Office patched.
- Do not store unrelated credentials in the account profile.
- Review workflow and evidence-policy changes carefully.
- Keep the Excel regression job's pull request status non-required: a skipped
  fork PR job would report success. The release workflow already requires a
  green Excel run of the exact candidate commit; see
  [Release certification](EXCEL_CERTIFICATION.md#exact-tag-certification).
