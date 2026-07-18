# Excel/VBA Regression CI

This repository can execute the complete VBA regression harness through Microsoft Excel on a self-hosted Windows GitHub Actions runner.

The workflow is defined in `.github/workflows/excel-vba-regression.yml` and invokes `ci/Run-ExcelVbaTests.ps1`.

## Requirements

- A dedicated Windows machine or VM
- Desktop Microsoft Excel installed and activated
- A self-hosted GitHub Actions runner labeled `excel`
- Excel Trust Center setting **Trust access to the VBA project object model** enabled for the runner account
- An interactive logged-in Windows session for reliable Office COM automation

GitHub-hosted `windows-latest` runners do not include Excel, so this workflow cannot use a standard hosted runner.

## Execution model

For each run, the PowerShell script:

1. creates an isolated Excel COM instance;
2. creates a temporary macro-enabled workbook;
3. imports the current modules from `src/`;
4. imports `tests/M_STATS_PROBDIST_TEST.bas`;
5. injects a CI-only bridge into the test module;
6. executes all suites in dependency order;
7. reads the private assertion counters from inside the same module;
8. maps the failure count to the PowerShell process exit code;
9. writes `artifacts/excel-vba-ci/test-result.txt`.

The runtime bridge is not committed to the production test module.

## Runner setup

1. In GitHub, open **Settings > Actions > Runners > New self-hosted runner**.
2. Install the runner under a dedicated Windows account.
3. Add the custom label `excel`.
4. Open Excel once under that account and complete all first-run, activation, privacy and update prompts.
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

Microsoft does not recommend unattended Office automation from a non-interactive Windows service. An interactive dedicated session is materially more reliable for Excel COM execution.

## Security model

Self-hosted runners execute repository code with the runner account's permissions. The workflow therefore does not execute pull requests from forks:

```yaml
if: >-
  github.event_name != 'pull_request' ||
  github.event.pull_request.head.repo.full_name == github.repository
```

For an external contribution, review the changes and copy or cherry-pick them to a maintainer-controlled branch before running the Excel workflow.

Do not switch this workflow to `pull_request_target` for untrusted code.

## Local execution

From a configured Windows/Excel machine:

```powershell
Set-Location C:\path\to\VBA-PROBABILITY-DISTRIBUTIONS
.\ci\Run-ExcelVbaTests.ps1
```

## Result contract

The injected VBA function returns:

```text
TOTAL=<count>;PASS=<count>;FAIL=<count>
```

The PowerShell process exits with code `0` only when all assertions pass. Missing modules, VBA compilation failures, Excel automation failures, invalid result counters and test failures all produce exit code `1`.

## Troubleshooting

### Programmatic access denied

Enable **Trust access to the VBA project object model** under the exact Windows account running the self-hosted runner.

### Excel COM class not registered

Confirm that desktop Excel is installed and activated. Repair Office if `New-Object -ComObject Excel.Application` fails.

### Workflow remains queued

Confirm that the runner is online and has all labels:

```text
self-hosted
Windows
X64
excel
```

### Orphaned Excel process

Check for modal Excel dialogs, first-run prompts, add-ins, Protected View prompts or Office update dialogs. The script closes the temporary workbook, calls `Excel.Quit()`, releases COM objects and forces finalization, but a modal dialog can still block Office automation.

## Operational controls

- Use a dedicated machine or VM.
- Limit the runner account's permissions.
- Keep Windows and Office patched.
- Do not store unrelated credentials in the account profile.
- Review workflow changes carefully.
- Add the Excel regression job as a required branch-protection check after the runner is stable.
