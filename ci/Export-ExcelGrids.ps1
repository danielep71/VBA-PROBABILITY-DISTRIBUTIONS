#==============================================================================
# Export-ExcelGrids.ps1
#------------------------------------------------------------------------------
# PURPOSE
#   Exports the selected accuracy grids through Excel COM, unattended, from the
#   exact checked-out commit, after ci/Run-ExcelVbaTests.ps1 has produced a
#   green regression record for that commit in the same job.
#
# WHAT IT DOES
#   1. Refuses unless the regression record names this commit (and, in GitHub
#      Actions, this run), every stage passed, every policy assertion passed,
#      and the retained log matches its digest.
#   2. Copies each selected grid, and always the main grid, from the checkout
#      into a private work folder, and saves a fresh macro-enabled workbook in
#      that folder. The exporters look for their grid in the workbook's folder,
#      so they resolve it with no dialog and no change to their source.
#   3. Imports the exact src/ modules and the unchanged exporter modules, plus a
#      CI-only bridge module that is never committed. The bridge defines a
#      Public MsgBox, which takes precedence over VBA.MsgBox for unqualified
#      calls in the project, so the exporters' notifications are recorded
#      instead of blocking. A "Could not locate" prompt makes the bridge raise
#      an error, so the exporter fails before it can open a file picker.
#   4. Runs Export_ExcelEnvironment and each selected exporter under a watchdog
#      that kills only this script's own Excel process, on a visible dialog or
#      at the timeout.
#   5. Checks every exported grid against the checked-out grid: the bytes are
#      what VBA's Print # writes, every in-scope row is filled, nothing outside
#      observed_vba of in-scope rows changed, and ERROR tokens are counted.
#   6. Writes the artifact: the grids exactly as Excel wrote them, the Excel
#      environment file completed with the commit SHA, the regression record
#      and log, the export log and export-manifest.json.
#
# WHAT IT NEVER DOES
#   It never changes a tracked file, never commits or pushes, and never touches
#   an Excel process it did not start.
#
# EXIT CODE
#   0 only when every selected grid was exported and verified. The artifact,
#   including export-manifest.json with status FAIL and the reason, is written
#   in every case.
#==============================================================================

param(
    [Parameter(Mandatory = $false)]
    [string]$RepositoryRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path,

    [Parameter(Mandatory = $false)]
    [string]$ArtifactDirectory = "",

    [Parameter(Mandatory = $false)]
    [string]$RegressionDirectory = "",

    # Workflow inputs arrive as "true"/"false" strings through cmd.
    [Parameter(Mandatory = $false)]
    [string]$Main = "false",

    [Parameter(Mandatory = $false)]
    [string]$Holdout = "false",

    [Parameter(Mandatory = $false)]
    [string]$HoldoutOlder = "false",

    [Parameter(Mandatory = $false)]
    [ValidateSet("full", "missing-only")]
    [string]$MainMode = "full",

    [Parameter(Mandatory = $false)]
    # Four macros at this limit plus the regression stay inside the workflow's
    # 120-minute export timeout, so the watchdog, not the job timeout, stops a
    # hung Excel (a job timeout would leave it running).
    [ValidateRange(1, 240)]
    [int]$MacroTimeoutMinutes = 20
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($ArtifactDirectory)) {
    $ArtifactDirectory = Join-Path $RepositoryRoot "artifacts\excel-grid-export"
}
if ([string]::IsNullOrWhiteSpace($RegressionDirectory)) {
    $RegressionDirectory = Join-Path $RepositoryRoot "artifacts\excel-vba-ci"
}
$ArtifactDirectory = [System.IO.Path]::GetFullPath($ArtifactDirectory)
$RegressionDirectory = [System.IO.Path]::GetFullPath($RegressionDirectory)
$WorkDirectory = [System.IO.Path]::GetFullPath((Join-Path $RepositoryRoot "artifacts\excel-grid-export-work"))

Import-Module (Join-Path $PSScriptRoot "ExcelGridExport.psm1") -Force

$AdapterName = "ci/Export-ExcelGrids.ps1"
$BridgeName = "M_CI_EXPORT_DIALOG_BRIDGE"
$DialogClasses = @("#32770", "bosa_sdm_XL9", "NUIDialog")
$Utf8NoBom = New-Object System.Text.UTF8Encoding($false)

$logPath = Join-Path $ArtifactDirectory "export-log.txt"
$manifestPath = Join-Path $ArtifactDirectory "export-manifest.json"
$workbookPath = Join-Path $WorkDirectory "VBA_Probability_Distributions_Export.xlsm"

$startedAt = [DateTime]::UtcNow
$exitCode = 1
$failure = $null
$candidateSha = $null
$policy = $null
$selected = @()
$gridResults = New-Object System.Collections.Generic.List[object]
$macroTimings = New-Object System.Collections.Generic.List[object]
$importedSources = New-Object System.Collections.Generic.List[object]
$baselines = @{}
$regressionSummary = $null
$excelSummary = [ordered]@{ version = "unavailable"; build = "unavailable"; office_bitness = "unavailable"; process_id = $null }
$environmentSummary = $null
$cleanupDetail = "No Excel instance was created"
$excelKilled = $false

$excel = $null
$workbook = $null
$vbProject = $null
$excelProcessId = 0

#------------------------------------------------------------------------------
# LOGGING
#------------------------------------------------------------------------------
function Write-ExportLog {
    param([string]$Message)

    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
    Write-Host $line
    if (Test-Path -LiteralPath $ArtifactDirectory -PathType Container) {
        [System.IO.File]::AppendAllText($logPath, $line + "`r`n", $Utf8NoBom)
    }
}

function Stop-Export {
    # Fail with a message that names the stage. The catch block records it.
    param([string]$Message)
    throw "EXPORT REFUSED: $Message"
}

#------------------------------------------------------------------------------
# HOST HELPERS (Git, files, Win32)
#------------------------------------------------------------------------------
function Invoke-Git {
    param([Parameter(Mandatory = $true)][string[]]$Arguments)

    # stderr is not redirected: under ErrorActionPreference Stop, Windows
    # PowerShell turns redirected native stderr into a terminating error.
    $output = & git -C $RepositoryRoot @Arguments
    $code = $LASTEXITCODE
    return [pscustomobject]@{ Code = $code; Output = (($output | ForEach-Object { [string]$_ }) -join "`n") }
}

function Get-GitBlobSha256 {
    # SHA-256 of the canonical bytes Git stores, the digest format of the
    # regression record's source inventory.
    param([Parameter(Mandatory = $true)][string]$Commit, [Parameter(Mandatory = $true)][string]$RelativePath)

    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = "git"
    $startInfo.Arguments = "show $Commit`:$RelativePath"
    $startInfo.WorkingDirectory = $RepositoryRoot
    $startInfo.UseShellExecute = $false
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.CreateNoWindow = $true
    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo
    if (-not $process.Start()) { throw "Unable to start git while hashing $RelativePath" }
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $hash = $sha.ComputeHash($process.StandardOutput.BaseStream)
        $stderr = $process.StandardError.ReadToEnd()
        $process.WaitForExit()
        if ($process.ExitCode -ne 0) { throw "git show failed for $RelativePath`: $stderr" }
        return "sha256:" + (($hash | ForEach-Object { $_.ToString("x2") }) -join "")
    }
    finally {
        $sha.Dispose()
        $process.Dispose()
    }
}

function Get-RepositoryPath {
    param([Parameter(Mandatory = $true)][string]$RelativePath)
    return Join-Path $RepositoryRoot ($RelativePath -replace '/', '\')
}

function Get-OfficeBitnessFromExecutable {
    # Same reading as ci/Run-ExcelVbaTests.ps1: the PE machine type of EXCEL.EXE.
    param([Parameter(Mandatory = $true)][string]$ExcelExecutable)

    $stream = [System.IO.File]::Open($ExcelExecutable, [System.IO.FileMode]::Open,
                                    [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    $reader = New-Object System.IO.BinaryReader($stream)
    try {
        $stream.Position = 0x3C
        $peOffset = $reader.ReadInt32()
        $stream.Position = $peOffset + 4
        $machine = $reader.ReadUInt16()
        switch ($machine) {
            0x014c { return "32-bit" }
            0x8664 { return "64-bit" }
            0xAA64 { return "64-bit" }
            default { throw ("Unsupported EXCEL.EXE PE machine type 0x{0:X4}" -f $machine) }
        }
    }
    finally {
        $reader.Dispose()
        $stream.Dispose()
    }
}

function Initialize-Win32 {
    if (([System.Management.Automation.PSTypeName]"CiExcelExportWin32").Type) { return }
    Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

public static class CiExcelExportWin32
{
    private delegate bool EnumProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumProc callback, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(IntPtr parent, EnumProc callback, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr hWnd, StringBuilder name, int capacity);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint msg, IntPtr wParam,
        StringBuilder lParam, uint flags, uint timeout, out IntPtr result);

    private const uint WM_GETTEXT = 0x000D;
    private const uint SMTO_ABORTIFHUNG = 0x0002;

    public static uint ProcessIdOfWindow(long hWnd)
    {
        uint processId;
        GetWindowThreadProcessId(new IntPtr(hWnd), out processId);
        return processId;
    }

    private static string TextOf(IntPtr hWnd)
    {
        StringBuilder text = new StringBuilder(2048);
        IntPtr ignored;
        SendMessageTimeout(hWnd, WM_GETTEXT, new IntPtr(text.Capacity), text, SMTO_ABORTIFHUNG, 1000, out ignored);
        return text.ToString();
    }

    public static string[] VisibleDialogs(uint processId, string[] classes)
    {
        List<string> found = new List<string>();
        EnumWindows(delegate(IntPtr hWnd, IntPtr lParam)
        {
            uint owner;
            GetWindowThreadProcessId(hWnd, out owner);
            if (owner != processId || !IsWindowVisible(hWnd)) { return true; }
            StringBuilder name = new StringBuilder(256);
            GetClassName(hWnd, name, name.Capacity);
            if (Array.IndexOf(classes, name.ToString()) < 0) { return true; }
            StringBuilder body = new StringBuilder();
            EnumChildWindows(hWnd, delegate(IntPtr child, IntPtr childParam)
            {
                string text = TextOf(child);
                if (text.Length > 0)
                {
                    if (body.Length > 0) { body.Append(" | "); }
                    body.Append(text);
                }
                return true;
            }, IntPtr.Zero);
            found.Add(name.ToString() + " '" + TextOf(hWnd) + "': " + body.ToString());
            return true;
        }, IntPtr.Zero);
        return found.ToArray();
    }
}
"@
}

function Start-ExcelWatchdog {
    # Poll the owned Excel process from a second runspace. On a visible dialog
    # or at the deadline, record why and kill that one process, which makes
    # the blocked Application.Run call fail on the main thread.
    param(
        [Parameter(Mandatory = $true)][uint32]$ProcessId,
        [Parameter(Mandatory = $true)][int]$TimeoutSeconds
    )

    $state = [hashtable]::Synchronized(@{ Stop = $false; Reason = $null; Killed = $false; KillError = $null })
    $shell = [PowerShell]::Create()
    [void]$shell.AddScript({
        param($ProcessId, $TimeoutSeconds, $State, $Classes)
        $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
        while (-not $State.Stop) {
            $dialogs = [CiExcelExportWin32]::VisibleDialogs($ProcessId, [string[]]$Classes)
            if ($dialogs.Length -gt 0) {
                $State.Reason = "interactive dialog in the owned Excel process: " + ($dialogs -join " || ")
                break
            }
            if ([DateTime]::UtcNow -ge $deadline) {
                $State.Reason = "no result after $TimeoutSeconds s"
                break
            }
            Start-Sleep -Milliseconds 1000
        }
        if ($null -ne $State.Reason -and -not $State.Stop) {
            try {
                Stop-Process -Id $ProcessId -Force -ErrorAction Stop
                $State.Killed = $true
            }
            catch {
                $State.KillError = $_.Exception.Message
            }
        }
    }).AddArgument($ProcessId).AddArgument($TimeoutSeconds).AddArgument($state).AddArgument($DialogClasses)
    $handle = $shell.BeginInvoke()
    return [pscustomobject]@{ Shell = $shell; Handle = $handle; State = $state }
}

function Stop-ExcelWatchdog {
    param([Parameter(Mandatory = $true)]$Watchdog)

    $Watchdog.State.Stop = $true
    try {
        [void]$Watchdog.Shell.EndInvoke($Watchdog.Handle)
    }
    finally {
        $Watchdog.Shell.Dispose()
    }
}

function Invoke-OwnedMacro {
    # Run one public macro of the temporary workbook under the watchdog and
    # return the dialogs the bridge recorded.
    param([Parameter(Mandatory = $true)][string]$Macro)

    # Everything that runs VBA runs under the watchdog: the first call also
    # compiles the project, and a compile error opens a VBE dialog.
    $watchdog = Start-ExcelWatchdog -ProcessId $excelProcessId -TimeoutSeconds ($MacroTimeoutMinutes * 60)
    $clock = [System.Diagnostics.Stopwatch]::StartNew()
    $runError = $null
    $log = $null
    try {
        [void]$excel.Run("'" + $workbook.Name + "'!CI_ExportBridge_Reset")
        [void]$excel.Run("'" + $workbook.Name + "'!" + $Macro)
        $log = [string]$excel.Run("'" + $workbook.Name + "'!CI_ExportBridge_Log")
    }
    catch {
        $runError = $_.Exception.Message
    }
    finally {
        $clock.Stop()
        Stop-ExcelWatchdog -Watchdog $watchdog
    }
    $seconds = [Math]::Round($clock.Elapsed.TotalSeconds, 2)
    $macroTimings.Add([ordered]@{ macro = $Macro; seconds = $seconds })

    if ($null -ne $watchdog.State.Reason) {
        $script:excelKilled = [bool]$watchdog.State.Killed
        $detail = "$Macro stopped by the watchdog after $seconds s: $($watchdog.State.Reason)"
        if ($watchdog.State.Killed) { $detail += "; killed owned Excel process $excelProcessId" }
        if ($null -ne $watchdog.State.KillError) { $detail += "; kill failed: $($watchdog.State.KillError)" }
        throw $detail
    }
    if ($null -ne $runError) {
        throw "Application.Run $Macro failed after $seconds s: $runError"
    }
    $dialogs = Split-BridgeDialogLog -Log $log
    Write-ExportLog "$Macro finished in $seconds s"
    foreach ($dialog in $dialogs) {
        Write-ExportLog ("  exporter message: " + (Format-DialogText -Dialog $dialog))
    }
    return , $dialogs
}

function Release-ComObjectSafely {
    param([object]$ComObject)
    if ($null -ne $ComObject -and [System.Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
    }
}

$bridgeCode = @'
Option Explicit

'==============================================================================
' M_CI_EXPORT_DIALOG_BRIDGE
'------------------------------------------------------------------------------
' CI-only. ci/Export-ExcelGrids.ps1 adds this module to its temporary export
' workbook; it is never committed. A Public MsgBox in a standard module takes
' precedence over VBA.MsgBox for unqualified calls in this project, so the
' exporters' notifications are recorded here instead of blocking the run.
' A "Could not locate" prompt precedes the exporters' file picker; raising an
' error there sends the exporter to its own error handler instead.
'==============================================================================

Private mLog As String
Private mCount As Long

Public Function MsgBox(ByVal Prompt As Variant, Optional ByVal Buttons As Variant, _
                       Optional ByVal Title As Variant, Optional ByVal HelpFile As Variant, _
                       Optional ByVal Context As Variant) As Long
    Dim TitleText As String
    If Not IsMissing(Title) Then TitleText = CStr(Title)
    If mCount > 0 Then mLog = mLog & Chr$(30)
    mLog = mLog & TitleText & Chr$(31) & CStr(Prompt)
    mCount = mCount + 1
    If InStr(1, CStr(Prompt), "Could not locate", vbBinaryCompare) = 1 Then
        Err.Raise vbObjectError + 1181, "CI export bridge", _
                  "file picker blocked: the grid is not beside the export workbook"
    End If
    MsgBox = 1
End Function

Public Sub CI_ExportBridge_Reset()
    mLog = vbNullString
    mCount = 0
End Sub

Public Function CI_ExportBridge_Log() As String
    CI_ExportBridge_Log = mLog
End Function
'@

#------------------------------------------------------------------------------
# EXPORT
#------------------------------------------------------------------------------
try {
    foreach ($directory in @($ArtifactDirectory, $WorkDirectory)) {
        if (Test-Path -LiteralPath $directory) { Remove-Item -LiteralPath $directory -Recurse -Force }
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }
    Write-ExportLog "Excel grid export adapter: $AdapterName"
    Write-ExportLog "Repository root: $RepositoryRoot"

    $wantMain = ConvertTo-BooleanSwitch -Value $Main -Name "Main"
    $wantHoldout = ConvertTo-BooleanSwitch -Value $Holdout -Name "Holdout"
    $wantHoldoutOlder = ConvertTo-BooleanSwitch -Value $HoldoutOlder -Name "HoldoutOlder"
    if ($wantMain) { $selected += Get-ExportGridSpec -Name "main" -MainMode $MainMode }
    if ($wantHoldout) { $selected += Get-ExportGridSpec -Name "holdout" }
    if ($wantHoldoutOlder) { $selected += Get-ExportGridSpec -Name "holdout_older" }
    if ($selected.Count -eq 0) { Stop-Export "no grid selected; pass -Main, -Holdout or -HoldoutOlder true" }
    Write-ExportLog ("Selected grids: " + (($selected | ForEach-Object { "$($_.Name) ($($_.Macro))" }) -join ", "))
    if (-not $wantMain -and $MainMode -ne "full") {
        Write-ExportLog "Main-grid mode $MainMode is ignored: the main grid is not selected"
    }

    $policyPath = Get-RepositoryPath ".github/excel-evidence-policy.json"
    $policy = Get-Content -LiteralPath $policyPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ([int]$policy.schema_version -ne 1) { Stop-Export "unsupported Excel evidence policy schema $($policy.schema_version)" }

    $head = Invoke-Git -Arguments @("rev-parse", "HEAD")
    $candidateSha = $head.Output.Trim()
    if ($head.Code -ne 0 -or $candidateSha -notmatch '^[0-9a-f]{40}$') { Stop-Export "cannot resolve the full checked-out SHA" }
    if (-not [string]::IsNullOrWhiteSpace($env:GITHUB_SHA) -and $candidateSha -ne $env:GITHUB_SHA) {
        Stop-Export "checked-out $candidateSha differs from GITHUB_SHA $($env:GITHUB_SHA)"
    }
    $clean = Invoke-Git -Arguments @("diff", "--quiet", "HEAD", "--")
    if ($clean.Code -ne 0) { Stop-Export "tracked files differ from $candidateSha; export only from a clean checkout" }
    Write-ExportLog "Candidate SHA: $candidateSha"

    #--------------------------------------------------------------------------
    # 1. The regression record of this commit (and run) must be green.
    #--------------------------------------------------------------------------
    $recordPath = Join-Path $RegressionDirectory "excel-certification.json"
    $regressionLogPath = Join-Path $RegressionDirectory "test-result.txt"
    foreach ($path in @($recordPath, $regressionLogPath)) {
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
            Stop-Export "regression evidence missing: $path; run ci\Run-ExcelVbaTests.ps1 first"
        }
    }
    $recordBytes = [System.IO.File]::ReadAllBytes($recordPath)
    $regressionLogBytes = [System.IO.File]::ReadAllBytes($regressionLogPath)
    $record = $Utf8NoBom.GetString($recordBytes).TrimStart([char]0xFEFF) | ConvertFrom-Json
    $runId = [Int64]0
    if (-not [string]::IsNullOrWhiteSpace($env:GITHUB_RUN_ID)) { $runId = [Int64]$env:GITHUB_RUN_ID }
    $reasons = Test-GreenRegressionRecord -Record $record -CandidateSha $candidateSha `
        -ExpectedAssertions ([int]$policy.expected_assertions) `
        -LogSha256 (Get-Sha256Text -Bytes $regressionLogBytes) -RunId $runId
    if ($reasons.Length -gt 0) {
        foreach ($reason in $reasons) { Write-ExportLog "  regression record: $reason" }
        Stop-Export "the regression record is not green for this commit; no grid is exported"
    }
    [System.IO.File]::WriteAllBytes((Join-Path $ArtifactDirectory "excel-certification.json"), $recordBytes)
    [System.IO.File]::WriteAllBytes((Join-Path $ArtifactDirectory "test-result.txt"), $regressionLogBytes)
    $regressionSummary = [ordered]@{
        record = "excel-certification.json"
        record_sha256 = Get-Sha256Text -Bytes $recordBytes
        log = "test-result.txt"
        assertions = [int]$record.harness.assertions
        passed = [int]$record.harness.passed
        failed = [int]$record.harness.failed
    }
    Write-ExportLog "Regression record is green: $($record.harness.passed)/$($record.harness.assertions) assertions for $candidateSha"

    #--------------------------------------------------------------------------
    # 2. Stage the grids beside the export workbook.
    #--------------------------------------------------------------------------
    $staged = @(Get-ExportGridSpec -Name "main" -MainMode $MainMode)
    foreach ($spec in $selected) {
        if ($spec.Name -ne "main") { $staged += $spec }
    }
    foreach ($spec in $staged) {
        $source = Get-RepositoryPath $spec.Path
        if (-not (Test-Path -LiteralPath $source -PathType Leaf)) { Stop-Export "grid not found: $($spec.Path)" }
        $bytes = [System.IO.File]::ReadAllBytes($source)
        $baselines[$spec.Name] = $bytes
        [System.IO.File]::WriteAllBytes((Join-Path $WorkDirectory $spec.FileName), $bytes)
        Write-ExportLog "Staged $($spec.Path) ($($bytes.Length) bytes, $(Get-Sha256Text -Bytes $bytes))"
    }

    #--------------------------------------------------------------------------
    # 3. Start an owned Excel instance and build the export workbook.
    #--------------------------------------------------------------------------
    Initialize-Win32
    $existing = @(Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $excel = New-Object -ComObject Excel.Application
    $cleanupDetail = "Excel was started"
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    # msoAutomationSecurityLow = 1 in this process only, as the regression adapter.
    $excel.AutomationSecurity = 1

    $excelProcessId = [CiExcelExportWin32]::ProcessIdOfWindow([long]$excel.Hwnd)
    if ($excelProcessId -eq 0) { Stop-Export "cannot identify the owned Excel process" }
    if ($existing -contains [int]$excelProcessId) {
        Stop-Export "Excel COM returned a pre-existing Excel process ($excelProcessId); the watchdog may only kill its own"
    }
    $excelSummary.process_id = [int]$excelProcessId
    $excelSummary.version = [string]$excel.Version
    $excelSummary.build = [string]$excel.Build
    $excelSummary.office_bitness = Get-OfficeBitnessFromExecutable -ExcelExecutable (Join-Path ([string]$excel.Path) "EXCEL.EXE")
    Write-ExportLog "Owned Excel process: $excelProcessId; Excel $($excelSummary.version) build $($excelSummary.build), $($excelSummary.office_bitness)"
    $recorded = $record.environment
    if ($excelSummary.version -cne [string]$recorded.excel_version -or
        $excelSummary.build -cne [string]$recorded.excel_build -or
        $excelSummary.office_bitness -cne [string]$recorded.office_bitness) {
        Stop-Export "this Excel differs from the regression record's ($($recorded.excel_version) build $($recorded.excel_build), $($recorded.office_bitness))"
    }

    $workbook = $excel.Workbooks.Add()
    $workbook.SaveAs($workbookPath, 52) # xlOpenXMLWorkbookMacroEnabled
    try {
        $vbProject = $workbook.VBProject
    }
    catch {
        Stop-Export "Excel denied programmatic access to the VBA project object model (Trust Center): $($_.Exception.Message)"
    }

    $modules = @($policy.regression_sources | ForEach-Object { [string]$_ } | Where-Object { $_.StartsWith("src/") })
    $modules += (Get-ExportGridSpec -Name "main").Exporter
    foreach ($spec in $selected) {
        if ($modules -notcontains $spec.Exporter) { $modules += $spec.Exporter }
    }
    foreach ($relativePath in $modules) {
        $absolutePath = Get-RepositoryPath $relativePath
        if (-not (Test-Path -LiteralPath $absolutePath -PathType Leaf)) { Stop-Export "module not found: $relativePath" }
        Write-ExportLog "Importing $relativePath"
        [void]$vbProject.VBComponents.Import($absolutePath)
        $importedSources.Add([ordered]@{ path = $relativePath; sha256 = Get-GitBlobSha256 -Commit $candidateSha -RelativePath $relativePath })
    }
    $bridge = $vbProject.VBComponents.Add(1) # vbext_ct_StdModule
    $bridge.Name = $BridgeName
    $bridgeModule = $bridge.CodeModule
    # A new module already holds Option Explicit when the VBE option "Require
    # Variable Declaration" is on; a second one would not compile.
    if ($bridgeModule.CountOfLines -gt 0) { $bridgeModule.DeleteLines(1, $bridgeModule.CountOfLines) }
    $bridgeModule.AddFromString($bridgeCode)
    Release-ComObjectSafely $bridgeModule
    Release-ComObjectSafely $bridge
    $workbook.Save()
    Write-ExportLog "Export workbook ready: $workbookPath (CI bridge $BridgeName injected, not committed)"

    #--------------------------------------------------------------------------
    # 4. Run the environment export and the selected exporters.
    #--------------------------------------------------------------------------
    $environmentPath = Join-Path $WorkDirectory "excel_environment.json"
    $dialogs = Invoke-OwnedMacro -Macro "Export_ExcelEnvironment"
    $problems = Test-ExportDialogs -Dialogs $dialogs -SuccessPrefix "Excel environment written to:" -ExpectedPath $environmentPath
    if ($problems.Length -gt 0) { Stop-Export ("Export_ExcelEnvironment: " + ($problems -join "; ")) }

    foreach ($spec in $selected) {
        $gridPath = Join-Path $WorkDirectory $spec.FileName
        $expectedCount = -1
        if ($spec.ReportsCount) {
            $expectedCount = 0
            $baselineLines = ConvertTo-GridLines -Bytes $baselines[$spec.Name]
            for ($k = 1; $k -lt $baselineLines.Length; $k++) {
                if (Test-GridRowInScope -Spec $spec -Line $baselineLines[$k]) { $expectedCount++ }
            }
        }
        $dialogs = Invoke-OwnedMacro -Macro $spec.Macro
        $problems = Test-ExportDialogs -Dialogs $dialogs -SuccessPrefix $spec.SuccessPrefix -ExpectedPath $gridPath -ExpectedCount $expectedCount
        if ($problems.Length -gt 0) { Stop-Export ("$($spec.Macro): " + ($problems -join "; ")) }
    }
}
catch {
    $failure = $_.Exception.Message
    Write-ExportLog "RESULT: EXPORT FAILED"
    Write-ExportLog $failure
}
finally {
    #--------------------------------------------------------------------------
    # Cleanup: close the owned workbook and Excel; never touch other processes.
    #--------------------------------------------------------------------------
    $cleanupErrors = New-Object System.Collections.Generic.List[string]
    if ($null -ne $excel -and -not $excelKilled) {
        # Close and Quit run under the watchdog too: Excel launched through COM
        # is not in the job's process tree, so a dialog here would outlive the
        # job's timeout.
        $closeWatchdog = $null
        if ($excelProcessId -ne 0) {
            $closeWatchdog = Start-ExcelWatchdog -ProcessId $excelProcessId -TimeoutSeconds 120
        }
        try {
            if ($null -ne $workbook) {
                try { $workbook.Close($false) } catch { $cleanupErrors.Add("workbook close: $($_.Exception.Message)") }
            }
            try {
                $excel.DisplayAlerts = $false
                $excel.Quit()
            }
            catch { $cleanupErrors.Add("Excel.Quit: $($_.Exception.Message)") }
        }
        finally {
            if ($null -ne $closeWatchdog) {
                Stop-ExcelWatchdog -Watchdog $closeWatchdog
                if ($null -ne $closeWatchdog.State.Reason) {
                    $excelKilled = [bool]$closeWatchdog.State.Killed
                    $cleanupErrors.Add("closing Excel: $($closeWatchdog.State.Reason)")
                }
            }
        }
    }
    foreach ($object in @($vbProject, $workbook, $excel)) {
        try { Release-ComObjectSafely $object } catch { $cleanupErrors.Add("COM release: $($_.Exception.Message)") }
    }
    $vbProject = $null
    $workbook = $null
    $excel = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    if ($excelProcessId -ne 0) {
        $owned = Get-Process -Id $excelProcessId -ErrorAction SilentlyContinue
        $waited = 0
        while ($null -ne $owned -and -not $owned.HasExited -and $waited -lt 30) {
            Start-Sleep -Seconds 1
            $waited++
            $owned = Get-Process -Id $excelProcessId -ErrorAction SilentlyContinue
        }
        if ($null -ne $owned -and -not $owned.HasExited) {
            Write-ExportLog "WARNING: owned Excel process $excelProcessId still running $waited s after Quit; killing it"
            try { Stop-Process -Id $excelProcessId -Force -ErrorAction Stop } catch { $cleanupErrors.Add("kill owned Excel: $($_.Exception.Message)") }
        }
        if ($excelKilled) {
            $cleanupDetail = "The watchdog killed owned Excel process $excelProcessId"
        }
        else {
            $cleanupDetail = "Closed the owned workbook and Excel process $excelProcessId"
        }
    }
    if ($cleanupErrors.Count -gt 0) {
        # As in the regression adapter, a cleanup failure makes the run non-green.
        $cleanupDetail += "; " + ($cleanupErrors -join "; ")
        Write-ExportLog "RESULT: CLEANUP FAILURE - $($cleanupErrors -join '; ')"
        if ($null -eq $failure) { $failure = "cleanup failed: " + ($cleanupErrors -join "; ") }
    }
}

#------------------------------------------------------------------------------
# 5. Verify every output against the checkout and write the artifact.
#------------------------------------------------------------------------------
try {
    if (-not (Test-Path -LiteralPath $ArtifactDirectory -PathType Container)) {
        New-Item -ItemType Directory -Path $ArtifactDirectory -Force | Out-Null
    }

    if ($null -eq $failure) {
        $rawEnvironmentBytes = [System.IO.File]::ReadAllBytes((Join-Path $WorkDirectory "excel_environment.json"))
        $environment = Read-ExcelEnvironmentFile -Bytes $rawEnvironmentBytes
        if ((ConvertTo-OfficeBitnessLabel -Value $environment.office_bitness) -cne $excelSummary.office_bitness -or
            $environment.excel_version -cne $excelSummary.version -or
            $environment.excel_build -cne $excelSummary.build) {
            Stop-Export "excel_environment.json disagrees with the owned Excel instance"
        }
        $completed = ConvertTo-CompletedEnvironmentJson -Environment $environment -CommitSha $candidateSha -RecordedBy $AdapterName
        $completedBytes = $Utf8NoBom.GetBytes($completed)
        [System.IO.File]::WriteAllBytes((Join-Path $ArtifactDirectory "excel_environment.excel.json"), $rawEnvironmentBytes)
        [System.IO.File]::WriteAllBytes((Join-Path $ArtifactDirectory "excel_environment.json"), $completedBytes)
        $environmentSummary = [ordered]@{
            path = "excel_environment.json"
            sha256 = Get-Sha256Text -Bytes $completedBytes
            excel_written = "excel_environment.excel.json"
            excel_written_sha256 = Get-Sha256Text -Bytes $rawEnvironmentBytes
            commit_sha = $candidateSha
        }
        Write-ExportLog "Environment file completed with commit_sha $candidateSha"

        if (-not ($selected | Where-Object { $_.Name -eq "main" })) {
            $untouched = [System.IO.File]::ReadAllBytes((Join-Path $WorkDirectory "probability_accuracy_grid.csv"))
            if ((Get-Sha256Text -Bytes $untouched) -cne (Get-Sha256Text -Bytes $baselines["main"])) {
                Stop-Export "Export_ExcelEnvironment changed the main grid, which was staged only for path resolution"
            }
        }

        $verificationFailures = New-Object System.Collections.Generic.List[string]
        foreach ($spec in $selected) {
            $exportedBytes = [System.IO.File]::ReadAllBytes((Join-Path $WorkDirectory $spec.FileName))
            $check = Test-ExportedGrid -Spec $spec -BaselineBytes $baselines[$spec.Name] -ExportedBytes $exportedBytes
            $artifactPath = "grids/" + $spec.Path
            $target = Join-Path $ArtifactDirectory ($artifactPath -replace '/', '\')
            New-Item -ItemType Directory -Path (Split-Path -Parent $target) -Force | Out-Null
            [System.IO.File]::WriteAllBytes($target, $exportedBytes)
            $timing = $macroTimings | Where-Object { $_.macro -eq $spec.Macro } | Select-Object -First 1
            $gridResults.Add([ordered]@{
                name = $spec.Name
                path = $spec.Path
                artifact_path = $artifactPath
                mode = $spec.Mode
                macro = $spec.Macro
                exporter = $spec.Exporter
                rows = $check.rows
                in_scope_rows = $check.in_scope_rows
                filled_rows = $check.filled_rows
                error_rows = $check.error_rows
                out_of_scope_rows = $check.out_of_scope_rows
                changed_observations = $check.changed_observations
                baseline_sha256 = Get-Sha256Text -Bytes $baselines[$spec.Name]
                raw_sha256 = Get-Sha256Text -Bytes $exportedBytes
                seconds = $timing.seconds
                verified = ($check.problems.Length -eq 0)
                problems = @($check.problems)
            })
            Write-ExportLog ("{0}: {1} rows, {2} in scope, {3} filled, {4} ERROR, {5} observation(s) changed, {6} out of scope unchanged" -f `
                $spec.Path, $check.rows, $check.in_scope_rows, $check.filled_rows, $check.error_rows, $check.changed_observations, $check.out_of_scope_rows)
            foreach ($problem in $check.problems) {
                Write-ExportLog "  $($spec.Path): $problem"
                $verificationFailures.Add("$($spec.Path): $problem")
            }
        }
        if ($verificationFailures.Count -gt 0) {
            Stop-Export ("the exported grids failed verification: " + ($verificationFailures -join "; "))
        }
    }
}
catch {
    $failure = $_.Exception.Message
    Write-ExportLog "RESULT: EXPORT VERIFICATION FAILED"
    Write-ExportLog $failure
}

if ($null -ne $failure) {
    # Keep what the exporters left behind for diagnosis; it is never a grid
    # that apply_excel_export.py accepts, because the manifest says FAIL.
    $diagnostics = Join-Path $ArtifactDirectory "diagnostics"
    New-Item -ItemType Directory -Path $diagnostics -Force | Out-Null
    Get-ChildItem -LiteralPath $WorkDirectory -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -in @(".csv", ".json") } |
        ForEach-Object { Copy-Item -LiteralPath $_.FullName -Destination $diagnostics -Force }
}
else {
    $exitCode = 0
    Write-ExportLog "RESULT: EXPORT PASSED"
}

$finishedAt = [DateTime]::UtcNow
try {
    $workflow = $null
    if (-not [string]::IsNullOrWhiteSpace($env:GITHUB_ACTIONS)) {
        $workflow = [ordered]@{
            repository = $env:GITHUB_REPOSITORY
            path = if ($null -ne $policy) { [string]$policy.workflow } else { $null }
            ref = $env:GITHUB_REF
            run_id = [Int64]$env:GITHUB_RUN_ID
            run_attempt = [Int64]$env:GITHUB_RUN_ATTEMPT
            runner = $env:RUNNER_NAME
        }
    }
    $manifest = [ordered]@{
        schema_version = 1
        kind = "excel-grid-export"
        adapter = $AdapterName
        repository = if ($null -ne $policy) { [string]$policy.repository } else { $null }
        candidate_sha = $candidateSha
        status = if ($exitCode -eq 0) { "PASS" } else { "FAIL" }
        failure = $failure
        requested = @($selected | ForEach-Object { $_.Name })
        main_mode = $MainMode
        workflow = $workflow
        regression = $regressionSummary
        excel = $excelSummary
        environment = $environmentSummary
        sources = @($importedSources.ToArray())
        grids = @($gridResults.ToArray())
        macros = @($macroTimings.ToArray())
        cleanup = $cleanupDetail
        started_at = $startedAt.ToString("o")
        finished_at = $finishedAt.ToString("o")
        total_seconds = [Math]::Round(($finishedAt - $startedAt).TotalSeconds, 2)
    }
    $json = $manifest | ConvertTo-Json -Depth 10
    [System.IO.File]::WriteAllText($manifestPath, $json + "`n", $Utf8NoBom)
    Write-Host "Export manifest: $manifestPath"
}
catch {
    $exitCode = 1
    Write-Error "Unable to write the export manifest: $($_.Exception.Message)"
}

exit $exitCode
