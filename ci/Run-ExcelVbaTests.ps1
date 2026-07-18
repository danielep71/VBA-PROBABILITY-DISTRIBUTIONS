param(
    [Parameter(Mandatory = $false)]
    [string]$RepositoryRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path,

    [Parameter(Mandatory = $false)]
    [string]$ArtifactDirectory = (Join-Path $RepositoryRoot "artifacts\excel-vba-ci")
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$excel = $null
$workbook = $null
$vbProject = $null
$testComponent = $null
$codeModule = $null

$workbookPath = Join-Path $ArtifactDirectory "VBA_Probability_Distributions_CI.xlsm"
$resultPath = Join-Path $ArtifactDirectory "test-result.txt"

function Write-CiLog {
    param([string]$Message)

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$timestamp] $Message"
    Write-Host $line
    Add-Content -LiteralPath $resultPath -Value $line -Encoding UTF8
}

function Release-ComObjectSafely {
    param([object]$ComObject)

    if ($null -ne $ComObject -and [System.Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
    }
}

try {
    New-Item -ItemType Directory -Path $ArtifactDirectory -Force | Out-Null
    Set-Content -LiteralPath $resultPath -Value "" -Encoding UTF8

    if (Test-Path -LiteralPath $workbookPath) {
        Remove-Item -LiteralPath $workbookPath -Force
    }

    Write-CiLog "Repository root: $RepositoryRoot"
    Write-CiLog "Creating isolated Excel workbook for VBA regression execution."

    $sourceFiles = @(
        "src\M_STATS_PROBDIST_CORE.bas",
        "src\M_STATS_PROBDIST_SPECIALFUNCS.bas",
        "src\M_STATS_PROBDIST_NORMALFAMILY.bas",
        "src\M_STATS_PROBDIST_TFAMILY.bas",
        "src\M_STATS_PROBDIST_CONTINUOUS.bas",
        "tests\M_STATS_PROBDIST_TEST.bas"
    )

    foreach ($relativePath in $sourceFiles) {
        $absolutePath = Join-Path $RepositoryRoot $relativePath
        if (-not (Test-Path -LiteralPath $absolutePath -PathType Leaf)) {
            throw "Required VBA module not found: $relativePath"
        }
    }

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false

    # msoAutomationSecurityLow = 1. This applies only to the isolated Excel
    # instance created by this script. The runner must still allow trusted
    # programmatic access to the VBA project object model.
    $excel.AutomationSecurity = 1

    Write-CiLog ("Excel version: " + $excel.Version)
    Write-CiLog ("Excel build: " + $excel.Build)

    $workbook = $excel.Workbooks.Add()
    $workbook.SaveAs($workbookPath, 52) # xlOpenXMLWorkbookMacroEnabled

    try {
        $vbProject = $workbook.VBProject
    }
    catch {
        throw @"
Excel denied programmatic access to the VBA project object model.
On the self-hosted runner, open Excel and enable:
File > Options > Trust Center > Trust Center Settings > Macro Settings >
Trust access to the VBA project object model.
Original error: $($_.Exception.Message)
"@
    }

    foreach ($relativePath in $sourceFiles) {
        $absolutePath = Join-Path $RepositoryRoot $relativePath
        Write-CiLog "Importing $relativePath"
        [void]$vbProject.VBComponents.Import($absolutePath)
    }

    $testComponent = $vbProject.VBComponents.Item("M_STATS_PROBDIST_TEST")
    $codeModule = $testComponent.CodeModule

    # Inject the CI bridge into the same module as the private counters and suite
    # drivers. It is not committed to the production .bas file and exists only
    # in this temporary workbook.
    $ciBridge = @'

Public Function Test_STATS_PROBDIST_RunAll_CI() As String
'
'==============================================================================
' Test_STATS_PROBDIST_RunAll_CI
'------------------------------------------------------------------------------
' PURPOSE
'   Executes the complete regression suite and returns machine-readable counters
'   to the PowerShell/COM GitHub Actions runner.
'
' NOTES
'   This procedure is injected at CI runtime into M_STATS_PROBDIST_TEST so it can
'   read the module-private assertion counters without widening their production
'   scope.
'==============================================================================
'
    On Error GoTo Err_Handler

    BeginRun "ALL SUITES - GITHUB ACTIONS"
    RunCoreSuite
    RunNormalFamilySuite
    RunTFamilySuite
    RunContinuousSuite
    EndRun

    Test_STATS_PROBDIST_RunAll_CI = _
        "TOTAL=" & CStr(mTestCount) & _
        ";PASS=" & CStr(mPassCount) & _
        ";FAIL=" & CStr(mFailCount)
    Exit Function

Err_Handler:
    Test_STATS_PROBDIST_RunAll_CI = _
        "ERROR=" & CStr(Err.Number) & ";DESCRIPTION=" & Err.Description
End Function


Public Function Test_STATS_PROBDIST_GetFailureLog_CI() As String
'
'==============================================================================
' Test_STATS_PROBDIST_GetFailureLog_CI
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the accumulated failure lines from the most recent run so the CI
'   runner can surface each failed assertion's name, actual, expected and
'   tolerance, not just the failure count. Read after Test_STATS_PROBDIST_RunAll_CI
'   in the same session; it does not reset the buffer.
'==============================================================================
'
    Test_STATS_PROBDIST_GetFailureLog_CI = mFailureLog
End Function
'@

    $insertLine = $codeModule.CountOfLines + 1
    $codeModule.InsertLines($insertLine, $ciBridge)

    $workbook.Save()
    Write-CiLog "Executing Test_STATS_PROBDIST_RunAll_CI"

    $macroName = "'" + $workbook.Name + "'!Test_STATS_PROBDIST_RunAll_CI"
    $rawResult = [string]$excel.Run($macroName)
    Write-CiLog "VBA result: $rawResult"

    if ($rawResult -match '^ERROR=(-?\d+);DESCRIPTION=(.*)$') {
        throw "The VBA test entry point raised error $($Matches[1]): $($Matches[2])"
    }

    if ($rawResult -notmatch '^TOTAL=(\d+);PASS=(\d+);FAIL=(\d+)$') {
        throw "Unexpected machine-readable VBA result: $rawResult"
    }

    $total = [int]$Matches[1]
    $passed = [int]$Matches[2]
    $failed = [int]$Matches[3]

    Write-CiLog "Assertions executed: $total"
    Write-CiLog "Assertions passed: $passed"
    Write-CiLog "Assertions failed: $failed"

    if ($failed -gt 0) {
        $logMacro = "'" + $workbook.Name + "'!Test_STATS_PROBDIST_GetFailureLog_CI"
        $failureLog = [string]$excel.Run($logMacro)

        if ([string]::IsNullOrWhiteSpace($failureLog)) {
            Write-CiLog "Failed assertions reported but the failure log was empty."
        }
        else {
            Write-CiLog "Failed assertions:"
            foreach ($line in ($failureLog -split "`r?`n")) {
                if (-not [string]::IsNullOrWhiteSpace($line)) {
                    Write-CiLog "  $line"
                }
            }
        }
    }

    if ($total -le 0) {
        throw "The VBA harness reported zero executed assertions."
    }

    if (($passed + $failed) -ne $total) {
        throw "Inconsistent VBA counters: PASS + FAIL does not equal TOTAL."
    }

    if ($failed -gt 0) {
        Write-CiLog "RESULT: TEST FAILURE"
        exit 1
    }

    Write-CiLog "RESULT: ALL TESTS PASSED"
    exit 0
}
catch {
    $message = $_.Exception.Message
    Write-CiLog "RESULT: CI EXECUTION ERROR"
    Write-CiLog $message
    Write-Error $message
    exit 1
}
finally {
    if ($null -ne $workbook) {
        try {
            $workbook.Close($false)
        }
        catch {
            Write-Warning "Unable to close the temporary workbook cleanly: $($_.Exception.Message)"
        }
    }

    if ($null -ne $excel) {
        try {
            $excel.DisplayAlerts = $false
            $excel.Quit()
        }
        catch {
            Write-Warning "Unable to quit Excel cleanly: $($_.Exception.Message)"
        }
    }

    Release-ComObjectSafely $codeModule
    Release-ComObjectSafely $testComponent
    Release-ComObjectSafely $vbProject
    Release-ComObjectSafely $workbook
    Release-ComObjectSafely $excel

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
