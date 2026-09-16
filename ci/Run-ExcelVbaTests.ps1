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
$certificationPath = Join-Path $ArtifactDirectory "excel-certification.json"
$policyPath = Join-Path $RepositoryRoot ".github\excel-evidence-policy.json"

$startedAtUtc = [DateTime]::UtcNow.ToString("o")
$finishedAtUtc = $null
$exitCode = 1
$rawResult = $null
$total = 0
$passed = 0
$failed = 0
$excelVersion = "unavailable"
$excelBuild = "unavailable"
$officeBitness = "unavailable"
$candidateSha = $null
$policy = $null
$sourceInventory = @()
$runReachedExcel = $false

$stages = [ordered]@{
    import = [ordered]@{ status = "NOT_RUN"; detail = "Not started" }
    compile = [ordered]@{ status = "NOT_RUN"; detail = "Not started" }
    regression = [ordered]@{ status = "NOT_RUN"; detail = "Not started" }
    cleanup = [ordered]@{ status = "NOT_RUN"; detail = "Not started" }
}

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

function Get-CanonicalGitBlobSha256 {
    param(
        [Parameter(Mandatory = $true)][string]$Commit,
        [Parameter(Mandatory = $true)][string]$RelativePath
    )

    # Hash the bytes stored in Git, not the CRLF working-tree representation of
    # exported VBA files. .gitattributes deliberately stores normalized LF
    # blobs and checks .bas files out as CRLF.
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
    if (-not $process.Start()) {
        throw "Unable to start git while hashing $RelativePath"
    }

    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $hashBytes = $sha.ComputeHash($process.StandardOutput.BaseStream)
        $stderr = $process.StandardError.ReadToEnd()
        $process.WaitForExit()
        if ($process.ExitCode -ne 0) {
            throw "git show failed for $RelativePath`: $stderr"
        }
        $hex = ($hashBytes | ForEach-Object { $_.ToString("x2") }) -join ""
        return "sha256:$hex"
    }
    finally {
        $sha.Dispose()
        $process.Dispose()
    }
}

function Get-OfficeBitnessFromExecutable {
    param([Parameter(Mandatory = $true)][string]$ExcelExecutable)

    # Read the Portable Executable COFF machine type from EXCEL.EXE. This
    # records Office bitness itself rather than inferring it from runner OS.
    $stream = [System.IO.File]::Open($ExcelExecutable, [System.IO.FileMode]::Open,
                                    [System.IO.FileAccess]::Read,
                                    [System.IO.FileShare]::ReadWrite)
    $reader = New-Object System.IO.BinaryReader($stream)
    try {
        $stream.Position = 0x3C
        $peOffset = $reader.ReadInt32()
        $stream.Position = $peOffset + 4
        $machine = $reader.ReadUInt16()
        switch ($machine) {
            0x014c { return "32-bit" } # IMAGE_FILE_MACHINE_I386
            0x8664 { return "64-bit" } # IMAGE_FILE_MACHINE_AMD64
            0xAA64 { return "64-bit" } # IMAGE_FILE_MACHINE_ARM64
            default { throw ("Unsupported EXCEL.EXE PE machine type 0x{0:X4}" -f $machine) }
        }
    }
    finally {
        $reader.Dispose()
        $stream.Dispose()
    }
}

function Set-FailureStageFromException {
    param([Parameter(Mandatory = $true)][string]$Message)

    if ($stages.import.status -ne "PASS") {
        $stages.import.status = "FAIL"
        $stages.import.detail = "Import/host setup failed: $Message"
    }
    elseif ($stages.compile.status -ne "PASS") {
        $stages.compile.status = "FAIL"
        $stages.compile.detail = "Excel could not execute the imported project: $Message"
    }
    elseif ($stages.regression.status -ne "PASS") {
        $stages.regression.status = "FAIL"
        $stages.regression.detail = "Regression execution failed: $Message"
    }
}

try {
    New-Item -ItemType Directory -Path $ArtifactDirectory -Force | Out-Null
    Set-Content -LiteralPath $resultPath -Value "" -Encoding UTF8
    if (Test-Path -LiteralPath $certificationPath) {
        Remove-Item -LiteralPath $certificationPath -Force
    }
    if (Test-Path -LiteralPath $workbookPath) {
        Remove-Item -LiteralPath $workbookPath -Force
    }

    Write-CiLog "Repository root: $RepositoryRoot"

    if (-not (Test-Path -LiteralPath $policyPath -PathType Leaf)) {
        throw "Excel evidence policy not found: $policyPath"
    }
    $policy = Get-Content -LiteralPath $policyPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ([int]$policy.schema_version -ne 1) {
        throw "Unsupported Excel evidence policy schema: $($policy.schema_version)"
    }

    $candidateSha = (& git -C $RepositoryRoot rev-parse HEAD).Trim()
    if ($LASTEXITCODE -ne 0 -or $candidateSha -notmatch '^[0-9a-f]{40}$') {
        throw "Unable to resolve a full candidate Git SHA"
    }
    if (-not [string]::IsNullOrWhiteSpace($env:GITHUB_SHA) -and $candidateSha -ne $env:GITHUB_SHA) {
        throw "Checked-out candidate $candidateSha differs from GITHUB_SHA $($env:GITHUB_SHA)"
    }

    & git -C $RepositoryRoot diff --quiet $candidateSha --
    if ($LASTEXITCODE -ne 0) {
        throw "Tracked working tree differs from exact candidate $candidateSha"
    }

    Write-CiLog "Candidate SHA: $candidateSha"
    Write-CiLog "Creating isolated Excel workbook for VBA regression execution."

    $sourceFiles = @($policy.regression_sources | ForEach-Object { [string]$_ })
    if ($sourceFiles.Count -le 0) {
        throw "Excel evidence policy declares no regression sources"
    }

    foreach ($relativePath in $sourceFiles) {
        $absolutePath = Join-Path $RepositoryRoot ($relativePath -replace '/', '\')
        if (-not (Test-Path -LiteralPath $absolutePath -PathType Leaf)) {
            throw "Required VBA module not found: $relativePath"
        }
        $sourceInventory += [ordered]@{
            path = $relativePath
            sha256 = Get-CanonicalGitBlobSha256 -Commit $candidateSha -RelativePath $relativePath
        }
    }
    $sourceInventory = @($sourceInventory | Sort-Object { $_.path })

    $excel = New-Object -ComObject Excel.Application
    $runReachedExcel = $true
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false

    # msoAutomationSecurityLow = 1. This applies only to the isolated Excel
    # process created by this script. It does not edit persistent Trust Center
    # settings; VBProject access must already be enabled for the runner account.
    $excel.AutomationSecurity = 1

    $excelVersion = [string]$excel.Version
    $excelBuild = [string]$excel.Build
    $excelExecutable = Join-Path ([string]$excel.Path) "EXCEL.EXE"
    $officeBitness = Get-OfficeBitnessFromExecutable -ExcelExecutable $excelExecutable

    Write-CiLog "Excel version: $excelVersion"
    Write-CiLog "Excel build: $excelBuild"
    Write-CiLog "Office bitness: $officeBitness"

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
        $absolutePath = Join-Path $RepositoryRoot ($relativePath -replace '/', '\')
        Write-CiLog "Importing $relativePath"
        [void]$vbProject.VBComponents.Import($absolutePath)
    }
    $stages.import.status = "PASS"
    $stages.import.detail = "Imported the policy-declared exact candidate source inventory into a fresh workbook"

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
'==============================================================================
'
    On Error GoTo Err_Handler

    BeginRun "ALL SUITES - GITHUB ACTIONS"
    RunCoreSuite
    RunNormalFamilySuite
    RunTFamilySuite
    RunContinuousSuite
    RunDiscreteSuite
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
    Test_STATS_PROBDIST_GetFailureLog_CI = mFailureLog
End Function
'@

    $insertLine = $codeModule.CountOfLines + 1
    $codeModule.InsertLines($insertLine, $ciBridge)

    $workbook.Save()
    Write-CiLog "Executing $($policy.entry_point)"

    $macroName = "'" + $workbook.Name + "'!" + [string]$policy.entry_point
    $rawResult = [string]$excel.Run($macroName)

    # A successful Application.Run means Excel accepted the complete imported
    # project for execution. This is execution-backed compile evidence; it does
    # not pretend that the VBE Debug > Compile command was separately observed.
    $stages.compile.status = "PASS"
    $stages.compile.detail = "Excel executed the imported project through Application.Run; execution-backed compile evidence (not a separately observed Debug > Compile command)"

    Write-CiLog "VBA result: $rawResult"

    if ($rawResult -match '^ERROR=(-?\d+);DESCRIPTION=(.*)$') {
        $stages.regression.status = "FAIL"
        $stages.regression.detail = "VBA test entry point raised error $($Matches[1]): $($Matches[2])"
        throw $stages.regression.detail
    }

    if ($rawResult -notmatch '^TOTAL=(\d+);PASS=(\d+);FAIL=(\d+)$') {
        $stages.regression.status = "FAIL"
        $stages.regression.detail = "Unexpected machine-readable VBA result: $rawResult"
        throw $stages.regression.detail
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

    $expectedAssertions = [int]$policy.expected_assertions
    if ($total -ne $expectedAssertions) {
        $stages.regression.status = "FAIL"
        $stages.regression.detail = "Regression completeness failure: expected $expectedAssertions assertions, observed $total"
        Write-CiLog $stages.regression.detail
        $exitCode = 1
    }
    elseif ($failed -gt 0) {
        $stages.regression.status = "FAIL"
        $stages.regression.detail = "$failed of $total assertions failed"
        $exitCode = 1
    }
    else {
        $stages.regression.status = "PASS"
        $stages.regression.detail = "Complete policy-declared regression: $passed/$total assertions passed"
        $exitCode = 0
    }
}
catch {
    $message = $_.Exception.Message
    Set-FailureStageFromException -Message $message
    if (Test-Path -LiteralPath $resultPath) {
        Write-CiLog "RESULT: CI EXECUTION ERROR"
        Write-CiLog $message
    }
    else {
        Write-Host "RESULT: CI EXECUTION ERROR"
        Write-Host $message
    }
    $exitCode = 1
}
finally {
    $cleanupErrors = New-Object System.Collections.Generic.List[string]

    if ($null -ne $workbook) {
        try {
            $workbook.Close($false)
        }
        catch {
            $cleanupErrors.Add("workbook close: $($_.Exception.Message)")
        }
    }

    if ($null -ne $excel) {
        try {
            $excel.DisplayAlerts = $false
            $excel.Quit()
        }
        catch {
            $cleanupErrors.Add("Excel.Quit: $($_.Exception.Message)")
        }
    }

    try { Release-ComObjectSafely $codeModule } catch { $cleanupErrors.Add("release codeModule: $($_.Exception.Message)") }
    try { Release-ComObjectSafely $testComponent } catch { $cleanupErrors.Add("release testComponent: $($_.Exception.Message)") }
    try { Release-ComObjectSafely $vbProject } catch { $cleanupErrors.Add("release vbProject: $($_.Exception.Message)") }
    try { Release-ComObjectSafely $workbook } catch { $cleanupErrors.Add("release workbook: $($_.Exception.Message)") }
    try { Release-ComObjectSafely $excel } catch { $cleanupErrors.Add("release Excel: $($_.Exception.Message)") }

    try {
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
    catch {
        $cleanupErrors.Add("COM finalization: $($_.Exception.Message)")
    }

    if ($cleanupErrors.Count -eq 0) {
        $stages.cleanup.status = "PASS"
        $stages.cleanup.detail = if ($runReachedExcel) {
            "Closed the owned temporary workbook/Excel instance and released owned COM references"
        } else {
            "No Excel instance was created; no owned host state remained to clean"
        }
    }
    else {
        $stages.cleanup.status = "FAIL"
        $stages.cleanup.detail = ($cleanupErrors -join "; ")
        $exitCode = 1
    }
}

try {
    if (-not (Test-Path -LiteralPath $resultPath)) {
        New-Item -ItemType Directory -Path $ArtifactDirectory -Force | Out-Null
        Set-Content -LiteralPath $resultPath -Value "" -Encoding UTF8
    }

    if ($stages.regression.status -eq "PASS" -and $stages.cleanup.status -eq "PASS") {
        Write-CiLog "RESULT: ALL TESTS PASSED"
    }
    elseif ($stages.regression.status -eq "FAIL") {
        Write-CiLog "RESULT: TEST FAILURE"
    }
    if ($stages.cleanup.status -eq "FAIL") {
        Write-CiLog "RESULT: CLEANUP FAILURE - $($stages.cleanup.detail)"
    }

    $finishedAtUtc = [DateTime]::UtcNow.ToString("o")
    $logDigest = "sha256:" + (Get-FileHash -LiteralPath $resultPath -Algorithm SHA256).Hash.ToLowerInvariant()

    if ($null -eq $candidateSha) {
        $candidateSha = (& git -C $RepositoryRoot rev-parse HEAD).Trim()
    }
    if ($null -eq $policy) {
        $policy = Get-Content -LiteralPath $policyPath -Raw -Encoding UTF8 | ConvertFrom-Json
    }

    $isActions = -not [string]::IsNullOrWhiteSpace($env:GITHUB_ACTIONS)
    if ($isActions) {
        $runnerClass = "self-hosted-excel"
        $runnerIdentity = if ([string]::IsNullOrWhiteSpace($env:RUNNER_NAME)) { [Environment]::MachineName } else { $env:RUNNER_NAME }
        $workflowRepository = if ([string]::IsNullOrWhiteSpace($env:GITHUB_REPOSITORY)) { [string]$policy.repository } else { $env:GITHUB_REPOSITORY }
        $runId = [Int64]$env:GITHUB_RUN_ID
        $runAttempt = [Int64]$env:GITHUB_RUN_ATTEMPT
        $workflow = [ordered]@{
            repository = $workflowRepository
            path = [string]$policy.workflow
            sha = $candidateSha
            run_id = $runId
            run_attempt = $runAttempt
        }
        $execution = "automated"
    }
    else {
        $runnerClass = "manual-interactive"
        $runnerIdentity = [Environment]::MachineName
        $workflow = $null
        $execution = "manual"
    }

    $gridRecords = [ordered]@{}
    foreach ($property in $policy.grid_exporters.PSObject.Properties) {
        $gridRecords[$property.Name] = [ordered]@{
            exported = $false
            sha256 = $null
            row_count = $null
            exported_utc = $null
        }
    }

    $environment = [ordered]@{
        excel_version = $excelVersion
        excel_build = $excelBuild
        office_bitness = $officeBitness
        windows = [System.Environment]::OSVersion.VersionString
        os_architecture = [System.Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString()
        automation_security = "msoAutomationSecurityLow=1 in isolated Excel process; persistent Trust Center settings unchanged"
        vba_project_access = "preconfigured runner setting; script does not modify Trust Center access"
    }

    $record = [ordered]@{
        schema_version = 1
        repository = [string]$policy.repository
        candidate_sha = $candidateSha
        execution = $execution
        started_at = $startedAtUtc
        finished_at = $finishedAtUtc
        runner = [ordered]@{
            class = $runnerClass
            identity = $runnerIdentity
            workflow = $workflow
        }
        environment = $environment
        sources = $sourceInventory
        stages = $stages
        harness = [ordered]@{
            entry_point = [string]$policy.entry_point
            assertions = $total
            passed = $passed
            failed = $failed
        }
        log = [ordered]@{
            path = "test-result.txt"
            sha256 = $logDigest
        }
        grids = $gridRecords
    }

    $json = $record | ConvertTo-Json -Depth 10
    Set-Content -LiteralPath $certificationPath -Value $json -Encoding UTF8
    Write-Host "Excel certification record: $certificationPath"
}
catch {
    $exitCode = 1
    Write-Error "Unable to write Excel certification record: $($_.Exception.Message)"
}

exit $exitCode
