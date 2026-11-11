param([string]$RepositoryRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$out = Join-Path $RepositoryRoot 'artifacts\kernel-profile'
$excel = $null
$workbook = $null
$report = [ordered]@{
    status = 'FAIL'; candidate_sha = ''; excel_version = ''; excel_build = ''
    calibration_us_per_enter_leave = ''; samples = @(); error = ''
}
$invariant = [Globalization.CultureInfo]::InvariantCulture
try {
    $report.candidate_sha = (& git -C $RepositoryRoot rev-parse HEAD).Trim()
    & git -C $RepositoryRoot diff --exit-code HEAD -- src benchmark/kernel_profile
    if ($LASTEXITCODE -ne 0) { throw 'Dirty profile inputs' }
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $report.excel_version = [string]$excel.Version
    $report.excel_build = [string]$excel.Build
    $workbook = $excel.Workbooks.Add()
    $workbook.SaveAs((Join-Path $out 'profile.xlsm'), 52)
    $project = $workbook.VBProject
    foreach ($file in (Get-ChildItem (Join-Path $RepositoryRoot 'src\*.bas') | Sort-Object Name)) {
        [void]$project.VBComponents.Import($file.FullName)
    }
    [void]$project.VBComponents.Import((Join-Path $PSScriptRoot 'M_STATS_PROBDIST_KPROFILE.bas'))
    $special = $project.VBComponents.Item('M_STATS_PROBDIST_SPECIALFUNCS').CodeModule
    $original = $special.Lines(1, $special.CountOfLines)
    $macro = "'" + $workbook.Name + "'!KP_Run"
    $report.calibration_us_per_enter_leave = [string]$excel.Run("'" + $workbook.Name + "'!KP_Calibrate")
    $points = @(@(0.0001, 0.5), @(0.5, 0.25), @(2.5, 1.0), @(2.5, 6.0),
                @(100.0, 90.0), @(100.0, 110.0), @(10000.0, 9900.0), @(10000.0, 10100.0))
    $cases = @()
    foreach ($point in $points) {
        foreach ($route in @('p', 'q')) {
            $cases += [pscustomobject]@{ a = $point[0]; x = $point[1]; route = $route }
        }
    }
    $cases += [pscustomobject]@{ a = 2.0; x = 0.5; route = 'inverse' }
    $cases += [pscustomobject]@{ a = 100.0; x = 0.9; route = 'inverse' }
    $baseline = @{}
    foreach ($mode in @('baseline', 'coarse', 'arithmetic', 'restored')) {
        if ($mode -ne 'baseline') {
            $replacement = $original
            if ($mode -ne 'restored') {
                $replacement = [IO.File]::ReadAllText((Join-Path $out ($mode + '.txt')))
            }
            $special.DeleteLines(1, $special.CountOfLines)
            $special.AddFromString($replacement)
        }
        foreach ($case in $cases) {
            $key = $case.route + ':' + $case.a.ToString($invariant) + ':' + $case.x.ToString($invariant)
            # Fixed work, three samples, warmup excluded. Slow inverse uses one
            # call per sample; forward calls use four. No adaptive selection.
            $n = 4
            if ($case.route -eq 'inverse') { $n = 1 }
            for ($sample = 1; $sample -le 3; $sample++) {
                $raw = [string]$excel.Run($macro, [double]$case.a, [double]$case.x, [string]$case.route, [int]$n)
                $lines = @($raw.Trim() -split "`r?`n")
                $total = $lines[0] -split ','
                if ($total.Count -ne 4 -or $total[0] -ne 'TOTAL') { throw "Invalid output: $raw" }
                $bits = $total[1]
                if (-not $baseline.ContainsKey($key)) {
                    if ($mode -ne 'baseline') { throw 'Missing baseline' }
                    $baseline[$key] = $bits
                }
                if ($bits -ne $baseline[$key]) { throw "Bit mismatch $mode $key $bits vs $($baseline[$key])" }
                $scopes = @()
                foreach ($line in $lines | Select-Object -Skip 1) {
                    $fields = $line -split ','
                    if ($fields.Count -ne 5 -or $fields[0] -ne 'SCOPE') { throw "Invalid scope $line" }
                    $scopes += [ordered]@{
                        id = [int]$fields[1]; calls = [int]$fields[2]
                        inclusive_us = [double]::Parse($fields[3], $invariant)
                        exclusive_us = [double]::Parse($fields[4], $invariant)
                    }
                }
                $row = [ordered]@{
                    mode = $mode; a = $case.a; x = $case.x; route = $case.route
                    sample = $sample; bits = $bits; calls = [int]$total[2]
                    elapsed_us = [double]::Parse($total[3], $invariant); scopes = $scopes
                }
                $report.samples += $row
                Write-Host "$mode $key sample=$sample bits=$bits us_per_call=$($row.elapsed_us / $n)"
                $report | ConvertTo-Json -Depth 12 | Set-Content (Join-Path $out 'profile.json') -Encoding UTF8
            }
        }
    }
    $report.status = 'PASS'
}
catch {
    $report.error = $_.Exception.Message
    throw
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    if ($null -ne $excel) {
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
    $report | ConvertTo-Json -Depth 12 | Set-Content (Join-Path $out 'profile.json') -Encoding UTF8
}
