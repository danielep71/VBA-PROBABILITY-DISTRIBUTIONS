param([string]$RepositoryRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$out = Join-Path $RepositoryRoot 'artifacts\gamma-cache'
$excel = $null
$workbook = $null
$report = [ordered]@{ status = 'FAIL'; candidate_sha = ''; excel_version = ''; excel_build = ''; private_checks = @{}; samples = @(); error = '' }
$invariant = [Globalization.CultureInfo]::InvariantCulture
try {
    $report.candidate_sha = (& git -C $RepositoryRoot rev-parse HEAD).Trim()
    & git -C $RepositoryRoot diff --exit-code HEAD -- src benchmark/gamma_cache_study
    if ($LASTEXITCODE -ne 0) { throw 'Dirty study inputs' }
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $report.excel_version = [string]$excel.Version
    $report.excel_build = [string]$excel.Build
    $workbook = $excel.Workbooks.Add()
    $workbook.SaveAs((Join-Path $out 'cache-study.xlsm'), 52)
    $project = $workbook.VBProject
    foreach ($file in (Get-ChildItem (Join-Path $RepositoryRoot 'src\*.bas') | Sort-Object Name)) {
        [void]$project.VBComponents.Import($file.FullName)
    }
    [void]$project.VBComponents.Import((Join-Path $PSScriptRoot 'M_STATS_PROBDIST_GCACHE.bas'))
    $special = $project.VBComponents.Item('M_STATS_PROBDIST_SPECIALFUNCS').CodeModule
    $macro = "'" + $workbook.Name + "'!GC_Sample"
    $points = @(@(0.0001, 0.5), @(0.5, 0.25), @(2.5, 1.0), @(2.5, 6.0),
                @(100.0, 90.0), @(100.0, 110.0), @(10000.0, 9900.0), @(10000.0, 10100.0))
    $cases = @()
    foreach ($point in $points) {
        foreach ($route in @('p', 'q')) { $cases += [pscustomobject]@{ a = $point[0]; x = $point[1]; route = $route } }
    }
    $cases += [pscustomobject]@{ a = 2.0; x = 0.5; route = 'inverse' }
    $cases += [pscustomobject]@{ a = 100.0; x = 0.9; route = 'inverse' }
    $bitsByCase = @{}
    for ($sample = 1; $sample -le 3; $sample++) {
        foreach ($case in $cases) {
            $key = $case.route + ':' + $case.a.ToString($invariant) + ':' + $case.x.ToString($invariant)
            $modes = @('uncached', 'cached')
            if ($sample -eq 2) { $modes = @('cached', 'uncached') }
            foreach ($mode in $modes) {
                $special.DeleteLines(1, $special.CountOfLines)
                $special.AddFromString([IO.File]::ReadAllText((Join-Path $out ($mode + '.txt'))))
                if (-not $report.private_checks.ContainsKey($mode)) {
                    $check = [string]$excel.Run("'" + $workbook.Name + "'!GC_PrivateChecks")
                    if ($check -notmatch '^PASS checks=40 preserved_error=[1-9][0-9]*$') { throw "Private checks failed: $check" }
                    $report.private_checks[$mode] = $check
                }
                foreach ($cold in @($true, $false)) {
                    $raw = [string]$excel.Run($macro, [double]$case.a, [double]$case.x, [string]$case.route, [bool]$cold)
                    $fields = $raw -split ','
                    if ($fields.Count -ne 3) { throw "Invalid sample: $raw" }
                    if (-not $bitsByCase.ContainsKey($key)) {
                        if ($mode -ne 'uncached') { throw 'Missing uncached reference' }
                        $bitsByCase[$key] = $fields[0]
                    }
                    if ($fields[0] -ne $bitsByCase[$key]) { throw "Bit mismatch: $key $mode cold=$cold" }
                    $row = [ordered]@{ sample = $sample; a = $case.a; x = $case.x; route = $case.route; mode = $mode; cold = $cold; bits = $fields[0]; calls = [int]$fields[1]; elapsed_us = [double]::Parse($fields[2], $invariant) }
                    $report.samples += $row
                    Write-Host "$key $mode cold=$cold sample=$sample bits=$($row.bits) us_per_call=$($row.elapsed_us / $row.calls)"
                    $report | ConvertTo-Json -Depth 12 | Set-Content (Join-Path $out 'cache-results.json') -Encoding UTF8
                }
            }
        }
    }
    $report.status = 'PASS'
}
catch { $report.error = $_.Exception.Message; throw }
finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    $report | ConvertTo-Json -Depth 12 | Set-Content (Join-Path $out 'cache-results.json') -Encoding UTF8
}
