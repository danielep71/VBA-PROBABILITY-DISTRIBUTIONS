param([string]$RepositoryRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$out = Join-Path $RepositoryRoot 'artifacts\gamma-pair'
$excel = $null
$book = $null
$report = [ordered]@{status='FAIL'; candidate_sha=''; excel_version=''; excel_build=''; rows=@(); timings=@(); error=''; cleanup_error=''}
$inv = [Globalization.CultureInfo]::InvariantCulture
try {
    $report.candidate_sha = (& git -C $RepositoryRoot rev-parse HEAD).Trim()
    & git -C $RepositoryRoot diff --exit-code HEAD -- src benchmark/gamma_route_study
    if ($LASTEXITCODE -ne 0) { throw 'Dirty study inputs' }
    $inputData = Get-Content (Join-Path $out 'inputs.json') -Raw | ConvertFrom-Json
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $report.excel_version = [string]$excel.Version
    $report.excel_build = [string]$excel.Build
    $book = $excel.Workbooks.Add()
    $book.SaveAs((Join-Path $out 'pair-study.xlsm'), 52)
    foreach ($file in (Get-ChildItem (Join-Path $RepositoryRoot 'src\*.bas') | Sort-Object Name)) {
        [void]$book.VBProject.VBComponents.Import($file.FullName)
    }
    $module = $book.VBProject.VBComponents.Item('M_STATS_PROBDIST_SPECIALFUNCS').CodeModule
    $module.AddFromString([IO.File]::ReadAllText((Join-Path $out 'candidate.txt')))
    [void]$book.VBProject.VBComponents.Import((Join-Path $out 'M_STATS_PROBDIST_B4.bas'))
    $macro = "'" + $book.Name + "'!B4_Sample"
    foreach ($case in $inputData.cases) {
        $raw = [string]$excel.Run($macro, [double]$case.a, [double]$case.x, $false, 0)
        $report.rows += [ordered]@{id=$case.id; raw=$raw}
        $fields = $raw -split ','
        if ($fields[0] -eq 'DECLINED') {
            if ($fields.Count -ne 3 -or ($fields[1]+':'+$fields[2]) -ne $case.id) { throw 'Decline echo mismatch' }
            $observed = 'DECLINED'
        } else {
            if ($fields.Count -ne 5) { throw "Malformed result: $raw" }
            if (($fields[3]+':'+$fields[4]) -ne $case.id) { throw 'Input echo mismatch' }
            $observed = $fields[0]+','+$fields[1]
        }
        if ($observed -ne $case.expected) { throw "Fidelity mismatch $($case.id): $observed vs $($case.expected)" }
        if ($report.rows.Count % 100 -eq 0) { Write-Host "Fidelity $($report.rows.Count) / $($inputData.cases.Count)" }
    }
    $points = @(@(0.0001,0.5),@(0.5,0.25),@(2.5,1.0),@(2.5,6.0),@(100.0,90.0),@(100.0,110.0),@(10000.0,9900.0),@(10000.0,10100.0))
    for ($round=1; $round -le 3; $round++) {
        foreach ($point in $points) {
            $modes = @($true,$false)
            if ($round -eq 2) { $modes = @($false,$true) }
            foreach ($baseline in $modes) {
                $raw = [string]$excel.Run($macro,[double]$point[0],[double]$point[1],$baseline,4)
                $f = $raw -split ','
                if ($f.Count -ne 5) { throw "Invalid timing: $raw" }
                $report.timings += [ordered]@{round=$round; a=$point[0]; x=$point[1]; baseline=$baseline; calls=4; p=$f[0]; q=$f[1]; elapsed_us=[double]::Parse($f[2],$inv)}
            }
        }
    }
    $report.status = 'PASS'
} catch { $report.error = $_.Exception.Message; throw }
finally {
    try { if ($null -ne $book) { $book.Close($false) } } catch { $report.cleanup_error += $_.Exception.Message; $report.status='FAIL' }
    try { if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) } } catch { $report.cleanup_error += $_.Exception.Message; $report.status='FAIL' }
    $report | ConvertTo-Json -Depth 12 | Set-Content (Join-Path $out 'pair-results.json') -Encoding UTF8
    if ($report.cleanup_error) { throw $report.cleanup_error }
}
