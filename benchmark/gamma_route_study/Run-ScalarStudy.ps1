param([string]$RepositoryRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$out = Join-Path $RepositoryRoot 'artifacts\gamma-scalar'
$excel = $null
$book = $null
$report = [ordered]@{status='FAIL'; candidate_sha=''; excel_version=''; excel_build=''; rows=@(); timings=@(); error=''; cleanup_error=''; normalization_check=''; arithmetic_check=''; current_case=''}
$inv = [Globalization.CultureInfo]::InvariantCulture
function ConvertFrom-DoubleBits([string]$Hex) {
    if ($Hex -cnotmatch '^[0-9A-F]{16}$') { throw 'Invalid IEEE Double token' }
    $word = [Convert]::ToUInt64($Hex, 16)
    return [BitConverter]::ToDouble([BitConverter]::GetBytes([UInt64]$word), 0)
}
try {
    $report.candidate_sha = (& git -C $RepositoryRoot rev-parse HEAD).Trim()
    & git -C $RepositoryRoot diff --exit-code HEAD -- src benchmark/gamma_route_study
    if ($LASTEXITCODE -ne 0) { throw 'Dirty study inputs' }
    $inputData = Get-Content (Join-Path $out 'inputs.json') -Raw | ConvertFrom-Json
    if ($inputData.study -ne 'B5_scalar' -or $inputData.baseline -ne 'B4') { throw 'Wrong study design' }
    if ($report.candidate_sha -cne $inputData.candidate_sha) { throw 'Candidate SHA mismatch' }
    foreach ($entry in @(@('candidate.txt','candidate_sha256'),@('probe.txt','probe_sha256'))) {
        $actual = (Get-FileHash (Join-Path $out $entry[0]) -Algorithm SHA256).Hash.ToLowerInvariant()
        if ($actual -cne $inputData.($entry[1])) { throw "Study payload hash mismatch: $($entry[0])" }
    }
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $report.excel_version = [string]$excel.Version
    $report.excel_build = [string]$excel.Build
    $book = $excel.Workbooks.Add()
    $book.SaveAs((Join-Path $out 'scalar-study.xlsm'), 52)
    foreach ($file in (Get-ChildItem (Join-Path $RepositoryRoot 'src\*.bas') | Sort-Object Name)) {
        [void]$book.VBProject.VBComponents.Import($file.FullName)
    }
    $module = $book.VBProject.VBComponents.Item('M_STATS_PROBDIST_SPECIALFUNCS').CodeModule
    $module.AddFromString([IO.File]::ReadAllText((Join-Path $out 'candidate.txt')))
    [IO.File]::WriteAllText((Join-Path $out 'M_STATS_PROBDIST_B5.bas'), [IO.File]::ReadAllText((Join-Path $out 'probe.txt')))
    [void]$book.VBProject.VBComponents.Import((Join-Path $out 'M_STATS_PROBDIST_B5.bas'))
    $macro = "'" + $book.Name + "'!B5_Sample"
    $report.normalization_check = [string]$excel.Run("'" + $book.Name + "'!B4_NormalizationSelfTest")
    if ($report.normalization_check -ne "PASS checks=11") { throw "Normalization check failed" }
    $report.arithmetic_check = [string]$excel.Run("'" + $book.Name + "'!B5_ArithmeticSelfTest")
    if ($report.arithmetic_check -ne "PASS checks=576") { throw "Arithmetic check failed: $($report.arithmetic_check)" }
    foreach ($case in $inputData.cases) {
        $report.current_case = $case.id
        $words = $case.id -split ':'
        if ($words.Count -ne 2) { throw 'Invalid exact-input ID' }
        $caseA = ConvertFrom-DoubleBits $words[0]
        $caseX = ConvertFrom-DoubleBits $words[1]
        $raw = [string]$excel.Run($macro, [double]$caseA, [double]$caseX, $false, 0)
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
                $echoA = ConvertFrom-DoubleBits $f[3]
                $echoX = ConvertFrom-DoubleBits $f[4]
                if ($echoA -ne [double]$point[0] -or $echoX -ne [double]$point[1]) { throw 'Timing echo mismatch' }
                $report.timings += [ordered]@{round=$round; a=$point[0]; x=$point[1]; baseline=$baseline; calls=4; p=$f[0]; q=$f[1]; elapsed_us=[double]::Parse($f[2],$inv)}
            }
        }
    }
    $report.status = 'PASS'
} catch { $report.error = $_.Exception.Message; throw }
finally {
    try { if ($null -ne $book) { $book.Close($false) } } catch { $report.cleanup_error += $_.Exception.Message; $report.status='FAIL' }
    try { if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) } } catch { $report.cleanup_error += $_.Exception.Message; $report.status='FAIL' }
    $report | ConvertTo-Json -Depth 12 | Set-Content (Join-Path $out 'scalar-results.json') -Encoding UTF8
    if ($report.cleanup_error) { throw $report.cleanup_error }
}
