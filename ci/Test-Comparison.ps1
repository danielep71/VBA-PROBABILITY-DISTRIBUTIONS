param([Parameter(Mandatory=$true)][string]$PackageDirectory)
# The Comparison sheet (Excel vs VBA vs Python) against its cases
# (examples/demo/comparison-cases.json: the benchmark/excel_comparison study, with mpmath
# references at 50 digits and stored SciPy values). For every case: the live Excel and library
# formulas are the study's, the correct digits the sheet shows are the ones computed here from the
# same values, the grades follow the thresholds, the library and SciPy are reference grade, and
# the detail rows show the selected case. Also: the layout the plan promises, the chart, other
# thresholds, an invalid threshold, the =PY() cell with a number and with text, and the reset.
# The report records what this Excel returned for every case, with its version and build.
$ErrorActionPreference = 'Stop'
$directory = (Resolve-Path $PackageDirectory).Path
$path = Join-Path $directory 'Probability-Distributions-Demo.xlsm'
$fixture = Get-Content (Join-Path $PSScriptRoot '../examples/demo/comparison-cases.json') -Raw | ConvertFrom-Json
$plan = Get-Content (Join-Path $directory 'comparison-plan.json') -Raw | ConvertFrom-Json
$report = [ordered]@{ status = 'FAIL'; sha256 = (Get-FileHash $path -Algorithm SHA256).Hash; cases = 0; observations = @(); notes = @() }
$excel = $null
$book = $null
$invariant = [Globalization.CultureInfo]::InvariantCulture
$full = [double]$plan.full_digits

function Is-Error($Value) { return ($Value -is [int]) -and ($Value -lt -2146820000) }
# One call site per value type: PowerShell binds a COM setter to the first type it sees.
function Set-Text($Sheet, [string]$Address, [string]$Value) { $Sheet.Range($Address).Value2 = $Value }
function Set-Number($Sheet, [string]$Address, [double]$Value) { $Sheet.Range($Address).Value2 = $Value }
# Correct digits as the sheet defines them: -log10 of the relative difference, from 0 to 16.
function Digits($Value, [double]$Reference) {
    if (-not ($Value -is [double])) { return $null }
    if ($Reference -eq 0) { if ($Value -eq 0) { return $full } else { return 0.0 } }
    $relative = [Math]::Abs($Value - $Reference) / [Math]::Abs($Reference)
    if ($relative -eq 0) { return $full }
    return [Math]::Min($full, [Math]::Max(0.0, -[Math]::Log10($relative)))
}
# Excel rewrites the numbers of a formula (1E-15 becomes 0.000000000000001): formulas are
# compared with every number written as its value.
function Same-Formula([string]$Actual, [string]$Expected) {
    $pattern = '(?<![A-Za-z_.\d])\d+(\.\d+)?(E[+-]?\d+)?'
    $evaluator = { param($found) [double]::Parse($found.Value, $invariant).ToString('R', $invariant) }
    return ([regex]::Replace($Actual, $pattern, $evaluator) -eq [regex]::Replace($Expected, $pattern, $evaluator))
}
function Close([double]$Actual, [double]$Expected) {
    return [Math]::Abs($Actual - $Expected) -le 1E-15 * [Math]::Abs($Expected)
}
function Grade($Value, $Digits, [double]$ReferenceGrade, [double]$ReportGrade) {
    if ($Value -is [string]) { return 'no function' }
    if (Is-Error $Value) { return 'error' }
    if ($Digits -le 0) { return 'wrong' }
    if ($Digits -ge $ReferenceGrade) { return 'reference' }
    if ($Digits -ge $ReportGrade) { return 'report' }
    return 'inadequate'
}
function Shown($Value) {
    if ($Value -is [double]) { return $Value.ToString('R', $invariant) }
    if (Is-Error $Value) { return 'error ' + $Value }
    return [string]$Value
}
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($path, 0, $false)
    # Excel may still be loading or calculating the sheets after the workbook opens: a COM call is rejected until it is ready.
    for ($wait = 0; $wait -lt 240; $wait++) { try { if ($excel.Ready) { break } } catch { }; Start-Sleep -Milliseconds 500 }
    $sheet = $book.Worksheets.Item('Comparison')
    $prefix = "'" + $book.Name + "'!"
    $report.excel_version = $excel.Version
    $report.excel_build = $excel.Build
    $report.excel_os = $excel.OperatingSystem
    $first = [int]$plan.layout.first_row
    $last = [int]$plan.layout.last_row
    $rows = $plan.layout.detail_rows

    # Structure: no implicit intersection, the promised cells, locked outputs, unlocked inputs, messages, the chart.
    foreach ($formula in $sheet.Range("A1:O$last").Formula2) {
        if ([string]$formula -like '=*' -and [string]$formula -match '@|_xlfn.SINGLE') { throw "Implicit intersection remains: $formula" }
    }
    if ($sheet.Range('D6').Validation.Formula1 -ne '=ComparisonLabels') { throw 'The case list is not backed by the table' }
    if ($sheet.Range('D6:D9').Locked) { throw 'D6:D9 must be editable' }
    foreach ($address in $plan.layout.outputs) {
        if (-not $sheet.Range($address).Locked) { throw "$address must be locked" }
    }
    foreach ($row in 6..9) {
        $rule = $sheet.Range("D$row").Validation
        foreach ($text in @($rule.InputTitle, $rule.InputMessage)) {
            if ([string]::IsNullOrWhiteSpace([string]$text)) { throw "D$row lacks an input message" }
        }
    }
    # The conditional format on the Excel grade: red for wrong, amber for no function, nothing else.
    $sheet.Activate()
    $gradeRules = $sheet.Range([string]$plan.layout.excel_grades).FormatConditions
    if ($gradeRules.Count -ne 2) { throw "The Excel grade column has $($gradeRules.Count) conditional rules, expected 2" }
    for ($index = 0; $index -lt $fixture.cases.Count; $index++) {
        $row = $first + $index
        $shown = [int]$sheet.Range("M$row").DisplayFormat.Interior.Color
        $grade = [string]$sheet.Range("M$row").Value2
        $expectedFill = switch ($grade) { 'wrong' { 0xCEC7FF } 'no function' { 0x9CEBFF } default { $null } }
        if ($null -ne $expectedFill -and $shown -ne $expectedFill) { throw "Excel grade '$grade' in M$row shows the fill $shown, expected $expectedFill" }
        if ($null -eq $expectedFill -and ($shown -eq 0xCEC7FF -or $shown -eq 0x9CEBFF)) { throw "Excel grade '$grade' in M$row is highlighted" }
    }
    if ($sheet.ChartObjects().Count -ne 1) { throw "The sheet has $($sheet.ChartObjects().Count) charts, expected 1" }
    $chart = $sheet.ChartObjects([string]$plan.chart.name).Chart
    if ($chart.SeriesCollection().Count -ne 3) { throw 'The chart does not have three series' }
    foreach ($index in 1..3) {
        if ($chart.SeriesCollection($index).Points().Count -ne $fixture.cases.Count) { throw "Series $index of the chart does not have a bar per case" }
    }
    foreach ($rule in @($sheet.Cells.FormatConditions)) {
        if ($rule.Type -eq 2 -and [string]$rule.Formula1 -match 'Explorer') { throw "A conditional rule reads the Explorer: $($rule.Formula1)" }
    }
    foreach ($other in @(@('Explorer', 2), @('Probabilities', 1), @('Parameters', 2))) {
        if ($book.Worksheets.Item($other[0]).ChartObjects().Count -ne $other[1]) { throw "The $($other[0]) sheet lost a chart" }
    }
    if (-not ([string]$sheet.Range("D$($rows.excel_build)").Value2).Contains('build')) { throw "The Excel build reads: $($sheet.Range("D$($rows.excel_build)").Value2)" }
    if ([string]$sheet.Range('D19').Value2 -ne 'OK') { throw "Status at the defaults: $($sheet.Range('D19').Value2)" }
    $report.structure = 'PASS'

    # Every case: the formulas, the values, the digits and the grades of the table.
    $excel.Calculate()
    for ($index = 0; $index -lt $fixture.cases.Count; $index++) {
        $case = $fixture.cases[$index]
        $row = $first + $index
        $where = "$($case.label) (row $row)"
        if ([string]$sheet.Range("C$row").Value2 -ne [string]$case.label) { throw "${where}: the label reads $($sheet.Range("C$row").Value2)" }
        # The stored numbers may differ from the decimal texts in their last bit (the JSON and the
        # parsing of PowerShell 5): within 1E-15. The digits are then computed, as on the sheet,
        # against the reference held in the cell.
        $reference = [double]$sheet.Range("D$row").Value2
        if (-not (Close $reference ([double]::Parse([string]$case.reference, $invariant)))) { throw "${where}: the reference is $reference, expected $($case.reference)" }
        if (-not (Close ([double]$sheet.Range("H$row").Value2) ([double]::Parse([string]$case.scipy, $invariant)))) { throw "${where}: the SciPy value is $($sheet.Range("H$row").Value2), expected $($case.scipy)" }
        if (-not (Same-Formula ([string]$sheet.Range("G$row").Formula2) ('=' + $case.library))) { throw "${where}: the library formula is $($sheet.Range("G$row").Formula2)" }
        if ($null -eq $case.excel) {
            if ([string]$sheet.Range("F$row").Value2 -ne 'none') { throw "${where}: Excel has no function, the cell reads $($sheet.Range("F$row").Value2)" }
        } elseif (-not (Same-Formula ([string]$sheet.Range("F$row").Formula2) ('=' + $case.excel))) { throw "${where}: the Excel formula is $($sheet.Range("F$row").Formula2)" }
        $observation = [ordered]@{ label = [string]$case.label }
        foreach ($side in @(@('excel', 'F', 'I', 'M'), @('library', 'G', 'K', 'N'), @('scipy', 'H', 'L', 'O'))) {
            $value = $sheet.Range($side[1] + $row).Value2
            $expected = Digits $value $reference
            $shownDigits = $sheet.Range($side[2] + $row).Value2
            if ($null -eq $expected) {
                if (-not (Is-Error $shownDigits)) { throw "${where}: $($side[0]) has no number but shows $shownDigits digits" }
            } elseif (-not ($shownDigits -is [double]) -or [Math]::Abs([double]$shownDigits - $expected) -gt 1E-9) {
                throw "${where}: $($side[0]) shows $shownDigits digits, expected $expected"
            }
            $grade = Grade $value $expected 12 6
            if ([string]$sheet.Range($side[3] + $row).Value2 -ne $grade) { throw "${where}: $($side[0]) grade is $($sheet.Range($side[3] + $row).Value2), expected $grade" }
            $observation[$side[0]] = Shown $value
            $observation[$side[0] + '_digits'] = $expected
            $observation[$side[0] + '_grade'] = $grade
        }
        # The library and SciPy are reference grade on every case of the study; Excel is only recorded.
        if ($observation.library_grade -ne 'reference') { throw "${where}: the library is $($observation.library_grade) grade, $($observation.library_digits) digits" }
        if ($observation.scipy_grade -ne 'reference') { throw "${where}: SciPy is $($observation.scipy_grade) grade" }
        $report.observations += $observation
        $report.cases++
    }

    # The selected case: every case through D6, with events on.
    $excel.EnableEvents = $true
    for ($index = 0; $index -lt $fixture.cases.Count; $index++) {
        $case = $fixture.cases[$index]
        $row = $first + $index
        Set-Text $sheet 'D6' $case.label
        $excel.Calculate()
        $where = "detail of $($case.label)"
        if ([string]$sheet.Range("D$($rows.library_formula)").Value2 -ne ('=' + $case.library)) { throw "${where}: the library formula reads $($sheet.Range("D$($rows.library_formula)").Value2)" }
        $excelText = [string]$sheet.Range("D$($rows.excel_formula)").Value2
        if ($null -eq $case.excel) { if (-not $excelText.StartsWith('none')) { throw "${where}: the Excel formula reads $excelText" } }
        elseif ($excelText -ne ('=' + $case.excel)) { throw "${where}: the Excel formula reads $excelText" }
        if ([string]$sheet.Range("D$($rows.reference)").Value2 -ne [string]$case.reference) { throw "${where}: the reference reads $($sheet.Range("D$($rows.reference)").Value2)" }
        if ([string]$sheet.Range("D$($rows.python_code)").Value2 -ne [string]$case.python) { throw "${where}: the Python code reads $($sheet.Range("D$($rows.python_code)").Value2)" }
        if ([string]$sheet.Range("D$($rows.why)").Value2 -ne [string]$case.why) { throw "${where}: the reason reads $($sheet.Range("D$($rows.why)").Value2)" }
        foreach ($pair in @(@('excel_grade', 'M'), @('library_grade', 'N'), @('scipy_grade', 'O'))) {
            $line = [string]$sheet.Range('D' + $rows.($pair[0])).Value2
            $grade = [string]$sheet.Range($pair[1] + $row).Value2
            if (-not $line.EndsWith($grade)) { throw "${where}: '$line' does not end with the grade $grade" }
        }
        if (-not ([string]$sheet.Range("D$($rows.python_grade)").Value2).StartsWith('not used')) { throw "${where}: the =PY() line reads $($sheet.Range("D$($rows.python_grade)").Value2)" }
    }
    $report.detail = 'PASS'

    # The =PY() cell: a number is graded like the other values; text is refused with a hint.
    Set-Text $sheet 'D6' $fixture.default_case
    $default = $fixture.cases | Where-Object { $_.label -eq $fixture.default_case }
    Set-Number $sheet 'D9' ([double]::Parse([string]$default.scipy, $invariant))
    $excel.Calculate()
    $line = [string]$sheet.Range("D$($rows.python_grade)").Value2
    if (-not $line.EndsWith('reference') -or -not $line.StartsWith('difference')) { throw "A =PY() value equal to SciPy's reads: $line" }
    Write-Output 'step: =PY() text'
    Set-Text $sheet 'D9' 'abc'
    $excel.Calculate()
    if (-not ([string]$sheet.Range("D$($rows.python_grade)").Value2).StartsWith('D9 is not a number')) { throw "Text in D9 reads: $($sheet.Range("D$($rows.python_grade)").Value2)" }
    $sheet.Range('D9').ClearContents() | Out-Null
    $report.python = 'PASS'

    # Other thresholds change the grades and the chart's title; contradictory ones are refused.
    Write-Output 'step: reference threshold 15'
    Set-Number $sheet 'D7' 15
    Write-Output 'step: report threshold 14'
    Set-Number $sheet 'D8' 14
    Write-Output 'step: calculate'
    $excel.Calculate()
    for ($index = 0; $index -lt $fixture.cases.Count; $index++) {
        $row = $first + $index
        $value = $sheet.Range("G$row").Value2
        $expected = Grade $value (Digits $value ([double]$sheet.Range("D$row").Value2)) 15 14
        if ([string]$sheet.Range("N$row").Value2 -ne $expected) { throw "Thresholds 15 and 14: $($fixture.cases[$index].label) is $($sheet.Range("N$row").Value2), expected $expected" }
    }
    if (-not ([string]$chart.ChartTitle.Text).Contains('15')) { throw "The chart's title does not follow the threshold: $($chart.ChartTitle.Text)" }
    Write-Output 'step: report threshold 16'
    Set-Number $sheet 'D8' 16
    $excel.Calculate()
    if ([string]$sheet.Range('D19').Value2 -ne 'Check: grade thresholds') { throw "A report threshold above the reference one gives: $($sheet.Range('D19').Value2)" }
    if ([string]$sheet.Range("N$first").Value2 -ne '-') { throw 'A grade is shown with contradictory thresholds' }
    Write-Output 'step: unknown case'
    Set-Text $sheet 'D6' 'no such case'
    $excel.Calculate()
    if ([string]$sheet.Range('D19').Value2 -ne 'Check: case') { throw "An unknown case gives: $($sheet.Range('D19').Value2)" }
    $report.thresholds = 'PASS'

    # Reset.
    Write-Output 'step: reset'
    [void]$excel.Run($prefix + 'ResetComparisonInputs', $false)
    $excel.Calculate()
    if ([string]$sheet.Range('D6').Value2 -ne [string]$fixture.default_case -or [double]$sheet.Range('D7').Value2 -ne 12 -or
        [double]$sheet.Range('D8').Value2 -ne 6 -or $null -ne $sheet.Range('D9').Value2 -or [string]$sheet.Range('D19').Value2 -ne 'OK') {
        throw 'The reset did not restore the defaults'
    }
    $report.reset = 'PASS'
    $report.status = 'PASS'
} catch {
    $report.error = $_.Exception.Message
    # When Excel itself stops (the RPC server is unavailable), Windows records why: the faulting module and
    # the exception code of EXCEL.EXE in the Application log.
    if ($_.Exception.Message -match '0x800706B[AE]') {
        try {
            $crashes = @(Get-WinEvent -FilterHashtable @{ LogName = 'Application'; StartTime = (Get-Date).AddMinutes(-10) } -ErrorAction Stop |
                Where-Object { $_.ProviderName -in @('Application Error', 'Windows Error Reporting', 'Application Hang') -and $_.Message -match 'EXCEL' } |
                Select-Object -First 3 | ForEach-Object { ($_.TimeCreated.ToString('s') + ' ' + $_.ProviderName + ': ' + ($_.Message -replace '\s+', ' ')) })
            $report.notes += $crashes
            $crashes | Write-Output
        } catch { Write-Output ('could not read the Application log: ' + $_.Exception.Message) }
    }
    throw
} finally {
    $report | ConvertTo-Json -Depth 5 | Set-Content -Encoding UTF8 (Join-Path $directory 'comparison-report.json')
    $report | ConvertTo-Json -Depth 5 | Write-Output
    if ($null -ne $book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
}
