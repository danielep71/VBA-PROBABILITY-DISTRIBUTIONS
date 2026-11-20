param([Parameter(Mandatory=$true)][string]$PackageDirectory)
# The Tails sheet (Extreme Tails and Numerical Limits) against its scenarios
# (examples/demo/tails-cases.json: 13 scenarios, 219 points, mpmath references). For every scenario,
# chosen through D6 with events on: the ladder the sheet spills (points, references, depth), the library's
# direct value and the alternative evaluated again through a scratch cell with the formula written out
# (so the SWITCH of the sheet is checked against the call it names), the correct digits and the
# statuses recomputed here from the values, and the colour Excel displays for the statuses. Also: the
# structure the plan promises, the two charts, other and contradictory thresholds, an unknown scenario
# and the reset. The report records what this Excel returned at every point, with its version and build.
# The library's accuracy at the far points is recorded, not asserted: only the body of every scenario and
# the Standard Normal tail, where 1 - CDF is known to fail, are asserted.
$ErrorActionPreference = 'Stop'
$directory = (Resolve-Path $PackageDirectory).Path
$path = Join-Path $directory 'Probability-Distributions-Demo.xlsm'
$fixture = Get-Content (Join-Path $PSScriptRoot '../examples/demo/tails-cases.json') -Raw | ConvertFrom-Json
$plan = Get-Content (Join-Path $directory 'tails-plan.json') -Raw | ConvertFrom-Json
$report = [ordered]@{ status = 'FAIL'; sha256 = (Get-FileHash $path -Algorithm SHA256).Hash; scenarios = 0; points = 0; notes = @(); observations = @() }
$excel = $null
$book = $null
$invariant = [Globalization.CultureInfo]::InvariantCulture
$full = [double]$plan.full_digits
$referenceGrade = [double]$fixture.grades.reference
$reportGrade = [double]$fixture.grades.report
$red = 0xCEC7FF
$amber = 0x9CEBFF

function Is-Error($Value) { return ($Value -is [int]) -and ($Value -lt -2146820000) }
# One call site per value type: PowerShell binds a COM setter to the first type it sees.
function Set-Text($Sheet, [string]$Address, [string]$Value) { $Sheet.Range($Address).Value2 = $Value }
function Set-Number($Sheet, [string]$Address, [double]$Value) { $Sheet.Range($Address).Value2 = $Value }
# The values a spilled formula holds, one per row (or $null when it does not spill).
function Spill-Values($Sheet, [string]$Address) {
    $spill = $Sheet.Range($Address).SpillingToRange
    if ($null -eq $spill) { return @($Sheet.Range($Address).Value2) }
    $array = $spill.Value2
    $list = @()
    for ($row = 1; $row -le $spill.Rows.Count; $row++) { $item = $array[$row, 1]; $list += , $item }
    return $list
}
# Correct digits as the sheet defines them: -log10 of the relative difference, from 0 to 16.
function Digits($Value, [double]$Reference) {
    if (-not ($Value -is [double])) { return $null }
    if ($Reference -eq 0) { return $null }
    if ($Value -eq $Reference) { return $full }
    $relative = [Math]::Abs($Value - $Reference) / [Math]::Abs($Reference)
    return [Math]::Min($full, [Math]::Max(0.0, -[Math]::Log10($relative)))
}
function Status($Value, [double]$Reference, $Digits, [bool]$Alternative, [double]$Reference12, [double]$Report) {
    if ($Value -is [string]) { if ($Value -ceq 'underflow') { return 'PMF underflow' } else { return 'error' } }
    if ($Reference -eq 0) { if ($Value -eq 0) { return 'zero: underflow' } else { return 'nonzero below range' } }
    if ($Value -eq 0) { if ($Alternative) { return 'zero: cancellation' } else { return 'zero: unexpected' } }
    $digs = -1.0
    if ($null -ne $Digits) { $digs = [double]$Digits }
    if ($digs -le 0) { return 'wrong' }
    if ($digs -ge $Reference12) { return 'reference grade' }
    if ($digs -ge $Report) { return 'report grade' }
    return 'inadequate'
}
# The stored numbers may differ from the decimal texts of the JSON in their last bit (PowerShell 5 parses
# them less exactly than Excel reads them): within 1E-15. The digits are graded, as on the sheet, against
# the reference value the sheet holds.
function Close([double]$Actual, [double]$Expected) {
    return [Math]::Abs($Actual - $Expected) -le 1E-15 * [Math]::Abs($Expected)
}
function Shown($Value) {
    if ($Value -is [double]) { return $Value.ToString('R', $invariant) }
    if (Is-Error $Value) { return 'error ' + $Value }
    return [string]$Value
}
function Same-Value($Actual, $Expected) {
    if ($Actual -is [string] -or $Expected -is [string]) { return ([string]$Actual -ceq [string]$Expected) }
    if (Is-Error $Actual -or Is-Error $Expected) { return (Is-Error $Actual) -and (Is-Error $Expected) }
    return [double]$Actual -eq [double]$Expected
}
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($path, 0, $false)
    # Excel may still be loading or calculating the sheets after the workbook opens: a COM call is rejected until it is ready.
    for ($wait = 0; $wait -lt 240; $wait++) { try { if ($excel.Ready) { break } } catch { }; Start-Sleep -Milliseconds 500 }
    $sheet = $book.Worksheets.Item('Tails')
    $prefix = "'" + $book.Name + "'!"
    $report.excel_version = $excel.Version
    $report.excel_build = $excel.Build
    $report.excel_os = $excel.OperatingSystem
    $first = [int]$plan.layout.first_row
    $rows = $plan.layout.detail_rows

    # Structure: no implicit intersection, the promised cells, locked outputs, unlocked inputs, messages, the charts.
    foreach ($formula in $sheet.Range('A1:AO80').Formula2) {
        if ([string]$formula -like '=*' -and [string]$formula -match '@|_xlfn.SINGLE') { throw "Implicit intersection remains: $formula" }
    }
    if ($sheet.Range('D6').Validation.Formula1 -ne '=TailsScenarioNames') { throw 'The scenario list is not backed by the table of scenarios' }
    if ($sheet.Range('D6:D8').Locked) { throw 'D6:D8 must be editable' }
    foreach ($address in $plan.layout.outputs) {
        if (-not $sheet.Range($address).Locked) { throw "$address must be locked" }
    }
    foreach ($row in 6..8) {
        $rule = $sheet.Range("D$row").Validation
        foreach ($text in @($rule.InputTitle, $rule.InputMessage, $rule.ErrorTitle, $rule.ErrorMessage)) {
            if ([string]::IsNullOrWhiteSpace([string]$text)) { throw "D$row lacks an input or error message" }
        }
    }
    if ($sheet.ChartObjects().Count -ne 2) { throw "The sheet has $($sheet.ChartObjects().Count) charts, expected 2" }
    foreach ($spec in $plan.charts) {
        $chart = $sheet.ChartObjects([string]$spec.name).Chart
        if ($chart.SeriesCollection().Count -ne @($spec.series).Count) { throw "$($spec.name) has $($chart.SeriesCollection().Count) series, expected $(@($spec.series).Count)" }
    }
    foreach ($rule in @($sheet.Cells.FormatConditions)) {
        if ($rule.Type -eq 2 -and [string]$rule.Formula1 -match 'Explorer') { throw "A conditional rule reads the Explorer: $($rule.Formula1)" }
    }
    foreach ($other in @(@('Explorer', 2), @('Probabilities', 1), @('Parameters', 2), @('Comparison', 1))) {
        if ($book.Worksheets.Item($other[0]).ChartObjects().Count -ne $other[1]) { throw "The $($other[0]) sheet lost a chart" }
    }
    if (-not ([string]$sheet.Range("D$($rows.excel_build)").Value2).Contains('build')) { throw "The Excel build reads: $($sheet.Range("D$($rows.excel_build)").Value2)" }
    if ([string]$sheet.Range('D19').Value2 -ne 'OK') { throw "Status at the defaults: $($sheet.Range('D19').Value2)" }
    if ([string]$sheet.Range('D6').Value2 -ne [string]$fixture.default_scenario) { throw "The default scenario is $($sheet.Range('D6').Value2)" }
    $report.structure = 'PASS'

    # Every scenario: the ladder, the values, the digits, the statuses and the colours.
    $excel.EnableEvents = $true
    $sheet.Activate()
    foreach ($scenario in $fixture.scenarios) {
        Set-Text $sheet 'D6' ([string]$scenario.name)
        $excel.Calculate()
        $where = [string]$scenario.name
        if ([string]$sheet.Range('D19').Value2 -ne 'OK') { throw "${where}: status is $($sheet.Range('D19').Value2)" }
        $points = @($scenario.points)
        $x = @(Spill-Values $sheet "C$first")
        $referenceText = @(Spill-Values $sheet "D$first")
        $direct = @(Spill-Values $sheet "F$first")
        $alternative = @(Spill-Values $sheet "G$first")
        $directDigits = @(Spill-Values $sheet "H$first")
        $alternativeDigits = @(Spill-Values $sheet "I$first")
        $directStatus = @(Spill-Values $sheet "K$first")
        $alternativeStatus = @(Spill-Values $sheet "L$first")
        $depth = @(Spill-Values $sheet "M$first")
        $held = @(Spill-Values $sheet "$($plan.helper_reference_value)$first")
        if ($held.Count -ne $points.Count) { throw "${where}: the reference values have $($held.Count) rows, expected $($points.Count)" }
        foreach ($name in @('x', 'referenceText', 'direct', 'alternative', 'directDigits', 'alternativeDigits', 'directStatus', 'alternativeStatus', 'depth')) {
            $count = @((Get-Variable $name).Value).Count
            if ($count -ne $points.Count) { throw "${where}: column $name has $count rows, expected $($points.Count)" }
        }
        # The scratch cells: both paths written out in full, one point per row, evaluated by the library.
        $scratchDirect = $sheet.Range("AQ1:AQ$($points.Count)")
        $scratchAlternative = $sheet.Range("AR1:AR$($points.Count)")
        $failure = 'error'
        if ($scenario.kind -eq 'log') { $failure = 'underflow' }
        for ($index = 0; $index -lt $points.Count; $index++) {
            $literal = ([double]$points[$index].x).ToString('R', $invariant)
            $sheet.Range("AQ$($index + 1)").Formula2 = '=IFERROR(' + ([string]$scenario.direct).Replace('{x}', $literal) + ',"error")'
            $sheet.Range("AR$($index + 1)").Formula2 = '=IFERROR(' + ([string]$scenario.alternative).Replace('{x}', $literal) + ',"' + $failure + '")'
        }
        $excel.Calculate()
        for ($index = 0; $index -lt $points.Count; $index++) {
            $point = $points[$index]
            $label = "$where point $($point.x)"
            if (([double]$x[$index]) -ne [double]$point.x) { throw "${label}: the ladder holds x = $($x[$index])" }
            if ([string]$referenceText[$index] -cne [string]$point.reference) { throw "${label}: the reference reads $($referenceText[$index])" }
            if ([Math]::Abs([double]$depth[$index] - [double]$point.depth) -gt 1E-9) { throw "${label}: the depth is $($depth[$index]), expected $($point.depth)" }
            $expectedDirect = $sheet.Range("AQ$($index + 1)").Value2
            $expectedAlternative = $sheet.Range("AR$($index + 1)").Value2
            if (-not (Same-Value $direct[$index] $expectedDirect)) { throw "${label}: the direct path shows $(Shown $direct[$index]), the formula written out gives $(Shown $expectedDirect)" }
            if (-not (Same-Value $alternative[$index] $expectedAlternative)) { throw "${label}: the alternative path shows $(Shown $alternative[$index]), the formula written out gives $(Shown $expectedAlternative)" }
            $truth = [double]$held[$index]
            if (-not (Close $truth ([double]$point.reference_value))) { throw "${label}: the reference value is $truth, expected $($point.reference_value)" }
            $observation = [ordered]@{ scenario = $where; x = Shown $x[$index]; reference = [string]$point.reference }
            foreach ($walk in @(@('direct', $direct, $directDigits, $directStatus, $false), @('alternative', $alternative, $alternativeDigits, $alternativeStatus, $true))) {
                $value = $walk[1][$index]
                $expectedDigits = Digits $value $truth
                $shownDigits = $walk[2][$index]
                if ($null -eq $expectedDigits) {
                    if (-not (Is-Error $shownDigits)) { throw "${label}: the $($walk[0]) path has nothing to grade but shows $shownDigits digits" }
                } elseif (-not ($shownDigits -is [double]) -or [Math]::Abs([double]$shownDigits - $expectedDigits) -gt 1E-9) {
                    throw "${label}: the $($walk[0]) path shows $shownDigits digits, expected $expectedDigits"
                }
                $expectedStatus = Status $value $truth $expectedDigits $walk[4] $referenceGrade $reportGrade
                if ([string]$walk[3][$index] -cne $expectedStatus) { throw "${label}: the $($walk[0]) status is '$($walk[3][$index])', expected '$expectedStatus'" }
                $observation[$walk[0]] = Shown $value
                $observation[$walk[0] + '_digits'] = $expectedDigits
                $observation[$walk[0] + '_status'] = $expectedStatus
            }
            $report.observations += $observation
            $report.points++
        }
        $sheet.Range("AQ1:AR$($points.Count)").ClearContents() | Out-Null
        # The body of every scenario is graded by the sheet: the first point is a probability of about 0.1 to 1.
        $bodyDigits = Digits $direct[0] ([double]$held[0])
        if ($null -eq $bodyDigits -or $bodyDigits -lt $reportGrade) { throw "${where}: the direct path at the body point has $bodyDigits correct digits" }
        # The colours Excel displays: red for wrong or cancellation, amber for an underflow, none for a grade.
        for ($index = 0; $index -lt $points.Count; $index++) {
            foreach ($pair in @(@('K', $directStatus), @('L', $alternativeStatus))) {
                $cellStatus = [string]$pair[1][$index]
                $shown = [int]$sheet.Range($pair[0] + ($first + $index)).DisplayFormat.Interior.Color
                $expectedFill = $null
                foreach ($rule in $plan.conditional.status_rules) { if ([string]$rule.text -ceq $cellStatus) { $expectedFill = [int]$rule.fill } }
                if ($null -ne $expectedFill -and $shown -ne $expectedFill) { throw "$where $($pair[0])$($first + $index) '$cellStatus' shows the fill $shown, expected $expectedFill" }
                if ($null -eq $expectedFill -and ($shown -eq $red -or $shown -eq $amber)) { throw "$where $($pair[0])$($first + $index) '$cellStatus' is highlighted" }
            }
        }
        # The summary lines of the scenario name what the ladder shows.
        foreach ($key in @('title', 'parameters', 'direct', 'alternative', 'reference', 'guarantee', 'alternative_usable', 'direct_usable', 'limit', 'count')) {
            $line = [string]$sheet.Range("D$($rows.$key)").Value2
            if ([string]::IsNullOrWhiteSpace($line) -or $line.StartsWith('#')) { throw "${where}: the line '$key' reads '$line'" }
        }
        if (-not ([string]$sheet.Range("D$($rows.direct)").Value2).StartsWith('=K_STATS_')) { throw "${where}: the direct formula reads $($sheet.Range("D$($rows.direct)").Value2)" }
        $chart = $sheet.ChartObjects('TailsDigits').Chart
        if ($chart.SeriesCollection(1).Points().Count -ne $points.Count) { throw "${where}: the chart of digits has $($chart.SeriesCollection(1).Points().Count) points, expected $($points.Count)" }
        $report.scenarios++
    }
    $report.ladder = 'PASS'

    # The Standard Normal tail: 1 - CDF fails where the direct path does not (z = 10 is a tail of 1E-23).
    Set-Text $sheet 'D6' 'Standard Normal'
    $excel.Calculate()
    $normal = $report.observations | Where-Object { $_.scenario -eq 'Standard Normal' -and $_.x -eq '10' }
    if ($null -eq $normal) { throw 'The Standard Normal ladder has no point z = 10' }
    if (@('wrong', 'zero: cancellation') -notcontains $normal.alternative_status) { throw "1 - CDF at z = 10 is '$($normal.alternative_status)'" }
    if ($null -eq $normal.direct_digits -or [double]$normal.direct_digits -lt $reportGrade) { throw "The direct survival at z = 10 has $($normal.direct_digits) correct digits" }
    $underflow = $report.observations | Where-Object { $_.scenario -eq 'Standard Normal' -and $_.x -eq '40' }
    if ($underflow.direct_status -cne 'zero: underflow' -and $underflow.direct_status -cne 'error') { $report.notes += "Standard Normal z = 40: the direct path is '$($underflow.direct_status)'" }
    $report.cancellation = 'PASS'

    # Other thresholds change the statuses; contradictory ones are refused; so is an unknown scenario.
    Set-Number $sheet 'D7' 15
    Set-Number $sheet 'D8' 14
    $excel.Calculate()
    $statusNow = @(Spill-Values $sheet "K$first")
    $valuesNow = @(Spill-Values $sheet "F$first")
    $heldNow = @(Spill-Values $sheet "$($plan.helper_reference_value)$first")
    $normalScenario = $fixture.scenarios | Where-Object { $_.name -eq 'Standard Normal' }
    for ($index = 0; $index -lt @($normalScenario.points).Count; $index++) {
        $point = @($normalScenario.points)[$index]
        $value = $valuesNow[$index]
        $truthNow = [double]$heldNow[$index]
        $expected = Status $value $truthNow (Digits $value $truthNow) $false 15 14
        if ([string]$statusNow[$index] -cne $expected) { throw "Thresholds 15 and 14: z = $($point.x) is '$($statusNow[$index])', expected '$expected'" }
    }
    if (-not ([string]$sheet.ChartObjects('TailsDigits').Chart.ChartTitle.Text).Contains('15')) { throw "The chart's title does not follow the threshold: $($sheet.ChartObjects('TailsDigits').Chart.ChartTitle.Text)" }
    Set-Number $sheet 'D8' 16
    $excel.Calculate()
    if ([string]$sheet.Range('D19').Value2 -ne 'Check: grade thresholds') { throw "A report threshold above the reference one gives: $($sheet.Range('D19').Value2)" }
    if (-not (Is-Error $sheet.Range("F$first").Value2)) { throw 'A value is shown with contradictory thresholds' }
    Set-Number $sheet 'D8' 6
    Set-Text $sheet 'D6' 'no such scenario'
    $excel.Calculate()
    if ([string]$sheet.Range('D19').Value2 -ne 'Check: scenario') { throw "An unknown scenario gives: $($sheet.Range('D19').Value2)" }
    $report.thresholds = 'PASS'

    # Reset.
    [void]$excel.Run($prefix + 'ResetTailsInputs', $false)
    $excel.Calculate()
    if ([string]$sheet.Range('D6').Value2 -ne [string]$fixture.default_scenario -or [double]$sheet.Range('D7').Value2 -ne 12 -or
        [double]$sheet.Range('D8').Value2 -ne 6 -or [string]$sheet.Range('D19').Value2 -ne 'OK') {
        throw 'The reset did not restore the defaults'
    }
    $report.reset = 'PASS'
    $report.status = 'PASS'
} catch {
    $report.error = $_.Exception.Message
    throw
} finally {
    $report | ConvertTo-Json -Depth 5 | Set-Content -Encoding UTF8 (Join-Path $directory 'tails-report.json')
    ([ordered]@{ status = $report.status; scenarios = $report.scenarios; points = $report.points; notes = $report.notes; error = $report.error;
                 excel_build = $report.excel_build }) | ConvertTo-Json -Depth 3 | Write-Output
    if ($null -ne $book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
}
