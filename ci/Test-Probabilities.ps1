param([Parameter(Mandatory=$true)][string]$PackageDirectory)
# The Probabilities sheet against independent references
# (examples/demo/probabilities-references.json, made with mpmath): every case sets
# the family, its parameters and the points a and b and the probability p through
# the sheet's own cells, with events on, and reads the results. Also: the layout
# the plan promises, the chart of each kind, the shaded columns, an invalid
# input, the number of decimals and the reset.
$ErrorActionPreference = 'Stop'
$directory = (Resolve-Path $PackageDirectory).Path
$path = Join-Path $directory 'Probability-Distributions-Demo.xlsm'
$fixture = Get-Content (Join-Path $PSScriptRoot '../examples/demo/probabilities-references.json') -Raw | ConvertFrom-Json
$plan = Get-Content (Join-Path $directory 'probabilities-plan.json') -Raw | ConvertFrom-Json
$report = [ordered]@{ status = 'FAIL'; sha256 = (Get-FileHash $path -Algorithm SHA256).Hash; cases = 0; roundtrips = 0; notes = @() }
$excel = $null
$book = $null

function Is-Error($Value) { return ($Value -is [int]) -and ($Value -lt -2146820000) }
function Read-Number($Sheet, [string]$Address) {
    $value = $Sheet.Range($Address).Value2
    if (Is-Error $value) { return $null }
    if (-not ($value -is [double])) { return $null }
    return [double]$value
}
function Near($Actual, $Expected) {
    if ($null -eq $Actual) { return $false }
    return [Math]::Abs($Actual - $Expected) -le ([double]$fixture.absolute_tolerance + [double]$fixture.relative_tolerance * [Math]::Abs($Expected))
}
# One call site per value type: PowerShell binds a COM setter to the first type it sees.
function Set-Text($Sheet, [string]$Address, [string]$Value) { $Sheet.Range($Address).Value2 = $Value }
function Set-Number($Sheet, [string]$Address, [double]$Value) { $Sheet.Range($Address).Value2 = $Value }
function Grid-Rows($Sheet) {
    $spill = $Sheet.Range('F20').SpillingToRange
    if ($null -eq $spill) { return 0 }
    return [int]$spill.Rows.Count
}
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($path, 0, $false)
    $sheet = $book.Worksheets.Item('Probabilities')
    $explorer = $book.Worksheets.Item('Explorer')
    $prefix = "'" + $book.Name + "'!"
    $report.excel_version = $excel.Version
    $report.excel_build = $excel.Build

    # Structure: no implicit intersection, the promised cells, locked outputs, unlocked inputs, messages.
    foreach ($formula in $sheet.Range('A1:P40').Formula2) {
        if ([string]$formula -like '=*' -and [string]$formula -match '@|_xlfn.SINGLE') { throw "Implicit intersection remains: $formula" }
    }
    $validation = $sheet.Range('D6').Validation
    if ($validation.Type -ne 3 -or $validation.Formula1 -ne '=ExplorerFamily') { throw 'The distribution list is not backed by the family table' }
    if ($sheet.Range('D6:D14').Locked) { throw 'D6:D14 must be editable' }
    foreach ($address in @('D17:D18', 'D20:D28', 'F20:I1020')) {
        if (-not $sheet.Range($address).Locked) { throw "$address must be locked" }
    }
    if ([double]$sheet.Range('D28').Font.Size -ne 8) { throw "The verdict cell has font size $($sheet.Range('D28').Font.Size), expected 8" }
    foreach ($row in @(6, 7, 8, 9, 10, 11, 12, 13, 14)) {
        $rule = $sheet.Range("D$row").Validation
        foreach ($text in @($rule.InputTitle, $rule.InputMessage, $rule.ErrorTitle, $rule.ErrorMessage)) {
            if ([string]::IsNullOrWhiteSpace([string]$text)) { throw "D$row lacks an input or error message" }
        }
    }
    if ($sheet.ChartObjects().Count -ne 1) { throw "The sheet has $($sheet.ChartObjects().Count) charts, expected 1" }
    $chart = $sheet.ChartObjects($plan.chart.name).Chart
    if ($chart.SeriesCollection().Count -ne 3) { throw "The chart has $($chart.SeriesCollection().Count) series, expected 3" }
    if ($explorer.ChartObjects().Count -ne 2) { throw 'The Explorer lost a chart' }
    $report.structure = 'PASS'

    # Every case: the results of the sheet against the references.

    $current = ''
    foreach ($case in $fixture.cases) {
        $excel.EnableEvents = $true
        if ($current -ne $case.family) {
            Set-Text $sheet 'D6' $case.family
            $current = [string]$case.family
        }
        for ($slot = 0; $slot -lt $case.parameters.Count; $slot++) { Set-Number $sheet ('D' + (7 + $slot)) ([double]$case.parameters[$slot]) }
        Set-Number $sheet 'D11' ([double]$case.a)
        Set-Number $sheet 'D12' ([double]$case.b)
        Set-Number $sheet 'D13' ([double]$case.p)
        $excel.Calculate()
        $where = "$($case.family) [$($case.label)] a=$($case.a) b=$($case.b) p=$($case.p)"
        if ([string]$sheet.Range('D17').Value2 -ne 'OK') { throw "${where}: status is $($sheet.Range('D17').Value2)" }
        $pairs = @(@('D20', $case.cdf_a), @('D21', $case.survival_a), @('D22', $case.cdf_b), @('D23', $case.survival_b), @('D24', $case.interval))
        foreach ($pair in $pairs) {
            if ($null -eq $pair[1]) { continue }
            $actual = Read-Number $sheet $pair[0]
            if (-not (Near $actual ([double]$pair[1]))) { throw "${where}: $($pair[0]) is $($sheet.Range($pair[0]).Value2), expected $($pair[1])" }
        }
        # The library's quantile contract covers p from 0.01 to 0.999 and the deep lower tail
        # (benchmark/accuracy_contracts.csv): above 0.999 the quantile is compared loosely and
        # the round-trip verdict is not asserted, only recorded.
        $upperEnd = [double]$case.p -gt 0.999
        $quantile = Read-Number $sheet 'D25'
        if ($null -ne $case.quantile) {
            $tolerance = if ($upperEnd) { 1e-6 } else { [double]$fixture.relative_tolerance }
            if ($null -eq $quantile -or [Math]::Abs($quantile - [double]$case.quantile) -gt ([double]$fixture.absolute_tolerance + $tolerance * [Math]::Abs([double]$case.quantile))) {
                throw "${where}: D25 is $($sheet.Range('D25').Value2), expected $($case.quantile)"
            }
            if ($upperEnd -and $quantile -ne [double]$case.quantile) {
                $report.notes += "$where : quantile $quantile, reference $($case.quantile)"
            }
        }
        if (Is-Error $sheet.Range('D28').Value2) { throw "${where}: the round-trip verdict is an error ($($sheet.Range('D28').Text))" }
        $verdict = [string]$sheet.Range('D28').Value2
        if (-not $upperEnd -and ($verdict -like 'OUTSIDE*' -or $verdict -like 'CHECK*')) { throw "${where}: round-trip verdict is '$verdict' (difference $($sheet.Range('D27').Value2))" }
        if ($upperEnd -and $verdict -like 'OUTSIDE*') { $report.notes += "$where : verdict $verdict" }
        if ($sheet.Range('D27').Value2 -is [double]) { $report.roundtrips++ }
        $report.cases++
    }

    # A family of each kind: the chart types, the shaded columns, the number of decimals.
    foreach ($kind in @(@('Normal', 75, 1), @('Binomial', 51, 51))) {
        $excel.EnableEvents = $true
        Set-Text $sheet 'D6' $kind[0]
        $excel.Calculate()
        $chart = $sheet.ChartObjects($plan.chart.name).Chart
        $rows = Grid-Rows $sheet
        if ($rows -lt 2) { throw "$($kind[0]): the grid has $rows rows" }
        foreach ($column in @('G', 'H', 'I')) {
            $spill = $sheet.Range($column + '20').SpillingToRange
            if ($null -eq $spill -or [int]$spill.Rows.Count -ne $rows) { throw "$($kind[0]): column $column does not spill over the same rows as x" }
        }
        if ($kind[0] -eq 'Normal') {
            if ($chart.SeriesCollection(1).ChartType -ne 4 -or $chart.SeriesCollection(2).ChartType -ne 1) { throw "Normal: chart types are $($chart.SeriesCollection(1).ChartType), $($chart.SeriesCollection(2).ChartType)" }
        } elseif ($chart.SeriesCollection(1).ChartType -ne 51 -or $chart.SeriesCollection(3).ChartType -ne 51) { throw 'Binomial: the chart is not made of columns' }
        # Where x <= a the first region equals the density, elsewhere 0; no cell is in both regions.
        $x = $sheet.Range('F20').SpillingToRange.Value2
        $f = $sheet.Range('G20').SpillingToRange.Value2
        $h = $sheet.Range('H20').SpillingToRange.Value2
        $i = $sheet.Range('I20').SpillingToRange.Value2
        $a = [double]$sheet.Range('D11').Value2
        $b = [double]$sheet.Range('D12').Value2
        if ($kind[0] -eq 'Binomial') { $a = [Math]::Floor($a); $b = [Math]::Floor($b) }
        for ($k = 1; $k -le $rows; $k++) {
            $xk = [double]$x[$k, 1]
            $fk = [double]$f[$k, 1]
            $hk = [double]$h[$k, 1]
            $ik = [double]$i[$k, 1]
            $inFirst = $xk -le $a
            $inSecond = ($xk -gt $a) -and ($xk -le $b)
            $wantFirst = 0.0
            $wantSecond = 0.0
            if ($inFirst) { $wantFirst = $fk }
            if ($inSecond) { $wantSecond = $fk }
            $offFirst = [Math]::Abs($hk - $wantFirst)
            $offSecond = [Math]::Abs($ik - $wantSecond)
            if ($offFirst -gt 1e-15 -or $offSecond -gt 1e-15) { throw "$($kind[0]): the shaded columns are wrong at x=$xk" }
        }
        Set-Number $sheet 'D14' 3
        $excel.Calculate()
        if (([string]$sheet.Range('D20').Text) -notmatch '^-?[0-9]+[.,][0-9]{3}$') { throw "Three decimals show D20 as '$($sheet.Range('D20').Text)'" }
        Set-Number $sheet 'D14' 6
    }
    $report.charts = 'PASS'

    # A paste over several cells that includes D6 keeps the pasted values and brings the
    # chart, the number formats and the parameter messages to the pasted family.
    $excel.EnableEvents = $true
    Set-Text $sheet 'D6' 'Normal'
    $excel.Calculate()
    $normalTitle = [string]$sheet.Range('D7').Validation.InputTitle
    $block = New-Object 'object[,]' 3, 1
    $block[0, 0] = 'Binomial'
    $block[1, 0] = [double]12
    $block[2, 0] = [double]0.25
    $sheet.Range('D6:D8').Value2 = $block
    $excel.Calculate()
    if ([double]$sheet.Range('D7').Value2 -ne 12 -or [double]$sheet.Range('D8').Value2 -ne 0.25) { throw 'A paste over D6:D8 lost the pasted parameters' }
    if ($sheet.ChartObjects($plan.chart.name).Chart.SeriesCollection(1).ChartType -ne 51) { throw 'A paste of a discrete family left the continuous chart' }
    if ([string]$sheet.Range('D7').Validation.InputTitle -eq $normalTitle) { throw 'A paste of another family left the parameter messages of the old one' }
    if ($sheet.Range('F20').NumberFormat -ne '0') { throw 'A paste of a discrete family left the continuous x format' }
    $report.paste = 'PASS'

    # An invalid parameter and an interval the wrong way round: the status names the first one to check.
    Set-Text $sheet 'D6' 'Normal'
    Set-Number $sheet 'D8' -1
    $excel.Calculate()
    if (-not ([string]$sheet.Range('D17').Value2).StartsWith('Check:')) { throw "An invalid scale gives: $($sheet.Range('D17').Value2)" }
    if (-not (Is-Error $sheet.Range('D20').Value2)) { throw 'A result is shown for an invalid input' }
    Set-Number $sheet 'D8' 1
    Set-Number $sheet 'D11' 2
    Set-Number $sheet 'D12' 1
    $excel.Calculate()
    if (([string]$sheet.Range('D17').Value2) -ne 'Check: a must not exceed b') { throw "a above b gives: $($sheet.Range('D17').Value2)" }
    $report.invalid = 'PASS'

    # Reset.
    [void]$excel.Run($prefix + 'ResetProbabilitiesInputs', $false)
    $excel.Calculate()
    if ([string]$sheet.Range('D6').Value2 -ne 'Normal' -or [double]$sheet.Range('D13').Value2 -ne 0.95 -or [string]$sheet.Range('D17').Value2 -ne 'OK') { throw 'The reset did not restore the defaults' }
    $report.reset = 'PASS'
    $report.status = 'PASS'
} catch {
    $report.error = $_.Exception.Message
    throw
} finally {
    $report | ConvertTo-Json -Depth 5 | Set-Content -Encoding UTF8 (Join-Path $directory 'probabilities-report.json')
    $report | ConvertTo-Json -Depth 5 | Write-Output
    if ($null -ne $book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
}
