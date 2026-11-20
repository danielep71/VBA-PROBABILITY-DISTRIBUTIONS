param([Parameter(Mandatory=$true)][string]$PackageDirectory)
# The Parameter Effects sheet against independent references
# (examples/demo/parameter-effects-references.json, made with mpmath): every case sets the
# family, the parameter to vary and its four values through the sheet's own cells, with events
# on, and reads the table of the four curves (mean, standard deviation, skewness, the tail
# beyond three standard deviations), the shared window and five points of the spilled grid.
# Also: the layout the plan promises, the two charts of each kind, an invalid curve value,
# a paste over several cells that changes the family, the choice of another parameter and the reset.
$ErrorActionPreference = 'Stop'
$directory = (Resolve-Path $PackageDirectory).Path
$path = Join-Path $directory 'Probability-Distributions-Demo.xlsm'
$fixture = Get-Content (Join-Path $PSScriptRoot '../examples/demo/parameter-effects-references.json') -Raw | ConvertFrom-Json
$plan = Get-Content (Join-Path $directory 'parameters-plan.json') -Raw | ConvertFrom-Json
$explorerPlan = Get-Content (Join-Path $directory 'explorer-plan.json') -Raw | ConvertFrom-Json
$report = [ordered]@{ status = 'FAIL'; sha256 = (Get-FileHash $path -Algorithm SHA256).Hash; cases = 0; curves = 0; points = 0; notes = @() }
$excel = $null
$book = $null

function Is-Error($Value) { return ($Value -is [int]) -and ($Value -lt -2146820000) }
function Read-Number($Sheet, [string]$Address) {
    $value = $Sheet.Range($Address).Value2
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
function Parameter-Label([string]$Family, [int]$Slot) {
    $entry = $explorerPlan.families | Where-Object { $_.name -eq $Family }
    return [string]$entry.parameters[$Slot - 1].label
}
function Spill-Rows($Sheet, [string]$Address) {
    $spill = $Sheet.Range($Address).SpillingToRange
    if ($null -eq $spill) { return 0 }
    return [int]$spill.Rows.Count
}
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($path, 0, $false)
    # Excel may still be loading or calculating the sheets after the workbook opens: a COM call is rejected until it is ready.
    for ($wait = 0; $wait -lt 240; $wait++) { try { if ($excel.Ready) { break } } catch { }; Start-Sleep -Milliseconds 500 }
    $sheet = $book.Worksheets.Item('Parameters')
    $explorer = $book.Worksheets.Item('Explorer')
    $prefix = "'" + $book.Name + "'!"
    $report.excel_version = $excel.Version
    $report.excel_build = $excel.Build
    $first = [int]$plan.layout.first_curve_row
    $x = [string]$plan.grid.x_column + [string]$plan.grid.first_row

    # Structure: no implicit intersection, the promised cells, locked outputs, unlocked inputs, messages.
    foreach ($formula in $sheet.Range('A1:P40').Formula2) {
        if ([string]$formula -like '=*' -and [string]$formula -match '@|_xlfn.SINGLE') { throw "Implicit intersection remains: $formula" }
    }
    if ($sheet.Range('D6').Validation.Formula1 -ne '=ExplorerFamily') { throw 'The distribution list is not backed by the family table' }
    if ($sheet.Range('D11').Validation.Formula1 -ne '=ParameterEffectsChoices') { throw 'The parameter to vary is not backed by the family table' }
    if ($sheet.Range('D6:D16').Locked) { throw 'D6:D16 must be editable' }
    foreach ($address in @('D19:D20', "C${first}:I$($first + 3)", 'C27')) {
        if (-not $sheet.Range($address).Locked) { throw "$address must be locked" }
    }
    foreach ($row in 6..16) {
        $rule = $sheet.Range("D$row").Validation
        foreach ($text in @($rule.InputTitle, $rule.InputMessage, $rule.ErrorTitle, $rule.ErrorMessage)) {
            if ([string]::IsNullOrWhiteSpace([string]$text)) { throw "D$row lacks an input or error message" }
        }
    }
    if ($sheet.ChartObjects().Count -ne 2) { throw "The sheet has $($sheet.ChartObjects().Count) charts, expected 2" }
    foreach ($name in @('ParameterEffectsDensity', 'ParameterEffectsCdf')) {
        if ($sheet.ChartObjects($name).Chart.SeriesCollection().Count -ne 4) { throw "$name does not have four series" }
    }
    if ($explorer.ChartObjects().Count -ne 2) { throw 'The Explorer lost a chart' }
    if ([double]$book.Worksheets.Item('Probabilities').ChartObjects().Count -ne 1) { throw 'The Probabilities sheet lost its chart' }
    $report.structure = 'PASS'

    # Every case: the table, the window and the grid against the references.
    $current = ''
    foreach ($case in $fixture.cases) {
        $excel.EnableEvents = $true
        if ($current -ne $case.family) {
            Set-Text $sheet 'D6' $case.family
            $current = [string]$case.family
        }
        # The base parameters of the family are its defaults, loaded with it.
        Set-Text $sheet 'D11' (Parameter-Label $case.family ([int]$case.vary))
        for ($k = 0; $k -lt 4; $k++) { Set-Number $sheet ('D' + (12 + $k)) ([double]$case.values[$k]) }
        $excel.Calculate()
        $where = "$($case.family) varying $(Parameter-Label $case.family ([int]$case.vary)) over $($case.values -join ', ') [$($case.label)]"
        if ([string]$sheet.Range('D19').Value2 -ne 'OK') { throw "${where}: status is $($sheet.Range('D19').Value2)" }
        for ($k = 0; $k -lt 4; $k++) {
            $row = $first + $k
            $curve = $case.curves[$k]
            if (-not (Near (Read-Number $sheet "D$row") ([double]$curve.value))) { throw "${where}: D$row is $($sheet.Range("D$row").Value2), expected $($curve.value)" }
            foreach ($pair in @(@('F', $curve.mean), @('G', $curve.sd), @('H', $curve.skewness), @('I', $curve.tail))) {
                $cell = $pair[0] + $row
                $expected = $pair[1]
                if ($expected -is [string]) {
                    if ([string]$sheet.Range($cell).Value2 -ne $expected) { throw "${where}: curve $($k + 1) $cell reads '$($sheet.Range($cell).Value2)', expected '$expected'" }
                } elseif (-not (Near (Read-Number $sheet $cell) ([double]$expected))) {
                    throw "${where}: curve $($k + 1) $cell is $($sheet.Range($cell).Value2), expected $expected"
                }
            }
            $report.curves++
        }
        $rows = Spill-Rows $sheet $x
        if ($rows -ne [int]$case.rows) { throw "${where}: the grid has $rows rows, expected $($case.rows)" }
        foreach ($name in @($plan.grid.density_columns + $plan.grid.cdf_columns)) {
            if ((Spill-Rows $sheet ([string]$name + [string]$plan.grid.first_row)) -ne $rows) { throw "${where}: column $name does not spill over the same rows as x" }
        }
        $xs = $sheet.Range($x).SpillingToRange.Value2
        foreach ($sample in $case.sample) {
            $index = [int]$sample.index + 1
            if (-not (Near ([double]$xs[$index, 1]) ([double]$sample.x))) { throw "${where}: grid x at index $($sample.index) is $($xs[$index, 1]), expected $($sample.x)" }
            for ($k = 0; $k -lt 4; $k++) {
                $density = $sheet.Range([string]$plan.grid.density_columns[$k] + [string]$plan.grid.first_row).SpillingToRange.Value2
                $cdf = $sheet.Range([string]$plan.grid.cdf_columns[$k] + [string]$plan.grid.first_row).SpillingToRange.Value2
                if ($null -ne $sample.density[$k] -and -not (Near ([double]$density[$index, 1]) ([double]$sample.density[$k]))) { throw "${where}: curve $($k + 1) density at x=$($sample.x) is $($density[$index, 1]), expected $($sample.density[$k])" }
                if ($null -ne $sample.cdf[$k] -and -not (Near ([double]$cdf[$index, 1]) ([double]$sample.cdf[$k]))) { throw "${where}: curve $($k + 1) cumulative at x=$($sample.x) is $($cdf[$index, 1]), expected $($sample.cdf[$k])" }
                $report.points++
            }
        }
        $report.cases++
    }

    # A family of each kind: the chart types of both charts, and the number of decimals.
    foreach ($kind in @(@('Normal', 4, 4), @('Binomial', 51, 65))) {
        $excel.EnableEvents = $true
        Set-Text $sheet 'D6' $kind[0]
        $excel.Calculate()
        $density = $sheet.ChartObjects('ParameterEffectsDensity').Chart
        $cdf = $sheet.ChartObjects('ParameterEffectsCdf').Chart
        for ($k = 1; $k -le 4; $k++) {
            if ($density.SeriesCollection($k).ChartType -ne $kind[1]) { throw "$($kind[0]): the density chart has type $($density.SeriesCollection($k).ChartType), expected $($kind[1])" }
            if ($cdf.SeriesCollection($k).ChartType -ne $kind[2]) { throw "$($kind[0]): the cumulative chart has type $($cdf.SeriesCollection($k).ChartType), expected $($kind[2])" }
        }
        Set-Number $sheet 'D16' 3
        $excel.Calculate()
        if (([string]$sheet.Range("G$first").Text) -notmatch '^-?[0-9]+[.,][0-9]{3}$') { throw "Three decimals show G$first as '$($sheet.Range("G$first").Text)'" }
        Set-Number $sheet 'D16' 6
    }
    $report.charts = 'PASS'

    # Choosing another parameter loads its four values; the varied base cell is not used.
    Set-Text $sheet 'D6' 'Normal'
    Set-Text $sheet 'D11' 'Mean'
    $excel.Calculate()
    $expectedMean = @(-2.0, -1.0, 0.0, 1.0)
    for ($k = 0; $k -lt 4; $k++) {
        $value = [double]$sheet.Range('D' + (12 + $k)).Value2
        if ($value -ne $expectedMean[$k]) { throw "The value of curve $($k + 1) for the Normal mean is $value, expected $($expectedMean[$k])" }
    }
    if ([string]$sheet.Range('D19').Value2 -ne 'OK') { throw "Status after choosing the mean: $($sheet.Range('D19').Value2)" }
    if (-not ([string]$sheet.Range('C27').Value2).StartsWith('Location')) { throw "The line on the mean reads: $($sheet.Range('C27').Value2)" }
    $report.vary = 'PASS'

    # An invalid curve value and a base parameter out of range: the status names the first one to check.
    Set-Text $sheet 'D11' 'Standard deviation'
    Set-Number $sheet 'D12' -1
    $excel.Calculate()
    if (([string]$sheet.Range('D19').Value2) -ne 'Check: curve 1') { throw "A negative scale gives: $($sheet.Range('D19').Value2)" }
    if (-not (Is-Error $sheet.Range("F$first").Value2)) { throw 'A statistic is shown for an invalid curve' }
    Set-Number $sheet 'D12' 0.5
    Set-Number $sheet 'D7' 1000001
    $excel.Calculate()
    if (([string]$sheet.Range('D19').Value2) -ne 'Check: Mean') { throw "A mean out of range gives: $($sheet.Range('D19').Value2)" }
    Set-Number $sheet 'D7' 0
    $report.invalid = 'PASS'

    # A paste over several cells that includes D6 keeps the pasted values and follows the family.
    $excel.EnableEvents = $true
    Set-Text $sheet 'D6' 'Normal'
    $normalTitle = [string]$sheet.Range('D7').Validation.InputTitle
    $block = New-Object 'object[,]' 3, 1
    $block[0, 0] = 'Binomial'
    $block[1, 0] = [double]12
    $block[2, 0] = [double]0.25
    $sheet.Range('D6:D8').Value2 = $block
    $excel.Calculate()
    if ([double]$sheet.Range('D7').Value2 -ne 12 -or [double]$sheet.Range('D8').Value2 -ne 0.25) { throw 'A paste over D6:D8 lost the pasted parameters' }
    if ($sheet.ChartObjects('ParameterEffectsDensity').Chart.SeriesCollection(1).ChartType -ne 51) { throw 'A paste of a discrete family left the continuous charts' }
    if ([string]$sheet.Range('D7').Validation.InputTitle -eq $normalTitle) { throw 'A paste of another family left the parameter messages of the old one' }
    if (@('Number of trials', 'Trials') -notcontains [string]$sheet.Range('D11').Value2 -and [string]$sheet.Range('D19').Value2 -ne 'OK') { throw "A paste left the parameter to vary at '$($sheet.Range('D11').Value2)' with status $($sheet.Range('D19').Value2)" }
    $report.paste = 'PASS'

    # Reset.
    [void]$excel.Run($prefix + 'ResetParameterEffectsInputs', $false)
    $excel.Calculate()
    if ([string]$sheet.Range('D6').Value2 -ne 'Normal' -or [string]$sheet.Range('D11').Value2 -ne 'Standard deviation' -or [string]$sheet.Range('D19').Value2 -ne 'OK') { throw 'The reset did not restore the defaults' }
    # Compared as numbers: joined as text they would follow the regional decimal separator.
    $expectedReset = @(0.5, 1.0, 1.5, 2.0)
    for ($k = 0; $k -lt 4; $k++) {
        $value = [double]$sheet.Range('D' + (12 + $k)).Value2
        if ($value -ne $expectedReset[$k]) { throw "The reset value of curve $($k + 1) is $value, expected $($expectedReset[$k])" }
    }
    $report.reset = 'PASS'
    $report.status = 'PASS'
} catch {
    $report.error = $_.Exception.Message
    throw
} finally {
    $report | ConvertTo-Json -Depth 5 | Set-Content -Encoding UTF8 (Join-Path $directory 'parameters-report.json')
    $report | ConvertTo-Json -Depth 5 | Write-Output
    if ($null -ne $book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
}
