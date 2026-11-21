param([Parameter(Mandatory=$true)][string]$PackageDirectory)
# The Risk sheet (value at risk, expected shortfall, tails and default counts) against its cases
# (examples/demo/risk-cases.json: inputs and mpmath references). For every case, entered through the input
# cells with events on: the status, the thirteen results against the references, the formula written beside
# each result evaluated again in a scratch cell (so the text the sheet shows is the formula it computes),
# the reading of the two models, the table of losses row by row (the losses, the two densities, the two
# exceedance probabilities and their ratio), and the charts. Then the structure the plan promises, invalid
# inputs one at a time, and the reset. The report records what this Excel returned for every case, with its
# version and build.
$ErrorActionPreference = 'Stop'
$directory = (Resolve-Path $PackageDirectory).Path
$path = Join-Path $directory 'Probability-Distributions-Demo.xlsm'
$fixture = Get-Content (Join-Path $PSScriptRoot '../examples/demo/risk-cases.json') -Raw | ConvertFrom-Json
$plan = Get-Content (Join-Path $directory 'risk-plan.json') -Raw | ConvertFrom-Json
$report = [ordered]@{ status = 'FAIL'; sha256 = (Get-FileHash $path -Algorithm SHA256).Hash; cases = 0; table_rows = 0;
                      largest_relative_error = 0.0; notes = @(); observations = @() }
$excel = $null
$book = $null
$invariant = [Globalization.CultureInfo]::InvariantCulture
# The loosest accuracy contract of the functions the sheet calls is 3E-12 (the Student-t quantile); the
# sheet composes them with a few operations. 1E-9 holds every value with a wide margin and still fails a
# wrong formula, a wrong argument or a wrong tail.
$tolerance = 1E-9

function Is-Error($Value) { return ($Value -is [int]) -and ($Value -lt -2146820000) }
# One call site per value type: PowerShell binds a COM setter to the first type it sees.
function Set-Text($Sheet, [string]$Address, [string]$Value) { $Sheet.Range($Address).Value2 = $Value }
function Set-Number($Sheet, [string]$Address, [double]$Value) { $Sheet.Range($Address).Value2 = $Value }
# The values a spilled formula holds, one per row (or the single value when it does not spill).
function Spill-Values($Sheet, [string]$Address) {
    $spill = $Sheet.Range($Address).SpillingToRange
    if ($null -eq $spill) { return @($Sheet.Range($Address).Value2) }
    $array = $spill.Value2
    $list = @()
    for ($row = 1; $row -le $spill.Rows.Count; $row++) { $item = $array[$row, 1]; $list += , $item }
    return $list
}
# The gap between a value and its reference, relative to the larger of the reference and a scale (the
# loss standard deviation for an amount, the reference itself for a probability or a density).
function Gap($Actual, [double]$Expected, [double]$Scale) {
    if (-not ($Actual -is [double])) { return [double]::PositiveInfinity }
    $size = [Math]::Max([Math]::Abs($Expected), $Scale)
    if ($size -eq 0) { return [Math]::Abs([double]$Actual) }
    return [Math]::Abs([double]$Actual - $Expected) / $size
}
function Check-Value([string]$Label, $Actual, [double]$Expected, [double]$Scale) {
    $gap = Gap $Actual $Expected $Scale
    if ($gap -gt $tolerance) { throw "${Label}: $(Shown $Actual), expected $Expected (relative error $gap)" }
    if ($gap -gt $report.largest_relative_error) { $report.largest_relative_error = $gap }
}
function Check-Probability([string]$Label, $Actual, [double]$Expected) {
    Check-Value $Label $Actual $Expected 0
    if ([double]$Actual -lt 0 -or [double]$Actual -gt 1) { throw "${Label}: $Actual is not a probability" }
}
function Same-Value($Actual, $Expected) {
    if ($Actual -is [string] -or $Expected -is [string]) { return ([string]$Actual -ceq [string]$Expected) }
    if ((Is-Error $Actual) -or (Is-Error $Expected)) { return (Is-Error $Actual) -and (Is-Error $Expected) }
    return [double]$Actual -eq [double]$Expected
}
function Shown($Value) {
    if ($Value -is [double]) { return $Value.ToString('R', $invariant) }
    if (Is-Error $Value) { return 'error ' + $Value }
    return [string]$Value
}
function Set-Inputs($Sheet, $Inputs) {
    foreach ($entry in $plan.layout.input_cells.PSObject.Properties) {
        Set-Number $Sheet ([string]$entry.Value) ([double]$Inputs.($entry.Name))
    }
}
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($path, 0, $false)
    # Excel may still be loading or calculating the sheets after the workbook opens: a COM call is rejected until it is ready.
    for ($wait = 0; $wait -lt 240; $wait++) { try { if ($excel.Ready) { break } } catch { }; Start-Sleep -Milliseconds 500 }
    $sheet = $book.Worksheets.Item('Risk')
    $prefix = "'" + $book.Name + "'!"
    $report.excel_version = $excel.Version
    $report.excel_build = $excel.Build
    $report.excel_os = $excel.OperatingSystem
    $first = [int]$plan.layout.first_row
    $rows = $plan.layout.result_rows
    $keys = @($plan.layout.value_keys)
    $amounts = @('mean_loss', 'sd_loss', 'normal_var', 'normal_es', 't_scale', 't_var', 't_es', 'default_var')
    $probabilities = @('normal_exceed', 't_exceed', 'default_exceed')

    # Structure: no implicit intersection, the promised cells, locked outputs, unlocked inputs, messages, the charts.
    foreach ($formula in $sheet.Range("A1:AK$($first + 2)").Formula2) {
        if ([string]$formula -like '=*' -and [string]$formula -match '@|_xlfn.SINGLE') { throw "Implicit intersection remains: $formula" }
    }
    foreach ($entry in $plan.validation.custom.PSObject.Properties) {
        if ($sheet.Range($entry.Name).Validation.Formula1 -ne [string]$entry.Value) { throw "$($entry.Name) is not validated by $($entry.Value)" }
    }
    if ($sheet.Range($plan.layout.inputs).Locked) { throw "$($plan.layout.inputs) must be editable" }
    foreach ($address in $plan.layout.outputs) {
        if (-not $sheet.Range($address).Locked) { throw "$address must be locked" }
    }
    foreach ($row in 6..15) {
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
    if ($sheet.ChartObjects('RiskTail').Chart.Axes(2).ScaleType -ne -4133) { throw 'The tail chart is not on a log scale' }
    foreach ($rule in @($sheet.Cells.FormatConditions)) {
        if ($rule.Type -eq 2 -and [string]$rule.Formula1 -match 'Explorer') { throw "A conditional rule reads the Explorer: $($rule.Formula1)" }
    }
    foreach ($other in @(@('Explorer', 2), @('Probabilities', 1), @('Parameters', 2), @('Comparison', 1), @('Discrete', 2))) {
        if ($book.Worksheets.Item($other[0]).ChartObjects().Count -ne $other[1]) { throw "The $($other[0]) sheet lost a chart" }
    }
    # Every run builds the Tails sheet, a change to the release branch too: it has its two charts.
    $tailsCharts = $book.Worksheets.Item('Tails').ChartObjects().Count
    if ($tailsCharts -ne 2) { throw "The Tails sheet has $tailsCharts charts, expected 2" }
    if (-not ([string]$sheet.Range("D$($rows.excel_build)").Value2).Contains('build')) { throw "The Excel build reads: $($sheet.Range("D$($rows.excel_build)").Value2)" }
    if ([string]$sheet.Range('D19').Value2 -ne 'OK') { throw "Status at the defaults: $($sheet.Range('D19').Value2)" }
    foreach ($entry in $plan.layout.input_cells.PSObject.Properties) {
        if ([double]$sheet.Range([string]$entry.Value).Value2 -ne [double]$plan.defaults.($entry.Name)) { throw "The sheet was not saved with the default $($entry.Name)" }
    }
    $report.structure = 'PASS'

    # Every case: the results, the formulas beside them, the reading, the table and the charts.
    $excel.EnableEvents = $true
    $sheet.Activate()
    foreach ($case in $fixture.cases) {
        $where = [string]$case.name
        Set-Inputs $sheet $case.inputs
        $excel.Calculate()
        if ([string]$sheet.Range('D19').Value2 -ne 'OK') { throw "${where}: status is $($sheet.Range('D19').Value2)" }
        $scale = [double]$case.results.sd_loss
        $observation = [ordered]@{ case = $where }
        foreach ($key in $keys) {
            $value = $sheet.Range("D$($rows.$key)").Value2
            $expected = [double]$case.results.$key
            $observation[$key] = Shown $value
            if ($key -eq 'default_quantile') {
                if (-not ($value -is [double]) -or [double]$value -ne $expected) { throw "${where}: the default count VaR is $(Shown $value), expected $expected" }
            } elseif ($amounts -contains $key) {
                Check-Value "${where} $key" $value $expected $scale
            } elseif ($probabilities -contains $key) {
                Check-Probability "${where} $key" $value $expected
            } else {
                Check-Value "${where} $key" $value $expected 0
            }
            # The formula shown beside the result, evaluated again in a scratch cell.
            $text = [string]$sheet.Range("F$($rows.$key)").Value2
            if (-not $text.StartsWith('=')) { throw "${where}: the formula of $key reads '$text'" }
            $sheet.Range('AQ1').Formula2 = $text
            $excel.Calculate()
            if (-not (Same-Value $sheet.Range('AQ1').Value2 $value)) {
                throw "${where}: $key shows $(Shown $value), its formula $text gives $(Shown $sheet.Range('AQ1').Value2)"
            }
            $sheet.Range('AQ1').ClearContents() | Out-Null
        }
        # VaR and ES in the order the definitions give them.
        foreach ($pair in @(@('normal_var', 'normal_es'), @('t_var', 't_es'))) {
            if ([double]$sheet.Range("D$($rows.($pair[1]))").Value2 -lt [double]$sheet.Range("D$($rows.($pair[0]))").Value2) { throw "${where}: $($pair[1]) is below $($pair[0])" }
        }
        $reading = [string]$sheet.Range("D$($rows.reading)").Value2
        if (-not $reading.Contains('ES is at least the VaR in both models: yes') -or -not $reading.Contains('<= 1 - confidence.')) { throw "${where}: the reading is '$reading'" }
        if ([string]::IsNullOrWhiteSpace([string]$sheet.Range("D$($rows.assumptions)").Value2)) { throw "${where}: no assumptions are shown" }

        # The table: 101 losses, then every column against the references.
        $losses = @(Spill-Values $sheet "C$first")
        $normalDensity = @(Spill-Values $sheet "D$first")
        $tDensity = @(Spill-Values $sheet "F$first")
        $normalExceed = @(Spill-Values $sheet "G$first")
        $tExceed = @(Spill-Values $sheet "H$first")
        $ratio = @(Spill-Values $sheet "I$first")
        $table = @($case.table)
        foreach ($name in @('losses', 'normalDensity', 'tDensity', 'normalExceed', 'tExceed', 'ratio')) {
            if (@((Get-Variable $name).Value).Count -ne $table.Count) { throw "${where}: column $name has $(@((Get-Variable $name).Value).Count) rows, expected $($table.Count)" }
        }
        for ($index = 0; $index -lt $table.Count; $index++) {
            $reference = $table[$index]
            $label = "$where row $($first + $index)"
            # The loss is arithmetic on the Double mean and standard deviation, as the references form it.
            if ((Gap $losses[$index] ([double]$reference.x) $scale) -gt 1E-14) { throw "${label}: the loss is $(Shown $losses[$index]), expected $($reference.x)" }
            Check-Value "$label Normal density" $normalDensity[$index] ([double]$reference.normal_density) 0
            Check-Value "$label Student-t density" $tDensity[$index] ([double]$reference.t_density) 0
            Check-Probability "$label Normal P(L > x)" $normalExceed[$index] ([double]$reference.normal_exceed)
            Check-Probability "$label Student-t P(L > x)" $tExceed[$index] ([double]$reference.t_exceed)
            Check-Value "$label ratio" $ratio[$index] ([double]$reference.t_exceed / [double]$reference.normal_exceed) 0
            $report.table_rows++
        }
        # The charts follow the table.
        $density = $sheet.ChartObjects('RiskDensity').Chart
        if ($density.SeriesCollection(1).Points().Count -ne $table.Count) { throw "${where}: the density chart has $($density.SeriesCollection(1).Points().Count) points" }
        if ($density.SeriesCollection(3).Points().Count -ne 2 -or $density.SeriesCollection(4).Points().Count -ne 2) { throw "${where}: a VaR line has not two points" }
        if (-not ([string]$density.ChartTitle.Text).Contains('Student t')) { throw "${where}: the density chart's title is '$($density.ChartTitle.Text)'" }
        $tail = $sheet.ChartObjects('RiskTail').Chart
        if ($tail.SeriesCollection(3).Points().Count -ne $table.Count) { throw "${where}: the line of 1 - confidence has $($tail.SeriesCollection(3).Points().Count) points" }
        $observation.top = [string]$case.top
        $observation.reading = $reading
        $report.observations += $observation
        $report.cases++
    }
    $report.cases_check = 'PASS'

    # Invalid inputs, one at a time from the default example: the status names the input and nothing is shown.
    [void]$excel.Run($prefix + 'ResetRiskInputs', $false)
    $nu = 'Check: degrees of freedom (above 2, or the variance is not finite)'
    $invalid = @(
        @('D6', 'number', 0, 'Check: position value V'),
        @('D7', 'number', 0.5, 'Check: mean daily return'),
        @('D8', 'number', 0, 'Check: daily volatility'),
        @('D9', 'number', 0, 'Check: horizon in days'),
        @('D9', 'number', 2.5, 'Check: horizon in days'),
        @('D10', 'number', 0.4, 'Check: confidence level'),
        @('D10', 'number', 1, 'Check: confidence level'),
        @('D11', 'number', 2, $nu),
        @('D11', 'number', 1.5, $nu),
        @('D12', 'text', 'large', 'Check: loss threshold'),
        @('D13', 'number', 0, 'Check: number of loans'),
        @('D13', 'number', 3.5, 'Check: number of loans'),
        @('D14', 'number', 0, 'Check: default probability'),
        @('D14', 'number', 1, 'Check: default probability'),
        @('D15', 'number', -1, 'Check: loss per default')
    )
    foreach ($case in $invalid) {
        $before = $sheet.Range($case[0]).Value2
        if ($case[1] -eq 'text') { Set-Text $sheet $case[0] ([string]$case[2]) } else { Set-Number $sheet $case[0] ([double]$case[2]) }
        $excel.Calculate()
        $status = [string]$sheet.Range('D19').Value2
        if ($status -ne $case[3]) { throw "$($case[0]) = $($case[2]) gives '$status', expected '$($case[3])'" }
        if (-not (Is-Error $sheet.Range('D23').Value2) -or -not (Is-Error $sheet.Range("C$first").Value2)) { throw "$($case[0]) = $($case[2]) still shows results" }
        if ([int]$sheet.Range('D19').DisplayFormat.Interior.Color -ne 0xCEC7FF) { throw "$($case[0]) = $($case[2]): the status is not red" }
        if ($before -is [string]) { Set-Text $sheet $case[0] ([string]$before) } else { Set-Number $sheet $case[0] ([double]$before) }
    }
    $report.invalid = 'PASS'

    # A zero mean return at confidence 0.5 gives a Normal VaR of 0: the reading states the two VaRs instead of
    # a ratio, and keeps the rest.
    Set-Number $sheet 'D7' 0
    Set-Number $sheet 'D10' 0.5
    $excel.Calculate()
    if ([string]$sheet.Range('D19').Value2 -ne 'OK') { throw "A zero mean at 0.5 gives the status $($sheet.Range('D19').Value2)" }
    if ([double]$sheet.Range("D$($rows.normal_var)").Value2 -ne 0) { throw "The Normal VaR at 0.5 with a zero mean is $($sheet.Range("D$($rows.normal_var)").Value2)" }
    $reading = [string]$sheet.Range("D$($rows.reading)").Value2
    if (-not $reading.Contains('against a Normal VaR of 0') -or -not $reading.Contains('ES is at least the VaR in both models: yes')) { throw "The reading with a Normal VaR of 0 is '$reading'" }
    $report.zero_var = 'PASS'

    # Reset.
    Set-Number $sheet 'D11' 30
    Set-Number $sheet 'D10' 0.95
    [void]$excel.Run($prefix + 'ResetRiskInputs', $false)
    $excel.Calculate()
    foreach ($entry in $plan.layout.input_cells.PSObject.Properties) {
        if ([double]$sheet.Range([string]$entry.Value).Value2 -ne [double]$plan.defaults.($entry.Name)) { throw "The reset did not restore $($entry.Name)" }
    }
    if ([string]$sheet.Range('D19').Value2 -ne 'OK') { throw 'The reset did not give a valid status' }
    $report.reset = 'PASS'
    $report.status = 'PASS'
} catch {
    $report.error = $_.Exception.Message
    throw
} finally {
    $report | ConvertTo-Json -Depth 5 | Set-Content -Encoding UTF8 (Join-Path $directory 'risk-report.json')
    ([ordered]@{ status = $report.status; cases = $report.cases; table_rows = $report.table_rows;
                 largest_relative_error = $report.largest_relative_error; notes = $report.notes; error = $report.error;
                 excel_build = $report.excel_build }) | ConvertTo-Json -Depth 3 | Write-Output
    if ($null -ne $book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
}
