param([Parameter(Mandatory=$true)][string]$PackageDirectory)
# The Discrete sheet (Binomial and Poisson counts) against its cases (examples/demo/discrete-cases.json:
# inputs and mpmath references). For every case, entered through the input cells with events on: the
# status, the ten results against the references, the formula written beside each result evaluated again
# in a scratch cell (so the text the sheet shows is the call it makes), the quantile's inequalities, the
# table of counts row by row (the window, the mass, the four readings of a threshold, and the shift that
# makes "fewer than k" the row above's "at most"), the charts and the colours Excel displays. Then the
# structure the plan promises, invalid inputs one at a time, a table too wide, and the reset. The report
# records what this Excel returned for every case, with its version and build.
$ErrorActionPreference = 'Stop'
$directory = (Resolve-Path $PackageDirectory).Path
$path = Join-Path $directory 'Probability-Distributions-Demo.xlsm'
$fixture = Get-Content (Join-Path $PSScriptRoot '../examples/demo/discrete-cases.json') -Raw | ConvertFrom-Json
$plan = Get-Content (Join-Path $directory 'discrete-plan.json') -Raw | ConvertFrom-Json
$report = [ordered]@{ status = 'FAIL'; sha256 = (Get-FileHash $path -Algorithm SHA256).Hash; cases = 0; table_rows = 0;
                      largest_relative_error = 0.0; notes = @(); observations = @() }
$excel = $null
$book = $null
$invariant = [Globalization.CultureInfo]::InvariantCulture
# The loosest accuracy contract of the Binomial and Poisson functions (Binomial survival, 5E-8): every
# probability on the sheet is within it of the mpmath reference, or the test fails.
$tolerance = 5E-8
$amber = 0x9CEBFF
$green = 0xCEEFC6
$grey = 0xF2F2F2

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
function Relative($Actual, [double]$Expected) {
    if (-not ($Actual -is [double])) { return [double]::PositiveInfinity }
    if ($Expected -eq 0) { return [Math]::Abs([double]$Actual) }
    return [Math]::Abs([double]$Actual - $Expected) / [Math]::Abs($Expected)
}
function Check-Probability([string]$Label, $Actual, [double]$Expected) {
    $gap = Relative $Actual $Expected
    if ($Expected -eq 0 -and $gap -ne 0) { throw "${Label}: $Actual, expected exactly 0" }
    if ($gap -gt $tolerance) { throw "${Label}: $Actual, expected $Expected (relative error $gap)" }
    if ([double]$Actual -lt 0 -or [double]$Actual -gt 1) { throw "${Label}: $Actual is not a probability" }
    if ($gap -gt $report.largest_relative_error) { $report.largest_relative_error = $gap }
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
function Set-Inputs($Sheet, $Inputs, [string]$Model) {
    Set-Text $Sheet 'D6' $Model
    Set-Number $Sheet 'D7' ([double]$Inputs.trials)
    Set-Number $Sheet 'D8' ([double]$Inputs.probability)
    Set-Number $Sheet 'D9' ([double]$Inputs.rate)
    Set-Number $Sheet 'D10' ([double]$Inputs.exposure)
    Set-Number $Sheet 'D11' ([double]$Inputs.threshold)
    Set-Number $Sheet 'D12' ([double]$Inputs.level)
}
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($path, 0, $false)
    # Excel may still be loading or calculating the sheets after the workbook opens: a COM call is rejected until it is ready.
    for ($wait = 0; $wait -lt 240; $wait++) { try { if ($excel.Ready) { break } } catch { }; Start-Sleep -Milliseconds 500 }
    $sheet = $book.Worksheets.Item('Discrete')
    $prefix = "'" + $book.Name + "'!"
    $report.excel_version = $excel.Version
    $report.excel_build = $excel.Build
    $report.excel_os = $excel.OperatingSystem
    $first = [int]$plan.layout.first_row
    $rows = $plan.layout.result_rows
    $keys = @('mean', 'sd', 'mass', 'at_most', 'fewer_than', 'at_least', 'more_than', 'quantile', 'before_quantile', 'at_quantile')

    # Structure: no implicit intersection, the promised cells, locked outputs, unlocked inputs, messages, the charts.
    foreach ($formula in $sheet.Range("A1:AK$($first + 2)").Formula2) {
        if ([string]$formula -like '=*' -and [string]$formula -match '@|_xlfn.SINGLE') { throw "Implicit intersection remains: $formula" }
    }
    if ($sheet.Range('D6').Validation.Formula1 -ne '=DiscreteModels') { throw 'The model list is not backed by the list of models' }
    foreach ($entry in $plan.validation.custom.PSObject.Properties) {
        if ($sheet.Range($entry.Name).Validation.Formula1 -ne [string]$entry.Value) { throw "$($entry.Name) is not validated by $($entry.Value)" }
    }
    if ($sheet.Range($plan.layout.inputs).Locked) { throw "$($plan.layout.inputs) must be editable" }
    foreach ($address in $plan.layout.outputs) {
        if (-not $sheet.Range($address).Locked) { throw "$address must be locked" }
    }
    foreach ($row in 6..12) {
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
    foreach ($other in @(@('Explorer', 2), @('Probabilities', 1), @('Parameters', 2), @('Comparison', 1), @('Tails', 2))) {
        if ($book.Worksheets.Item($other[0]).ChartObjects().Count -ne $other[1]) { throw "The $($other[0]) sheet lost a chart" }
    }
    if (-not ([string]$sheet.Range("D$($rows.excel_build)").Value2).Contains('build')) { throw "The Excel build reads: $($sheet.Range("D$($rows.excel_build)").Value2)" }
    if ([string]$sheet.Range('D19').Value2 -ne 'OK') { throw "Status at the defaults: $($sheet.Range('D19').Value2)" }
    if ([string]$sheet.Range('D6').Value2 -ne [string]$plan.defaults.model -or [double]$sheet.Range('D7').Value2 -ne [double]$plan.defaults.trials -or
        [double]$sheet.Range('D11').Value2 -ne [double]$plan.defaults.threshold -or [double]$sheet.Range('D12').Value2 -ne [double]$plan.defaults.level) {
        throw 'The sheet was not saved with the default example'
    }
    $report.structure = 'PASS'

    # Every case: the results, the formulas beside them, the table, the charts and the colours.
    $excel.EnableEvents = $true
    $sheet.Activate()
    foreach ($case in $fixture.cases) {
        $where = [string]$case.name
        Set-Inputs $sheet $case.inputs ([string]$case.model)
        $excel.Calculate()
        if ([string]$sheet.Range('D19').Value2 -ne 'OK') { throw "${where}: status is $($sheet.Range('D19').Value2)" }
        $k = [int]$case.inputs.threshold
        $expectedQuantile = [int]$case.results.quantile
        $observation = [ordered]@{ case = $where; model = [string]$case.model }
        foreach ($key in $keys) {
            $value = $sheet.Range("D$($rows.$key)").Value2
            $expected = [double]$case.results.$key
            $observation[$key] = Shown $value
            if ($key -eq 'quantile') {
                if (-not ($value -is [double]) -or [double]$value -ne $expectedQuantile) { throw "${where}: the quantile is $(Shown $value), expected $expectedQuantile" }
            } elseif ($key -eq 'mean' -or $key -eq 'sd') {
                if ((Relative $value $expected) -gt 1E-13) { throw "${where}: the $key is $(Shown $value), expected $expected" }
            } else {
                Check-Probability "${where} $key" $value $expected
            }
            # The formula shown beside the result, evaluated again in a scratch cell.
            $text = [string]$sheet.Range("F$($rows.$key)").Value2
            $match = [regex]::Match($text, '=K_STATS_\w+\([^)]*\)')
            if (-not $match.Success) { throw "${where}: the formula of $key reads '$text'" }
            $skipped = ($k -eq 0 -and @('fewer_than', 'at_least') -contains $key) -or ($expectedQuantile -eq 0 -and $key -eq 'before_quantile')
            if (-not $skipped) {
                $sheet.Range('AQ1').Formula2 = $match.Value
                $excel.Calculate()
                if (-not (Same-Value $sheet.Range('AQ1').Value2 $value)) {
                    throw "${where}: $key shows $(Shown $value), its formula $($match.Value) gives $(Shown $sheet.Range('AQ1').Value2)"
                }
                $sheet.Range('AQ1').ClearContents() | Out-Null
            }
        }
        $check = [string]$sheet.Range("D$($rows.check)").Value2
        if (-not $check.Contains('is the smallest k with F(k) >= q')) { throw "${where}: the quantile check reads '$check'" }
        if ([string]::IsNullOrWhiteSpace([string]$sheet.Range("D$($rows.assumptions)").Value2)) { throw "${where}: no assumptions are shown" }

        # The table: the window of counts, then every reading against the references.
        $counts = @(Spill-Values $sheet "C$first")
        $mass = @(Spill-Values $sheet "D$first")
        $atMost = @(Spill-Values $sheet "F$first")
        $fewer = @(Spill-Values $sheet "G$first")
        $atLeast = @(Spill-Values $sheet "H$first")
        $more = @(Spill-Values $sheet "I$first")
        $lower = [int]$case.window.lower
        $upper = [int]$case.window.upper
        if ($counts.Count -ne $upper - $lower + 1) { throw "${where}: the table has $($counts.Count) rows, expected $($upper - $lower + 1) ($lower to $upper)" }
        $byCount = @{}
        foreach ($row in $case.table) { $byCount[[int]$row.k] = $row }
        foreach ($name in @('mass', 'atMost', 'fewer', 'atLeast', 'more')) {
            if (@((Get-Variable $name).Value).Count -ne $counts.Count) { throw "${where}: column $name has $(@((Get-Variable $name).Value).Count) rows" }
        }
        for ($index = 0; $index -lt $counts.Count; $index++) {
            $count = [int]$counts[$index]
            $label = "$where k = $count"
            if ($count -ne $lower + $index) { throw "${label}: expected the count $($lower + $index)" }
            $reference = $byCount[$count]
            Check-Probability "$label P(X = k)" $mass[$index] ([double]$reference.mass)
            Check-Probability "$label P(X <= k)" $atMost[$index] ([double]$reference.at_most)
            Check-Probability "$label P(X > k)" $more[$index] ([double]$reference.more_than)
            if ($count -eq 0) {
                if ([double]$fewer[$index] -ne 0 -or [double]$atLeast[$index] -ne 1) { throw "${label}: fewer than 0 is $($fewer[$index]) and at least 0 is $($atLeast[$index])" }
            } else {
                $previous = $byCount[$count - 1]
                Check-Probability "$label P(X < k)" $fewer[$index] ([double]$previous.at_most)
                Check-Probability "$label P(X >= k)" $atLeast[$index] ([double]$previous.more_than)
                if ($index -gt 0 -and ((-not (Same-Value $fewer[$index] $atMost[$index - 1])) -or (-not (Same-Value $atLeast[$index] $more[$index - 1])))) {
                    throw "${label}: P(X < k) and P(X >= k) are not the row above's P(X <= k) and P(X > k)"
                }
            }
            $report.table_rows++
        }
        # The charts follow the table, and their titles name the model.
        $massChart = $sheet.ChartObjects('DiscreteMass').Chart
        if ($massChart.SeriesCollection(1).Points().Count -ne $counts.Count) { throw "${where}: the mass chart has $($massChart.SeriesCollection(1).Points().Count) columns" }
        if (-not ([string]$massChart.ChartTitle.Text).Contains([string]$case.model)) { throw "${where}: the mass chart's title is '$($massChart.ChartTitle.Text)'" }
        $cumulative = $sheet.ChartObjects('DiscreteCumulative').Chart
        if ($cumulative.SeriesCollection(3).Points().Count -ne $counts.Count) { throw "${where}: the line of q has $($cumulative.SeriesCollection(3).Points().Count) points" }
        # The colours: the threshold's count in amber, the quantile's in green, nothing below the table;
        # the inputs of the other model greyed.
        $thresholdFill = [int]$sheet.Range("C$($first + $k - $lower)").DisplayFormat.Interior.Color
        if ($thresholdFill -ne $amber) { throw "${where}: the count k = $k shows the fill $thresholdFill" }
        if ($expectedQuantile -ne $k) {
            $quantileFill = [int]$sheet.Range("C$($first + $expectedQuantile - $lower)").DisplayFormat.Interior.Color
            if ($quantileFill -ne $green) { throw "${where}: the count Q = $expectedQuantile shows the fill $quantileFill" }
        }
        for ($index = 0; $index -lt $counts.Count + 3; $index++) {
            $count = $lower + $index
            if ($count -eq $k -or $count -eq $expectedQuantile) { continue }
            $fill = [int]$sheet.Range("C$($first + $index)").DisplayFormat.Interior.Color
            if ($fill -eq $amber -or $fill -eq $green) { throw "${where}: row $($first + $index) is marked ($fill)" }
        }
        $greyed = @('D9', 'D7')
        if ([string]$case.model -eq 'Poisson') { $greyed = @('D7', 'D9') }
        if ([int]$sheet.Range($greyed[0]).DisplayFormat.Interior.Color -ne $grey) { throw "${where}: $($greyed[0]) is not greyed" }
        if ([int]$sheet.Range($greyed[1]).DisplayFormat.Interior.Color -eq $grey) { throw "${where}: $($greyed[1]) is greyed" }
        $observation.window = "$lower to $upper"
        $observation.check = $check
        $report.observations += $observation
        $report.cases++
    }
    $report.cases_check = 'PASS'

    # Invalid inputs, one at a time from the default example: the status names the input and nothing is shown.
    [void]$excel.Run($prefix + 'ResetDiscreteInputs', $false)
    $invalid = @(
        @('D6', 'text', 'Negative binomial', 'Check: model'),
        @('D7', 'number', 10.5, 'Check: trials n'),
        @('D7', 'number', 0, 'Check: trials n'),
        @('D8', 'number', 0, 'Check: probability of success p'),
        @('D8', 'number', 1, 'Check: probability of success p'),
        @('D8', 'number', 1.5, 'Check: probability of success p'),
        @('D11', 'number', -1, 'Check: threshold k'),
        @('D11', 'number', 3.5, 'Check: threshold k'),
        @('D11', 'number', 101, 'Check: threshold k'),
        @('D12', 'number', 0, 'Check: probability q'),
        @('D12', 'number', 1, 'Check: probability q')
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
    # The Poisson inputs: they are checked only when the Poisson is chosen.
    Set-Number $sheet 'D9' 0
    $excel.Calculate()
    if ([string]$sheet.Range('D19').Value2 -ne 'OK') { throw 'The Binomial reads the Poisson rate' }
    Set-Text $sheet 'D6' 'Poisson'
    $excel.Calculate()
    if ([string]$sheet.Range('D19').Value2 -ne 'Check: event rate') { throw "A zero rate gives: $($sheet.Range('D19').Value2)" }
    Set-Number $sheet 'D9' 0.02
    Set-Number $sheet 'D10' -1
    $excel.Calculate()
    if ([string]$sheet.Range('D19').Value2 -ne 'Check: exposure') { throw "A negative exposure gives: $($sheet.Range('D19').Value2)" }
    Set-Number $sheet 'D9' 1000
    Set-Number $sheet 'D10' 1000
    $excel.Calculate()
    if ([string]$sheet.Range('D19').Value2 -ne 'Check: rate x exposure') { throw "A mean of 1E6 gives: $($sheet.Range('D19').Value2)" }
    # A table that would pass its rows: 10000 trials at p = 0.5 with k = 0 spans 0 to 5250.
    [void]$excel.Run($prefix + 'ResetDiscreteInputs', $false)
    Set-Number $sheet 'D7' 10000
    Set-Number $sheet 'D8' 0.5
    Set-Number $sheet 'D11' 0
    $excel.Calculate()
    if ([string]$sheet.Range('D19').Value2 -ne "Check: the table would pass $($plan.max_rows) rows") { throw "A table too wide gives: $($sheet.Range('D19').Value2)" }
    $report.invalid = 'PASS'

    # Reset.
    [void]$excel.Run($prefix + 'ResetDiscreteInputs', $false)
    $excel.Calculate()
    if ([string]$sheet.Range('D6').Value2 -ne [string]$plan.defaults.model -or [double]$sheet.Range('D7').Value2 -ne [double]$plan.defaults.trials -or
        [double]$sheet.Range('D8').Value2 -ne [double]$plan.defaults.probability -or [double]$sheet.Range('D11').Value2 -ne [double]$plan.defaults.threshold -or
        [string]$sheet.Range('D19').Value2 -ne 'OK') {
        throw 'The reset did not restore the defaults'
    }
    $report.reset = 'PASS'
    $report.status = 'PASS'
} catch {
    $report.error = $_.Exception.Message
    throw
} finally {
    $report | ConvertTo-Json -Depth 5 | Set-Content -Encoding UTF8 (Join-Path $directory 'discrete-report.json')
    ([ordered]@{ status = $report.status; cases = $report.cases; table_rows = $report.table_rows;
                 largest_relative_error = $report.largest_relative_error; notes = $report.notes; error = $report.error;
                 excel_build = $report.excel_build }) | ConvertTo-Json -Depth 3 | Write-Output
    if ($null -ne $book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
}
