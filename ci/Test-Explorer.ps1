param([Parameter(Mandatory=$true)][string]$PackageDirectory)
# The Explorer against independent references (examples/demo/explorer-references.json,
# made with mpmath): every family's probe values, its default grid, the charts that
# follow the spilled columns, the validation rules, an invalid input and the reset.
$ErrorActionPreference = 'Stop'
$directory = (Resolve-Path $PackageDirectory).Path
$path = Join-Path $directory 'Probability-Distributions-Demo.xlsm'
$fixture = Get-Content (Join-Path $PSScriptRoot '../examples/demo/explorer-references.json') -Raw | ConvertFrom-Json
$plan = Get-Content (Join-Path $directory 'explorer-plan.json') -Raw | ConvertFrom-Json
$report = [ordered]@{ status = 'FAIL'; sha256 = (Get-FileHash $path -Algorithm SHA256).Hash; cases = 0; quantiles = 0; grids = @(); notes = @() }
$excel = $null
$book = $null
$other = $null

function Is-Error($Value) { return ($Value -is [int]) -and ($Value -lt -2146820000) }
function Read-Number($Sheet, [string]$Address) {
    $value = $Sheet.Range($Address).Value2
    if (Is-Error $value) { return $null }
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
function Chart-Points($Sheet, [string]$Name) {
    return @($Sheet.ChartObjects($Name).Chart.SeriesCollection(1).Values).Count
}
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($path, 0, $false)
    $sheet = $book.Worksheets.Item('Explorer')
    $prefix = "'" + $book.Name + "'!"
    $report.excel_version = $excel.Version
    $report.excel_build = $excel.Build
    $families = @($plan.families | ForEach-Object { $_.name })
    if ($families.Count -ne 16) { throw "The plan lists $($families.Count) families, expected 16" }

    # Structure: no implicit intersection, one list of 16 names, locked outputs, unlocked inputs.
    foreach ($formula in $sheet.Range('A1:P40').Formula2) {
        if ([string]$formula -like '=*' -and [string]$formula -match '@|_xlfn.SINGLE') { throw "Implicit intersection remains: $formula" }
    }
    $validation = $sheet.Range('D6').Validation
    if ($validation.Type -ne 3 -or $validation.Formula1 -ne '=ExplorerFamily') { throw 'The distribution list is not backed by the family table' }
    $listed = @($sheet.Range('ExplorerFamily').Cells | ForEach-Object { [string]$_.Value2 })
    if (($listed -join '|') -ne ($families -join '|')) { throw 'The family table does not match the plan' }
    foreach ($address in @('D6:D16')) {
        if ($sheet.Range($address).Locked) { throw "$address must be editable" }
    }
    foreach ($address in @('D19:D25', 'F20:I1019')) {
        if (-not $sheet.Range($address).Locked) { throw "$address must be locked" }
    }
    if ($sheet.Range('C26').MergeArea.Address($false, $false) -ne 'C26:D27') { throw "The support text is not one cell over C26:D27: $($sheet.Range('C26').MergeArea.Address($false, $false))" }
    if ($sheet.Range('D20').FormatConditions.Count -ne 1 -or $sheet.Range('D19').FormatConditions.Count -ne 2) { throw 'The status or macros cell has duplicated conditional formats' }
    $report.structure = 'PASS'

    # Every family at its defaults: window, spilled grid, sample points, charts.
    foreach ($grid in $fixture.grids) {
        $excel.EnableEvents = $true
        Set-Text $sheet 'D6' $grid.family
        $excel.Calculate()
        $row = [ordered]@{ family = $grid.family; status = 'FAIL' }
        if ([string]$sheet.Range('D19').Value2 -ne 'OK') { throw "$($grid.family): status is $($sheet.Range('D19').Value2)" }
        for ($slot = 0; $slot -lt $grid.parameters.Count; $slot++) {
            if ((Read-Number $sheet ('D' + (7 + $slot))) -ne [double]$grid.parameters[$slot]) { throw "$($grid.family): default parameter $($slot + 1) differs from the reference set" }
        }
        if (-not (Near (Read-Number $sheet 'D13') $grid.lower) -or -not (Near (Read-Number $sheet 'D14') $grid.upper)) {
            throw "$($grid.family): plot window $($sheet.Range('D13').Value2) .. $($sheet.Range('D14').Value2), expected $($grid.lower) .. $($grid.upper)"
        }
        $rows = Grid-Rows $sheet
        if ($rows -ne [int]$grid.rows) { throw "$($grid.family): the grid has $rows rows, expected $($grid.rows)" }
        foreach ($column in @('G', 'H', 'I')) {
            $spill = $sheet.Range($column + '20').SpillingToRange
            if ($null -eq $spill -or [int]$spill.Rows.Count -ne $rows) { throw "$($grid.family): column $column does not spill over the same rows as x" }
        }
        $errors = 0
        $values = $sheet.Range('G20').SpillingToRange.Value2
        foreach ($value in $values) { if (Is-Error $value) { $errors++ } }
        $row.density_errors = $errors
        foreach ($sample in $grid.sample) {
            $address = 20 + [int]$sample.index
            if (-not (Near (Read-Number $sheet "F$address") $sample.x)) { throw "$($grid.family): grid x at row $address is $($sheet.Range("F$address").Value2), expected $($sample.x)" }
            if ($null -eq $sample.density) { continue }
            if (-not (Near (Read-Number $sheet "G$address") $sample.density)) { throw "$($grid.family): grid density at x=$($sample.x) is $($sheet.Range("G$address").Value2), expected $($sample.density)" }
            if (-not (Near (Read-Number $sheet "I$address") $sample.cdf)) { throw "$($grid.family): grid CDF at x=$($sample.x) is $($sheet.Range("I$address").Value2), expected $($sample.cdf)" }
            if (-not (Near (Read-Number $sheet "H$address") $sample.survival)) { throw "$($grid.family): grid survival at x=$($sample.x) is $($sheet.Range("H$address").Value2), expected $($sample.survival)" }
        }
        $discrete = ([string]$sheet.Range('ExplorerKind').Cells($families.IndexOf($grid.family) + 1, 1).Value2) -eq 'Discrete'
        $types = if ($discrete) { @(51, -4169) } else { @(75, 75) }
        foreach ($pair in @(@('ExplorerDensity', $types[0]), @('ExplorerCDF', $types[1]))) {
            $chart = $sheet.ChartObjects($pair[0]).Chart
            if ($chart.ChartType -ne $pair[1]) { throw "$($grid.family): $($pair[0]) has chart type $($chart.ChartType), expected $($pair[1])" }
            if ((Chart-Points $sheet $pair[0]) -ne $rows) { throw "$($grid.family): $($pair[0]) plots $(Chart-Points $sheet $pair[0]) points, expected $rows" }
        }
        $row.rows = $rows
        $row.status = 'PASS'
        $report.grids += $row
    }

    # Probe values: every case of the fixture, several parameter sets per family.
    $excel.EnableEvents = $false
    Set-Number $sheet 'D15' 2.0
    Set-Number $sheet 'D13' 0.0
    Set-Number $sheet 'D14' 3.0
    foreach ($case in $fixture.cases) {
        Set-Text $sheet 'D6' $case.family
        for ($slot = 0; $slot -lt $case.parameters.Count; $slot++) { Set-Number $sheet ('D' + (7 + $slot)) ([double]$case.parameters[$slot]) }
        Set-Number $sheet 'D11' ([double]$case.x)
        $excel.Calculate()
        if ([string]$sheet.Range('D19').Value2 -ne 'OK') { throw "$($case.family) $($case.parameters -join ','): a valid case is rejected: $($sheet.Range('D19').Value2)" }
        $label = "$($case.family)($($case.parameters -join ',')) at $($case.x)"
        $density = Read-Number $sheet 'D21'
        $cdf = Read-Number $sheet 'D22'
        $survival = Read-Number $sheet 'D23'
        if (-not (Near $density $case.density)) { throw "$label density $($sheet.Range('D21').Value2), expected $($case.density)" }
        if (-not (Near $cdf $case.cdf)) { throw "$label CDF $($sheet.Range('D22').Value2), expected $($case.cdf)" }
        if (-not (Near $survival $case.survival)) { throw "$label survival $($sheet.Range('D23').Value2), expected $($case.survival)" }
        $report.cases++
    }
    foreach ($case in $fixture.quantiles) {
        Set-Text $sheet 'D6' $case.family
        for ($slot = 0; $slot -lt $case.parameters.Count; $slot++) { Set-Number $sheet ('D' + (7 + $slot)) ([double]$case.parameters[$slot]) }
        Set-Number $sheet 'D12' ([double]$case.p)
        $excel.Calculate()
        $quantile = Read-Number $sheet 'D24'
        if (-not (Near $quantile $case.quantile)) { throw "$($case.family)($($case.parameters -join ',')) quantile at $($case.p) is $($sheet.Range('D24').Value2), expected $($case.quantile)" }
        $report.quantiles++
    }
    Set-Number $sheet 'D12' 1.5
    $excel.Calculate()
    if (-not (Is-Error $sheet.Range('D24').Value2) -or [string]$sheet.Range('D19').Value2 -ne 'OK') { throw 'A probability above 1 gives a quantile, or stops the other results' }
    Set-Number $sheet 'D12' 0.5
    # A negative fractional lower bound of a discrete grid starts at the next integer up.
    Set-Text $sheet 'D6' 'Discrete uniform'
    Set-Number $sheet 'D7' -5.0
    Set-Number $sheet 'D8' 5.0
    Set-Number $sheet 'D13' -2.5
    Set-Number $sheet 'D14' 2.0
    $excel.Calculate()
    if ((Read-Number $sheet 'F20') -ne -2.0 -or (Grid-Rows $sheet) -ne 5) { throw "Discrete window -2.5 to 2 starts at $($sheet.Range('F20').Value2) with $(Grid-Rows $sheet) rows" }
    Set-Number $sheet 'D13' -0.5
    Set-Number $sheet 'D14' -0.1
    $excel.Calculate()
    if ([string]$sheet.Range('D19').Value2 -ne 'Check: plot window') { throw 'A discrete window with no integer is accepted' }
    $excel.EnableEvents = $true

    # The input event: a parameter edit moves the plot window and the spilled grid with it.
    [void]$excel.Run($prefix + 'ResetExplorerInputs', $false)
    Set-Number $sheet 'D8' 2.0
    if ([Math]::Abs((Read-Number $sheet 'D21') - 0.19947114020071635) -gt 1e-12) { throw 'Input event did not recalculate the Normal density' }
    if ((Read-Number $sheet 'D13') -ne -8.0 -or (Read-Number $sheet 'D14') -ne 8.0) { throw 'The plot window did not follow sigma' }
    if ((Read-Number $sheet 'F20') -ne -8.0 -or (Read-Number $sheet 'F220') -ne 8.0) { throw 'The grid did not follow the plot window' }
    $values = @($sheet.ChartObjects('ExplorerDensity').Chart.SeriesCollection(1).Values)
    if ([Math]::Abs([double]$values[100] - (Read-Number $sheet 'G120')) -gt 1e-12) { throw 'The density chart did not follow the grid' }
    Set-Number $sheet 'D15' 50.0
    if ((Grid-Rows $sheet) -ne 50 -or (Chart-Points $sheet 'ExplorerCDF') -ne 50) { throw 'The grid and the charts did not follow the number of points' }
    $report.edits_move_grid_and_charts = 'PASS'

    # An invalid input is named and the results stop, instead of showing wrong numbers.
    Set-Number $sheet 'D8' -1.0
    if ([string]$sheet.Range('D19').Value2 -ne 'Check: Standard deviation') { throw "Invalid sigma gives: $($sheet.Range('D19').Value2)" }
    if (-not (Is-Error $sheet.Range('D21').Value2)) { throw 'A result is shown for an invalid input' }
    # A write through COM bypasses data validation, and Validation.Value does not
    # evaluate a custom rule, so the rule is read as the cell formula it is.
    $rule = $sheet.Range('NE1')
    function Rule-Holds([string]$Address) {
        $slot = [int]$Address.Substring(1) - 6
        $rule.Formula2 = "=ExplorerParameterOk$slot"
        return [bool]$rule.Value2
    }
    foreach ($address in @('D7', 'D8', 'D9', 'D10')) {
        if ($sheet.Range($address).Validation.Type -ne 7 -or $sheet.Range($address).Validation.Formula1 -ne ('=ExplorerParameterOk' + ([int]$address.Substring(1) - 6))) { throw "$address has no parameter rule" }
    }
    foreach ($test in @(@('D8', 0.0, $false), @('D8', 3.0, $true), @('D8', 20000.0, $false), @('D7', 2000000.0, $false), @('D7', -5.0, $true))) {
        Set-Number $sheet $test[0] $test[1]
        if ((Rule-Holds $test[0]) -ne $test[2]) { throw "The rule of $($test[0]) = $($test[1]) should be $($test[2])" }
    }
    Set-Text $sheet 'D6' 'Binomial'
    Set-Number $sheet 'D7' 12.5
    if (Rule-Holds 'D7') { throw 'A fractional number of trials passes the rule' }
    Set-Number $sheet 'D7' 12.0
    if (-not (Rule-Holds 'D7')) { throw 'A whole number of trials fails the rule' }
    Set-Number $sheet 'D8' 1.5
    if (Rule-Holds 'D8') { throw 'A probability above 1 passes the rule' }
    if ([string]$sheet.Range('D19').Value2 -ne 'Check: Probability of success') { throw "Invalid probability gives: $($sheet.Range('D19').Value2)" }
    $rule.ClearContents() | Out-Null
    Set-Text $sheet 'D6' 'Hypergeometric'
    Set-Number $sheet 'D7' 10.0
    Set-Number $sheet 'D8' 30.0
    Set-Number $sheet 'D9' 20.0
    if (([string]$sheet.Range('D19').Value2) -notlike 'Check:*') { throw 'A sample larger than the population is accepted' }
    Set-Text $sheet 'D6' 'Normal,Exponential'
    if ($sheet.Range('D6').Validation.Value) { throw 'The validation accepts a made-up distribution name' }
    $report.invalid_input = 'PASS'

    # The number of decimals sets the format of the results and the grid, through conditional formats.
    [void]$excel.Run($prefix + 'ResetExplorerInputs', $false)
    $dec = [string]$excel.International(3)
    foreach ($test in @(@(3.0, ('0' + $dec + '399')), @(0.0, '0'), @(10.0, ('0' + $dec + '3989422804')))) {
        Set-Number $sheet 'D16' $test[0]
        if ([string]$sheet.Range('D21').Text -ne $test[1]) { throw "With $($test[0]) decimals D21 shows $($sheet.Range('D21').Text), expected $($test[1])" }
    }
    Set-Number $sheet 'D16' 2.0
    if ([string]$sheet.Range('F20').Text -ne ('-4' + $dec + '00') -or [string]$sheet.Range('H120').Text -ne ('0' + $dec + '50')) { throw "The grid does not follow the decimals: $($sheet.Range('F20').Text), $($sheet.Range('H120').Text)" }
    Set-Text $sheet 'D6' 'Poisson'
    if ([string]$sheet.Range('F21').Text -ne '1') { throw "A discrete x shows $($sheet.Range('F21').Text), not a whole number" }
    $report.decimals = 'PASS'

    # The reset restores Normal(0,1) and touches no other workbook.
    $other = $excel.Workbooks.Add()
    $other.Worksheets.Item(1).Range('D6').Value2 = 'sentinel'
    $other.Activate()
    [void]$excel.Run($prefix + 'ResetExplorerInputs', $false)
    if ($other.Worksheets.Item(1).Range('D6').Value2 -ne 'sentinel') { throw 'Reset changed another workbook' }
    $sheet = $book.Worksheets.Item('Explorer')
    if ($sheet.Range('D6').Value2 -ne 'Normal' -or $sheet.Range('D8').Value2 -ne 1.0 -or (Read-Number $sheet 'D15') -ne 201.0 -or (Read-Number $sheet 'D16') -ne 6.0 -or (Read-Number $sheet 'D12') -ne 0.95) { throw 'Reset did not restore the defaults' }
    if ((Grid-Rows $sheet) -ne 201) { throw 'Reset did not restore the grid' }
    $report.reset_isolation = 'PASS'
    if (-not $excel.EnableEvents) { throw 'Events were left off' }
    $report.status = 'PASS'
} catch {
    $report.error = $_.Exception.Message
    throw
} finally {
    if ($null -ne $other) { $other.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($other) }
    if ($null -ne $book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    $report | ConvertTo-Json -Depth 6 | Set-Content -Encoding UTF8 (Join-Path $directory 'explorer-report.json')
    $report | ConvertTo-Json -Depth 6 | Write-Output
}
