param([Parameter(Mandatory=$true)][string]$PackageDirectory)
# The Diagnostics sheet (errors and diagnostics) against its cases (examples/demo/diagnostics-cases.json: the
# outcome the documentation gives each call, and mpmath references for the numbers). The case table: every
# call returns its documented outcome (a number within 1E-12 relative, or the documented error), its check
# reads As documented in green, its Status is the library's (empty for a number, naming the argument or the
# condition for an error, "Not called" for a text), and the call as written gives the same value when it is
# entered again in a scratch cell. The try-it area: each trial typed into the inputs with events on (a new
# function loads its example), the value, its kind, the Status, the meaning, the arguments Excel converted,
# the sweep and the chart. Then invalid inputs, the Excel state and the navigation after them, and the reset.
# The report records what this Excel returned for every call, with its version and build.
$ErrorActionPreference = 'Stop'
$directory = (Resolve-Path $PackageDirectory).Path
$path = Join-Path $directory 'Probability-Distributions-Demo.xlsm'
$fixture = Get-Content (Join-Path $PSScriptRoot '../examples/demo/diagnostics-cases.json') -Raw | ConvertFrom-Json
$plan = Get-Content (Join-Path $directory 'diagnostics-plan.json') -Raw | ConvertFrom-Json
$report = [ordered]@{ status = 'FAIL'; sha256 = (Get-FileHash $path -Algorithm SHA256).Hash; cases = 0; trials = 0;
                      largest_relative_error = 0.0; notes = @(); observations = @() }
$excel = $null
$book = $null
$invariant = [Globalization.CultureInfo]::InvariantCulture
# The numbers of the case table are closed forms, limits and exact zeros: 1E-12 relative holds the loosest
# contract of those functions and fails a wrong argument or a wrong tail.
$tolerance = 1E-12
$scratch = 'AZ2'
# The values COM gives for the worksheet errors.
$errorCodes = @{ '#NUM!' = -2146826252; '#VALUE!' = -2146826273; '#N/A' = -2146826246 }

function Is-Error($Value) { return ($Value -is [int]) -and ($Value -lt -2146820000) }
# Each write runs the sheet's macro and recalculates; Excel rejects a COM call while it is still busy
# (RPC_E_CALL_REJECTED). A write is retried until Excel accepts it, for up to five minutes.
function Invoke-WhenReady([scriptblock]$Action) {
    for ($attempt = 1; ; $attempt++) {
        try { return & $Action }
        catch {
            $text = [string]$_.Exception + ' ' + [string]$_.Exception.InnerException
            $busy = $_.Exception.HResult -eq -2147418111 -or $text -match '0x80010001|RPC_E_CALL_REJECTED'
            if (-not $busy -or $attempt -ge 600) { throw }
            Start-Sleep -Milliseconds 500
        }
    }
}
# One call site per value type: PowerShell binds a COM setter to the first type it sees.
function Set-Text($Sheet, [string]$Address, [string]$Value) { Invoke-WhenReady { $Sheet.Range($Address).Value2 = $Value } }
function Set-Number($Sheet, [string]$Address, [double]$Value) { Invoke-WhenReady { $Sheet.Range($Address).Value2 = $Value } }
function Set-Argument($Sheet, [string]$Address, $Value) {
    if ($Value -is [System.Management.Automation.PSCustomObject]) {
        if ($Value.empty) { $Sheet.Range($Address).ClearContents() | Out-Null } else { Set-Text $Sheet $Address ([string]$Value.text) }
    } else { Set-Number $Sheet $Address ([double]$Value) }
}
function Shown($Value) {
    if ($Value -is [double]) { return $Value.ToString('R', $invariant) }
    if (Is-Error $Value) { foreach ($entry in $errorCodes.GetEnumerator()) { if ($Value -eq $entry.Value) { return $entry.Key } }; return 'error ' + $Value }
    return [string]$Value
}
function Same-Value($Actual, $Expected) {
    if ($Actual -is [string] -or $Expected -is [string]) { return ([string]$Actual -ceq [string]$Expected) }
    if ((Is-Error $Actual) -or (Is-Error $Expected)) { return ($Actual -eq $Expected) }
    return [double]$Actual -eq [double]$Expected
}
# A value against its documented outcome: the error itself, or a number within the tolerance.
function Check-Outcome([string]$Label, $Actual, $Expected) {
    if ($Expected.kind -ne 'number') {
        if (-not (Is-Error $Actual) -or $Actual -ne $errorCodes[[string]$Expected.kind]) { throw "${Label}: $(Shown $Actual), documented $($Expected.kind)" }
        return
    }
    if (-not ($Actual -is [double])) { throw "${Label}: $(Shown $Actual), documented the number $($Expected.value)" }
    $reference = [double]$Expected.value
    $gap = [Math]::Abs([double]$Actual - $reference)
    if ($reference -ne 0) { $gap = $gap / [Math]::Abs($reference) }
    if ($gap -gt $tolerance) { throw "${Label}: $(Shown $Actual), documented $reference (relative error $gap)" }
    if ($gap -gt $report.largest_relative_error) { $report.largest_relative_error = $gap }
}
# The Status written for an outcome: empty for a number, the fragment of the library's message for an error.
function Check-Status([string]$Label, [string]$Status, $Expected, $Fragment, [string]$Empty) {
    if ($Expected.kind -eq 'number') {
        if ($Status -ne $Empty) { throw "${Label}: a number came back with the Status '$Status'" }
        return
    }
    if ([string]::IsNullOrWhiteSpace($Status) -or $Status -eq $Empty) { throw "${Label}: $($Expected.kind) came back with no Status" }
    if ($Status.StartsWith('Unexpected')) { throw "${Label}: the library reports an unexpected failure: $Status" }
    if ($Fragment -and -not $Status.Contains([string]$Fragment)) { throw "${Label}: the Status '$Status' does not name '$Fragment'" }
}
function Evaluate-Again($Sheet, [string]$Text) {
    $Sheet.Range($scratch).Formula2 = $Text
    $excel.Calculate()
    $value = $Sheet.Range($scratch).Value2
    $Sheet.Range($scratch).ClearContents() | Out-Null
    return $value
}
function Excel-State() {
    return [ordered]@{ events = $excel.EnableEvents; screen = $excel.ScreenUpdating; calculation = $excel.Calculation;
                       alerts = $excel.DisplayAlerts; workbook = $excel.ActiveWorkbook.Name }
}
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($path, 0, $false)
    # Excel may still be loading or calculating the sheets after the workbook opens: a COM call is rejected until it is ready.
    for ($wait = 0; $wait -lt 240; $wait++) { try { if ($excel.Ready) { break } } catch { }; Start-Sleep -Milliseconds 500 }
    $sheet = $book.Worksheets.Item('Diagnostics')
    $prefix = "'" + $book.Name + "'!"
    $report.excel_version = $excel.Version
    $report.excel_build = $excel.Build
    $report.excel_os = $excel.OperatingSystem
    $at = $plan.layout
    $first = [int]$at.first_row
    $rows = $at.result_rows
    $columns = $at.table

    # Structure: no implicit intersection, the promised cells, the catalogue, locked outputs, unlocked inputs,
    # the messages, the chart.
    foreach ($formula in $sheet.Range("A1:AX$($at.last_row)").Formula2) {
        if ([string]$formula -like '=*' -and [string]$formula -match '@|_xlfn.SINGLE') { throw "Implicit intersection remains: $formula" }
    }
    if ($sheet.Range($at.function).Validation.Formula1 -ne [string]$plan.validation.function_list) { throw 'The function is not chosen from the list' }
    foreach ($entry in $plan.validation.custom.PSObject.Properties) {
        if ($sheet.Range($entry.Name).Validation.Formula1 -ne [string]$entry.Value) { throw "$($entry.Name) is not validated by $($entry.Value)" }
    }
    foreach ($address in $plan.validation.input_only) {
        $rule = $sheet.Range([string]$address).Validation
        if ($rule.Type -ne 0) { throw "$address must take any value (type $($rule.Type))" }
    }
    foreach ($entry in $plan.messages.PSObject.Properties) {
        if ([string]::IsNullOrWhiteSpace([string]$sheet.Range($entry.Name).Validation.InputMessage)) { throw "$($entry.Name) lacks an input message" }
    }
    if ($sheet.Range($at.inputs).Locked) { throw "$($at.inputs) must be editable" }
    foreach ($address in $at.outputs) {
        if (-not $sheet.Range($address).Locked) { throw "$address must be locked" }
    }
    $names = @($book.Names.Item('DiagFunctionList').RefersToRange.Value2)
    if ($names.Count -ne @($fixture.functions).Count) { throw "The catalogue holds $($names.Count) functions, expected $(@($fixture.functions).Count)" }
    for ($index = 0; $index -lt $names.Count; $index++) {
        if ([string]$names[$index] -ne [string]$fixture.functions[$index].name) { throw "Catalogue row $index is $($names[$index])" }
    }
    if ($sheet.ChartObjects().Count -ne 1) { throw "The sheet has $($sheet.ChartObjects().Count) charts, expected 1" }
    $chart = $sheet.ChartObjects('DiagSweep').Chart
    if ($chart.SeriesCollection().Count -ne @($plan.charts[0].series).Count) { throw "The chart has $($chart.SeriesCollection().Count) series, expected $(@($plan.charts[0].series).Count)" }
    foreach ($rule in @($sheet.Cells.FormatConditions)) {
        if ($rule.Type -eq 2 -and [string]$rule.Formula1 -match 'Explorer') { throw "A conditional rule reads the Explorer: $($rule.Formula1)" }
    }
    foreach ($other in @(@('Explorer', 2), @('Probabilities', 1), @('Parameters', 2), @('Comparison', 1), @('Discrete', 2), @('Risk', 2))) {
        if ($book.Worksheets.Item($other[0]).ChartObjects().Count -ne $other[1]) { throw "The $($other[0]) sheet lost a chart" }
    }
    # Every run builds the Tails sheet, a change to the release branch too: it has its two charts.
    $tailsCharts = $book.Worksheets.Item('Tails').ChartObjects().Count
    if ($tailsCharts -ne 2) { throw "The Tails sheet has $tailsCharts charts, expected 2" }
    if (-not ([string]$sheet.Range("D$($rows.excel_build)").Value2).Contains('build')) { throw "The Excel build reads: $($sheet.Range("D$($rows.excel_build)").Value2)" }
    if ([string]$sheet.Range($at.status).Value2 -ne 'OK') { throw "Status at the defaults: $($sheet.Range($at.status).Value2)" }
    if ([string]$sheet.Range($at.function).Value2 -ne [string]$plan.defaults.function) { throw 'The sheet was not saved with the default function' }
    $report.structure = 'PASS'

    # The case table: every call against its documented outcome.
    $excel.Calculate()
    $index = 0
    foreach ($case in $fixture.cases) {
        $row = $first + $index
        $label = "$($case.id) $($case.title)"
        $returned = $sheet.Range("$($columns.returned)$row").Value2
        Check-Outcome $label $returned $case.expected
        $check = [string]$sheet.Range("$($columns.check)$row").Value2
        if ($check -ne 'As documented') { throw "${label}: the check reads '$check'" }
        if ([int]$sheet.Range("$($columns.check)$row").DisplayFormat.Interior.Color -ne 0xCEEFC6) { throw "${label}: the check is not green" }
        $status = [string]$sheet.Range("$($columns.status)$row").Value2
        if ($case.group -eq 'excel' -and $case.expected.kind -eq '#VALUE!') {
            if (-not $status.StartsWith('Not called')) { throw "${label}: a text argument gives the Status '$status'" }
        } else {
            Check-Status $label $status $case.expected $case.status_contains '(empty)'
        }
        $text = [string]$sheet.Range("$($columns.call)$row").Value2
        if (-not $text.StartsWith('=' + [string]$case.function + '(')) { throw "${label}: the call reads '$text'" }
        $again = Evaluate-Again $sheet $text
        if (-not (Same-Value $again $returned)) { throw "${label}: the cell holds $(Shown $returned), the call $text gives $(Shown $again)" }
        if ([string]::IsNullOrWhiteSpace([string]$sheet.Range("$($columns.cause)$row").Value2)) { throw "${label}: no cause is given" }
        $report.observations += [ordered]@{ case = [string]$case.id; call = $text; returned = (Shown $returned); status = $status }
        $report.cases++
        $index++
    }
    $count = @($fixture.cases).Count
    $summary = [string]$sheet.Range("D$($rows.cases)").Value2
    if (-not $summary.Contains("$count of $count calls as documented, 0 unexpected")) { throw "The case table line reads '$summary'" }
    $report.cases_check = 'PASS'

    # The try-it area, with events on: a new function loads its example, then the trial's arguments.
    $excel.EnableEvents = $true
    $sheet.Activate()
    $empty = '(empty: the library wrote no failure)'
    foreach ($trial in $fixture.trials) {
        $label = "trial '$($trial.name)'"
        Set-Text $sheet $at.function ([string]$trial.function)
        $excel.Calculate()
        $catalogue = $fixture.functions | Where-Object { $_.name -eq $trial.function }
        for ($slot = 0; $slot -lt @($catalogue.example).Count; $slot++) {
            if ([double]$sheet.Range([string]$at.arguments[$slot]).Value2 -ne [double]$catalogue.example[$slot]) { throw "${label}: the function did not load its example argument $($slot + 1)" }
        }
        if ([double]$sheet.Range([string]$at.sweep[0]).Value2 -ne [double]$catalogue.sweep[0] -or
            [double]$sheet.Range([string]$at.sweep[1]).Value2 -ne [double]$catalogue.sweep[1]) { throw "${label}: the function did not load its sweep" }
        for ($slot = 0; $slot -lt @($trial.args).Count; $slot++) { Set-Argument $sheet ([string]$at.arguments[$slot]) $trial.args[$slot] }
        $excel.Calculate()
        if ([string]$sheet.Range($at.status).Value2 -ne 'OK') { throw "${label}: status is $($sheet.Range($at.status).Value2)" }
        $returned = $sheet.Range("D$($rows.returned)").Value2
        Check-Outcome $label $returned $trial.expected
        $kind = [string]$sheet.Range("D$($rows.kind)").Value2
        $documented = if ($trial.expected.kind -eq 'number') { 'a number' } else { [string]$trial.expected.kind }
        if ($kind -ne $documented) { throw "${label}: the kind reads '$kind', expected '$documented'" }
        $status = [string]$sheet.Range("D$($rows.status)").Value2
        if ($trial.status_contains -eq 'Not called') {
            if (-not $status.StartsWith('Not called')) { throw "${label}: the Status reads '$status'" }
        } else {
            Check-Status $label $status $trial.expected $trial.status_contains $empty
        }
        foreach ($key in @('meaning', 'action', 'domain', 'converted', 'sweep')) {
            if ([string]::IsNullOrWhiteSpace([string]$sheet.Range("D$($rows.$key)").Value2)) { throw "${label}: $key is empty" }
        }
        $call = [string]$sheet.Range("D$($rows.call)").Value2
        if (-not $call.StartsWith('=' + [string]$trial.function + '(')) { throw "${label}: the call reads '$call'" }
        $formula = [string]$sheet.Range("F$($rows.returned)").Value2
        $again = Evaluate-Again $sheet $formula
        if (-not (Same-Value $again $returned)) { throw "${label}: the value is $(Shown $returned), its formula $formula gives $(Shown $again)" }
        $converted = [string]$sheet.Range("D$($rows.converted)").Value2
        $hasEmpty = @($trial.args | Where-Object { $_ -is [System.Management.Automation.PSCustomObject] -and $_.empty }).Count -gt 0
        $hasText = @($trial.args | Where-Object { $_ -is [System.Management.Automation.PSCustomObject] -and $_.text }).Count -gt 0
        if ($hasEmpty -ne $converted.Contains('is empty')) { throw "${label}: the converted arguments read '$converted'" }
        if ($hasText -ne $converted.Contains('is text')) { throw "${label}: the converted arguments read '$converted'" }
        if (-not $hasEmpty -and -not $hasText -and -not $converted.StartsWith('None')) { throw "${label}: the converted arguments read '$converted'" }
        $points = $chart.SeriesCollection(1).Points().Count
        if ($points -ne [int]$at.sweep_points) { throw "${label}: the chart has $points points" }
        if (-not ([string]$chart.ChartTitle.Text).Contains([string]$trial.function)) { throw "${label}: the chart's title is '$($chart.ChartTitle.Text)'" }
        $report.observations += [ordered]@{ trial = [string]$trial.name; call = $call; returned = (Shown $returned); kind = $kind;
                                            status = $status; sweep = [string]$sheet.Range("D$($rows.sweep)").Value2 }
        $report.trials++
    }
    $report.trials_check = 'PASS'

    # The default sweep: the probability from 0 to 1 is refused at both ends.
    [void]$excel.Run($prefix + 'ResetDiagnosticsInputs', $false)
    $excel.Calculate()
    $sweep = [string]$sheet.Range("D$($rows.sweep)").Value2
    $inside = [int]$at.sweep_points - 2
    if (-not $sweep.Contains(": $inside numbers, 2 #NUM!")) { throw "The default sweep reads '$sweep'" }
    $raw = $sheet.Range('AG' + $first).SpillingToRange.Value2
    # The sweep holds a number, or the name of the error as a text: the LAMBDA returns no error value.
    for ($k = 1; $k -le [int]$at.sweep_points; $k++) {
        $refused = ($k -eq 1) -or ($k -eq [int]$at.sweep_points)
        if ($refused -and [string]$raw[$k, 1] -ne '#NUM!') { throw "Sweep row ${k}: $(Shown $raw[$k, 1]), expected the text #NUM!" }
        if (-not $refused -and -not ($raw[$k, 1] -is [double])) { throw "Sweep row ${k}: $(Shown $raw[$k, 1]), expected a number" }
    }
    $report.sweep = 'PASS'

    # A function name typed in another case is the same function: Excel accepts it, and so do the example
    # loader and the Status reader.
    $lower = ([string]$plan.defaults.function).ToLowerInvariant()
    Set-Text $sheet $at.function $lower
    Set-Number $sheet ([string]$at.arguments[0]) 1
    $excel.Calculate()
    if ([string]$sheet.Range($at.status).Value2 -ne 'OK') { throw "The name $lower gives the status $($sheet.Range($at.status).Value2)" }
    if ([double]$sheet.Range([string]$at.arguments[1]).Value2 -ne [double]$plan.defaults.args[1]) { throw "The name $lower did not load its example" }
    $status = [string]$sheet.Range("D$($rows.status)").Value2
    if (-not $status.Contains('Probability')) { throw "The name $lower gives the Status '$status'" }
    [void]$excel.Run($prefix + 'ResetDiagnosticsInputs', $false)
    $excel.Calculate()
    $report.case_insensitive = 'PASS'

    # An argument that holds an error value: the function never runs, and the three lines that read the
    # arguments say so alike.
    $sheet.Range([string]$at.arguments[0]).Formula2 = '=NA()'
    $excel.Calculate()
    $status = [string]$sheet.Range("D$($rows.status)").Value2
    $converted = [string]$sheet.Range("D$($rows.converted)").Value2
    $call = [string]$sheet.Range("D$($rows.call)").Value2
    if (-not $status.StartsWith('Not called')) { throw "An error value as argument 1 gives the Status '$status'" }
    if (-not $converted.Contains('holds an error value')) { throw "An error value as argument 1 gives the converted arguments '$converted'" }
    if (-not $call.Contains('[error value]')) { throw "An error value as argument 1 gives the call '$call'" }
    if (-not (Is-Error $sheet.Range("D$($rows.returned)").Value2)) { throw 'An error value as argument 1 returns a number' }
    $report.observations += [ordered]@{ trial = 'an error value'; call = $call; returned = (Shown $sheet.Range("D$($rows.returned)").Value2);
                                        kind = [string]$sheet.Range("D$($rows.kind)").Value2; status = $status }
    [void]$excel.Run($prefix + 'ResetDiagnosticsInputs', $false)
    $excel.Calculate()
    $report.error_value = 'PASS'

    # Invalid inputs, one at a time from the default example: the status names the input, nothing is shown,
    # the Excel state is unchanged and the sheet can still be left and reached.
    $state = Excel-State
    $shown = $excel.Run($prefix + 'NavigationBarShown')
    $invalid = @(
        @('D6', 'text', 'NOT_A_FUNCTION', 'Check: choose a function from the list'),
        @('D11', 'text', 'low', 'Check: sweep from (a number, at most 1E15 in size)'),
        @('D11', 'number', 2E15, 'Check: sweep from (a number, at most 1E15 in size)'),
        @('D12', 'number', -3E15, 'Check: sweep to (a number, at most 1E15 in size)'),
        @('D12', 'number', -1, 'Check: sweep from must be below sweep to')
    )
    foreach ($case in $invalid) {
        $before = $sheet.Range($case[0]).Value2
        if ($case[1] -eq 'text') { Set-Text $sheet $case[0] ([string]$case[2]) } else { Set-Number $sheet $case[0] ([double]$case[2]) }
        $excel.Calculate()
        $status = [string]$sheet.Range($at.status).Value2
        if ($status -ne $case[3]) { throw "$($case[0]) = $($case[2]) gives '$status', expected '$($case[3])'" }
        if (-not (Is-Error $sheet.Range("D$($rows.returned)").Value2) -or -not (Is-Error $sheet.Range("AF$first").Value2)) { throw "$($case[0]) = $($case[2]) still shows results" }
        if ([int]$sheet.Range($at.status).DisplayFormat.Interior.Color -ne 0xCEC7FF) { throw "$($case[0]) = $($case[2]): the status is not red" }
        if ($case[0] -eq 'D6') {
            # The function cell restored by hand loads its example again: restore the sweep too.
            Set-Text $sheet 'D6' ([string]$before)
        } elseif ($before -is [string]) { Set-Text $sheet $case[0] ([string]$before) } else { Set-Number $sheet $case[0] ([double]$before) }
        $excel.Calculate()
        if ([string]$sheet.Range($at.status).Value2 -ne 'OK') { throw "Restoring $($case[0]) gives '$($sheet.Range($at.status).Value2)'" }
    }
    $after = Excel-State
    foreach ($key in $state.Keys) { if ($state[$key] -ne $after[$key]) { throw "The invalid inputs changed the Excel state: $key was $($state[$key]), is $($after[$key])" } }
    [void]$excel.Run($prefix + 'ToggleNavigationBar')
    [void]$excel.Run($prefix + 'ToggleNavigationBar')
    if ($excel.Run($prefix + 'NavigationBarShown') -ne $shown) { throw 'The navigation bar did not come back to its state' }
    if ($excel.Run($prefix + 'RibbonSheet', 'pdGoDiagnostics') -ne 'shDiagnostics') { throw 'The Ribbon does not lead to the Diagnostics sheet' }
    $book.Worksheets.Item('Explorer').Activate()
    $sheet.Activate()
    if ($excel.ActiveSheet.CodeName -ne 'shDiagnostics') { throw 'The Diagnostics sheet cannot be reached again' }
    $report.invalid = 'PASS'

    # Reset.
    Set-Text $sheet 'D6' 'K_STATS_Beta_Density'
    Set-Number $sheet 'D7' 0
    [void]$excel.Run($prefix + 'ResetDiagnosticsInputs', $false)
    $excel.Calculate()
    if ([string]$sheet.Range($at.function).Value2 -ne [string]$plan.defaults.function) { throw 'The reset did not restore the function' }
    for ($slot = 0; $slot -lt $at.arguments.Count; $slot++) {
        $value = $sheet.Range([string]$at.arguments[$slot]).Value2
        $default = $plan.defaults.args[$slot]
        if ($null -eq $default) {
            if ($null -ne $value) { throw "The reset left argument $($slot + 1) as $value" }
        } elseif ([double]$value -ne [double]$default) { throw "The reset did not restore argument $($slot + 1)" }
    }
    if ([string]$sheet.Range($at.status).Value2 -ne 'OK') { throw 'The reset did not give a valid status' }
    $after = Excel-State
    foreach ($key in $state.Keys) { if ($state[$key] -ne $after[$key]) { throw "The reset changed the Excel state: $key" } }
    $report.reset = 'PASS'
    $report.status = 'PASS'
} catch {
    $report.error = $_.Exception.Message
    throw
} finally {
    $report | ConvertTo-Json -Depth 5 | Set-Content -Encoding UTF8 (Join-Path $directory 'diagnostics-report.json')
    ([ordered]@{ status = $report.status; cases = $report.cases; trials = $report.trials;
                 largest_relative_error = $report.largest_relative_error; notes = $report.notes; error = $report.error;
                 excel_build = $report.excel_build }) | ConvertTo-Json -Depth 3 | Write-Output
    if ($null -ne $book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
}
