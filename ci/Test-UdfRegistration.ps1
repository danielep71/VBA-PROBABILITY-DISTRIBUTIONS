param([Parameter(Mandatory=$true)][string]$PackageDirectory)
$ErrorActionPreference = 'Stop'
$directory = (Resolve-Path $PackageDirectory).Path
$path = Join-Path $directory 'Probability-Distributions-Demo.xlsm'
$report = @{status='FAIL'; sha256=(Get-FileHash $path -Algorithm SHA256).Hash}
$excel = $null
$book = $null
$other = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($path,0,$true)
    # Excel may still be loading or calculating the sheets after the workbook opens: a COM call is rejected until it is ready.
    for ($wait = 0; $wait -lt 240; $wait++) { try { if ($excel.Ready) { break } } catch { }; Start-Sleep -Milliseconds 500 }
    $prefix = "'" + $book.Name.Replace("'", "''") + "'!"
    $api = @(Get-Content (Join-Path $PSScriptRoot '../docs/PUBLIC_API.txt') | Where-Object { $_ -and -not $_.StartsWith('#') })
    $expected = $api.Count
    $count = [int]$excel.Run($prefix + 'PROB_Register_LastCount')
    if ($count -ne $expected) { throw "Workbook_Open registered $count of $expected functions" }
    $report.automatic_open = 'PASS'
    if ([string]$excel.Run($prefix + 'PROB_Register_LastReport') -ne '') { throw 'Automatic registration reported failures' }
    if ([int]$excel.Run($prefix + 'PROB_Register_ManifestCount') -ne $expected) { throw 'Manifest count differs from public API' }
    for ($i=1; $i -le $expected; $i++) {
        $fields = $api[$i-1].Split("`t")
        if ($excel.Run($prefix + 'PROB_Register_ManifestName',$i) -ne $fields[2]) { throw 'Manifest name/order differs' }
        $args = [regex]::Matches($fields[3], '(?:ByVal|ByRef)\s+(\w+)')
        if ([int]$excel.Run($prefix + 'PROB_Register_ManifestArgCount',$i) -ne $args.Count) { throw 'Argument count differs' }
        $descriptions = @($excel.Run($prefix + 'PROB_Register_ManifestArgDescriptions',$i))
        if ($descriptions.Count -ne $args.Count) { throw 'Argument descriptions are incomplete' }
    }
    foreach ($saved in @($true,$false)) {
        $book.Saved = $saved
        foreach ($operation in @('PROB_Register_PublicUDFs','PROB_Register_PublicUDFs','PROB_Register_ClearPublicUDFs','PROB_Register_ClearPublicUDFs','PROB_Register_PublicUDFs')) {
            if ([int]$excel.Run($prefix + $operation) -ne $expected) { throw "$operation rejected entries: $($excel.Run($prefix + 'PROB_Register_LastReport'))" }
            if ($book.Saved -ne $saved) { throw "$operation changed the Saved flag" }
        }
    }
    if ([int]$excel.Run($prefix + 'PROB_Register_ManifestArgCount',0) -ne -1) { throw 'Invalid manifest index not rejected' }
    $other = $excel.Workbooks.Add()
    $other.Saved = $true
    $other.Activate()
    $beforeEvents = $excel.EnableEvents
    $beforeCalc = $excel.Calculation
    $beforeAlerts = $excel.DisplayAlerts
    $accepted = [int]$excel.Run($prefix + 'PROB_Register_PublicUDFs')
    if ($excel.ActiveWorkbook.FullName -ne $other.FullName -or -not $other.Saved) { throw 'Registration changed another workbook' }
    if ($excel.EnableEvents -ne $beforeEvents -or $excel.Calculation -ne $beforeCalc -or $excel.DisplayAlerts -ne $beforeAlerts) { throw 'Registration leaked application state' }
    $report.inactive_workbook_accepted = $accepted
    $report.inactive_workbook_report = [string]$excel.Run($prefix + 'PROB_Register_LastReport')
    $book.Activate()
    if ([int]$excel.Run($prefix + 'PROB_Register_PublicUDFs') -ne $expected) { throw 'Registration did not recover' }
    if ([Math]::Abs([double]$excel.Run($prefix + 'K_STATS_Normal_Cumulative',0.0,0.0,1.0)-0.5) -gt 1e-12) { throw 'Registration affected calculations' }
    $report.registered = $expected
    $report.repeat_register_clear = 'PASS'
    $report.saved_and_application_state = 'PASS'
    $report.signature_metadata = 'PASS'
    $report.excel_version = $excel.Version
    $report.excel_build = $excel.Build
    $report.status = 'PASS'
} catch {
    $report.error = $_.Exception.Message
    throw
} finally {
    if ($null -ne $other) { $other.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($other) }
    if ($null -ne $book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    $report | ConvertTo-Json -Depth 5 | Set-Content -Encoding UTF8 (Join-Path $directory 'udf-registration-report.json')
    $report | ConvertTo-Json -Depth 5 | Write-Output
}
