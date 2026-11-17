param(
    [Parameter(Mandatory=$true)][string]$PackageDirectory,
    [ValidateSet('Com', 'NativeUI')][string]$SaveMode = 'NativeUI'
)
$ErrorActionPreference = 'Stop'
$directory = (Resolve-Path $PackageDirectory).Path
$path = Join-Path $directory 'Probability-Distributions-Demo.xlsm'
$copy = Join-Path $directory 'Excel-roundtrip.xlsm'
$report = @{ status = 'FAIL'; workbook = $path; sha256 = (Get-FileHash $path -Algorithm SHA256).Hash }
$excel = $null
$book = $null
$other = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $report.excel_version = $excel.Version
    $report.excel_build = $excel.Build
    # Normal open only: never request Excel's repair/extract-data modes.
    $book = $excel.Workbooks.Open($path, 0, $false)
    if ($book.Worksheets.Count -ne 10) { throw 'Expected ten worksheets' }
    if ($book.HasVBProject -ne $true) { throw 'VBA project missing' }
    $report.open = 'PASS'
    $book.SaveCopyAs($copy)
    $book.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
    $book = $null
    $book = $excel.Workbooks.Open($copy, 0, $true)
    if ($book.Worksheets.Count -ne 10) { throw 'Round-trip lost worksheets' }
    if ($book.HasVBProject -ne $true) { throw 'Round-trip lost VBA project' }
    $report.roundtrip = 'PASS'
    $book.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
    $book = $null
    # Do not reuse a VBA project first loaded with macros forcibly disabled.
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    # Macro execution is confined to this owned Excel process; Trust Center is unchanged.
    $excel.Visible = ($SaveMode -eq 'NativeUI')
    $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($path, 0, $false)
    $prefix = "'" + $book.Name + "'!"
    $value = $excel.Run($prefix + 'K_STATS_NormalStandard_Cumulative', 0.0)
    if ([Math]::Abs([double]$value - 0.5) -gt 1e-15) { throw 'Embedded VBA smoke result differs from 0.5' }
    $report.embedded_vba = 'PASS'
    $sheetMap = Get-Content (Join-Path $PSScriptRoot '..\examples\demo\sheets.json') -Raw | ConvertFrom-Json
    foreach ($section in $sheetMap.sections) {
        foreach ($entry in $section.sheets) {
            $control = 'pdGo' + $entry.code.Substring(2)
            $actual = $excel.Run($prefix + 'RibbonSheet', $control)
            if ($actual -ne $entry.code) { throw "Wrong Ribbon mapping: $control" }
            $target = $book.Names.Item('nav_' + $entry.code).RefersToRange
            if ($target.Worksheet.CodeName -ne $entry.code) { throw "Wrong navigation name: $control" }
        }
    }
    # COM marshals VBA vbNullString as $null, rather than an empty .NET string.
    $unknownTarget = $excel.Run($prefix + 'RibbonSheet', 'unknown-control')
    if (-not [string]::IsNullOrEmpty([string]$unknownTarget)) { throw "Unknown control has a destination: $unknownTarget" }
    $report.navigation_targets = 'PASS'
    $other = $excel.Workbooks.Add()
    $other.Worksheets.Item(1).Range('A1').Value2 = 'isolation sentinel'
    $other.Activate()
    $savedScreen = $excel.ScreenUpdating
    $savedFlag = $book.Saved
    $excel.Run($prefix + 'ShowNavigationBar', $false)
    foreach ($sheet in $book.Worksheets) {
        if (-not $sheet.Columns.Item(1).Hidden) { throw "Navigation column still visible: $($sheet.Name)" }
        foreach ($name in @('NavBarTop','NavBar')) {
            if ($sheet.Shapes.Item($name).Visible -ne 0) { throw "Navigation shape still visible: $name" }
        }
        if (-not $sheet.ProtectContents) { throw 'Navigation toggle removed sheet protection' }
    }
    if ($excel.ScreenUpdating -ne $savedScreen -or $book.Saved -ne $savedFlag) { throw 'Navigation toggle leaked state' }
    if ($other.Worksheets.Item(1).Columns.Item(1).Hidden) { throw 'Navigation toggle changed another workbook' }
    if ($other.Worksheets.Item(1).Range('A1').Value2 -ne 'isolation sentinel') { throw 'Other workbook changed' }
    $report.hide_and_isolation = 'PASS'
    $excel.Run($prefix + 'RecalculateWorkbook')
    $report.events_enabled = $excel.EnableEvents
    $report.workbook_codename = $book.CodeName
    try {
        $report.vba_references = @($book.VBProject.References | ForEach-Object { $_.Name + ':' + $_.Guid })
        $component = $book.VBProject.VBComponents.Item('ThisWorkbook')
        $report.workbook_module_type = $component.Type
        $report.workbook_module_source = $component.CodeModule.Lines(1, $component.CodeModule.CountOfLines)
    } catch {
        $report.project_inspection = $_.Exception.Message
    }
    $excel.Run($prefix + 'NavigationBarBeforeSave')
    if ($book.Worksheets.Item(1).Columns.Item(1).Hidden) { throw 'Before-save helper did not show the navigation' }
    $excel.Run($prefix + 'NavigationBarAfterSave')
    if (-not $book.Worksheets.Item(1).Columns.Item(1).Hidden) { throw 'After-save helper did not restore session state' }
    $report.save_helpers = 'PASS'
    # Require an actual save even when the session-only toggle preserved Saved.
    $book.Saved = $false
    $book.Activate()
    $report.save_mode = $SaveMode
    if ($SaveMode -eq 'NativeUI') {
        if (-not $excel.CommandBars.GetEnabledMso('FileSave')) { throw 'Native FileSave is not available' }
        Write-Output 'Invoking native FileSave in the visible test instance.'
        $excel.CommandBars.ExecuteMso('FileSave')
    } else {
        $book.Save()
    }
    $report.saved_path = $book.FullName
    $report.read_only = $book.ReadOnly
    Copy-Item -LiteralPath $path -Destination (Join-Path $directory 'Excel-after-save.xlsm')
    if (-not $book.Worksheets.Item(1).Columns.Item(1).Hidden) { throw 'Save did not restore session-hidden bar' }
    $report.session_state_after_save = 'PASS'
    $book.Close($false)
    Copy-Item -LiteralPath $path -Destination (Join-Path $directory 'Excel-after-close.xlsm')
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
    $book = $null
    $other.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($other)
    $other = $null
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($path, 0, $true)
    foreach ($sheet in $book.Worksheets) {
        if ($sheet.Columns.Item(1).Hidden) { throw 'Saved file hides its navigation with macros disabled' }
        foreach ($name in @('NavBarTop','NavBar')) {
            if ($sheet.Shapes.Item($name).Visible -eq 0) { throw 'Saved file hides a navigation group' }
        }
    }
    $report.saved_navigation = 'PASS'
    $report.save_events = 'PASS'
    $report.scope = 'Embedded VBA, navigation mapping, toggle, isolation, save events; Ribbon clicks remain manual.'
    $report.status = 'PASS'
} catch {
    $report.error = $_.Exception.Message
    throw
} finally {
    if ($null -ne $book) {
        $book.Close($false)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
    }
    if ($null -ne $other) {
        $other.Close($false)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($other)
    }
    if ($null -ne $excel) {
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
    $report | ConvertTo-Json | Set-Content -Encoding UTF8 (Join-Path $directory 'excel-open-report.json')
    $report | ConvertTo-Json | Write-Output
}
