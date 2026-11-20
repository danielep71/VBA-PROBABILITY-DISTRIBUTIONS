param([Parameter(Mandatory=$true)][string]$OutputDirectory)
# tools/Sync-PROBDIST-VBA.ps1 on desktop Excel: build a production workbook from
# src/workbook/PROBDIST_Template.xlsx, open it with macros enabled and check
# that it holds every component, that Workbook_Open registered the functions
# and that the Explorer calculates; then synchronize the same workbook again
# as a development workbook and check that the regression module was added;
# last, import from a folder (-SourceDir, the mode of Sync-PROBDIST-VBA-Downloads.cmd)
# and check that it takes a changed copy of a library module and a new module with
# the project prefix, ignores a file that is neither, and refuses a folder with nothing
# to import without touching the workbook.
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$null = New-Item -ItemType Directory -Force -Path $OutputDirectory
$directory = (Resolve-Path $OutputDirectory).Path
$path = Join-Path $directory 'Probability-Distributions-Sync.xlsm'
$sync = Join-Path $root 'tools\Sync-PROBDIST-VBA.ps1'
$report = [ordered]@{ status = 'FAIL'; runs = @() }

function Invoke-Sync([string[]]$Arguments) {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $sync -WorkbookPath $path -RepoRoot $root @Arguments | Write-Host
    $code = $LASTEXITCODE
    $script:report.runs += [ordered]@{ arguments = ($Arguments -join ' '); exit_code = $code }
    if ($code -ne 0) { throw "Sync-PROBDIST-VBA.ps1 $($Arguments -join ' ') returned $code" }
}

function Get-Components([string]$Mode) {
    $folders = @('src\modules', 'src\workbook')
    if ($Mode -eq 'Development') { $folders += 'tests' }
    $names = foreach ($folder in $folders) {
        Get-ChildItem -LiteralPath (Join-Path $root $folder) -File -Filter '*.bas' | ForEach-Object { $_.BaseName }
    }
    return @($names | Sort-Object)
}

function Test-Workbook([string]$Mode) {
    $excel = $null
    $book = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $excel.DisplayAlerts = $false
        $excel.AutomationSecurity = 1
        $book = $excel.Workbooks.Open($path, 0, $true)
        # Excel may still be loading or calculating the sheets after the workbook opens: a COM call is rejected until it is ready.
        for ($wait = 0; $wait -lt 240; $wait++) { try { if ($excel.Ready) { break } } catch { }; Start-Sleep -Milliseconds 500 }
        $prefix = "'" + $book.Name.Replace("'", "''") + "'!"
        $present = @($book.VBProject.VBComponents | Where-Object { $_.Type -eq 1 } | ForEach-Object { $_.Name } | Sort-Object)
        $expected = Get-Components $Mode
        if (($present -join '|') -ne ($expected -join '|')) { throw "$Mode workbook holds $($present -join ', '); expected $($expected -join ', ')" }
        $code = $book.VBProject.VBComponents.Item('ThisWorkbook').CodeModule
        $text = $code.Lines(1, $code.CountOfLines)
        foreach ($call in @('M_STATS_PROBDIST_EXPLORER.ExplorerOpened', 'M_STATS_PROBDIST_REGISTER.PROB_Register_PublicUDFs')) {
            if ($text -notmatch [regex]::Escape($call)) { throw "ThisWorkbook does not call $call" }
        }
        if ($text -match '^\s*Attribute\s') { throw 'ThisWorkbook holds an Attribute line as code' }
        $api = @(Get-Content (Join-Path $root 'docs\PUBLIC_API.txt') | Where-Object { $_ -and -not $_.StartsWith('#') })
        $registered = [int]$excel.Run($prefix + 'PROB_Register_LastCount')
        if ($registered -ne $api.Count) { throw "Workbook_Open registered $registered of $($api.Count) functions" }
        [void]$excel.Run($prefix + 'ResetExplorerInputs', $false)
        $excel.CalculateFull()
        $sheet = $book.Worksheets.Item('Explorer')
        $density = $sheet.Range('D21').Value2
        $expectedDensity = $excel.WorksheetFunction.Norm_Dist([double]$sheet.Range('D11').Value2, [double]$sheet.Range('D7').Value2, [double]$sheet.Range('D8').Value2, $false)
        if ([string]$sheet.Range('D6').Value2 -ne 'Normal') { throw "The reset family is $($sheet.Range('D6').Value2)" }
        if (-not ($density -is [double]) -or [Math]::Abs($density - $expectedDensity) -gt 1e-12) { throw "The Explorer shows $density for the Normal density, Excel $expectedDensity" }
        return [ordered]@{ mode = $Mode; modules = $present.Count; registered = $registered; normal_density = $density }
    } finally {
        if ($null -ne $book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
        if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

# The VB_Name a file declares is its identity, never its file name.
function Write-Module([string]$File, [string]$Text) {
    [IO.File]::WriteAllBytes($File, [Text.Encoding]::GetEncoding(28591).GetBytes($Text))
}

function Test-SourceDir {
    $folder = Join-Path $directory 'sourcedir-modules'
    $empty = Join-Path $directory 'sourcedir-empty'
    foreach ($item in @($folder, $empty)) {
        if (Test-Path -LiteralPath $item) { Remove-Item -LiteralPath $item -Recurse -Force }
        $null = New-Item -ItemType Directory -Force -Path $item
    }
    $marker = "' sourcedir marker"
    $library = Join-Path $root 'src\modules\M_STATS_PROBDIST_DISCRETE.bas'
    $text = [Text.Encoding]::GetEncoding(28591).GetString([IO.File]::ReadAllBytes($library))
    # A changed copy of a library module, under a file name that is not its VB_Name.
    Write-Module (Join-Path $folder 'M_STATS_PROBDIST_DISCRETE (1).bas') ($text.TrimEnd() + "`r`n$marker`r`n")
    # A new module with the project prefix.
    Write-Module (Join-Path $folder 'probe.bas') ("Attribute VB_Name = `"M_STATS_PROBDIST_SRCDIRPROBE`"`r`nOption Explicit`r`n`r`nPublic Function PROB_SourceDirProbe() As Long`r`n    PROB_SourceDirProbe = 42`r`nEnd Function`r`n")
    # Neither a repository component nor a project module.
    Write-Module (Join-Path $folder 'other.bas') ("Attribute VB_Name = `"Unrelated_Module`"`r`nOption Explicit`r`n")
    Write-Module (Join-Path $empty 'other.bas') ("Attribute VB_Name = `"Unrelated_Module`"`r`nOption Explicit`r`n")

    $before = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $sync -WorkbookPath $path -RepoRoot $root -SourceDir $empty -Force | Write-Host
    $refused = $LASTEXITCODE
    $script:report.runs += [ordered]@{ arguments = '-SourceDir <folder with nothing to import> -Force'; exit_code = $refused }
    if ($refused -eq 0) { throw 'A folder with nothing to import was accepted' }
    if ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -ne $before) { throw 'A refused folder changed the workbook' }

    Invoke-Sync @('-SourceDir', $folder, '-Force')

    $excel = $null
    $book = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $excel.DisplayAlerts = $false
        $excel.AutomationSecurity = 1
        $book = $excel.Workbooks.Open($path, 0, $true)
        # Excel may still be loading or calculating the sheets after the workbook opens: a COM call is rejected until it is ready.
        for ($wait = 0; $wait -lt 240; $wait++) { try { if ($excel.Ready) { break } } catch { }; Start-Sleep -Milliseconds 500 }
        $prefix = "'" + $book.Name.Replace("'", "''") + "'!"
        $present = @($book.VBProject.VBComponents | Where-Object { $_.Type -eq 1 } | ForEach-Object { $_.Name } | Sort-Object)
        $expected = @(@(Get-Components 'Development') + @('M_STATS_PROBDIST_SRCDIRPROBE') | Sort-Object)
        if (($present -join '|') -ne ($expected -join '|')) { throw "After the folder import the workbook holds $($present -join ', '); expected $($expected -join ', ')" }
        $discrete = $book.VBProject.VBComponents.Item('M_STATS_PROBDIST_DISCRETE').CodeModule
        if (-not $discrete.Lines(1, $discrete.CountOfLines).Contains($marker)) { throw 'The changed copy of the library module was not imported' }
        $probe = [int]$excel.Run($prefix + 'PROB_SourceDirProbe')
        if ($probe -ne 42) { throw "The new module returned $probe, expected 42" }
        return [ordered]@{ modules = $present.Count; replaced = 'M_STATS_PROBDIST_DISCRETE'; added = 'M_STATS_PROBDIST_SRCDIRPROBE'; probe = $probe }
    } finally {
        if ($null -ne $book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
        if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

try {
    Invoke-Sync @('-Mode', 'Production', '-FromTemplate')
    $report.production = Test-Workbook 'Production'
    Invoke-Sync @('-Mode', 'Development', '-IncludeSheets')
    $report.development = Test-Workbook 'Development'
    $report.source_dir = Test-SourceDir
    $backups = @(Get-ChildItem -LiteralPath $directory -Filter '*.pre-vba-sync.*')
    if ($backups.Count -ne 0) { throw "The safety backup was left behind: $($backups.Name -join ', ')" }
    $report.status = 'PASS'
} catch {
    $report.error = $_.Exception.Message
    throw
} finally {
    $report | ConvertTo-Json -Depth 5 | Set-Content -Encoding UTF8 (Join-Path $directory 'sync-report.json')
    $report | ConvertTo-Json -Depth 5 | Write-Output
}
