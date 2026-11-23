param([Parameter(Mandatory=$true)][string]$PackageDirectory)
# Applies home-plan.json to the Home and README sheets of the package in desktop
# Excel, after the other sheets: the subtitle in C3, then from row 5 a heading
# and a table per entry of the plan, in the template's own table styles. The
# styles are taken from Home as the template leaves it (a heading in row 5, a
# table in rows 6 to 8, a spacer in row 9, linked rows 19, 20 and 27), so the
# script runs on a fresh package only. The legend's labels take the format of
# the role they name: an input and a result of the Explorer, its heading, and
# the colours of the checks. The plan is made by tools/demobuild/home_plan.py.
$ErrorActionPreference = 'Stop'
$directory = (Resolve-Path $PackageDirectory).Path
$path = Join-Path $directory 'Probability-Distributions-Demo.xlsm'
$plan = Get-Content (Join-Path $directory 'home-plan.json') -Raw | ConvertFrom-Json
$manifest = Get-Content (Join-Path $directory 'manifest.json') -Raw | ConvertFrom-Json
$excel = $null
$book = $null
$stash = 200

function Step([string]$Text) {
    if ($null -ne $excel) { $excel.DisplayAlerts = $false }
    Write-Output ('[' + (Get-Date -Format 'HH:mm:ss') + '] ' + $Text)
}

function Copy-Format($Source, $Target) {
    # The Windows clipboard is shared with every program on the machine: another one can take it between the
    # copy and the paste, which Excel then refuses. Copy again and retry, up to five times.
    for ($attempt = 1; ; $attempt++) {
        try {
            [void]$Source.Copy()
            [void]$Target.PasteSpecial(-4122)
            return
        } catch {
            if ($attempt -ge 5) { throw }
            Write-Host "  the paste of formats was refused (attempt $attempt); copying again"
            Start-Sleep -Milliseconds 500
        }
    }
}

# The machine field of EXCEL.EXE's PE header: the bitness of Excel itself, not of PowerShell.
function Excel-Bitness($Application) {
    $reader = [IO.BinaryReader]::new([IO.File]::OpenRead((Join-Path $Application.Path 'EXCEL.EXE')))
    try {
        $reader.BaseStream.Position = 0x3c
        $reader.BaseStream.Position = $reader.ReadInt32()
        if ($reader.ReadUInt32() -ne 0x4550) { throw 'Invalid Excel PE signature' }
        switch ($reader.ReadUInt16()) {
            0x8664 { return '64-bit' }
            0x014c { return '32-bit' }
            0xaa64 { return 'ARM64' }
            default { throw 'Unrecognized Excel machine type' }
        }
    } finally { $reader.Dispose() }
}

function Fill([string]$Text) {
    return $Text.Replace('{version}', [string]$plan.version).Replace('{commit}', [string]$manifest.commit).Replace('{excel}', $script:excelLine)
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 1
    $excel.ScreenUpdating = $false
    $book = $excel.Workbooks.Open($path, 0, $false)
    $script:excelLine = 'Excel ' + $excel.Version + ' build ' + $excel.Build + ', ' + (Excel-Bitness $excel)
    $structure = $book.ProtectStructure
    try { $book.Unprotect() } catch { }
    $homeSheet = $book.Worksheets.Item('Home')
    $explorer = $book.Worksheets.Item('Explorer')
    if ([string]$homeSheet.Range('C5').Value2 -ne 'What this workbook shows' -or [string]$homeSheet.Range('C18').Value2 -ne 'Sheets') {
        throw 'Home is not laid out as the template leaves it: run this script on a fresh package'
    }
    try { $homeSheet.Unprotect() } catch { }

    Step 'keep the template table styles of Home'
    # Rows of the stash, from $stash: heading, first, middle, last, spacer, linked first, middle, last.
    $sources = @('C5:D5', 'C6:D6', 'C7:D7', 'C8:D8', 'C9:D9', 'C19:D19', 'C20:D20', 'C27:D27')
    for ($index = 0; $index -lt $sources.Count; $index++) {
        Copy-Format $homeSheet.Range($sources[$index]) $homeSheet.Range('C' + ($stash + $index) + ':D' + ($stash + $index))
    }
    try { $excel.CutCopyMode = $false } catch { }
    $style = @{ heading = 0; first = 1; middle = 2; last = 3; spacer = 4; link_first = 5; link_middle = 6; link_last = 7 }
    function Style-Row([string]$Name) { return $homeSheet.Range('C' + ($stash + $style[$Name]) + ':D' + ($stash + $style[$Name])) }

    $written = [ordered]@{}
    foreach ($spec in $plan.sheets) {
        $sheet = $book.Worksheets.Item([string]$spec.sheet)
        Step ('write ' + $spec.sheet)
        try { $sheet.Unprotect() } catch { }
        $area = $sheet.Range('C5:D' + ($stash - 1))
        $area.ClearHyperlinks()
        $area.Clear()
        $sheet.Range('C3').Value2 = [string]$spec.subtitle
        $row = [int]$plan.first_row
        $tableCount = 0
        foreach ($table in $spec.tables) {
            if ($tableCount -gt 0) {
                Copy-Format (Style-Row 'spacer') $sheet.Range("C${row}:D$row")
                $sheet.Rows.Item($row).RowHeight = $homeSheet.Rows.Item(9).RowHeight
                $row++
            }
            Copy-Format (Style-Row 'heading') $sheet.Range("C${row}:D$row")
            $sheet.Range("C$row").Value2 = [string]$table.heading
            $sheet.Rows.Item($row).RowHeight = $homeSheet.Rows.Item(5).RowHeight
            $row++
            $count = @($table.rows).Count
            for ($index = 0; $index -lt $count; $index++) {
                $entry = @($table.rows)[$index]
                $place = if ($index -eq 0) { 'first' } elseif ($index -eq $count - 1) { 'last' } else { 'middle' }
                $label = $sheet.Range("C$row")
                $text = $sheet.Range("D$row")
                if ($table.kind -eq 'links') {
                    [void]$sheet.Hyperlinks.Add($label, '', [string]$entry.target, [string]$entry.text, [string]$entry.label)
                    Copy-Format (Style-Row ('link_' + $place)) $sheet.Range("C${row}:D$row")
                } else {
                    Copy-Format (Style-Row $place) $sheet.Range("C${row}:D$row")
                }
                $label.Value2 = [string]$entry.label
                $text.Value2 = Fill ([string]$entry.text)
                switch ([string]$entry.swatch) {
                    'input' { Copy-Format $explorer.Range('D7') $label }
                    'result' { Copy-Format $explorer.Range('D22') $label }
                    'heading' { Copy-Format $explorer.Range('C5') $label }
                    'pass' { $label.Interior.Color = 0xCEEFC6; $label.Font.Color = 0x006100 }
                    'warning' { $label.Interior.Color = 0x9CEBFF; $label.Font.Color = 0x00659C }
                    'fail' { $label.Interior.Color = 0xC7CEFF; $label.Font.Color = 0x06009C }
                }
                if ($entry.swatch) {
                    $label.Value2 = [string]$entry.label
                    $label.HorizontalAlignment = -4108
                    $label.VerticalAlignment = -4108
                }
                $sheet.Rows.Item($row).RowHeight = [double]$entry.height
                $row++
            }
            $tableCount++
        }
        $written[[string]$spec.sheet] = $row - 1
    }

    Step 'clear the stash and protect'
    try { $excel.CutCopyMode = $false } catch { }
    foreach ($spec in $plan.sheets) {
        $sheet = $book.Worksheets.Item([string]$spec.sheet)
        $last = [int]$written[[string]$spec.sheet]
        $rest = $sheet.Range('C' + ($last + 1) + ':D' + ($stash + 10))
        $rest.ClearHyperlinks()
        $rest.Clear()
        $sheet.Rows.Item(($last + 1).ToString() + ':' + ($stash + 10)).RowHeight = $sheet.StandardHeight
        $sheet.Activate()
        $excel.ActiveWindow.DisplayGridlines = $false
        $sheet.Range('C6').Select()
        $sheet.ExportAsFixedFormat(0, (Join-Path $directory ($spec.sheet + '-preview.pdf')))
        $sheet.Protect()
    }
    if ($structure) { $book.Protect([Type]::Missing, $true) }

    # Read the sheets back: every field filled, the nine links in place.
    foreach ($spec in $plan.sheets) {
        $sheet = $book.Worksheets.Item([string]$spec.sheet)
        foreach ($cell in $sheet.Range('C3:D' + $written[[string]$spec.sheet]).Cells) {
            if ([string]$cell.Value2 -match '\{(version|commit|excel)\}') { throw "An unfilled field remains in $($spec.sheet)!$($cell.Address(0, 0))" }
        }
    }
    $links = @($homeSheet.Hyperlinks | Where-Object { $_.Type -eq 0 }).Count
    if ($links -ne 9) { throw "Home holds $links links, expected 9" }

    Step 'save'
    $explorer.Activate()
    $excel.ScreenUpdating = $true
    $excel.CommandBars.ExecuteMso('FileSave')
    $book.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
    $book = $null
    Write-Output ('Home and README evolved: ' + (Get-FileHash $path -Algorithm SHA256).Hash)
} catch {
    Write-Output $_.InvocationInfo.PositionMessage
    Write-Output $_.ScriptStackTrace
    throw
} finally {
    if ($null -ne $book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
}
