<#
.SYNOPSIS
Synchronize the repository's VBA source into a local Probability Distributions
.xlsm workbook.

.DESCRIPTION
Updates VBA code only. A full synchronization imports the components of one
of two workbooks, and asks which one unless -Mode gives the answer:

- Production: the library modules (src\modules\*.bas) and the modules of the
  Explorer workbook (src\workbook\*.bas), with the code of ThisWorkbook
  (src\workbook\ThisWorkbook.cls), for a workbook that is used to calculate.
- Development: the production components and the regression module
  (tests\M_STATS_PROBDIST_TEST.bas), for a workbook that runs the VBA test
  suites. The tests write to a sheet of their own; they do not need the
  Explorer.

With -SourceDir only the modules found in that folder are imported (see the
fourth example). Standard modules are replaced from source. ThisWorkbook and
worksheet document modules keep their workbook objects and only their code
text is replaced. The script does not write cells, formulas, names,
formatting, tables, worksheets, or workbook structure.

Before the workbook is opened the script lists the worksheet and
ThisWorkbook modules in the source and asks whether to replace their code as
well. Answer N for a workbook whose sheet code must stay as it is.
-IncludeSheets or -SkipSheets gives the answer in advance.

-FromTemplate builds the workbook first: it saves
src\workbook\PROBDIST_Template.xlsx, the Explorer workbook without its code,
as an .xlsm at -WorkbookPath, replacing a file of that name, so the sheets
are the current template's. The sheet code is then always replaced from the
source, without asking: the template holds no code. It is refused with
-SourceDir and with -SkipSheets. tools\Sync-PROBDIST-VBA.cmd started without a
workbook uses it to build test-results\Probability-Distributions-Development.xlsm.

Exit codes: 0 synchronized; 2 cancelled or nothing to import, the workbook
not changed; 1 failed, the workbook closed without saving.

The script follows the synchronization tool of the K-SACCR repository, by
the same author.

.REQUIREMENTS
- Windows desktop Excel.
- Excel > Options > Trust Center > Macro Settings >
  "Trust access to the VBA project object model" enabled.
- VBA project not password-locked.

.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Sync-PROBDIST-VBA.ps1 -WorkbookPath ".\test-results\Probability-Distributions-Development.xlsm" -Mode Development -FromTemplate

Builds a fresh development workbook from the template, with the code of
ThisWorkbook filled.

.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Sync-PROBDIST-VBA.ps1 -WorkbookPath "C:\Work\Probability-Distributions.xlsm"

Asks whether the workbook is a production or a development workbook, then
whether to replace the sheet code as well.

.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Sync-PROBDIST-VBA.ps1 -WorkbookPath "C:\Work\Probability-Distributions.xlsm" -Mode Production -SkipSheets

Imports the standard modules only and leaves the code of every worksheet and
of ThisWorkbook as it is, without asking about it.

.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Sync-PROBDIST-VBA.ps1 -WorkbookPath "C:\Work\Probability-Distributions.xlsm" -SourceDir "$env:USERPROFILE\Downloads"

With -SourceDir, only the modules found in that folder are imported, for
example modules downloaded for a test, without copying them into the
repository. A file is identified only by the Attribute VB_Name it declares,
never by its file name: a file without exactly one valid declaration is
ignored. It is taken when that VB_Name is a component of the repository
source, or when it is a standard module (.bas) whose name has the project
prefix M_STATS_PROBDIST_, such as a module added by the pull request under
test; it is listed apart as a new module. Other files in the folder are
ignored. When several files carry the same VB_Name (for example
"M_STATS_PROBDIST_CORE (1).bas"), the newest is taken. The script lists the
files and asks for confirmation unless -Force is given, converts their line
endings to CRLF and imports them. Every other module in the workbook is left
as it is.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$WorkbookPath,

    [Parameter(Mandatory = $false)]
    [string]$RepoRoot,

    [Parameter(Mandatory = $false)]
    [switch]$NoBackup,

    [Parameter(Mandatory = $false)]
    [switch]$Visible,

    [Parameter(Mandatory = $false)]
    [string]$SourceDir,

    [Parameter(Mandatory = $false)]
    [switch]$Force,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeSheets,

    [Parameter(Mandatory = $false)]
    [switch]$SkipSheets,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Production", "Development")]
    [string]$Mode,

    [Parameter(Mandatory = $false)]
    [switch]$FromTemplate
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Resolve the repository root after parameter binding. $PSScriptRoot is not
# reliable as a parameter default in every Windows PowerShell invocation mode.
if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $scriptPath = $MyInvocation.MyCommand.Path
    if ([string]::IsNullOrWhiteSpace($scriptPath)) {
        throw "Cannot determine the script path. Supply -RepoRoot explicitly."
    }

    $scriptDirectory = Split-Path -Parent $scriptPath
    $RepoRoot = Split-Path -Parent $scriptDirectory
}

$vbext_ct_StdModule = 1
$vbext_ct_ClassModule = 2
$vbext_ct_MSForm = 3
$vbext_ct_Document = 100
$xlCalculationManual = -4135
$xlCalculationAutomatic = -4105
$msoAutomationSecurityForceDisable = 3
$xlOpenXMLWorkbookMacroEnabled = 52

# With -SourceDir, a standard module the repository checkout does not have
# is offered for import when its name has the project prefix; any other
# unknown file is ignored.
$NewModulePattern = '^M_STATS_PROBDIST_[A-Za-z0-9_]+$'

function Resolve-FullPath {
    param([Parameter(Mandatory = $true)][string]$PathValue)
    return [System.IO.Path]::GetFullPath((Resolve-Path -LiteralPath $PathValue).Path)
}

function Get-VbaComponentName {
    param([Parameter(Mandatory = $true)][string]$SourcePath)

    foreach ($line in [System.IO.File]::ReadLines($SourcePath)) {
        if ($line -match '^\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"\s*$') {
            return $Matches[1]
        }
    }

    return [System.IO.Path]::GetFileNameWithoutExtension($SourcePath)
}

function Get-DeclaredVbaName {
    param([Parameter(Mandatory = $true)][string]$SourcePath)

    # The identity a file declares, for files from -SourceDir: exactly
    # one Attribute VB_Name line holding a valid VBA name. The file name is
    # never used instead, so a "M_STATS_PROBDIST_CORE.bas" without the attribute
    # cannot replace M_STATS_PROBDIST_CORE. Returns the name and, when there is none, why.
    # A list, not a set: a declaration repeated identically, or with another
    # case, is a second declaration too.
    $names = New-Object System.Collections.Generic.List[string]
    foreach ($line in [System.IO.File]::ReadLines($SourcePath)) {
        if ($line -match '^\s*Attribute\s+VB_Name\s*=\s*"([^"]*)"\s*$') {
            $names.Add($Matches[1])
        }
        elseif ($line -match '^\s*Attribute\s+VB_Name\b') {
            return [pscustomobject]@{ Name = ""; Problem = "malformed Attribute VB_Name line" }
        }
    }

    if ($names.Count -eq 0) {
        return [pscustomobject]@{ Name = ""; Problem = "no Attribute VB_Name line; the file name is not used as its identity" }
    }
    if ($names.Count -gt 1) {
        return [pscustomobject]@{ Name = ""; Problem = "several Attribute VB_Name lines" }
    }

    $name = $names[0]
    if ($name -notmatch '^[A-Za-z][A-Za-z0-9_]{0,30}$') {
        return [pscustomobject]@{ Name = ""; Problem = "VB_Name '$name' is not a valid VBA name" }
    }

    return [pscustomobject]@{ Name = $name; Problem = "" }
}

function Get-VbaEncoding {
    # Exported VBA is read in the Windows ANSI code page, cp1252, not as
    # UTF-8; the repository's modules are ASCII (tools/check_vba_source.py),
    # which cp1252 reads unchanged. Windows PowerShell has the code page
    # built in; PowerShell 7 may need the code-page provider registered.
    try {
        return [System.Text.Encoding]::GetEncoding(1252)
    }
    catch {
        [System.Text.Encoding]::RegisterProvider([System.Text.CodePagesEncodingProvider]::Instance)
        return [System.Text.Encoding]::GetEncoding(1252)
    }
}

function Test-DocumentSource {
    param([Parameter(Mandatory = $true)][string]$SourcePath)

    # A document module (ThisWorkbook, a worksheet or a chart sheet) is
    # exported with VB_Base naming Excel's Workbook (00020819), Worksheet
    # (00020820) or Chart (00020821) class, whatever its code name; an
    # ordinary class module has no VB_Base. The names Sync-VbaSource
    # recognizes (ThisWorkbook, sh*) also count, for a file without it.
    if ([System.IO.Path]::GetExtension($SourcePath).ToLowerInvariant() -ne ".cls") { return $false }
    foreach ($line in [System.IO.File]::ReadLines($SourcePath)) {
        if ($line -match '^\s*Attribute\s+VB_Base\s*=\s*"0\{(00020819|00020820|00020821)-') {
            return $true
        }
    }
    $name = Get-VbaComponentName -SourcePath $SourcePath
    return ($name -eq "ThisWorkbook" -or $name -like "sh*")
}

function Get-VbaCodeBody {
    param([Parameter(Mandatory = $true)][string]$SourcePath)

    $lines = [System.IO.File]::ReadAllLines($SourcePath, (Get-VbaEncoding))
    if ($lines.Count -eq 0) { return "" }

    # The module header is Attribute VB_<name> = ...; a member attribute
    # has a dot after the member's name, which may itself start with VB_,
    # so it never counts as the header and cuts no code above it.
    $lastAttribute = -1
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match '^\s*Attribute\s+VB_[^\s.=]*\s*=') {
            $lastAttribute = $i
        }
    }
    if (($lastAttribute + 1) -ge $lines.Count) { return "" }
    $body = $lines[($lastAttribute + 1)..($lines.Count - 1)]

    # An exported member attribute, such as
    # Attribute Worksheet_Change.VB_Description = "...", follows its
    # procedure's first line and is not valid as code, so it is dropped
    # wherever it is, with or without a module header. The member
    # name may hold code-page letters, so it is any run of characters
    # other than space, dot and equals sign.
    $code = @($body | Where-Object { $_ -notmatch '^\s*Attribute\s+[^\s.=]+\.VB_' })
    return [string]::Join([Environment]::NewLine, $code)
}

function Copy-WithCrlf {
    param(
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [Parameter(Mandatory = $true)][string]$TargetPath
    )

    # ISO-8859-1 maps every byte to one character and back, so the cp1252
    # bytes of an exported module are kept exactly; only line endings change.
    $encoding = [System.Text.Encoding]::GetEncoding(28591)
    $text = $encoding.GetString([System.IO.File]::ReadAllBytes($SourcePath))
    $text = $text.Replace("`r`n", "`n").Replace("`r", "`n").Replace("`n", "`r`n")
    [System.IO.File]::WriteAllBytes($TargetPath, $encoding.GetBytes($text))
}

function Select-FolderSources {
    param(
        [Parameter(Mandatory = $true)][string]$FolderPath,
        [Parameter(Mandatory = $true)][hashtable]$RepositoryNames,
        [Parameter(Mandatory = $true)][string]$StagePath
    )

    $picked = @{}
    $newNames = @{}
    $ignored = New-Object System.Collections.Generic.List[string]
    $candidates = @(
        Get-ChildItem -LiteralPath $FolderPath -File |
            Where-Object { $_.Extension.ToLowerInvariant() -in @(".bas", ".cls", ".frm") }
    )

    foreach ($file in $candidates) {
        $declared = Get-DeclaredVbaName -SourcePath $file.FullName
        if (-not [string]::IsNullOrEmpty($declared.Problem)) {
            $ignored.Add("$($file.Name) ($($declared.Problem))")
            continue
        }
        $name = $declared.Name
        $key = $name.ToLowerInvariant()
        if ($RepositoryNames.ContainsKey($key)) {
            $repoExtension = [System.IO.Path]::GetExtension($RepositoryNames[$key]).ToLowerInvariant()
            if ($file.Extension.ToLowerInvariant() -ne $repoExtension) {
                $ignored.Add("$($file.Name) ('$name' is a $repoExtension component in the repository)")
                continue
            }
        }
        elseif ($file.Extension.ToLowerInvariant() -eq ".bas" -and $name -match $NewModulePattern) {
            # A standard module the checkout does not have yet, for example a
            # module added by the pull request under test: listed apart.
            $newNames[$key] = $true
        }
        else {
            $ignored.Add("$($file.Name) (VB_Name '$name' is not a repository component)")
            continue
        }

        if ($picked.ContainsKey($key)) {
            if ($file.LastWriteTime -gt $picked[$key].LastWriteTime) {
                $ignored.Add("$($picked[$key].Name) (older copy of '$name')")
                $picked[$key] = $file
            }
            else {
                $ignored.Add("$($file.Name) (older copy of '$name')")
            }
        }
        else {
            $picked[$key] = $file
        }
    }

    if ($picked.Count -eq 0) {
        # Say why each file was passed over before stopping.
        if ($ignored.Count -gt 0) {
            Write-Host "Ignored:"
            foreach ($line in $ignored) { Write-Host "  $line" }
            Write-Host ""
        }
        throw "No repository VBA module found in $FolderPath."
    }

    Write-Host "Modules to import from $FolderPath"
    foreach ($key in ($picked.Keys | Sort-Object)) {
        if ($newNames.ContainsKey($key)) { continue }
        $file = $picked[$key]
        Write-Host ("  {0,-26} {1,-34} {2:yyyy-MM-dd HH:mm}" -f (Get-VbaComponentName -SourcePath $file.FullName), $file.Name, $file.LastWriteTime)
    }
    if ($newNames.Count -gt 0) {
        Write-Host "New modules, not in the repository checkout (imported as new standard modules):"
        foreach ($key in ($picked.Keys | Sort-Object)) {
            if (-not $newNames.ContainsKey($key)) { continue }
            $file = $picked[$key]
            Write-Host ("  {0,-26} {1,-34} {2:yyyy-MM-dd HH:mm}" -f (Get-VbaComponentName -SourcePath $file.FullName), $file.Name, $file.LastWriteTime)
        }
    }
    if ($ignored.Count -gt 0) {
        Write-Host "Ignored:"
        foreach ($line in $ignored) { Write-Host "  $line" }
    }
    Write-Host ""

    $null = New-Item -ItemType Directory -Path $StagePath -Force
    $staged = @(
        foreach ($key in ($picked.Keys | Sort-Object)) {
            $file = $picked[$key]
            $name = Get-VbaComponentName -SourcePath $file.FullName
            $target = Join-Path $StagePath ($name + $file.Extension.ToLowerInvariant())
            Copy-WithCrlf -SourcePath $file.FullName -TargetPath $target
            Get-Item -LiteralPath $target
        }
    )

    return $staged
}

function ConvertTo-SyncMode {
    # The answer to the production/development question: Production,
    # Development, or $null for anything else.
    param([AllowNull()][AllowEmptyString()][string]$Answer)

    switch -Regex (([string]$Answer).Trim()) {
        '^(?i)(p|prod|production)$' { return "Production" }
        '^(?i)(d|dev|development)$' { return "Development" }
    }
    return $null
}

function Get-ModeSources {
    # The components of one workbook: the library modules (src\modules), the modules
    # and the ThisWorkbook code of the Explorer workbook (src\workbook) and,
    # for a development workbook, the regression module (tests). Sorted by
    # path.
    param(
        [Parameter(Mandatory = $true)][string]$RepoPath,
        [Parameter(Mandatory = $true)][ValidateSet("Production", "Development")][string]$Mode
    )

    $folders = @("src\modules", "src\workbook")
    if ($Mode -eq "Development") { $folders += "tests" }
    $files = @(
        foreach ($folder in $folders) {
            Get-ChildItem -LiteralPath (Join-Path $RepoPath $folder) -File |
                Where-Object { $_.Extension.ToLowerInvariant() -in @(".bas", ".cls", ".frm") }
        }
    )
    return @($files | Sort-Object FullName)
}

function Find-VbaComponent {
    param(
        [Parameter(Mandatory = $true)]$VBProject,
        [Parameter(Mandatory = $true)][string]$Name
    )

    foreach ($component in $VBProject.VBComponents) {
        if ([string]::Equals(
            [string]$component.Name,
            $Name,
            [System.StringComparison]::OrdinalIgnoreCase
        )) {
            return $component
        }
    }

    return $null
}

function Get-SheetSnapshot {
    param([Parameter(Mandatory = $true)]$Workbook)

    $items = New-Object System.Collections.Generic.List[string]
    foreach ($sheet in $Workbook.Sheets) {
        $codeName = ""
        try { $codeName = [string]$sheet.CodeName }
        catch { $codeName = "<no-codename>" }

        $items.Add("$($sheet.Name)|$codeName")
    }

    return @($items)
}

function Assert-SheetSnapshotUnchanged {
    param(
        [Parameter(Mandatory = $true)][string[]]$Before,
        [Parameter(Mandatory = $true)][string[]]$After
    )

    if ($Before.Count -ne $After.Count) {
        throw "Workbook sheet count changed. Save aborted."
    }

    for ($i = 0; $i -lt $Before.Count; $i++) {
        if ($Before[$i] -cne $After[$i]) {
            throw "Workbook sheet structure changed. Save aborted."
        }
    }
}

function Replace-DocumentCode {
    param(
        [Parameter(Mandatory = $true)]$Component,
        [Parameter(Mandatory = $true)][string]$SourcePath
    )

    if ([int]$Component.Type -ne $vbext_ct_Document) {
        throw "Component '$($Component.Name)' is not a document module."
    }

    $module = $Component.CodeModule
    $count = [int]$module.CountOfLines
    if ($count -gt 0) {
        $module.DeleteLines(1, $count)
    }

    $code = Get-VbaCodeBody -SourcePath $SourcePath
    if (-not [string]::IsNullOrWhiteSpace($code)) {
        $module.AddFromString($code)
    }
}

function Sync-VbaSource {
    param(
        [Parameter(Mandatory = $true)]$VBProject,
        [Parameter(Mandatory = $true)][string]$SourcePath
    )

    $extension = [System.IO.Path]::GetExtension($SourcePath).ToLowerInvariant()
    $name = Get-VbaComponentName -SourcePath $SourcePath
    $existing = Find-VbaComponent -VBProject $VBProject -Name $name

    if ($null -ne $existing -and [int]$existing.Type -eq $vbext_ct_Document) {
        if ($extension -ne ".cls") {
            throw "Document module '$name' must come from a .cls source."
        }

        Replace-DocumentCode -Component $existing -SourcePath $SourcePath
        Write-Host ("{0,-30} document code replaced" -f $name)
        return
    }

    if ($extension -eq ".cls" -and $null -eq $existing) {
        if ($name -eq "ThisWorkbook" -or $name -like "sh*") {
            throw "Workbook document module '$name' was not found. Save aborted."
        }
    }

    switch ($extension) {
        ".bas" {
            if ($null -ne $existing) {
                if ([int]$existing.Type -ne $vbext_ct_StdModule) {
                    throw "Existing '$name' is not a standard module."
                }
                $VBProject.VBComponents.Remove($existing)
            }

            $imported = $VBProject.VBComponents.Import($SourcePath)
            Write-Host ("{0,-30} standard module imported" -f $imported.Name)
        }

        ".cls" {
            if ($null -ne $existing) {
                if ([int]$existing.Type -ne $vbext_ct_ClassModule) {
                    throw "Existing '$name' is not a class module."
                }
                $VBProject.VBComponents.Remove($existing)
            }

            $imported = $VBProject.VBComponents.Import($SourcePath)
            Write-Host ("{0,-30} class module imported" -f $imported.Name)
        }

        ".frm" {
            if ($null -ne $existing) {
                if ([int]$existing.Type -ne $vbext_ct_MSForm) {
                    throw "Existing '$name' is not a UserForm."
                }
                $VBProject.VBComponents.Remove($existing)
            }

            $imported = $VBProject.VBComponents.Import($SourcePath)
            Write-Host ("{0,-30} UserForm imported" -f $imported.Name)
        }

        default {
            throw "Unsupported VBA source extension '$extension'."
        }
    }
}

function New-WorkbookFromTemplate {
    # Save the macro-free template as an .xlsm at TargetPath, in an Excel
    # instance of its own. A file already at TargetPath is replaced.
    param(
        [Parameter(Mandatory = $true)][string]$TemplatePath,
        [Parameter(Mandatory = $true)][string]$TargetPath
    )

    if (-not (Test-Path -LiteralPath $TemplatePath -PathType Leaf)) {
        throw "Template not found: $TemplatePath"
    }
    $null = New-Item -ItemType Directory -Force -Path ([System.IO.Path]::GetDirectoryName($TargetPath))
    if (Test-Path -LiteralPath $TargetPath -PathType Leaf) {
        Remove-Item -LiteralPath $TargetPath -Force
    }

    $builder = $null
    $book = $null
    try {
        $builder = New-Object -ComObject Excel.Application
        $builder.Visible = $false
        $builder.DisplayAlerts = $false
        $builder.AutomationSecurity = $msoAutomationSecurityForceDisable
        $book = $builder.Workbooks.Open($TemplatePath, 0, $true)
        $book.SaveAs($TargetPath, $xlOpenXMLWorkbookMacroEnabled)
        $book.Close($false)
        $book = $null
    }
    finally {
        if ($null -ne $book) {
            try { $book.Close($false) } catch {}
        }
        if ($null -ne $builder) {
            try { $builder.Quit() } catch {}
            try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($builder) } catch {}
        }
    }
}

$excel = $null
$workbook = $null
$saved = $false
$backupPath = $null
$stagePath = $null
$syncMode = $null

try {
    # A workbook built from the template need not exist yet.
    if ($FromTemplate) {
        $workbookFullPath = [System.IO.Path]::GetFullPath($ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($WorkbookPath))
    }
    else {
        $workbookFullPath = Resolve-FullPath -PathValue $WorkbookPath
    }
    $repoFullPath = Resolve-FullPath -PathValue $RepoRoot

    if ([System.IO.Path]::GetExtension($workbookFullPath).ToLowerInvariant() -ne ".xlsm") {
        throw "WorkbookPath must point to an .xlsm file."
    }

    # -FromTemplate builds a whole workbook: every component of the mode,
    # and the sheet code, which the template does not hold.
    if ($FromTemplate) {
        if (-not [string]::IsNullOrWhiteSpace($SourceDir)) {
            throw "Use -FromTemplate only with a full synchronization, not with -SourceDir."
        }
        if ($SkipSheets) {
            throw "Use -FromTemplate without -SkipSheets: the template's sheet modules hold no code."
        }
        $IncludeSheets = [switch]$true
    }

    $sourceRoots = @(
        (Join-Path $repoFullPath "src\modules"),
        (Join-Path $repoFullPath "src\workbook"),
        (Join-Path $repoFullPath "tests")
    )

    foreach ($root in $sourceRoots) {
        if (-not (Test-Path -LiteralPath $root -PathType Container)) {
            throw "Required source directory not found: $root"
        }
    }

    $sourceFiles = @(
        foreach ($root in $sourceRoots) {
            Get-ChildItem -LiteralPath $root -File |
                Where-Object {
                    $_.Extension.ToLowerInvariant() -in @(".bas", ".cls", ".frm")
                }
        }
    ) | Sort-Object FullName

    if ($sourceFiles.Count -eq 0) {
        throw "No VBA source files found."
    }

    $seen = @{}
    foreach ($file in $sourceFiles) {
        $name = Get-VbaComponentName -SourcePath $file.FullName
        $key = $name.ToLowerInvariant()

        if ($seen.ContainsKey($key)) {
            throw "Duplicate VBA component name '$name' in repository source."
        }

        $seen[$key] = $file.FullName
    }

    # Full synchronization: the production or the development component set.
    # -SourceDir instead takes what the folder holds.
    if (-not [string]::IsNullOrWhiteSpace($SourceDir) -and -not [string]::IsNullOrWhiteSpace($Mode)) {
        throw "Use -Mode only for a full synchronization, not with -SourceDir."
    }

    if ([string]::IsNullOrWhiteSpace($SourceDir)) {
        $syncMode = $Mode
        if ([string]::IsNullOrWhiteSpace($syncMode)) {
            Write-Host "Which workbook is this?"
            Write-Host "  P = production: library, Explorer modules and ThisWorkbook code"
            Write-Host "  D = development: production components and the regression module"
            $syncMode = ConvertTo-SyncMode -Answer (Read-Host "Production or development? [P/D]")
            if ($null -eq $syncMode) {
                Write-Host "Cancelled: the workbook was not changed."
                exit 2
            }
        }

        $sourceFiles = @(Get-ModeSources -RepoPath $repoFullPath -Mode $syncMode)
        Write-Host "Mode: $syncMode ($($sourceFiles.Count) components)"
        Write-Host ""
    }

    # -SourceDir: import the repository modules found in that folder, and new
    # modules with the project prefix, after the user has seen the list. The repository source is used above
    # only to know which component names are valid.
    $sourceLabel = $repoFullPath
    if (-not [string]::IsNullOrWhiteSpace($SourceDir)) {
        $sourceDirFull = Resolve-FullPath -PathValue $SourceDir
        if (-not (Test-Path -LiteralPath $sourceDirFull -PathType Container)) {
            throw "SourceDir is not a folder: $sourceDirFull"
        }

        $stagePath = Join-Path ([System.IO.Path]::GetTempPath()) ("probdist-vba-sync-" + (Get-Date -Format "yyyyMMdd-HHmmss"))
        $sourceFiles = @(Select-FolderSources -FolderPath $sourceDirFull -RepositoryNames $seen -StagePath $stagePath)
        $sourceLabel = $sourceDirFull

        if (-not $Force) {
            $answer = Read-Host "Import these $($sourceFiles.Count) module(s) into the workbook? [Y/N]"
            if ($answer -notmatch '^\s*(y|yes|s|si)\s*$') {
                Write-Host "Cancelled: the workbook was not changed."
                exit 2
            }
        }
    }

    # Worksheet and ThisWorkbook modules: their code is replaced only when the
    # user agrees, so that a workbook can keep its own sheet code.
    if ($IncludeSheets -and $SkipSheets) {
        throw "Use -IncludeSheets or -SkipSheets, not both."
    }

    $documentFiles = @($sourceFiles | Where-Object { Test-DocumentSource -SourcePath $_.FullName })
    if ($documentFiles.Count -gt 0) {
        $importSheets = [bool]$IncludeSheets
        if (-not $IncludeSheets -and -not $SkipSheets) {
            Write-Host "Worksheet and ThisWorkbook modules in the source:"
            foreach ($file in $documentFiles) {
                Write-Host ("  {0}" -f (Get-VbaComponentName -SourcePath $file.FullName))
            }
            $answer = Read-Host "Replace the code of these $($documentFiles.Count) sheet module(s) as well? [Y/N]"
            $importSheets = ($answer -match '^\s*(y|yes|s|si)\s*$')
        }

        if ($importSheets) {
            Write-Host "Sheet modules: code replaced from the source."
        }
        else {
            $sourceFiles = @($sourceFiles | Where-Object { -not (Test-DocumentSource -SourcePath $_.FullName) })
            Write-Host "Sheet modules: skipped; the workbook keeps its own worksheet and ThisWorkbook code."
            if ($sourceFiles.Count -eq 0) {
                Write-Host "Nothing else to import: the workbook was not changed."
                exit 2
            }
        }
        Write-Host ""
    }

    # -FromTemplate: the workbook is built anew, so there is nothing to back up.
    if ($FromTemplate) {
        $templatePath = Join-Path $repoFullPath "src\workbook\PROBDIST_Template.xlsx"
        New-WorkbookFromTemplate -TemplatePath $templatePath -TargetPath $workbookFullPath
        Write-Host "Template:   $templatePath saved as $workbookFullPath"
    }
    elseif (-not $NoBackup) {
        $directory = [System.IO.Path]::GetDirectoryName($workbookFullPath)
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($workbookFullPath)
        $extension = [System.IO.Path]::GetExtension($workbookFullPath)
        $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $backupPath = Join-Path $directory "$baseName.pre-vba-sync.$timestamp$extension"

        Copy-Item -LiteralPath $workbookFullPath -Destination $backupPath -Force
        Write-Host "Backup: $backupPath"
    }

    Write-Host "Workbook:   $workbookFullPath"
    Write-Host "Repository: $repoFullPath"
    Write-Host "Source:     $sourceLabel"
    Write-Host "Sources:    $($sourceFiles.Count)"
    Write-Host ""

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = [bool]$Visible
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = $msoAutomationSecurityForceDisable

    # Some Excel builds reject Application.Calculation changes until at least
    # one workbook is open (HRESULT 0x800A03EC). Open first, then disable
    # calculation-before-save for this isolated Excel instance.
    $workbook = $excel.Workbooks.Open($workbookFullPath, 0, $false)

    if ($workbook.ReadOnly) {
        throw "Workbook opened read-only. Close any other Excel instance using it."
    }

    $excel.Calculation = $xlCalculationManual
    $excel.CalculateBeforeSave = $false

    $before = Get-SheetSnapshot -Workbook $workbook

    try {
        $vbProject = $workbook.VBProject
        $null = $vbProject.VBComponents.Count
    }
    catch {
        throw (
            "Cannot access the VBA project. Enable Excel Trust Center > Macro Settings > " +
            "'Trust access to the VBA project object model' and ensure the VBA project " +
            "is not password-locked."
        )
    }

    # A production synchronization never removes modules: the regression
    # module already in the workbook stays, and the user is told.
    if ($syncMode -eq "Production") {
        $testModules = @(
            foreach ($component in $vbProject.VBComponents) {
                if ($component.Name -eq "M_STATS_PROBDIST_TEST") { $component.Name }
            }
        )
        if ($testModules.Count -gt 0) {
            Write-Warning ("This workbook already holds the regression module, which a production " +
                "synchronization does not remove: " + ($testModules -join ", ") +
                ". For a production workbook, build it from the template.")
        }
    }

    foreach ($file in $sourceFiles) {
        Sync-VbaSource -VBProject $vbProject -SourcePath $file.FullName
    }

    $after = Get-SheetSnapshot -Workbook $workbook
    Assert-SheetSnapshotUnchanged -Before $before -After $after

    # The synchronization uses manual calculation only to avoid recalculation
    # while modules are being replaced. The workbook must leave this isolated
    # Excel instance in Automatic mode before it is saved, in either mode.
    $excel.Calculation = $xlCalculationAutomatic
    $excel.CalculateBeforeSave = $true

    $workbook.Save()
    $saved = $true

    # The backup is only a transactional safety copy. Remove it after a
    # successful save so repeated synchronizations do not accumulate files.
    if (-not [string]::IsNullOrWhiteSpace($backupPath) -and
        (Test-Path -LiteralPath $backupPath -PathType Leaf)) {
        Remove-Item -LiteralPath $backupPath -Force
        Write-Host "Backup removed: $backupPath"
    }

    Write-Host ""
    if ($FromTemplate) {
        Write-Host "SUCCESS: workbook built from the template and VBA synchronized."
    }
    else {
        Write-Host "SUCCESS: VBA synchronized. Workbook sheets were not removed or recreated."
    }
}
catch {
    Write-Error $_
    if ($null -ne $workbook -and -not $saved) {
        Write-Warning "Synchronization failed. Workbook will close without saving."
    }

    if (-not [string]::IsNullOrWhiteSpace($backupPath) -and
        (Test-Path -LiteralPath $backupPath -PathType Leaf)) {
        Write-Warning "Safety backup retained: $backupPath"
    }

    exit 1
}
finally {
    if (-not [string]::IsNullOrWhiteSpace($stagePath) -and
        (Test-Path -LiteralPath $stagePath -PathType Container)) {
        try { Remove-Item -LiteralPath $stagePath -Recurse -Force } catch {}
    }

    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {}
        try {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
        } catch {}
    }

    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        try {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        } catch {}
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
