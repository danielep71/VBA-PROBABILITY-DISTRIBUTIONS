param([int]$Minutes = 60)
# Prints what Windows recorded about EXCEL.EXE stopping in the last $Minutes minutes: the Application log's
# "Application Error" (faulting module, exception code, offset), "Windows Error Reporting" and "Application
# Hang" events. A test that fails with 0x800706BA or 0x800706BE ("the RPC server is unavailable", "the remote
# procedure call failed") lost its Excel process; these events name the module that stopped it. Read-only.
$ErrorActionPreference = 'Continue'
try {
    $events = @(Get-WinEvent -FilterHashtable @{ LogName = 'Application'; StartTime = (Get-Date).AddMinutes(-$Minutes) } -ErrorAction Stop |
        Where-Object { $_.ProviderName -in @('Application Error', 'Windows Error Reporting', 'Application Hang') -and $_.Message -match 'EXCEL' })
} catch {
    Write-Output ('The Application log could not be read: ' + $_.Exception.Message)
    exit 0
}
if ($events.Count -eq 0) { Write-Output "No EXCEL.EXE fault in the Application log of the last $Minutes minutes."; exit 0 }
foreach ($item in $events | Select-Object -First 6) {
    Write-Output ('--- ' + $item.TimeCreated.ToString('s') + ' ' + $item.ProviderName + ' (event ' + $item.Id + ')')
    Write-Output (($item.Message -split "`r?`n" | Where-Object { $_.Trim() } | Select-Object -First 14) -join "`n")
}
