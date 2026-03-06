Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -Force -DisableNameChecking

$ErrorActionPreference = 'Stop'
$config = Get-ShinsaConfig -ScriptPath $MyInvocation.MyCommand.Path
$paths = Get-ShinsaDataPaths -Config $config

Ensure-ShinsaState -Paths $paths

if ([string]::IsNullOrWhiteSpace($paths.SharePointLedgerPath) -or -not (Test-Path $paths.SharePointLedgerPath)) {
    throw "Ledger source path is invalid: $($paths.SharePointLedgerPath)"
}

if ([string]::IsNullOrWhiteSpace($paths.SharePointCaseRoot) -or -not (Test-Path $paths.SharePointCaseRoot)) {
    throw "Case folder root is invalid: $($paths.SharePointCaseRoot)"
}

$ledgerRecords = @(Import-ShinsaLedgerRecords -Config $config -Paths $paths | Sort-Object case_id)
$mailRecords = @(Import-ShinsaMailRecords -Paths $paths)
$folderRecords = @(Import-ShinsaFolderRecords -Paths $paths)

Write-ShinsaJson -Path $paths.LedgerJsonPath -Data $ledgerRecords
Write-ShinsaJson -Path $paths.MailsJsonPath -Data $mailRecords
Write-ShinsaJson -Path $paths.FoldersJsonPath -Data $folderRecords

$cache = Read-ShinsaCache -Paths $paths
$uiState = ConvertTo-ShinsaMap -InputObject $cache.ui_state
$uiState['last_sync_at'] = (Get-Date).ToString('o')
$cache.ui_state = [pscustomobject]$uiState
Save-ShinsaCache -Paths $paths -Cache $cache

Write-Host ''
Write-Host 'sync completed' -ForegroundColor Cyan
Write-Host ("  ledger : {0} cases" -f $ledgerRecords.Count)
Write-Host ("  mails  : {0} records" -f $mailRecords.Count)
Write-Host ("  folders: {0} files" -f $folderRecords.Count)
Write-Host ("  json   : {0}" -f $paths.JsonRoot)
