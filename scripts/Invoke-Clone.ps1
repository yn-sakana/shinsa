$ErrorActionPreference = 'Stop'
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$appRoot = Split-Path -Parent $scriptRoot
$config = & (Join-Path $scriptRoot 'Get-Config.ps1')

$cloneRoot = $config.paths.cloneRoot
$ledgerClone = Join-Path $cloneRoot 'ledger'
$caseClone = Join-Path $cloneRoot 'cases'
$mailClone = Join-Path $cloneRoot 'mail'

New-Item -ItemType Directory -Force $cloneRoot, $ledgerClone, $caseClone, $mailClone | Out-Null
Copy-Item (Join-Path $config.paths.onedriveLedgerRoot '*') $ledgerClone -Recurse -Force
Copy-Item (Join-Path $config.paths.onedriveCaseRoot '*') $caseClone -Recurse -Force
Copy-Item (Join-Path $config.paths.mailSourceRoot '*') $mailClone -Recurse -Force

Write-Host "Clone refresh complete."
Write-Host "  Ledger: $ledgerClone"
Write-Host "  Cases : $caseClone"
Write-Host "  Mail  : $mailClone"
