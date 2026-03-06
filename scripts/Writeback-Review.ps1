param(
    [switch]$Force
)

Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -Force -DisableNameChecking

$ErrorActionPreference = 'Stop'
$config = Get-ShinsaConfig -ScriptPath $MyInvocation.MyCommand.Path
$paths = Get-ShinsaDataPaths -Config $config

$plan = Get-ShinsaLedgerWritebackPlan -Config $config -Paths $paths

if ($plan.case_count -eq 0) {
    Write-Host 'writeback skipped: no ledger changes.' -ForegroundColor Yellow
    return
}

Write-Host ''
Write-Host 'pending writeback' -ForegroundColor Cyan
foreach ($change in @($plan.changes)) {
    $fields = @($change.changes.PSObject.Properties.Name) -join ', '
    Write-Host ("  {0}: {1}" -f $change.case_id, $fields)
}
Write-Host ("  total : {0} cases / {1} fields" -f $plan.case_count, $plan.change_count)

if (-not $Force) {
    $answer = Read-Host 'write changes back to the source ledger? [y/N]'
    if ($answer -notmatch '^(?i)y(es)?$') {
        Write-Host 'writeback cancelled.' -ForegroundColor Yellow
        return
    }
}

Invoke-ShinsaLedgerWriteback -Config $config -Paths $paths -Plan $plan
Write-ShinsaJson -Path $paths.LedgerJsonPath -Data @(Import-ShinsaLedgerRecords -Config $config -Paths $paths | Sort-Object case_id)

Write-Host 'writeback completed.' -ForegroundColor Green
