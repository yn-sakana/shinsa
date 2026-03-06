param(
    [switch]$Force
)

Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -Force -DisableNameChecking

$ErrorActionPreference = 'Stop'
$config = Get-ShinsaConfig -ScriptPath $MyInvocation.MyCommand.Path
$paths = Get-ShinsaDataPaths -Config $config

$sourceNames = Get-ShinsaSourceNames -Config $config
$totalCases = 0

foreach ($name in $sourceNames) {
    $src = $config.sources[$name]
    if ($null -eq $src.editable_columns -or @($src.editable_columns).Count -eq 0) { continue }

    $plan = Get-ShinsaWritebackPlan -Config $config -Paths $paths -SourceName $name

    if ($plan.case_count -eq 0) { continue }

    Write-Host ''
    Write-Host ("pending writeback: {0}" -f $name) -ForegroundColor Cyan
    foreach ($change in @($plan.changes)) {
        $fields = @($change.changes.PSObject.Properties.Name) -join ', '
        Write-Host ("  {0}: {1}" -f $change.key_value, $fields)
    }
    Write-Host ("  total: {0} records / {1} fields" -f $plan.case_count, $plan.change_count)

    if (-not $Force) {
        $answer = Read-Host "write changes back to '$name'? [y/N]"
        if ($answer -notmatch '^(?i)y(es)?$') {
            Write-Host 'skipped.' -ForegroundColor Yellow
            continue
        }
    }

    Invoke-ShinsaWriteback -Config $config -SourceName $name -Plan $plan

    $sourcePath = Get-ShinsaSourcePath -Config $config -SourceName $name
    $refreshed = @(Import-ShinsaFieldsRecords -SourceConfig $src -SourcePath $sourcePath)
    $jsonPath = Join-Path $paths.JsonRoot $src.file
    Write-ShinsaJson -Path $jsonPath -Data $refreshed

    Write-Host ("writeback completed: {0}" -f $name) -ForegroundColor Green
    $totalCases += $plan.case_count
}

if ($totalCases -eq 0) {
    Write-Host 'writeback skipped: no changes.' -ForegroundColor Yellow
}
