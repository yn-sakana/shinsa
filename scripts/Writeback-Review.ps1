Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -Force -DisableNameChecking

$config = Get-ShinsaConfig -ScriptPath $MyInvocation.MyCommand.Path
$paths = Get-ShinsaPaths -Config $config

$reviewState = Get-ReviewState -Paths $paths
$source = Read-ShinsaJson -Path $paths.SourceCasesPath

foreach ($review in $reviewState.reviews) {
    $target = $source.organizations | Where-Object { $_.case_id -eq $review.case_id } | Select-Object -First 1
    if (-not $target) { continue }

    foreach ($field in $config.ledger.editableFields) {
        if ($review.PSObject.Properties.Name -contains $field) {
            $target.$field = $review.$field
        }
    }
}

Write-ShinsaJson -Path $paths.SourceCasesPath -Data $source
Write-ShinsaLog -Message 'Review data written back to OneDrive source ledger.' -ScriptPath $MyInvocation.MyCommand.Path

