Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -Force -DisableNameChecking

$config = Get-ShinsaConfig -ScriptPath $MyInvocation.MyCommand.Path
$paths = Get-ShinsaPaths -Config $config
Ensure-ShinsaDirectory -Paths @($paths.CloneRoot, $paths.CloneLedgerRoot, $paths.CloneCaseRoot, $paths.CloneMailRoot)

Copy-Item (Join-Path $config.paths.onedriveLedgerRoot '*') $paths.CloneLedgerRoot -Recurse -Force
Copy-Item (Join-Path $config.paths.onedriveCaseRoot '*') $paths.CloneCaseRoot -Recurse -Force
if (Test-Path $config.paths.mailSourceRoot) {
    Copy-Item (Join-Path $config.paths.mailSourceRoot '*') $paths.CloneMailRoot -Recurse -Force
}

Write-ShinsaLog -Message 'Clone sync completed from OneDrive/local sources.' -ScriptPath $MyInvocation.MyCommand.Path

