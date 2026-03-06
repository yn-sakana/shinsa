$ErrorActionPreference = 'Stop'
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$appRoot = Split-Path -Parent $scriptRoot

function Read-JsonFile {
    param([string]$Path)
    Get-Content $Path -Raw -Encoding UTF8 | ConvertFrom-Json
}

$base = Read-JsonFile (Join-Path $appRoot 'config\config.base.json')
$local = Read-JsonFile (Join-Path $appRoot 'config\config.local.json')

$config = [ordered]@{
    app = $base.app
    paths = @{}
    ledger = $base.ledger
    gui = $base.gui
}

foreach ($name in $base.paths.PSObject.Properties.Name) {
    $config.paths[$name] = $base.paths.$name
}
foreach ($name in $local.paths.PSObject.Properties.Name) {
    $config.paths[$name] = $local.paths.$name
}

[pscustomobject]$config
