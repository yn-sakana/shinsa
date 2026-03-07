Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -Force -DisableNameChecking

$ErrorActionPreference = 'Stop'
$config = Get-ShinsaConfig -ScriptPath $MyInvocation.MyCommand.Path
$paths = Get-ShinsaDataPaths -Config $config

$mailSourcePath = Get-ShinsaSourcePath -Config $config -SourceName 'mail'
if ([string]::IsNullOrWhiteSpace($mailSourcePath)) {
    throw "Mail source_path is not configured. Set it in config/config.local.json under sources.mail.source_path"
}

$selfAddress = [string]$config.mail.self_address
if ([string]::IsNullOrWhiteSpace($selfAddress)) {
    throw "mail.self_address is not configured. Set it in config/config.local.json"
}

$stateFilePath = Join-Path $paths.JsonRoot 'outlook_exported.json'

# Resolve to absolute paths for VBA
$mailSourcePath = [System.IO.Path]::GetFullPath($mailSourcePath)
$stateFilePath = [System.IO.Path]::GetFullPath($stateFilePath)

try {
    $outlook = New-Object -ComObject Outlook.Application
}
catch {
    throw 'Outlook COM could not be started.'
}

try {
    $count = $outlook.Run('Shinsa_ExportMail', $mailSourcePath, $stateFilePath, $selfAddress)
}
catch {
    $modulePath = Join-Path (Get-ShinsaAppRoot -ScriptPath $MyInvocation.MyCommand.Path) 'VBA\ShinsaOutlookExport.bas'
    $message = @(
        "Outlook VBA macro 'Shinsa_ExportMail' is not available."
        "Import this module into Outlook VBA first: $modulePath"
    ) -join ' '
    throw $message
}

Write-Host ("Outlook export completed. New items: {0}" -f $count) -ForegroundColor Cyan
