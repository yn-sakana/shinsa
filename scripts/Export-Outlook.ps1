Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -Force -DisableNameChecking

$ErrorActionPreference = 'Stop'
$appRoot = Get-ShinsaAppRoot -ScriptPath $MyInvocation.MyCommand.Path
$mailConfigPath = Join-Path $appRoot 'config\mail_accounts.txt'

if (-not (Test-Path $mailConfigPath)) {
    throw "Mail account list is missing: $mailConfigPath"
}

try {
    $outlook = New-Object -ComObject Outlook.Application
}
catch {
    throw 'Outlook COM could not be started.'
}

try {
    $count = $outlook.Run('Shinsa_ExportRegisteredMailboxes', $appRoot)
}
catch {
    $modulePath = Join-Path $appRoot 'VBA\ShinsaOutlookExport.bas'
    $message = @(
        "Outlook VBA macro 'Shinsa_ExportRegisteredMailboxes' is not available."
        "Import this module into Outlook VBA first: $modulePath"
        "Mail account list: $mailConfigPath"
    ) -join ' '
    throw $message
}

Write-Host ("Outlook export completed via VBA. Exported items: {0}" -f $count) -ForegroundColor Cyan
