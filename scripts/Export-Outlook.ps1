Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -Force -DisableNameChecking

$config = Get-ShinsaConfig -ScriptPath $MyInvocation.MyCommand.Path
$appRoot = Split-Path -Parent $PSScriptRoot
$mailConfigPath = Join-Path $appRoot 'config\mail_accounts.txt'

if (-not (Test-Path $mailConfigPath)) {
    throw "Mail account list is missing: $mailConfigPath"
}

try {
    $outlook = New-Object -ComObject Outlook.Application
} catch {
    Write-ShinsaLog -Message 'Outlook COM could not be started.' -Level ERROR -ScriptPath $MyInvocation.MyCommand.Path
    throw
}

try {
    $count = $outlook.Run('Shinsa_ExportRegisteredMailboxes', $appRoot)
} catch {
    $modulePath = Join-Path $appRoot 'VBA\ShinsaOutlookExport.bas'
    $message = @(
        "Outlook VBA macro 'Shinsa_ExportRegisteredMailboxes' is not available."
        "Import this module into Outlook VBA first: $modulePath"
        "Mail account list: $mailConfigPath"
    ) -join ' '
    Write-ShinsaLog -Message $message -Level ERROR -ScriptPath $MyInvocation.MyCommand.Path
    throw $message
}

Write-ShinsaLog -Message ("Outlook export completed via VBA. Exported items: {0}" -f $count) -ScriptPath $MyInvocation.MyCommand.Path
