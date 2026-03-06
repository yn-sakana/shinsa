Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -Force -DisableNameChecking

$config = Get-ShinsaConfig -ScriptPath $MyInvocation.MyCommand.Path
$paths = Get-ShinsaPaths -Config $config
Ensure-ShinsaDirectory -Paths @(
    $config.paths.onedriveLedgerRoot,
    $config.paths.onedriveCaseRoot,
    $config.paths.mailSourceRoot,
    $paths.CloneRoot,
    $config.paths.indexRoot,
    $config.paths.workRoot,
    $config.paths.logRoot
)

$cases = [pscustomobject]@{
    organizations = @(
        [pscustomobject]@{
            case_id = 'CASE-0001'
            organization_name = 'Sample Foundation'
            contact_name = 'Hanako Sato'
            contact_email = 'grant@example.org'
            status = 'received'
            assigned_to = 'unassigned'
            review_note = ''
            missing_documents = @()
        }
    )
}
$contacts = [pscustomobject]@{
    contacts = @(
        [pscustomobject]@{
            organization_name = 'Sample Foundation'
            contact_name = 'Hanako Sato'
            contact_email = 'grant@example.org'
            role = 'Primary'
        }
    )
}

Write-ShinsaJson -Path (Join-Path $config.paths.onedriveLedgerRoot $config.ledger.casesFileName) -Data $cases
Write-ShinsaJson -Path (Join-Path $config.paths.onedriveLedgerRoot $config.ledger.contactsFileName) -Data $contacts

$caseFolder = Join-Path $config.paths.onedriveCaseRoot 'CASE-0001'
Ensure-ShinsaDirectory -Paths @($caseFolder)
Set-Content -Path (Join-Path $caseFolder 'application.txt') -Value 'Application form placeholder' -Encoding UTF8

$mailDir = Join-Path $config.paths.mailSourceRoot 'sample_mail_0001'
Ensure-ShinsaDirectory -Paths @($mailDir, (Join-Path $mailDir 'attachments'))
Write-ShinsaJson -Path (Join-Path $mailDir 'meta.json') -Data ([pscustomobject]@{
    mail_id = 'MAIL-0001'
    entry_id = 'MAIL-0001'
    case_id = 'CASE-0001'
    sender_name = 'Hanako Sato'
    sender_email = 'grant@example.org'
    subject = 'Grant application submission'
    received_at = '2026-03-06T10:15:30+09:00'
    body_path = 'body.txt'
    msg_path = ''
    attachments = @('application.pdf', 'budget.xlsx')
})
Set-Content -Path (Join-Path $mailDir 'body.txt') -Value 'Please find the attached application documents.' -Encoding UTF8
Set-Content -Path (Join-Path $mailDir 'attachments\application.pdf') -Value 'PDF placeholder' -Encoding UTF8
Set-Content -Path (Join-Path $mailDir 'attachments\budget.xlsx') -Value 'XLSX placeholder' -Encoding UTF8

Write-ShinsaLog -Message 'Sample data initialized.' -ScriptPath $MyInvocation.MyCommand.Path

