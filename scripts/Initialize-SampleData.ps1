Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -Force -DisableNameChecking

$ErrorActionPreference = 'Stop'
$config = Get-ShinsaConfig -ScriptPath $MyInvocation.MyCommand.Path
$paths = Get-ShinsaDataPaths -Config $config

Ensure-ShinsaDirectory -Paths @(
    (Split-Path -Parent $paths.SharePointLedgerPath),
    $paths.SharePointCaseRoot,
    $paths.MailArchiveRoot,
    $paths.JsonRoot
)

Write-ShinsaJson -Path $paths.SharePointLedgerPath -Data @(
    [pscustomobject]@{
        case_id = 'CASE-0001'
        receipt_no = 'R-2026-0001'
        organization_name = 'Sample Foundation'
        contact_name = 'Hanako Sato'
        contact_email = 'grant@example.org'
        status = 'checked'
        assigned_to = 'reviewer1'
        missing_documents = ''
        review_note_public = 'ready'
    }
)

$caseFolder = Join-Path $paths.SharePointCaseRoot 'CASE-0001'
Ensure-ShinsaDirectory -Paths @($caseFolder)
Set-Content -Path (Join-Path $caseFolder 'application.txt') -Value 'Application form placeholder' -Encoding UTF8

$mailFolder = $paths.MailArchiveRoot
Ensure-ShinsaDirectory -Paths @($mailFolder)
Set-Content -Path (Join-Path $mailFolder 'body.txt') -Value 'Please find the attached application documents.' -Encoding UTF8
Set-Content -Path (Join-Path $mailFolder 'application.pdf') -Value 'PDF placeholder' -Encoding UTF8
Set-Content -Path (Join-Path $mailFolder 'budget.xlsx') -Value 'XLSX placeholder' -Encoding UTF8
Write-ShinsaJson -Path (Join-Path $mailFolder 'mail_0001.json') -Data ([pscustomobject]@{
        mail_id = 'MAIL-0001'
        entry_id = 'ENTRY-0001'
        mailbox_address = 'grants@example.org'
        folder_path = 'Inbox/Grant Applications'
        received_at = '2026-03-06T10:15:30+09:00'
        sender_name = 'Hanako Sato'
        sender_email = 'grant@example.org'
        subject = 'Grant application submission'
        body_path = 'body.txt'
        msg_path = ''
        attachment_paths = @('application.pdf', 'budget.xlsx')
    })

Ensure-ShinsaState -Paths $paths
Write-Host 'Sample data initialized.' -ForegroundColor Cyan
