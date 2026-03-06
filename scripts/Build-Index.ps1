Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -Force -DisableNameChecking

$config = Get-ShinsaConfig -ScriptPath $MyInvocation.MyCommand.Path
$paths = Get-ShinsaPaths -Config $config
Ensure-ShinsaDirectory -Paths @($config.paths.indexRoot, $config.paths.workRoot)

$caseSource = Read-ShinsaJson -Path $paths.LedgerCloneCasesPath
$contactSource = Read-ShinsaJson -Path $paths.LedgerCloneContactsPath
$reviewState = Get-ReviewState -Paths $paths

$mailMetaFiles = Get-ChildItem -Path $paths.CloneMailRoot -Recurse -Filter 'meta.json' -File -ErrorAction SilentlyContinue
$mails = foreach ($mailMetaFile in $mailMetaFiles) {
    $meta = Read-ShinsaJson -Path $mailMetaFile.FullName
    $relativeBodyPath = Join-Path $mailMetaFile.DirectoryName $meta.body_path
    [pscustomobject]@{
        mail_id = $meta.mail_id
        entry_id = $meta.entry_id
        case_id = $meta.case_id
        sender_name = $meta.sender_name
        sender_email = $meta.sender_email
        subject = $meta.subject
        received_at = $meta.received_at
        body_path = $relativeBodyPath
        msg_path = if ($meta.msg_path) { Join-Path $mailMetaFile.DirectoryName $meta.msg_path } else { '' }
        attachments = @($meta.attachments)
        folder_path = $mailMetaFile.DirectoryName
    }
}

$cases = foreach ($case in $caseSource.organizations) {
    $caseMails = @($mails | Where-Object { $_.case_id -eq $case.case_id -or $_.sender_email -eq $case.contact_email })
    $contact = $contactSource.contacts | Where-Object { $_.contact_email -eq $case.contact_email } | Select-Object -First 1
    $review = $reviewState.reviews | Where-Object { $_.case_id -eq $case.case_id } | Select-Object -First 1
    $attachmentSum = ($caseMails | ForEach-Object { @($_.attachments).Count } | Measure-Object -Sum).Sum
    if (-not $attachmentSum) { $attachmentSum = 0 }
    [pscustomobject]@{
        case_id = $case.case_id
        organization_name = $case.organization_name
        contact_name = if ($contact) { $contact.contact_name } else { $case.contact_name }
        contact_email = $case.contact_email
        status = if ($review) { $review.status } else { $case.status }
        assigned_to = if ($review) { $review.assigned_to } else { $case.assigned_to }
        review_note = if ($review) { $review.review_note } else { $case.review_note }
        missing_documents = if ($review) { @($review.missing_documents) } else { @($case.missing_documents) }
        mail_count = $caseMails.Count
        attachment_count = $attachmentSum
        case_folder_path = Join-Path $paths.CloneCaseRoot $case.case_id
        latest_mail_at = ($caseMails | Sort-Object received_at -Descending | Select-Object -First 1).received_at
    }
}

Write-ShinsaJson -Path $paths.CaseIndexPath -Data $cases
Write-ShinsaJson -Path $paths.MailIndexPath -Data $mails
Write-ShinsaJson -Path $paths.ContactsIndexPath -Data $contactSource.contacts

Write-ShinsaLog -Message 'Indexes rebuilt.' -ScriptPath $MyInvocation.MyCommand.Path

