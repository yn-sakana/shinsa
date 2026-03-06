Set-StrictMode -Version Latest

function Get-ShinsaAppRoot {
    param([Parameter(Mandatory = $true)][string]$ScriptPath)

    $scriptDirectory = Split-Path -Parent $ScriptPath
    if (Test-Path (Join-Path $scriptDirectory 'config\config.base.json')) {
        return $scriptDirectory
    }

    Split-Path -Parent $scriptDirectory
}

function ConvertTo-ShinsaConfigValue {
    param($Value)

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [System.Management.Automation.PSCustomObject] -or $Value -is [System.Collections.IDictionary]) {
        return ConvertTo-ShinsaMap -InputObject $Value
    }

    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        return @($Value | ForEach-Object { ConvertTo-ShinsaConfigValue -Value $_ })
    }

    $Value
}

function ConvertTo-ShinsaMap {
    param($InputObject)

    $map = [ordered]@{}
    if ($null -eq $InputObject) {
        return $map
    }

    if ($InputObject -is [System.Collections.IDictionary]) {
        foreach ($key in $InputObject.Keys) {
            $map[[string]$key] = ConvertTo-ShinsaConfigValue -Value $InputObject[$key]
        }
        return $map
    }

    foreach ($property in $InputObject.PSObject.Properties) {
        $map[$property.Name] = ConvertTo-ShinsaConfigValue -Value $property.Value
    }

    $map
}

function Merge-ShinsaMap {
    param(
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Base,
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Overlay
    )

    foreach ($key in $Overlay.Keys) {
        if (
            $Base.Contains($key) -and
            $Base[$key] -is [System.Collections.IDictionary] -and
            $Overlay[$key] -is [System.Collections.IDictionary]
        ) {
            Merge-ShinsaMap -Base $Base[$key] -Overlay $Overlay[$key]
            continue
        }

        $Base[$key] = ConvertTo-ShinsaConfigValue -Value $Overlay[$key]
    }
}

function Read-ShinsaJson {
    param([Parameter(Mandatory = $true)][string]$Path)

    Get-Content -Path $Path -Raw -Encoding UTF8 | ConvertFrom-Json
}

function Write-ShinsaJson {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)]$Data
    )

    $directory = Split-Path -Parent $Path
    if (-not [string]::IsNullOrWhiteSpace($directory) -and -not (Test-Path $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }

    if (
        $Data -is [System.Collections.IEnumerable] -and
        -not ($Data -is [string]) -and
        -not ($Data -is [System.Collections.IDictionary]) -and
        -not ($Data -is [System.Management.Automation.PSCustomObject])
    ) {
        $items = @($Data)
        if ($items.Count -eq 0) {
            $json = '[]'
        }
        else {
            $json = "[`r`n" + (($items | ForEach-Object { $_ | ConvertTo-Json -Depth 16 }) -join ",`r`n") + "`r`n]"
        }
    }
    else {
        $json = ConvertTo-Json -InputObject $Data -Depth 16
    }

    Set-Content -Path $Path -Value $json -Encoding UTF8
}

function Resolve-ShinsaPath {
    param(
        [Parameter(Mandatory = $true)][string]$AppRoot,
        [string]$PathValue
    )

    if ([string]::IsNullOrWhiteSpace($PathValue)) {
        return ''
    }

    if ([System.IO.Path]::IsPathRooted($PathValue)) {
        return [System.IO.Path]::GetFullPath($PathValue)
    }

    [System.IO.Path]::GetFullPath((Join-Path $AppRoot $PathValue))
}

function Ensure-ShinsaDirectory {
    param([string[]]$Paths)

    foreach ($path in @($Paths)) {
        if ([string]::IsNullOrWhiteSpace($path)) {
            continue
        }

        if (-not (Test-Path $path)) {
            New-Item -ItemType Directory -Path $path -Force | Out-Null
        }
    }
}

function Get-ShinsaConfig {
    param([Parameter(Mandatory = $true)][string]$ScriptPath)

    $appRoot = Get-ShinsaAppRoot -ScriptPath $ScriptPath
    $basePath = Join-Path $appRoot 'config\config.base.json'
    $localPath = Join-Path $appRoot 'config\config.local.json'
    $samplePath = Join-Path $appRoot 'config\config.local.sample.json'

    if (-not (Test-Path $localPath)) {
        Copy-Item -Path $samplePath -Destination $localPath -Force
    }

    $baseConfig = ConvertTo-ShinsaMap -InputObject (Read-ShinsaJson -Path $basePath)
    $localConfig = ConvertTo-ShinsaMap -InputObject (Read-ShinsaJson -Path $localPath)
    Merge-ShinsaMap -Base $baseConfig -Overlay $localConfig
    $baseConfig['app_root'] = $appRoot

    [pscustomobject]$baseConfig
}

function Get-ShinsaDataPaths {
    param([Parameter(Mandatory = $true)]$Config)

    $appRoot = [string]$Config.app_root
    $jsonRoot = Resolve-ShinsaPath -AppRoot $appRoot -PathValue $Config.paths.json_root

    [pscustomobject]@{
        AppRoot = $appRoot
        JsonRoot = $jsonRoot
        LedgerJsonPath = Join-Path $jsonRoot 'ledger.json'
        MailsJsonPath = Join-Path $jsonRoot 'mails.json'
        FoldersJsonPath = Join-Path $jsonRoot 'folders.json'
        CacheJsonPath = Join-Path $jsonRoot 'cache.json'
        MailArchiveRoot = Resolve-ShinsaPath -AppRoot $appRoot -PathValue $Config.paths.mail_archive_root
        SharePointLedgerPath = Resolve-ShinsaPath -AppRoot $appRoot -PathValue $Config.paths.sharepoint_ledger_path
        SharePointCaseRoot = Resolve-ShinsaPath -AppRoot $appRoot -PathValue $Config.paths.sharepoint_case_root
    }
}

function Get-DefaultCacheState {
    [ordered]@{
        mail_links = @()
        mail_progress = @()
        ui_state = [ordered]@{}
    }
}

function Ensure-ShinsaCacheShape {
    param($Cache)

    $normalized = ConvertTo-ShinsaMap -InputObject $Cache

    if (-not $normalized.Contains('mail_links')) {
        $normalized['mail_links'] = @()
    }
    elseif ($normalized['mail_links'] -is [System.Collections.IDictionary]) {
        $links = @()
        foreach ($mailId in $normalized['mail_links'].Keys) {
            $links += [pscustomobject]@{
                mail_id = [string]$mailId
                case_id = [string]$normalized['mail_links'][$mailId]
                mode = 'manual'
                updated_at = ''
            }
        }
        $normalized['mail_links'] = $links
    }
    else {
        $normalized['mail_links'] = @($normalized['mail_links'])
    }

    if (-not $normalized.Contains('mail_progress')) {
        $normalized['mail_progress'] = @()
    }
    else {
        $normalized['mail_progress'] = @($normalized['mail_progress'])
    }

    if (-not $normalized.Contains('ui_state') -or -not ($normalized['ui_state'] -is [System.Collections.IDictionary])) {
        $normalized['ui_state'] = [ordered]@{}
    }

    [pscustomobject]$normalized
}

function Ensure-ShinsaState {
    param([Parameter(Mandatory = $true)]$Paths)

    Ensure-ShinsaDirectory -Paths @($Paths.JsonRoot)
    if (-not (Test-Path $Paths.CacheJsonPath)) {
        Write-ShinsaJson -Path $Paths.CacheJsonPath -Data (Get-DefaultCacheState)
    }
}

function Read-ShinsaCache {
    param([Parameter(Mandatory = $true)]$Paths)

    Ensure-ShinsaState -Paths $Paths
    Ensure-ShinsaCacheShape -Cache (Read-ShinsaJson -Path $Paths.CacheJsonPath)
}

function Save-ShinsaCache {
    param(
        [Parameter(Mandatory = $true)]$Paths,
        [Parameter(Mandatory = $true)]$Cache
    )

    Write-ShinsaJson -Path $Paths.CacheJsonPath -Data (Ensure-ShinsaCacheShape -Cache $Cache)
}

function ConvertTo-ShinsaString {
    param($Value)

    if ($null -eq $Value) {
        return ''
    }

    if ($Value -is [datetime]) {
        return $Value.ToString('o')
    }

    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        return (@($Value | ForEach-Object { ConvertTo-ShinsaString -Value $_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join ', ')
    }

    [string]$Value
}

function Test-ShinsaRecordProperty {
    param(
        [Parameter(Mandatory = $true)]$Record,
        [Parameter(Mandatory = $true)][string]$Name
    )

    if ($Record -is [System.Collections.IDictionary]) {
        return $Record.Contains($Name)
    }

    return $Record.PSObject.Properties.Name -contains $Name
}

function Get-ShinsaRecordValue {
    param(
        [Parameter(Mandatory = $true)]$Record,
        [Parameter(Mandatory = $true)][string]$Name
    )

    if ($Record -is [System.Collections.IDictionary]) {
        if ($Record.Contains($Name)) {
            return $Record[$Name]
        }

        return $null
    }

    $property = $Record.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $null
    }

    $property.Value
}

function Set-ShinsaRecordValue {
    param(
        [Parameter(Mandatory = $true)]$Record,
        [Parameter(Mandatory = $true)][string]$Name,
        $Value
    )

    if ($Record -is [System.Collections.IDictionary]) {
        $Record[$Name] = $Value
        return
    }

    if ($Record.PSObject.Properties.Name -contains $Name) {
        $Record.$Name = $Value
    }
    else {
        $Record | Add-Member -NotePropertyName $Name -NotePropertyValue $Value
    }
}

function Copy-ShinsaRecord {
    param([Parameter(Mandatory = $true)]$Record)

    [pscustomobject](ConvertTo-ShinsaMap -InputObject $Record)
}

function Convert-SourceRowToLedgerRecord {
    param(
        [Parameter(Mandatory = $true)]$Row,
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)][string]$LedgerPath,
        [string]$SheetName = '',
        [int]$RowId = 0
    )

    $record = [ordered]@{}

    foreach ($propertyName in (ConvertTo-ShinsaMap -InputObject $Row).Keys) {
        $record[$propertyName] = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $Row -Name $propertyName)
    }

    foreach ($logicalName in $Config.ledger.columns.Keys) {
        $sourceName = [string]$Config.ledger.columns[$logicalName]
        if (Test-ShinsaRecordProperty -Record $Row -Name $sourceName) {
            $record[$logicalName] = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $Row -Name $sourceName)
        }
        elseif (Test-ShinsaRecordProperty -Record $Row -Name $logicalName) {
            $record[$logicalName] = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $Row -Name $logicalName)
        }
    }

    $keyColumn = [string]$Config.ledger.key_column
    $keyValue = ''
    if ($record.Contains($keyColumn)) {
        $keyValue = ConvertTo-ShinsaString -Value $record[$keyColumn]
    }
    elseif ($record.Contains('case_id')) {
        $keyValue = ConvertTo-ShinsaString -Value $record['case_id']
    }

    if ([string]::IsNullOrWhiteSpace($keyValue)) {
        return $null
    }

    $record['case_id'] = $keyValue
    $record['ledger_path'] = $LedgerPath
    $record['ledger_sheet'] = $SheetName
    $record['ledger_row_id'] = $RowId

    foreach ($requiredField in @('receipt_no', 'organization_name', 'contact_name', 'contact_email', 'status', 'assigned_to', 'missing_documents', 'review_note_public')) {
        if (-not $record.Contains($requiredField)) {
            $record[$requiredField] = ''
        }
    }

    [pscustomobject]$record
}

function Import-ShinsaLedgerFromJson {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)]$Paths
    )

    $source = Read-ShinsaJson -Path $Paths.SharePointLedgerPath
    if ($source -is [System.Collections.IEnumerable] -and -not ($source -is [string]) -and -not ($source -is [System.Management.Automation.PSCustomObject])) {
        $rows = @($source)
    }
    elseif (Test-ShinsaRecordProperty -Record $source -Name 'organizations') {
        $rows = @($source.organizations)
    }
    elseif (Test-ShinsaRecordProperty -Record $source -Name 'rows') {
        $rows = @($source.rows)
    }
    else {
        $rows = @($source)
    }

    $records = @()
    $rowIndex = 2
    foreach ($row in $rows) {
        $record = Convert-SourceRowToLedgerRecord -Row $row -Config $Config -LedgerPath $Paths.SharePointLedgerPath -RowId $rowIndex
        if ($null -ne $record) {
            $records += $record
        }
        $rowIndex += 1
    }

    $records
}

function ConvertFrom-ExcelCellValue {
    param(
        $Value,
        [string]$HeaderName
    )

    if ($null -eq $Value) {
        return ''
    }

    if ($Value -is [double] -and $HeaderName -match '(?i)(date|time)') {
        return ([datetime]::FromOADate($Value)).ToString('o')
    }

    if ($Value -is [double] -and [math]::Floor($Value) -eq $Value) {
        return [string][int64]$Value
    }

    [string]$Value
}

function Import-ShinsaLedgerFromExcel {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)]$Paths
    )

    $excel = $null
    $workbook = $null
    $worksheet = $null
    $usedRange = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($Paths.SharePointLedgerPath, $false, $true)

        if ([string]::IsNullOrWhiteSpace([string]$Config.ledger.sheet_name)) {
            $worksheet = $workbook.Worksheets.Item(1)
        }
        else {
            $worksheet = $workbook.Worksheets.Item([string]$Config.ledger.sheet_name)
        }

        $usedRange = $worksheet.UsedRange
        $rowCount = [int]$usedRange.Rows.Count
        $columnCount = [int]$usedRange.Columns.Count
        $headerRow = [int]$Config.ledger.header_row

        if ($rowCount -lt $headerRow) {
            throw "Ledger sheet does not contain header row $headerRow."
        }

        $headers = @{}
        for ($columnIndex = 1; $columnIndex -le $columnCount; $columnIndex++) {
            $headerValue = ConvertTo-ShinsaString -Value $worksheet.Cells.Item($headerRow, $columnIndex).Text
            if (-not [string]::IsNullOrWhiteSpace($headerValue)) {
                $headers[$columnIndex] = $headerValue.Trim()
            }
        }

        $records = @()
        for ($rowIndex = $headerRow + 1; $rowIndex -le $rowCount; $rowIndex++) {
            $row = [ordered]@{}
            $hasValue = $false

            foreach ($columnIndex in $headers.Keys) {
                $headerName = $headers[$columnIndex]
                $cellValue = ConvertFrom-ExcelCellValue -Value $worksheet.Cells.Item($rowIndex, $columnIndex).Value2 -HeaderName $headerName
                if (-not [string]::IsNullOrWhiteSpace($cellValue)) {
                    $hasValue = $true
                }
                $row[$headerName] = $cellValue
            }

            if (-not $hasValue) {
                continue
            }

            $record = Convert-SourceRowToLedgerRecord -Row ([pscustomobject]$row) -Config $Config -LedgerPath $Paths.SharePointLedgerPath -SheetName $worksheet.Name -RowId $rowIndex
            if ($null -ne $record) {
                $records += $record
            }
        }

        $records
    }
    finally {
        if ($null -ne $workbook) {
            $workbook.Close($false) | Out-Null
        }
        if ($null -ne $excel) {
            $excel.Quit()
        }

        foreach ($comObject in @($usedRange, $worksheet, $workbook, $excel)) {
            if ($null -ne $comObject) {
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($comObject)
            }
        }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Import-ShinsaLedgerRecords {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)]$Paths
    )

    if (-not (Test-Path $Paths.SharePointLedgerPath)) {
        throw "Ledger source was not found: $($Paths.SharePointLedgerPath)"
    }

    $extension = [System.IO.Path]::GetExtension($Paths.SharePointLedgerPath).ToLowerInvariant()
    switch ($extension) {
        '.json' { return @(Import-ShinsaLedgerFromJson -Config $Config -Paths $Paths) }
        '.xlsx' { return @(Import-ShinsaLedgerFromExcel -Config $Config -Paths $Paths) }
        '.xlsm' { return @(Import-ShinsaLedgerFromExcel -Config $Config -Paths $Paths) }
        '.xlsb' { return @(Import-ShinsaLedgerFromExcel -Config $Config -Paths $Paths) }
        '.xls' { return @(Import-ShinsaLedgerFromExcel -Config $Config -Paths $Paths) }
        default { throw "Unsupported ledger source format: $extension" }
    }
}

function Resolve-ShinsaArchiveItemPath {
    param(
        [Parameter(Mandatory = $true)][string]$MailDirectory,
        [string]$RelativePath
    )

    if ([string]::IsNullOrWhiteSpace($RelativePath)) {
        return ''
    }

    if ([System.IO.Path]::IsPathRooted($RelativePath)) {
        return [System.IO.Path]::GetFullPath($RelativePath)
    }

    $candidate = Join-Path $MailDirectory $RelativePath
    if (Test-Path $candidate) {
        return [System.IO.Path]::GetFullPath($candidate)
    }

    $attachmentsCandidate = Join-Path (Join-Path $MailDirectory 'attachments') $RelativePath
    if (Test-Path $attachmentsCandidate) {
        return [System.IO.Path]::GetFullPath($attachmentsCandidate)
    }

    [System.IO.Path]::GetFullPath($candidate)
}

function Import-ShinsaMailRecords {
    param([Parameter(Mandatory = $true)]$Paths)

    if (-not (Test-Path $Paths.MailArchiveRoot)) {
        return @()
    }

    $manifestFiles = @(Get-ChildItem -Path $Paths.MailArchiveRoot -Recurse -File -Include '*.json' | Where-Object {
            $_.Name -eq 'meta.json' -or $_.Name -like 'mail_*.json'
        })

    $records = @()
    foreach ($manifestFile in $manifestFiles) {
        $manifest = Read-ShinsaJson -Path $manifestFile.FullName
        if (-not (Test-ShinsaRecordProperty -Record $manifest -Name 'mail_id')) {
            continue
        }

        $mailDirectory = Split-Path -Parent $manifestFile.FullName
        $attachmentSource = @()
        if (Test-ShinsaRecordProperty -Record $manifest -Name 'attachment_paths') {
            $attachmentSource = @($manifest.attachment_paths)
        }
        elseif (Test-ShinsaRecordProperty -Record $manifest -Name 'attachments') {
            $attachmentSource = @($manifest.attachments)
        }

        $attachmentPaths = @()
        foreach ($attachment in $attachmentSource) {
            if ($attachment -is [string]) {
                $attachmentPaths += Resolve-ShinsaArchiveItemPath -MailDirectory $mailDirectory -RelativePath $attachment
                continue
            }

            if (Test-ShinsaRecordProperty -Record $attachment -Name 'path') {
                $attachmentPaths += Resolve-ShinsaArchiveItemPath -MailDirectory $mailDirectory -RelativePath ([string]$attachment.path)
            }
        }

        if (Test-ShinsaRecordProperty -Record $manifest -Name 'folder_path') {
            $folderPath = ConvertTo-ShinsaString -Value $manifest.folder_path
        }
        else {
            $folderPath = $mailDirectory.Substring($Paths.MailArchiveRoot.Length).TrimStart('\')
        }

        $records += [pscustomobject]@{
            mail_id = ConvertTo-ShinsaString -Value $manifest.mail_id
            entry_id = if (Test-ShinsaRecordProperty -Record $manifest -Name 'entry_id') { ConvertTo-ShinsaString -Value $manifest.entry_id } else { ConvertTo-ShinsaString -Value $manifest.mail_id }
            mailbox_address = if (Test-ShinsaRecordProperty -Record $manifest -Name 'mailbox_address') { ConvertTo-ShinsaString -Value $manifest.mailbox_address } else { '' }
            folder_path = $folderPath
            received_at = if (Test-ShinsaRecordProperty -Record $manifest -Name 'received_at') { ConvertTo-ShinsaString -Value $manifest.received_at } else { '' }
            sender_name = if (Test-ShinsaRecordProperty -Record $manifest -Name 'sender_name') { ConvertTo-ShinsaString -Value $manifest.sender_name } else { '' }
            sender_email = if (Test-ShinsaRecordProperty -Record $manifest -Name 'sender_email') { ConvertTo-ShinsaString -Value $manifest.sender_email } else { '' }
            subject = if (Test-ShinsaRecordProperty -Record $manifest -Name 'subject') { ConvertTo-ShinsaString -Value $manifest.subject } else { '' }
            body_path = if (Test-ShinsaRecordProperty -Record $manifest -Name 'body_path') { Resolve-ShinsaArchiveItemPath -MailDirectory $mailDirectory -RelativePath ([string]$manifest.body_path) } else { '' }
            msg_path = if (Test-ShinsaRecordProperty -Record $manifest -Name 'msg_path') { Resolve-ShinsaArchiveItemPath -MailDirectory $mailDirectory -RelativePath ([string]$manifest.msg_path) } else { '' }
            attachment_paths = @($attachmentPaths)
        }
    }

    @($records | Sort-Object received_at, mail_id -Descending)
}

function Import-ShinsaFolderRecords {
    param([Parameter(Mandatory = $true)]$Paths)

    if (-not (Test-Path $Paths.SharePointCaseRoot)) {
        return @()
    }

    $root = [System.IO.Path]::GetFullPath($Paths.SharePointCaseRoot)
    $records = @()
    foreach ($file in Get-ChildItem -Path $root -Recurse -File) {
        $fullPath = [System.IO.Path]::GetFullPath($file.FullName)
        $relativePath = $fullPath.Substring($root.Length).TrimStart('\')
        if ([string]::IsNullOrWhiteSpace($relativePath)) {
            continue
        }

        $segments = $relativePath -split '\\'
        $caseId = if ($segments.Count -gt 0) { $segments[0] } else { '' }
        if ([string]::IsNullOrWhiteSpace($caseId)) {
            continue
        }

        $records += [pscustomobject]@{
            case_id = $caseId
            folder_path = Join-Path $root $caseId
            file_path = $fullPath
            relative_path = if ($segments.Count -gt 1) { ($segments[1..($segments.Count - 1)] -join '\') } else { $file.Name }
            file_name = $file.Name
            extension = $file.Extension
            modified_at = $file.LastWriteTime.ToString('o')
            size = [int64]$file.Length
        }
    }

    @($records | Sort-Object case_id, relative_path)
}

function Get-ShinsaMailLinkMap {
    param([Parameter(Mandatory = $true)]$Cache)

    $map = @{}
    foreach ($link in @($Cache.mail_links)) {
        if (-not (Test-ShinsaRecordProperty -Record $link -Name 'mail_id')) {
            continue
        }

        $mailId = ConvertTo-ShinsaString -Value $link.mail_id
        if ([string]::IsNullOrWhiteSpace($mailId)) {
            continue
        }

        $map[$mailId] = $link
    }

    $map
}

function Get-ShinsaRelatedMails {
    param(
        [Parameter(Mandatory = $true)]$CaseRecord,
        [Parameter(Mandatory = $true)]$Mails,
        [Parameter(Mandatory = $true)]$Cache
    )

    $caseId = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $CaseRecord -Name 'case_id')
    $contactEmail = (ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $CaseRecord -Name 'contact_email')).ToLowerInvariant()
    $linkMap = Get-ShinsaMailLinkMap -Cache $Cache
    $records = @()

    foreach ($mail in @($Mails)) {
        $mailId = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $mail -Name 'mail_id')
        $matchType = ''

        if ($linkMap.ContainsKey($mailId)) {
            if ((ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $linkMap[$mailId] -Name 'case_id')) -ne $caseId) {
                continue
            }
            $matchType = 'manual'
        }
        elseif (-not [string]::IsNullOrWhiteSpace($contactEmail)) {
            $senderEmail = (ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $mail -Name 'sender_email')).ToLowerInvariant()
            if ($senderEmail -ne $contactEmail) {
                continue
            }
            $matchType = 'email'
        }
        else {
            continue
        }

        $record = Copy-ShinsaRecord -Record $mail
        Set-ShinsaRecordValue -Record $record -Name 'match_type' -Value $matchType
        $records += $record
    }

    @($records | Sort-Object received_at, mail_id -Descending)
}

function Get-ShinsaLedgerWritebackPlan {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)]$Paths
    )

    if (-not (Test-Path $Paths.LedgerJsonPath)) {
        throw "Local ledger.json was not found. Run sync first."
    }

    $currentLedger = @(Read-ShinsaJson -Path $Paths.LedgerJsonPath)
    $sourceLedger = @(Import-ShinsaLedgerRecords -Config $Config -Paths $Paths)
    $editableColumns = @($Config.ledger.editable_columns)

    $sourceByCaseId = @{}
    foreach ($record in $sourceLedger) {
        $caseId = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $record -Name 'case_id')
        if (-not [string]::IsNullOrWhiteSpace($caseId)) {
            $sourceByCaseId[$caseId] = $record
        }
    }

    $changes = @()
    foreach ($record in $currentLedger) {
        $caseId = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $record -Name 'case_id')
        if ([string]::IsNullOrWhiteSpace($caseId) -or -not $sourceByCaseId.ContainsKey($caseId)) {
            continue
        }

        $sourceRecord = $sourceByCaseId[$caseId]
        $fieldChanges = [ordered]@{}
        foreach ($columnName in $editableColumns) {
            $targetValue = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $record -Name $columnName)
            $sourceValue = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $sourceRecord -Name $columnName)
            if ($targetValue -ne $sourceValue) {
                $fieldChanges[$columnName] = [ordered]@{
                    from = $sourceValue
                    to = $targetValue
                }
            }
        }

        if ($fieldChanges.Count -eq 0) {
            continue
        }

        $changes += [pscustomobject]@{
            case_id = $caseId
            ledger_row_id = [int](Get-ShinsaRecordValue -Record $record -Name 'ledger_row_id')
            changes = [pscustomobject]$fieldChanges
        }
    }

    $changeCount = (@($changes | ForEach-Object { @($_.changes.PSObject.Properties).Count } | Measure-Object -Sum).Sum)
    if ($null -eq $changeCount) {
        $changeCount = 0
    }

    [pscustomobject]@{
        source_kind = [System.IO.Path]::GetExtension($Paths.SharePointLedgerPath).ToLowerInvariant()
        changes = $changes
        change_count = [int]$changeCount
        case_count = @($changes).Count
    }
}

function Invoke-ShinsaJsonLedgerWriteback {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)]$Paths,
        [Parameter(Mandatory = $true)]$Plan
    )

    $source = Read-ShinsaJson -Path $Paths.SharePointLedgerPath
    if ($source -is [System.Collections.IEnumerable] -and -not ($source -is [string]) -and -not ($source -is [System.Management.Automation.PSCustomObject])) {
        $rows = @($source)
    }
    elseif (Test-ShinsaRecordProperty -Record $source -Name 'organizations') {
        $rows = @($source.organizations)
    }
    elseif (Test-ShinsaRecordProperty -Record $source -Name 'rows') {
        $rows = @($source.rows)
    }
    else {
        $rows = @($source)
    }

    $sourceKeyName = if ($Config.ledger.columns.Contains($Config.ledger.key_column)) { [string]$Config.ledger.columns[$Config.ledger.key_column] } else { [string]$Config.ledger.key_column }
    $rowByCaseId = @{}
    foreach ($row in $rows) {
        $caseId = ''
        if (Test-ShinsaRecordProperty -Record $row -Name $sourceKeyName) {
            $caseId = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $row -Name $sourceKeyName)
        }
        elseif (Test-ShinsaRecordProperty -Record $row -Name 'case_id') {
            $caseId = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $row -Name 'case_id')
        }

        if (-not [string]::IsNullOrWhiteSpace($caseId)) {
            $rowByCaseId[$caseId] = $row
        }
    }

    foreach ($change in @($Plan.changes)) {
        if (-not $rowByCaseId.ContainsKey($change.case_id)) {
            continue
        }

        $row = $rowByCaseId[$change.case_id]
        foreach ($property in $change.changes.PSObject.Properties) {
            $logicalName = $property.Name
            $sourceName = if ($Config.ledger.columns.Contains($logicalName)) { [string]$Config.ledger.columns[$logicalName] } else { $logicalName }
            Set-ShinsaRecordValue -Record $row -Name $sourceName -Value $property.Value.to
            if ($sourceName -ne $logicalName -and (Test-ShinsaRecordProperty -Record $row -Name $logicalName)) {
                Set-ShinsaRecordValue -Record $row -Name $logicalName -Value $property.Value.to
            }
        }
    }

    Write-ShinsaJson -Path $Paths.SharePointLedgerPath -Data $source
}

function Invoke-ShinsaExcelLedgerWriteback {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)]$Paths,
        [Parameter(Mandatory = $true)]$Plan
    )

    $excel = $null
    $workbook = $null
    $worksheet = $null
    $usedRange = $null
    $saved = $false

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($Paths.SharePointLedgerPath, $false, $false)

        if ([string]::IsNullOrWhiteSpace([string]$Config.ledger.sheet_name)) {
            $worksheet = $workbook.Worksheets.Item(1)
        }
        else {
            $worksheet = $workbook.Worksheets.Item([string]$Config.ledger.sheet_name)
        }

        $headerRow = [int]$Config.ledger.header_row
        $usedRange = $worksheet.UsedRange
        $columnCount = [int]$usedRange.Columns.Count
        $headerMap = @{}
        for ($columnIndex = 1; $columnIndex -le $columnCount; $columnIndex++) {
            $headerName = ConvertTo-ShinsaString -Value $worksheet.Cells.Item($headerRow, $columnIndex).Text
            if (-not [string]::IsNullOrWhiteSpace($headerName)) {
                $headerMap[$headerName.Trim()] = $columnIndex
            }
        }

        foreach ($change in @($Plan.changes)) {
            $rowId = [int]$change.ledger_row_id
            foreach ($property in $change.changes.PSObject.Properties) {
                $logicalName = $property.Name
                $headerName = if ($Config.ledger.columns.Contains($logicalName)) { [string]$Config.ledger.columns[$logicalName] } else { $logicalName }
                if (-not $headerMap.ContainsKey($headerName)) {
                    throw "Ledger column '$headerName' was not found in worksheet '$($worksheet.Name)'."
                }

                $worksheet.Cells.Item($rowId, $headerMap[$headerName]).Value2 = $property.Value.to
            }
        }

        $workbook.Save()
        $saved = $true
    }
    finally {
        if ($null -ne $workbook) {
            $workbook.Close($saved) | Out-Null
        }
        if ($null -ne $excel) {
            $excel.Quit()
        }

        foreach ($comObject in @($usedRange, $worksheet, $workbook, $excel)) {
            if ($null -ne $comObject) {
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($comObject)
            }
        }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Invoke-ShinsaLedgerWriteback {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)]$Paths,
        [Parameter(Mandatory = $true)]$Plan
    )

    if ($Plan.case_count -eq 0) {
        return
    }

    switch ($Plan.source_kind) {
        '.json' { Invoke-ShinsaJsonLedgerWriteback -Config $Config -Paths $Paths -Plan $Plan }
        '.xlsx' { Invoke-ShinsaExcelLedgerWriteback -Config $Config -Paths $Paths -Plan $Plan }
        '.xlsm' { Invoke-ShinsaExcelLedgerWriteback -Config $Config -Paths $Paths -Plan $Plan }
        '.xlsb' { Invoke-ShinsaExcelLedgerWriteback -Config $Config -Paths $Paths -Plan $Plan }
        '.xls' { Invoke-ShinsaExcelLedgerWriteback -Config $Config -Paths $Paths -Plan $Plan }
        default { throw "Writeback is not supported for source type '$($Plan.source_kind)'." }
    }
}

function Start-ShinsaItem {
    param([Parameter(Mandatory = $true)][string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw 'Path is empty.'
    }

    if (-not (Test-Path $Path)) {
        throw "Path was not found: $Path"
    }

    Start-Process -FilePath $Path | Out-Null
}

Export-ModuleMember -Function *-Shinsa*
