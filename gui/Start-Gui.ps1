Import-Module (Join-Path $PSScriptRoot '..\scripts\Common.psm1') -Force -DisableNameChecking
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = 'Stop'
$script:AppRoot = Split-Path -Parent $PSScriptRoot
$script:Config = Get-ShinsaConfig -ScriptPath $MyInvocation.MyCommand.Path
$script:Paths = Get-ShinsaDataPaths -Config $script:Config

Ensure-ShinsaState -Paths $script:Paths

foreach ($requiredPath in @($script:Paths.LedgerJsonPath, $script:Paths.MailsJsonPath, $script:Paths.FoldersJsonPath)) {
    if (-not (Test-Path $requiredPath)) {
        [System.Windows.Forms.MessageBox]::Show('Local JSON was not found. Run `sync` from shinsa first.', 'shinsa') | Out-Null
        exit 1
    }
}

$script:ledgerRecords = @()
$script:mailRecords = @()
$script:folderRecords = @()
$script:cacheState = $null
$script:currentCase = $null
$script:displayedMailRecords = @()
$script:displayedFolderRecords = @()
$script:detailControls = @{}
$script:suppressSelection = $false

function New-DetailLabel {
    param([string]$Text)

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.AutoSize = $true
    $label.Margin = New-Object System.Windows.Forms.Padding(4, 8, 4, 4)
    $label
}

function Get-FieldLabelText {
    param([string]$FieldName)

    switch ($FieldName) {
        'case_id' { 'Case ID' }
        'receipt_no' { 'Receipt No' }
        'organization_name' { 'Organization' }
        'contact_name' { 'Contact' }
        'contact_email' { 'Email' }
        'status' { 'Status' }
        'assigned_to' { 'Assigned To' }
        'missing_documents' { 'Missing Documents' }
        'review_note_public' { 'Public Note' }
        default { ($FieldName -replace '_', ' ') }
    }
}

function Get-WindowStateName {
    param([System.Windows.Forms.FormWindowState]$State)

    switch ($State) {
        'Maximized' { 'Maximized' }
        'Minimized' { 'Minimized' }
        default { 'Normal' }
    }
}

function Get-UiState {
    ConvertTo-ShinsaMap -InputObject $script:cacheState.ui_state
}

function Save-UiState {
    $uiState = Get-UiState
    $bounds = if ($form.WindowState -eq 'Normal') { $form.Bounds } else { $form.RestoreBounds }

    $uiState['search_text'] = $searchBox.Text
    $uiState['mail_filter'] = $mailFilterBox.Text
    $uiState['mail_scope'] = [string]$mailScopeCombo.SelectedItem
    $uiState['last_case_id'] = if ($script:currentCase) { ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $script:currentCase -Name 'case_id') } else { '' }
    $uiState['window_left'] = $bounds.Left
    $uiState['window_top'] = $bounds.Top
    $uiState['window_width'] = $bounds.Width
    $uiState['window_height'] = $bounds.Height
    $uiState['window_state'] = Get-WindowStateName -State $form.WindowState
    $uiState['main_splitter_distance'] = $mainSplit.SplitterDistance
    $uiState['detail_splitter_distance'] = $detailSplit.SplitterDistance
    $script:cacheState.ui_state = [pscustomobject]$uiState
    Save-ShinsaCache -Paths $script:Paths -Cache $script:cacheState
}

function Set-StatusText {
    param([string]$Text)

    $statusLabel.Text = $Text
}

function Format-FileSize {
    param([long]$Bytes)

    if ($Bytes -ge 1GB) { return ('{0:N1} GB' -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ('{0:N1} MB' -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ('{0:N1} KB' -f ($Bytes / 1KB)) }
    '{0} B' -f $Bytes
}

function Load-ShinsaData {
    $script:ledgerRecords = @((Read-ShinsaJson -Path $script:Paths.LedgerJsonPath))
    $script:mailRecords = @((Read-ShinsaJson -Path $script:Paths.MailsJsonPath))
    $script:folderRecords = @((Read-ShinsaJson -Path $script:Paths.FoldersJsonPath))
    $script:cacheState = Read-ShinsaCache -Paths $script:Paths
}

function Save-CurrentCase {
    param([switch]$Quiet)

    if ($null -eq $script:currentCase) {
        return
    }

    foreach ($fieldName in @($script:Config.ledger.editable_columns)) {
        if (-not $script:detailControls.ContainsKey($fieldName)) {
            continue
        }

        Set-ShinsaRecordValue -Record $script:currentCase -Name $fieldName -Value $script:detailControls[$fieldName].Text.Trim()
    }

    Write-ShinsaJson -Path $script:Paths.LedgerJsonPath -Data $script:ledgerRecords
    if (-not $Quiet) {
        Set-StatusText -Text ("Saved local edits for {0}" -f (Get-ShinsaRecordValue -Record $script:currentCase -Name 'case_id'))
    }
}

function Get-SelectedCaseIdFromGrid {
    if ($caseGrid.SelectedRows.Count -eq 0) {
        return ''
    }

    [string]$caseGrid.SelectedRows[0].Cells['case_id'].Value
}

function Get-CaseRecordById {
    param([string]$CaseId)

    $script:ledgerRecords | Where-Object { (Get-ShinsaRecordValue -Record $_ -Name 'case_id') -eq $CaseId } | Select-Object -First 1
}

function Apply-WindowState {
    $uiState = Get-UiState

    if ($uiState.Contains('window_width') -and $uiState.Contains('window_height')) {
        $form.StartPosition = 'Manual'
        $form.Size = New-Object System.Drawing.Size([int]$uiState.window_width, [int]$uiState.window_height)
        if ($uiState.Contains('window_left') -and $uiState.Contains('window_top')) {
            $form.Location = New-Object System.Drawing.Point([int]$uiState.window_left, [int]$uiState.window_top)
        }
    }
    else {
        $form.StartPosition = 'CenterScreen'
        $form.Size = New-Object System.Drawing.Size([int]$script:Config.gui.window_width, [int]$script:Config.gui.window_height)
    }
}

function Set-SafeSplitterLayout {
    param(
        [Parameter(Mandatory = $true)][System.Windows.Forms.SplitContainer]$SplitControl,
        [int]$Panel1MinSize,
        [int]$Panel2MinSize,
        [int]$PreferredDistance
    )

    $SplitControl.Panel1MinSize = [Math]::Max(0, $Panel1MinSize)
    $availableSize =
        if ($SplitControl.Orientation -eq [System.Windows.Forms.Orientation]::Vertical) {
            $SplitControl.ClientSize.Width
        }
        else {
            $SplitControl.ClientSize.Height
        }

    if ($availableSize -le 0) {
        return
    }

    $safePanel2Min = [Math]::Max(0, [Math]::Min($Panel2MinSize, [Math]::Max(0, $availableSize - $SplitControl.Panel1MinSize - 4)))
    $SplitControl.Panel2MinSize = $safePanel2Min

    $maxDistance = [Math]::Max($SplitControl.Panel1MinSize, $availableSize - $SplitControl.Panel2MinSize - 4)
    $safeDistance = [Math]::Max($SplitControl.Panel1MinSize, [Math]::Min($PreferredDistance, $maxDistance))
    $SplitControl.SplitterDistance = $safeDistance
}

function Build-CaseTable {
    param([string]$FilterText)

    $table = New-Object System.Data.DataTable
    foreach ($columnName in @($script:Config.ledger.display_columns)) {
        [void]$table.Columns.Add($columnName)
    }

    foreach ($record in @($script:ledgerRecords)) {
        $searchText = (@($script:Config.ledger.display_columns | ForEach-Object { ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $record -Name $_) }) -join ' ')
        if (-not [string]::IsNullOrWhiteSpace($FilterText) -and $searchText -notmatch [regex]::Escape($FilterText)) {
            continue
        }

        $row = $table.NewRow()
        foreach ($columnName in @($script:Config.ledger.display_columns)) {
            $row[$columnName] = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $record -Name $columnName)
        }
        [void]$table.Rows.Add($row)
    }

    return ,$table
}

function Refresh-CaseGrid {
    param([string]$SelectedCaseId = '')

    $script:suppressSelection = $true
    $caseGrid.DataSource = Build-CaseTable -FilterText $searchBox.Text

    if ([string]::IsNullOrWhiteSpace($SelectedCaseId)) {
        $uiState = Get-UiState
        if ($uiState.Contains('last_case_id')) {
            $SelectedCaseId = [string]$uiState['last_case_id']
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($SelectedCaseId)) {
        foreach ($row in $caseGrid.Rows) {
            if ([string]$row.Cells['case_id'].Value -eq $SelectedCaseId) {
                $row.Selected = $true
                $caseGrid.CurrentCell = $row.Cells['case_id']
                break
            }
        }
    }

    if ($caseGrid.SelectedRows.Count -eq 0 -and $caseGrid.Rows.Count -gt 0) {
        $caseGrid.Rows[0].Selected = $true
        $caseGrid.CurrentCell = $caseGrid.Rows[0].Cells['case_id']
    }
    $script:suppressSelection = $false

    if ($caseGrid.SelectedRows.Count -gt 0) {
        Select-Case -CaseId ([string]$caseGrid.SelectedRows[0].Cells['case_id'].Value) -SkipGridSelection
    }
    else {
        $script:currentCase = $null
        foreach ($fieldName in $script:detailControls.Keys) {
            $script:detailControls[$fieldName].Text = ''
        }
        $mailList.Items.Clear()
        $attachmentList.Items.Clear()
        $mailPreviewBox.Text = ''
        $fileList.Items.Clear()
    }
}

function Get-MailLinkStateLabel {
    param($MailRecord)

    $mailId = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $MailRecord -Name 'mail_id')
    $linkMap = Get-ShinsaMailLinkMap -Cache $script:cacheState
    if ($linkMap.ContainsKey($mailId)) {
        return ('manual:{0}' -f (Get-ShinsaRecordValue -Record $linkMap[$mailId] -Name 'case_id'))
    }

    if ($null -ne $script:currentCase) {
        $contactEmail = (ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $script:currentCase -Name 'contact_email')).ToLowerInvariant()
        $senderEmail = (ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $MailRecord -Name 'sender_email')).ToLowerInvariant()
        if (-not [string]::IsNullOrWhiteSpace($contactEmail) -and $contactEmail -eq $senderEmail) {
            return 'email'
        }
    }

    ''
}

function Load-MailPreview {
    param($MailRecord)

    $attachmentList.Items.Clear()
    $mailPreviewBox.Text = ''

    if ($null -eq $MailRecord) {
        return
    }

    $bodyPath = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $MailRecord -Name 'body_path')
    if (-not [string]::IsNullOrWhiteSpace($bodyPath) -and (Test-Path $bodyPath)) {
        $mailPreviewBox.Text = Get-Content -Path $bodyPath -Raw -Encoding UTF8
    }

    foreach ($attachmentPath in @((Get-ShinsaRecordValue -Record $MailRecord -Name 'attachment_paths'))) {
        $attachmentList.Items.Add([pscustomobject]@{
                name = [System.IO.Path]::GetFileName([string]$attachmentPath)
                path = [string]$attachmentPath
            }) | Out-Null
    }
}

function Get-SelectedMailRecord {
    if ($mailList.SelectedItems.Count -eq 0) {
        return $null
    }

    $mailId = [string]$mailList.SelectedItems[0].Tag
    $script:displayedMailRecords | Where-Object { (Get-ShinsaRecordValue -Record $_ -Name 'mail_id') -eq $mailId } | Select-Object -First 1
}

function Refresh-MailList {
    $mailList.Items.Clear()
    $script:displayedMailRecords = @()

    if ($null -eq $script:currentCase) {
        return
    }

    $scope = [string]$mailScopeCombo.SelectedItem
    if ($scope -eq 'All mails') {
        $records = @($script:mailRecords)
    }
    else {
        $records = @(Get-ShinsaRelatedMails -CaseRecord $script:currentCase -Mails $script:mailRecords -Cache $script:cacheState)
    }

    $filterText = $mailFilterBox.Text
    if (-not [string]::IsNullOrWhiteSpace($filterText)) {
        $records = @($records | Where-Object {
                $candidate = @(
                    Get-ShinsaRecordValue -Record $_ -Name 'mail_id'
                    Get-ShinsaRecordValue -Record $_ -Name 'sender_email'
                    Get-ShinsaRecordValue -Record $_ -Name 'sender_name'
                    Get-ShinsaRecordValue -Record $_ -Name 'subject'
                ) -join ' '
                $candidate -match [regex]::Escape($filterText)
            })
    }

    foreach ($record in @($records | Sort-Object received_at, mail_id -Descending)) {
        $mailId = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $record -Name 'mail_id')
        $item = New-Object System.Windows.Forms.ListViewItem((ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $record -Name 'received_at')))
        [void]$item.SubItems.Add((ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $record -Name 'sender_email')))
        [void]$item.SubItems.Add((ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $record -Name 'subject')))
        [void]$item.SubItems.Add((Get-MailLinkStateLabel -MailRecord $record))
        $item.Tag = $mailId
        $mailList.Items.Add($item) | Out-Null
        $script:displayedMailRecords += $record
    }

    if ($mailList.Items.Count -gt 0) {
        $mailList.Items[0].Selected = $true
    }
}

function Refresh-FileList {
    $fileList.Items.Clear()
    $script:displayedFolderRecords = @()

    if ($null -eq $script:currentCase) {
        return
    }

    $caseId = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $script:currentCase -Name 'case_id')
    $script:displayedFolderRecords = @($script:folderRecords | Where-Object { (Get-ShinsaRecordValue -Record $_ -Name 'case_id') -eq $caseId })

    foreach ($record in @($script:displayedFolderRecords)) {
        $item = New-Object System.Windows.Forms.ListViewItem((ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $record -Name 'relative_path')))
        [void]$item.SubItems.Add((ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $record -Name 'modified_at')))
        [void]$item.SubItems.Add((Format-FileSize -Bytes ([int64](Get-ShinsaRecordValue -Record $record -Name 'size'))))
        $item.Tag = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $record -Name 'file_path')
        $fileList.Items.Add($item) | Out-Null
    }
}

function Select-Case {
    param(
        [Parameter(Mandatory = $true)][string]$CaseId,
        [switch]$SkipGridSelection
    )

    $record = Get-CaseRecordById -CaseId $CaseId
    if ($null -eq $record) {
        return
    }

    $script:currentCase = $record
    foreach ($fieldName in $script:detailControls.Keys) {
        $script:detailControls[$fieldName].Text = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $record -Name $fieldName)
    }

    Refresh-MailList
    Refresh-FileList
    Set-StatusText -Text ("Loaded {0}" -f $CaseId)

    if (-not $SkipGridSelection) {
        foreach ($row in $caseGrid.Rows) {
            if ([string]$row.Cells['case_id'].Value -eq $CaseId) {
                $row.Selected = $true
                $caseGrid.CurrentCell = $row.Cells['case_id']
                break
            }
        }
    }
}

function Set-MailLink {
    param(
        [Parameter(Mandatory = $true)][string]$MailId,
        [Parameter(Mandatory = $true)][string]$CaseId
    )

    $updated = @()
    foreach ($link in @($script:cacheState.mail_links)) {
        if ((Get-ShinsaRecordValue -Record $link -Name 'mail_id') -eq $MailId) {
            continue
        }
        $updated += $link
    }

    $updated += [pscustomobject]@{
        mail_id = $MailId
        case_id = $CaseId
        mode = 'manual'
        updated_at = (Get-Date).ToString('o')
    }
    $script:cacheState.mail_links = $updated
    Save-ShinsaCache -Paths $script:Paths -Cache $script:cacheState
}

function Clear-MailLink {
    param([Parameter(Mandatory = $true)][string]$MailId)

    $script:cacheState.mail_links = @($script:cacheState.mail_links | Where-Object { (Get-ShinsaRecordValue -Record $_ -Name 'mail_id') -ne $MailId })
    Save-ShinsaCache -Paths $script:Paths -Cache $script:cacheState
}

function Get-SelectedAttachmentPath {
    if ($attachmentList.SelectedItem -eq $null) {
        return ''
    }

    [string]$attachmentList.SelectedItem.path
}

function Get-SelectedFilePath {
    if ($fileList.SelectedItems.Count -eq 0) {
        return ''
    }

    [string]$fileList.SelectedItems[0].Tag
}

function Invoke-GuiSync {
    Save-CurrentCase -Quiet
    & (Join-Path $script:AppRoot 'scripts\Sync-Data.ps1')
    Load-ShinsaData
    Refresh-CaseGrid -SelectedCaseId (Get-SelectedCaseIdFromGrid)
    Set-StatusText -Text 'Sync completed.'
}

function Invoke-GuiWriteback {
    Save-CurrentCase -Quiet
    $plan = Get-ShinsaLedgerWritebackPlan -Config $script:Config -Paths $script:Paths
    if ($plan.case_count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('No ledger changes to write back.', 'shinsa') | Out-Null
        return
    }

    $summary = @(
        'Write back the following local ledger edits?'
        ''
    )
    foreach ($change in @($plan.changes)) {
        $summary += ('{0}: {1}' -f $change.case_id, (@($change.changes.PSObject.Properties.Name) -join ', '))
    }
    $summary += ''
    $summary += ('Total: {0} cases / {1} fields' -f $plan.case_count, $plan.change_count)

    $result = [System.Windows.Forms.MessageBox]::Show(($summary -join [Environment]::NewLine), 'shinsa', [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
        return
    }

    Invoke-ShinsaLedgerWriteback -Config $script:Config -Paths $script:Paths -Plan $plan
    Write-ShinsaJson -Path $script:Paths.LedgerJsonPath -Data @(Import-ShinsaLedgerRecords -Config $script:Config -Paths $script:Paths | Sort-Object case_id)
    Load-ShinsaData
    Refresh-CaseGrid -SelectedCaseId (Get-SelectedCaseIdFromGrid)
    Set-StatusText -Text 'Writeback completed.'
}

$form = New-Object System.Windows.Forms.Form
$form.Text = $script:Config.gui.title
$form.MinimumSize = New-Object System.Drawing.Size(1100, 700)
$form.Font = New-Object System.Drawing.Font($script:Config.gui.font_name, [single]$script:Config.gui.font_size)
$form.AutoScaleMode = 'Dpi'

$root = New-Object System.Windows.Forms.TableLayoutPanel
$root.Dock = 'Fill'
$root.RowCount = 3
$root.ColumnCount = 1
[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))

$toolbar = New-Object System.Windows.Forms.FlowLayoutPanel
$toolbar.Dock = 'Fill'
$toolbar.WrapContents = $false
$toolbar.AutoSize = $true
$toolbar.Padding = New-Object System.Windows.Forms.Padding(6)

$searchLabel = New-DetailLabel -Text 'Search'
$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Width = 240

$reloadButton = New-Object System.Windows.Forms.Button
$reloadButton.Text = 'Reload'
$syncButton = New-Object System.Windows.Forms.Button
$syncButton.Text = 'Sync'
$saveButton = New-Object System.Windows.Forms.Button
$saveButton.Text = 'Save Local'
$writebackButton = New-Object System.Windows.Forms.Button
$writebackButton.Text = 'Writeback'
$openLedgerButton = New-Object System.Windows.Forms.Button
$openLedgerButton.Text = 'Open Ledger'
$openCaseFolderButton = New-Object System.Windows.Forms.Button
$openCaseFolderButton.Text = 'Open Case Folder'

$toolbar.Controls.AddRange(@(
    $searchLabel,
    $searchBox,
    $reloadButton,
    $syncButton,
    $saveButton,
    $writebackButton,
    $openLedgerButton,
    $openCaseFolderButton
))

$mainSplit = New-Object System.Windows.Forms.SplitContainer
$mainSplit.Dock = 'Fill'
$mainSplit.Orientation = 'Vertical'

$caseGrid = New-Object System.Windows.Forms.DataGridView
$caseGrid.Dock = 'Fill'
$caseGrid.ReadOnly = $true
$caseGrid.AllowUserToAddRows = $false
$caseGrid.AllowUserToDeleteRows = $false
$caseGrid.MultiSelect = $false
$caseGrid.SelectionMode = 'FullRowSelect'
$caseGrid.RowHeadersVisible = $false
$caseGrid.AutoSizeColumnsMode = 'Fill'
$caseGrid.BackgroundColor = [System.Drawing.SystemColors]::Window
$mainSplit.Panel1.Controls.Add($caseGrid)

$detailSplit = New-Object System.Windows.Forms.SplitContainer
$detailSplit.Dock = 'Fill'
$detailSplit.Orientation = 'Horizontal'

$detailPanel = New-Object System.Windows.Forms.TableLayoutPanel
$detailPanel.Dock = 'Fill'
$detailPanel.AutoScroll = $true
$detailPanel.Padding = New-Object System.Windows.Forms.Padding(10)
$detailPanel.ColumnCount = 2
[void]$detailPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 130)))
[void]$detailPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))

foreach ($fieldName in @($script:Config.ledger.detail_columns)) {
    $isMultiline = $fieldName -in @('missing_documents', 'review_note_public')
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Dock = 'Fill'
    $textBox.Multiline = $isMultiline
    $textBox.ScrollBars = if ($isMultiline) { 'Vertical' } else { 'None' }
    $textBox.ReadOnly = $fieldName -notin @($script:Config.ledger.editable_columns)
    $textBox.Height = if ($fieldName -eq 'review_note_public') { 96 } elseif ($isMultiline) { 56 } else { 24 }

    [void]$detailPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $rowIndex = $detailPanel.RowCount
    $detailPanel.Controls.Add((New-DetailLabel -Text (Get-FieldLabelText -FieldName $fieldName)), 0, $rowIndex)
    $detailPanel.Controls.Add($textBox, 1, $rowIndex)
    $detailPanel.RowCount += 1
    $script:detailControls[$fieldName] = $textBox
}

$detailSplit.Panel1.Controls.Add($detailPanel)

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = 'Fill'

$mailsTab = New-Object System.Windows.Forms.TabPage
$mailsTab.Text = 'Mails'
$mailRoot = New-Object System.Windows.Forms.TableLayoutPanel
$mailRoot.Dock = 'Fill'
$mailRoot.RowCount = 2
$mailRoot.ColumnCount = 1
[void]$mailRoot.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$mailRoot.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))

$mailToolbar = New-Object System.Windows.Forms.FlowLayoutPanel
$mailToolbar.Dock = 'Fill'
$mailToolbar.WrapContents = $false
$mailToolbar.AutoSize = $true
$mailToolbar.Padding = New-Object System.Windows.Forms.Padding(4)

$mailScopeLabel = New-DetailLabel -Text 'Scope'
$mailScopeCombo = New-Object System.Windows.Forms.ComboBox
$mailScopeCombo.Width = 100
$mailScopeCombo.DropDownStyle = 'DropDownList'
[void]$mailScopeCombo.Items.Add('Related mails')
[void]$mailScopeCombo.Items.Add('All mails')
$mailScopeCombo.SelectedIndex = 0

$mailFilterLabel = New-DetailLabel -Text 'Mail Filter'
$mailFilterBox = New-Object System.Windows.Forms.TextBox
$mailFilterBox.Width = 180

$linkMailButton = New-Object System.Windows.Forms.Button
$linkMailButton.Text = 'Link to Case'
$clearMailLinkButton = New-Object System.Windows.Forms.Button
$clearMailLinkButton.Text = 'Clear Link'
$openMsgButton = New-Object System.Windows.Forms.Button
$openMsgButton.Text = 'Open MSG'
$openBodyButton = New-Object System.Windows.Forms.Button
$openBodyButton.Text = 'Open Body'
$openAttachmentButton = New-Object System.Windows.Forms.Button
$openAttachmentButton.Text = 'Open Attachment'

$mailToolbar.Controls.AddRange(@(
    $mailScopeLabel,
    $mailScopeCombo,
    $mailFilterLabel,
    $mailFilterBox,
    $linkMailButton,
    $clearMailLinkButton,
    $openMsgButton,
    $openBodyButton,
    $openAttachmentButton
))

$mailSplit = New-Object System.Windows.Forms.SplitContainer
$mailSplit.Dock = 'Fill'
$mailSplit.Orientation = 'Horizontal'

$mailList = New-Object System.Windows.Forms.ListView
$mailList.Dock = 'Fill'
$mailList.View = 'Details'
$mailList.FullRowSelect = $true
$mailList.HideSelection = $false
[void]$mailList.Columns.Add('Received', 160)
[void]$mailList.Columns.Add('Sender', 200)
[void]$mailList.Columns.Add('Subject', 340)
[void]$mailList.Columns.Add('Link', 120)

$mailBottomSplit = New-Object System.Windows.Forms.SplitContainer
$mailBottomSplit.Dock = 'Fill'
$mailBottomSplit.Orientation = 'Vertical'

$mailPreviewBox = New-Object System.Windows.Forms.TextBox
$mailPreviewBox.Dock = 'Fill'
$mailPreviewBox.Multiline = $true
$mailPreviewBox.ScrollBars = 'Vertical'
$mailPreviewBox.ReadOnly = $true

$attachmentList = New-Object System.Windows.Forms.ListBox
$attachmentList.Dock = 'Fill'
$attachmentList.DisplayMember = 'name'

$mailBottomSplit.Panel1.Controls.Add($mailPreviewBox)
$mailBottomSplit.Panel2.Controls.Add($attachmentList)
$mailSplit.Panel1.Controls.Add($mailList)
$mailSplit.Panel2.Controls.Add($mailBottomSplit)
$mailRoot.Controls.Add($mailToolbar, 0, 0)
$mailRoot.Controls.Add($mailSplit, 0, 1)
$mailsTab.Controls.Add($mailRoot)

$filesTab = New-Object System.Windows.Forms.TabPage
$filesTab.Text = 'Case Files'
$fileRoot = New-Object System.Windows.Forms.TableLayoutPanel
$fileRoot.Dock = 'Fill'
$fileRoot.RowCount = 2
$fileRoot.ColumnCount = 1
[void]$fileRoot.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$fileRoot.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))

$fileToolbar = New-Object System.Windows.Forms.FlowLayoutPanel
$fileToolbar.Dock = 'Fill'
$fileToolbar.WrapContents = $false
$fileToolbar.AutoSize = $true
$fileToolbar.Padding = New-Object System.Windows.Forms.Padding(4)

$openFileButton = New-Object System.Windows.Forms.Button
$openFileButton.Text = 'Open File'
$openFolderButton = New-Object System.Windows.Forms.Button
$openFolderButton.Text = 'Open Folder'
$fileToolbar.Controls.AddRange(@($openFileButton, $openFolderButton))

$fileList = New-Object System.Windows.Forms.ListView
$fileList.Dock = 'Fill'
$fileList.View = 'Details'
$fileList.FullRowSelect = $true
$fileList.HideSelection = $false
[void]$fileList.Columns.Add('Relative Path', 420)
[void]$fileList.Columns.Add('Modified', 180)
[void]$fileList.Columns.Add('Size', 100)

$fileRoot.Controls.Add($fileToolbar, 0, 0)
$fileRoot.Controls.Add($fileList, 0, 1)
$filesTab.Controls.Add($fileRoot)

[void]$tabs.TabPages.Add($mailsTab)
[void]$tabs.TabPages.Add($filesTab)
$detailSplit.Panel2.Controls.Add($tabs)
$mainSplit.Panel2.Controls.Add($detailSplit)

$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Spring = $true
$statusLabel.TextAlign = 'MiddleLeft'
[void]$statusStrip.Items.Add($statusLabel)

$root.Controls.Add($toolbar, 0, 0)
$root.Controls.Add($mainSplit, 0, 1)
$root.Controls.Add($statusStrip, 0, 2)
$form.Controls.Add($root)

Load-ShinsaData
Apply-WindowState

$uiState = Get-UiState
if ($uiState.Contains('search_text')) { $searchBox.Text = [string]$uiState['search_text'] }
if ($uiState.Contains('mail_filter')) { $mailFilterBox.Text = [string]$uiState['mail_filter'] }
if ($uiState.Contains('mail_scope') -and $mailScopeCombo.Items.Contains([string]$uiState['mail_scope'])) {
    $mailScopeCombo.SelectedItem = [string]$uiState['mail_scope']
}
if ($uiState.Contains('window_state')) {
    $form.WindowState = [System.Enum]::Parse([System.Windows.Forms.FormWindowState], [string]$uiState['window_state'])
}
else {
    $form.WindowState = 'Normal'
}
$script:mainSplitDistance = if ($uiState.Contains('main_splitter_distance')) { [int]$uiState['main_splitter_distance'] } else { 520 }
$script:detailSplitDistance = if ($uiState.Contains('detail_splitter_distance')) { [int]$uiState['detail_splitter_distance'] } else { 300 }
$script:mailSplitDistance = 180
$script:mailBottomSplitDistance = 430

$searchBox.Add_TextChanged({ Refresh-CaseGrid -SelectedCaseId (Get-SelectedCaseIdFromGrid) })
$reloadButton.Add_Click({
        Save-CurrentCase -Quiet
        Load-ShinsaData
        Refresh-CaseGrid -SelectedCaseId (Get-SelectedCaseIdFromGrid)
        Set-StatusText -Text 'Reloaded local JSON.'
    })
$syncButton.Add_Click({
        try {
            Invoke-GuiSync
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null
        }
    })
$saveButton.Add_Click({ Save-CurrentCase })
$writebackButton.Add_Click({
        try {
            Invoke-GuiWriteback
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null
        }
    })
$openLedgerButton.Add_Click({
        try {
            $targetPath = if ($script:currentCase) { ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $script:currentCase -Name 'ledger_path') } else { $script:Paths.SharePointLedgerPath }
            Start-ShinsaItem -Path $targetPath
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null
        }
    })
$openCaseFolderButton.Add_Click({
        try {
            if ($null -eq $script:currentCase) { return }
            $caseId = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $script:currentCase -Name 'case_id')
            $folderPath = Join-Path $script:Paths.SharePointCaseRoot $caseId
            Start-ShinsaItem -Path $folderPath
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null
        }
    })

$caseGrid.Add_SelectionChanged({
        if ($script:suppressSelection -or $caseGrid.SelectedRows.Count -eq 0) {
            return
        }

        $newCaseId = [string]$caseGrid.SelectedRows[0].Cells['case_id'].Value
        if ($script:currentCase -and (Get-ShinsaRecordValue -Record $script:currentCase -Name 'case_id') -ne $newCaseId) {
            Save-CurrentCase -Quiet
        }
        Select-Case -CaseId $newCaseId -SkipGridSelection
    })

$mailScopeCombo.Add_SelectedIndexChanged({ Refresh-MailList })
$mailFilterBox.Add_TextChanged({ Refresh-MailList })
$mailList.Add_SelectedIndexChanged({ Load-MailPreview -MailRecord (Get-SelectedMailRecord) })

$linkMailButton.Add_Click({
        try {
            if ($null -eq $script:currentCase) { return }
            $mail = Get-SelectedMailRecord
            if ($null -eq $mail) { return }
            Set-MailLink -MailId (Get-ShinsaRecordValue -Record $mail -Name 'mail_id') -CaseId (Get-ShinsaRecordValue -Record $script:currentCase -Name 'case_id')
            Load-ShinsaData
            Select-Case -CaseId (Get-ShinsaRecordValue -Record $script:currentCase -Name 'case_id') -SkipGridSelection
            Set-StatusText -Text 'Mail linked manually.'
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null
        }
    })
$clearMailLinkButton.Add_Click({
        try {
            $mail = Get-SelectedMailRecord
            if ($null -eq $mail) { return }
            Clear-MailLink -MailId (Get-ShinsaRecordValue -Record $mail -Name 'mail_id')
            Load-ShinsaData
            if ($script:currentCase) {
                Select-Case -CaseId (Get-ShinsaRecordValue -Record $script:currentCase -Name 'case_id') -SkipGridSelection
            }
            Set-StatusText -Text 'Mail link cleared.'
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null
        }
    })
$openMsgButton.Add_Click({
        try {
            $mail = Get-SelectedMailRecord
            if ($null -eq $mail) { return }
            Start-ShinsaItem -Path (Get-ShinsaRecordValue -Record $mail -Name 'msg_path')
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null
        }
    })
$openBodyButton.Add_Click({
        try {
            $mail = Get-SelectedMailRecord
            if ($null -eq $mail) { return }
            Start-ShinsaItem -Path (Get-ShinsaRecordValue -Record $mail -Name 'body_path')
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null
        }
    })
$openAttachmentButton.Add_Click({
        try {
            Start-ShinsaItem -Path (Get-SelectedAttachmentPath)
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null
        }
    })
$openFileButton.Add_Click({
        try {
            Start-ShinsaItem -Path (Get-SelectedFilePath)
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null
        }
    })
$openFolderButton.Add_Click({
        try {
            if ($fileList.SelectedItems.Count -gt 0) {
                Start-ShinsaItem -Path ([System.IO.Path]::GetDirectoryName([string]$fileList.SelectedItems[0].Tag))
            }
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null
        }
    })
$form.Add_Shown({
        Set-SafeSplitterLayout -SplitControl $mainSplit -Panel1MinSize 380 -Panel2MinSize 420 -PreferredDistance $script:mainSplitDistance
        Set-SafeSplitterLayout -SplitControl $detailSplit -Panel1MinSize 260 -Panel2MinSize 240 -PreferredDistance $script:detailSplitDistance
        Set-SafeSplitterLayout -SplitControl $mailSplit -Panel1MinSize 140 -Panel2MinSize 120 -PreferredDistance $script:mailSplitDistance
        Set-SafeSplitterLayout -SplitControl $mailBottomSplit -Panel1MinSize 280 -Panel2MinSize 180 -PreferredDistance $script:mailBottomSplitDistance
        Refresh-CaseGrid
        Set-StatusText -Text ("Cases {0} / Mails {1} / Files {2}" -f $script:ledgerRecords.Count, $script:mailRecords.Count, $script:folderRecords.Count)
    })
$form.Add_FormClosing({
        Save-CurrentCase -Quiet
        Save-UiState
    })

Refresh-CaseGrid
Set-StatusText -Text ("Cases {0} / Mails {1} / Files {2}" -f $script:ledgerRecords.Count, $script:mailRecords.Count, $script:folderRecords.Count)
[void]$form.ShowDialog()
