Import-Module (Join-Path $PSScriptRoot '..\scripts\Common.psm1') -Force -DisableNameChecking
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Ensure the WinForms window appears even when launched from a hidden process
Add-Type -Name NativeMethods -Namespace Win32 -MemberDefinition @'
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
'@

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

# --- State ---
$script:ledgerRecords = @()
$script:mailRecords = @()
$script:folderRecords = @()
$script:cacheState = $null
$script:currentCase = $null
$script:displayedMailRecords = @()
$script:displayedFolderRecords = @()
$script:detailControls = @{}
$script:suppressSelection = $false

# --- Helpers ---

function Get-FieldLabelText {
    param([string]$FieldName)
    switch ($FieldName) {
        'case_id'             { 'Case ID' }
        'receipt_no'          { 'Receipt No' }
        'organization_name'   { 'Organization' }
        'contact_name'        { 'Contact' }
        'contact_email'       { 'Email' }
        'status'              { 'Status' }
        'assigned_to'         { 'Assigned To' }
        'missing_documents'   { 'Missing Docs' }
        'review_note_public'  { 'Public Note' }
        default               { ($FieldName -replace '_', ' ') }
    }
}

function Get-WindowStateName {
    param([System.Windows.Forms.FormWindowState]$State)
    switch ($State) {
        'Maximized' { 'Maximized' }
        'Minimized' { 'Minimized' }
        default      { 'Normal' }
    }
}

function Get-UiState {
    ConvertTo-ShinsaMap -InputObject $script:cacheState.ui_state
}

function Save-UiState {
    $uiState = Get-UiState
    $bounds = if ($form.WindowState -eq 'Normal') { $form.Bounds } else { $form.RestoreBounds }

    $uiState['search_text']             = $searchBox.Text
    $uiState['mail_filter']             = $mailFilterBox.Text
    $uiState['mail_scope']              = [string]$mailScopeCombo.SelectedItem
    $uiState['last_case_id']            = if ($script:currentCase) { ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $script:currentCase -Name 'case_id') } else { '' }
    $uiState['window_left']             = $bounds.Left
    $uiState['window_top']              = $bounds.Top
    $uiState['window_width']            = $bounds.Width
    $uiState['window_height']           = $bounds.Height
    $uiState['window_state']            = Get-WindowStateName -State $form.WindowState
    $uiState['main_splitter_distance']  = $mainSplit.SplitterDistance
    $uiState['mail_splitter_distance']  = $mailSplit.SplitterDistance
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
    $script:ledgerRecords  = @((Read-ShinsaJson -Path $script:Paths.LedgerJsonPath))
    $script:mailRecords    = @((Read-ShinsaJson -Path $script:Paths.MailsJsonPath))
    $script:folderRecords  = @((Read-ShinsaJson -Path $script:Paths.FoldersJsonPath))
    $script:cacheState     = Read-ShinsaCache -Paths $script:Paths
}

function Save-CurrentCase {
    param([switch]$Quiet)
    if ($null -eq $script:currentCase) { return }

    foreach ($fieldName in @($script:Config.ledger.editable_columns)) {
        if (-not $script:detailControls.ContainsKey($fieldName)) { continue }
        Set-ShinsaRecordValue -Record $script:currentCase -Name $fieldName -Value $script:detailControls[$fieldName].Text.Trim()
    }
    Write-ShinsaJson -Path $script:Paths.LedgerJsonPath -Data $script:ledgerRecords

    if (-not $Quiet) {
        Set-StatusText -Text ("Saved: {0}" -f (Get-ShinsaRecordValue -Record $script:currentCase -Name 'case_id'))
    }
}

function Get-SelectedCaseIdFromGrid {
    if ($caseGrid.SelectedRows.Count -eq 0) { return '' }
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
        [Parameter(Mandatory)][System.Windows.Forms.SplitContainer]$SplitControl,
        [int]$Panel1MinSize,
        [int]$Panel2MinSize,
        [int]$PreferredDistance
    )
    $SplitControl.Panel1MinSize = [Math]::Max(0, $Panel1MinSize)
    $availableSize = if ($SplitControl.Orientation -eq [System.Windows.Forms.Orientation]::Vertical) {
        $SplitControl.ClientSize.Width
    } else {
        $SplitControl.ClientSize.Height
    }
    if ($availableSize -le 0) { return }
    $safePanel2Min = [Math]::Max(0, [Math]::Min($Panel2MinSize, [Math]::Max(0, $availableSize - $SplitControl.Panel1MinSize - 4)))
    $SplitControl.Panel2MinSize = $safePanel2Min
    $maxDistance  = [Math]::Max($SplitControl.Panel1MinSize, $availableSize - $SplitControl.Panel2MinSize - 4)
    $safeDistance = [Math]::Max($SplitControl.Panel1MinSize, [Math]::Min($PreferredDistance, $maxDistance))
    $SplitControl.SplitterDistance = $safeDistance
}

function Build-CaseTable {
    param([string]$FilterText)
    $table = New-Object System.Data.DataTable
    foreach ($col in @($script:Config.ledger.display_columns)) {
        [void]$table.Columns.Add($col)
    }
    foreach ($record in @($script:ledgerRecords)) {
        $text = (@($script:Config.ledger.display_columns | ForEach-Object { ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $record -Name $_) }) -join ' ')
        if (-not [string]::IsNullOrWhiteSpace($FilterText) -and $text -notmatch [regex]::Escape($FilterText)) { continue }
        $row = $table.NewRow()
        foreach ($col in @($script:Config.ledger.display_columns)) {
            $row[$col] = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $record -Name $col)
        }
        [void]$table.Rows.Add($row)
    }
    return ,$table
}

function Update-CaseHeader {
    if ($null -eq $script:currentCase) {
        $caseHeaderTitle.Text = '(No case selected)'
        $caseHeaderInfo.Text  = ''
        return
    }
    $id  = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $script:currentCase -Name 'case_id')
    $org = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $script:currentCase -Name 'organization_name')
    $caseHeaderTitle.Text = if ([string]::IsNullOrWhiteSpace($org)) { $id } else { "$id  -  $org" }

    $parts = @()
    $contact  = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $script:currentCase -Name 'contact_name')
    $email    = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $script:currentCase -Name 'contact_email')
    $status   = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $script:currentCase -Name 'status')
    $assigned = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $script:currentCase -Name 'assigned_to')
    if (-not [string]::IsNullOrWhiteSpace($contact)) { $parts += $contact }
    if (-not [string]::IsNullOrWhiteSpace($email))   { $parts += "<$email>" }
    if (-not [string]::IsNullOrWhiteSpace($status))   { $parts += "Status: $status" }
    if (-not [string]::IsNullOrWhiteSpace($assigned)) { $parts += "Assigned: $assigned" }
    $caseHeaderInfo.Text = $parts -join '  |  '
}

function Refresh-CaseGrid {
    param([string]$SelectedCaseId = '')
    $script:suppressSelection = $true
    $caseGrid.DataSource = Build-CaseTable -FilterText $searchBox.Text

    if ([string]::IsNullOrWhiteSpace($SelectedCaseId)) {
        $uiState = Get-UiState
        if ($uiState.Contains('last_case_id')) { $SelectedCaseId = [string]$uiState['last_case_id'] }
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
        Update-CaseHeader
        foreach ($fn in $script:detailControls.Keys) { $script:detailControls[$fn].Text = '' }
        $mailList.Items.Clear()
        $attachmentList.Items.Clear()
        $mailPreviewBox.Text = ''
        $fileList.Items.Clear()
    }
}

function Get-MailLinkStateLabel {
    param($MailRecord)
    $mailId  = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $MailRecord -Name 'mail_id')
    $linkMap = Get-ShinsaMailLinkMap -Cache $script:cacheState
    if ($linkMap.ContainsKey($mailId)) {
        return ('manual:{0}' -f (Get-ShinsaRecordValue -Record $linkMap[$mailId] -Name 'case_id'))
    }
    if ($null -ne $script:currentCase) {
        $ce = (ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $script:currentCase -Name 'contact_email')).ToLowerInvariant()
        $se = (ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $MailRecord -Name 'sender_email')).ToLowerInvariant()
        if (-not [string]::IsNullOrWhiteSpace($ce) -and $ce -eq $se) { return 'email' }
    }
    ''
}

function Load-MailPreview {
    param($MailRecord)
    $attachmentList.Items.Clear()
    $mailPreviewBox.Text = ''
    if ($null -eq $MailRecord) { return }

    $bodyPath = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $MailRecord -Name 'body_path')
    if (-not [string]::IsNullOrWhiteSpace($bodyPath) -and (Test-Path $bodyPath)) {
        $mailPreviewBox.Text = Get-Content -Path $bodyPath -Raw -Encoding UTF8
    }
    foreach ($ap in @((Get-ShinsaRecordValue -Record $MailRecord -Name 'attachment_paths'))) {
        $attachmentList.Items.Add([pscustomobject]@{
            name = [System.IO.Path]::GetFileName([string]$ap)
            path = [string]$ap
        }) | Out-Null
    }
}

function Get-SelectedMailRecord {
    if ($mailList.SelectedItems.Count -eq 0) { return $null }
    $mailId = [string]$mailList.SelectedItems[0].Tag
    $script:displayedMailRecords | Where-Object { (Get-ShinsaRecordValue -Record $_ -Name 'mail_id') -eq $mailId } | Select-Object -First 1
}

function Refresh-MailList {
    $mailList.Items.Clear()
    $script:displayedMailRecords = @()
    if ($null -eq $script:currentCase) { return }

    $scope = [string]$mailScopeCombo.SelectedItem
    if ($scope -eq 'All') {
        $records = @($script:mailRecords)
    }
    else {
        $records = @(Get-ShinsaRelatedMails -CaseRecord $script:currentCase -Mails $script:mailRecords -Cache $script:cacheState)
    }

    $ft = $mailFilterBox.Text
    if (-not [string]::IsNullOrWhiteSpace($ft)) {
        $records = @($records | Where-Object {
            $c = @(
                Get-ShinsaRecordValue -Record $_ -Name 'mail_id'
                Get-ShinsaRecordValue -Record $_ -Name 'sender_email'
                Get-ShinsaRecordValue -Record $_ -Name 'sender_name'
                Get-ShinsaRecordValue -Record $_ -Name 'subject'
            ) -join ' '
            $c -match [regex]::Escape($ft)
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
    if ($mailList.Items.Count -gt 0) { $mailList.Items[0].Selected = $true }
}

function Refresh-FileList {
    $fileList.Items.Clear()
    $script:displayedFolderRecords = @()
    if ($null -eq $script:currentCase) { return }

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
        [Parameter(Mandatory)][string]$CaseId,
        [switch]$SkipGridSelection
    )
    $record = Get-CaseRecordById -CaseId $CaseId
    if ($null -eq $record) { return }

    $script:currentCase = $record
    Update-CaseHeader
    foreach ($fn in $script:detailControls.Keys) {
        $script:detailControls[$fn].Text = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $record -Name $fn)
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
    param([Parameter(Mandatory)][string]$MailId, [Parameter(Mandatory)][string]$CaseId)
    $updated = @()
    foreach ($link in @($script:cacheState.mail_links)) {
        if ((Get-ShinsaRecordValue -Record $link -Name 'mail_id') -eq $MailId) { continue }
        $updated += $link
    }
    $updated += [pscustomobject]@{ mail_id = $MailId; case_id = $CaseId; mode = 'manual'; updated_at = (Get-Date).ToString('o') }
    $script:cacheState.mail_links = $updated
    Save-ShinsaCache -Paths $script:Paths -Cache $script:cacheState
}

function Clear-MailLink {
    param([Parameter(Mandatory)][string]$MailId)
    $script:cacheState.mail_links = @($script:cacheState.mail_links | Where-Object { (Get-ShinsaRecordValue -Record $_ -Name 'mail_id') -ne $MailId })
    Save-ShinsaCache -Paths $script:Paths -Cache $script:cacheState
}

function Get-SelectedAttachmentPath {
    if ($null -eq $attachmentList.SelectedItem) { return '' }
    [string]$attachmentList.SelectedItem.path
}

function Get-SelectedFilePath {
    if ($fileList.SelectedItems.Count -eq 0) { return '' }
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
    $summary = @('Write back the following local ledger edits?', '')
    foreach ($change in @($plan.changes)) {
        $summary += ('{0}: {1}' -f $change.case_id, (@($change.changes.PSObject.Properties.Name) -join ', '))
    }
    $summary += ''
    $summary += ('Total: {0} cases / {1} fields' -f $plan.case_count, $plan.change_count)
    $result = [System.Windows.Forms.MessageBox]::Show(($summary -join [Environment]::NewLine), 'shinsa', [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) { return }
    Invoke-ShinsaLedgerWriteback -Config $script:Config -Paths $script:Paths -Plan $plan
    Write-ShinsaJson -Path $script:Paths.LedgerJsonPath -Data @(Import-ShinsaLedgerRecords -Config $script:Config -Paths $script:Paths | Sort-Object case_id)
    Load-ShinsaData
    Refresh-CaseGrid -SelectedCaseId (Get-SelectedCaseIdFromGrid)
    Set-StatusText -Text 'Writeback completed.'
}

# --- UI helpers ---

function New-ToolbarSeparator {
    $sep = New-Object System.Windows.Forms.Label
    $sep.Width = 2
    $sep.Height = 26
    $sep.BorderStyle = 'Fixed3D'
    $sep.AutoSize = $false
    $sep.Margin = New-Object System.Windows.Forms.Padding(6, 4, 6, 0)
    $sep
}

function New-ToolbarButton {
    param([string]$Text, [string]$Tip = '')
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $Text
    $btn.AutoSize = $true
    $btn.Padding = New-Object System.Windows.Forms.Padding(6, 1, 6, 1)
    $btn.Margin = New-Object System.Windows.Forms.Padding(2)
    $btn.FlatStyle = 'Standard'
    if ($Tip) { $toolTip.SetToolTip($btn, $Tip) }
    $btn
}

# ============================================================
#  UI Construction
# ============================================================

$toolTip = New-Object System.Windows.Forms.ToolTip

$form = New-Object System.Windows.Forms.Form
$form.Text = $script:Config.gui.title
$form.MinimumSize = New-Object System.Drawing.Size(1100, 700)
$form.Font = New-Object System.Drawing.Font($script:Config.gui.font_name, [single]$script:Config.gui.font_size)
$form.AutoScaleMode = 'Dpi'
$form.KeyPreview = $true

# --- Root layout: toolbar | main | status ---
$root = New-Object System.Windows.Forms.TableLayoutPanel
$root.Dock = 'Fill'
$root.RowCount = 3
$root.ColumnCount = 1
[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))

# === Toolbar (left group | right group) ===
$toolbar = New-Object System.Windows.Forms.TableLayoutPanel
$toolbar.Dock = 'Fill'
$toolbar.AutoSize = $true
$toolbar.ColumnCount = 2
$toolbar.RowCount = 1
[void]$toolbar.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$toolbar.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$toolbar.Padding = New-Object System.Windows.Forms.Padding(4, 4, 4, 2)

$toolbarLeft = New-Object System.Windows.Forms.FlowLayoutPanel
$toolbarLeft.Dock = 'Fill'
$toolbarLeft.WrapContents = $false
$toolbarLeft.AutoSize = $true

$searchLabel = New-Object System.Windows.Forms.Label
$searchLabel.Text = 'Search'
$searchLabel.AutoSize = $true
$searchLabel.Margin = New-Object System.Windows.Forms.Padding(4, 7, 2, 0)

$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Width = 200
$searchBox.Margin = New-Object System.Windows.Forms.Padding(2, 4, 8, 0)

$syncButton   = New-ToolbarButton -Text 'Sync'   -Tip 'Rebuild JSON from sources (F5)'
$reloadButton = New-ToolbarButton -Text 'Reload'  -Tip 'Reload local JSON'

$toolbarLeft.Controls.AddRange(@($searchLabel, $searchBox, (New-ToolbarSeparator), $syncButton, $reloadButton))

$toolbarRight = New-Object System.Windows.Forms.FlowLayoutPanel
$toolbarRight.Dock = 'Fill'
$toolbarRight.WrapContents = $false
$toolbarRight.AutoSize = $true
$toolbarRight.FlowDirection = 'RightToLeft'

$saveButton           = New-ToolbarButton -Text 'Save'        -Tip 'Save local edits (Ctrl+S)'
$writebackButton      = New-ToolbarButton -Text 'Writeback'   -Tip 'Write edits back to source ledger'
$openLedgerButton     = New-ToolbarButton -Text 'Open Ledger' -Tip 'Open the source ledger file'
$openCaseFolderButton = New-ToolbarButton -Text 'Open Folder' -Tip 'Open the case folder in Explorer'

# RightToLeft: add in reverse visual order
$toolbarRight.Controls.AddRange(@($openCaseFolderButton, $openLedgerButton, (New-ToolbarSeparator), $writebackButton, $saveButton))

$toolbar.Controls.Add($toolbarLeft,  0, 0)
$toolbar.Controls.Add($toolbarRight, 1, 0)

# === Main split: case list | right panel ===
$mainSplit = New-Object System.Windows.Forms.SplitContainer
$mainSplit.Dock = 'Fill'
$mainSplit.Orientation = 'Vertical'
$mainSplit.BorderStyle = 'Fixed3D'

# -- Case grid (left) --
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
$caseGrid.BorderStyle = 'None'
$caseGrid.GridColor = [System.Drawing.Color]::FromArgb(210, 210, 210)
$caseGrid.EnableHeadersVisualStyles = $false
$caseGrid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.SystemColors]::Control
$caseGrid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.SystemColors]::ControlText
$caseGrid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font($script:Config.gui.font_name, [single]$script:Config.gui.font_size, [System.Drawing.FontStyle]::Bold)
$caseGrid.ColumnHeadersBorderStyle = 'Raised'
$caseGrid.CellBorderStyle = 'SingleHorizontal'
$caseGrid.DefaultCellStyle.SelectionBackColor = [System.Drawing.SystemColors]::Highlight
$caseGrid.DefaultCellStyle.SelectionForeColor = [System.Drawing.SystemColors]::HighlightText
$caseGrid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(248, 248, 244)
$mainSplit.Panel1.Controls.Add($caseGrid)

# -- Right panel (header + tabs) --
$rightPanel = New-Object System.Windows.Forms.TableLayoutPanel
$rightPanel.Dock = 'Fill'
$rightPanel.RowCount = 2
$rightPanel.ColumnCount = 1
[void]$rightPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$rightPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))

# Case header
$caseHeaderPanel = New-Object System.Windows.Forms.Panel
$caseHeaderPanel.Dock = 'Fill'
$caseHeaderPanel.BackColor = [System.Drawing.SystemColors]::Window
$caseHeaderPanel.BorderStyle = 'Fixed3D'
$caseHeaderPanel.Padding = New-Object System.Windows.Forms.Padding(10, 6, 10, 4)
$caseHeaderPanel.Height = 52
$caseHeaderPanel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 2)

$caseHeaderTitle = New-Object System.Windows.Forms.Label
$caseHeaderTitle.Dock = 'Top'
$caseHeaderTitle.AutoSize = $true
$caseHeaderTitle.Font = New-Object System.Drawing.Font($script:Config.gui.font_name, 11, [System.Drawing.FontStyle]::Bold)
$caseHeaderTitle.Text = '(No case selected)'
$caseHeaderTitle.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 2)

$caseHeaderInfo = New-Object System.Windows.Forms.Label
$caseHeaderInfo.Dock = 'Top'
$caseHeaderInfo.AutoSize = $true
$caseHeaderInfo.ForeColor = [System.Drawing.SystemColors]::GrayText
$caseHeaderInfo.Text = ''

# Dock=Top: add bottom-first so title renders above info
$caseHeaderPanel.Controls.Add($caseHeaderInfo)
$caseHeaderPanel.Controls.Add($caseHeaderTitle)

$rightPanel.Controls.Add($caseHeaderPanel, 0, 0)

# === Tabs ===
$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = 'Fill'

# ---- Detail tab ----
$detailTab = New-Object System.Windows.Forms.TabPage
$detailTab.Text = 'Detail'
$detailTab.Padding = New-Object System.Windows.Forms.Padding(8)

$detailPanel = New-Object System.Windows.Forms.TableLayoutPanel
$detailPanel.Dock = 'Fill'
$detailPanel.AutoScroll = $true
$detailPanel.ColumnCount = 4
[void]$detailPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 110)))
[void]$detailPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
[void]$detailPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 100)))
[void]$detailPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$detailPanel.RowCount = 12
for ($i = 0; $i -lt 12; $i++) {
    [void]$detailPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
}

# Fields that take full width (span across both column pairs)
$fullWidthFields = @('contact_email', 'missing_documents', 'review_note_public')
$multilineFields = @{ 'missing_documents' = 56; 'review_note_public' = 96 }

$detailRow = 0
$detailCol = 0
foreach ($fieldName in @($script:Config.ledger.detail_columns)) {
    $isFullWidth = $fieldName -in $fullWidthFields
    $isMultiline = $multilineFields.ContainsKey($fieldName)

    if ($isFullWidth -and $detailCol -ne 0) {
        $detailRow++
        $detailCol = 0
    }

    $label = New-Object System.Windows.Forms.Label
    $label.Text = Get-FieldLabelText -FieldName $fieldName
    $label.AutoSize = $true
    $label.Margin = New-Object System.Windows.Forms.Padding(4, 8, 4, 4)
    $label.ForeColor = [System.Drawing.SystemColors]::ControlDarkDark

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Dock = 'Fill'
    $textBox.ReadOnly = $fieldName -notin @($script:Config.ledger.editable_columns)
    if ($textBox.ReadOnly) { $textBox.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 240) }
    $textBox.Margin = New-Object System.Windows.Forms.Padding(2, 4, 8, 4)

    if ($isMultiline) {
        $textBox.Multiline = $true
        $textBox.ScrollBars = 'Vertical'
        $textBox.Height = $multilineFields[$fieldName]
    }

    if ($isFullWidth) {
        $detailPanel.Controls.Add($label, 0, $detailRow)
        $detailPanel.SetColumnSpan($textBox, 3)
        $detailPanel.Controls.Add($textBox, 1, $detailRow)
        $detailRow++
        $detailCol = 0
    }
    else {
        $detailPanel.Controls.Add($label, $detailCol, $detailRow)
        $detailPanel.Controls.Add($textBox, $detailCol + 1, $detailRow)
        $detailCol += 2
        if ($detailCol -ge 4) {
            $detailRow++
            $detailCol = 0
        }
    }

    $script:detailControls[$fieldName] = $textBox
}

$detailTab.Controls.Add($detailPanel)

# ---- Mails tab ----
$mailsTab = New-Object System.Windows.Forms.TabPage
$mailsTab.Text = 'Mails'

$mailLayout = New-Object System.Windows.Forms.TableLayoutPanel
$mailLayout.Dock = 'Fill'
$mailLayout.RowCount = 2
$mailLayout.ColumnCount = 1
[void]$mailLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$mailLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))

$mailToolbar = New-Object System.Windows.Forms.FlowLayoutPanel
$mailToolbar.Dock = 'Fill'
$mailToolbar.WrapContents = $false
$mailToolbar.AutoSize = $true
$mailToolbar.Padding = New-Object System.Windows.Forms.Padding(2)

$mailScopeLabel = New-Object System.Windows.Forms.Label
$mailScopeLabel.Text = 'Scope'
$mailScopeLabel.AutoSize = $true
$mailScopeLabel.Margin = New-Object System.Windows.Forms.Padding(4, 7, 2, 0)

$mailScopeCombo = New-Object System.Windows.Forms.ComboBox
$mailScopeCombo.Width = 90
$mailScopeCombo.DropDownStyle = 'DropDownList'
[void]$mailScopeCombo.Items.Add('Related')
[void]$mailScopeCombo.Items.Add('All')
$mailScopeCombo.SelectedIndex = 0

$mailFilterLabel = New-Object System.Windows.Forms.Label
$mailFilterLabel.Text = 'Filter'
$mailFilterLabel.AutoSize = $true
$mailFilterLabel.Margin = New-Object System.Windows.Forms.Padding(8, 7, 2, 0)

$mailFilterBox = New-Object System.Windows.Forms.TextBox
$mailFilterBox.Width = 160

$linkMailButton      = New-ToolbarButton -Text 'Link'       -Tip 'Link selected mail to current case'
$clearMailLinkButton = New-ToolbarButton -Text 'Unlink'     -Tip 'Remove manual link'
$openMsgButton       = New-ToolbarButton -Text 'MSG'        -Tip 'Open .msg file'
$openBodyButton      = New-ToolbarButton -Text 'Body'       -Tip 'Open body text file'
$openAttachmentButton = New-ToolbarButton -Text 'Attachment' -Tip 'Open selected attachment'

$mailToolbar.Controls.AddRange(@(
    $mailScopeLabel, $mailScopeCombo,
    $mailFilterLabel, $mailFilterBox,
    (New-ToolbarSeparator),
    $linkMailButton, $clearMailLinkButton,
    (New-ToolbarSeparator),
    $openMsgButton, $openBodyButton, $openAttachmentButton
))

# Mail split: list (top) | preview+attachments (bottom)
$mailSplit = New-Object System.Windows.Forms.SplitContainer
$mailSplit.Dock = 'Fill'
$mailSplit.Orientation = 'Horizontal'

$mailList = New-Object System.Windows.Forms.ListView
$mailList.Dock = 'Fill'
$mailList.View = 'Details'
$mailList.FullRowSelect = $true
$mailList.HideSelection = $false
$mailList.BorderStyle = 'Fixed3D'
[void]$mailList.Columns.Add('Received', 140)
[void]$mailList.Columns.Add('Sender',   180)
[void]$mailList.Columns.Add('Subject',  300)
[void]$mailList.Columns.Add('Link',     100)

# Bottom: preview (left 65%) | attachments (right 35%)
$mailBottomPanel = New-Object System.Windows.Forms.TableLayoutPanel
$mailBottomPanel.Dock = 'Fill'
$mailBottomPanel.ColumnCount = 2
$mailBottomPanel.RowCount = 1
[void]$mailBottomPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 65)))
[void]$mailBottomPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 35)))

$mailPreviewBox = New-Object System.Windows.Forms.TextBox
$mailPreviewBox.Dock = 'Fill'
$mailPreviewBox.Multiline = $true
$mailPreviewBox.ScrollBars = 'Vertical'
$mailPreviewBox.ReadOnly = $true
$mailPreviewBox.BackColor = [System.Drawing.SystemColors]::Window

$attachmentGroup = New-Object System.Windows.Forms.GroupBox
$attachmentGroup.Text = 'Attachments'
$attachmentGroup.Dock = 'Fill'

$attachmentList = New-Object System.Windows.Forms.ListBox
$attachmentList.Dock = 'Fill'
$attachmentList.DisplayMember = 'name'
$attachmentGroup.Controls.Add($attachmentList)

$mailBottomPanel.Controls.Add($mailPreviewBox, 0, 0)
$mailBottomPanel.Controls.Add($attachmentGroup, 1, 0)

$mailSplit.Panel1.Controls.Add($mailList)
$mailSplit.Panel2.Controls.Add($mailBottomPanel)

$mailLayout.Controls.Add($mailToolbar, 0, 0)
$mailLayout.Controls.Add($mailSplit, 0, 1)
$mailsTab.Controls.Add($mailLayout)

# ---- Files tab ----
$filesTab = New-Object System.Windows.Forms.TabPage
$filesTab.Text = 'Files'

$fileLayout = New-Object System.Windows.Forms.TableLayoutPanel
$fileLayout.Dock = 'Fill'
$fileLayout.RowCount = 2
$fileLayout.ColumnCount = 1
[void]$fileLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$fileLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))

$fileToolbar = New-Object System.Windows.Forms.FlowLayoutPanel
$fileToolbar.Dock = 'Fill'
$fileToolbar.WrapContents = $false
$fileToolbar.AutoSize = $true
$fileToolbar.Padding = New-Object System.Windows.Forms.Padding(2)

$openFileButton   = New-ToolbarButton -Text 'Open File'   -Tip 'Open selected file'
$openFolderButton = New-ToolbarButton -Text 'Open Folder' -Tip 'Open containing folder'
$fileToolbar.Controls.AddRange(@($openFileButton, $openFolderButton))

$fileList = New-Object System.Windows.Forms.ListView
$fileList.Dock = 'Fill'
$fileList.View = 'Details'
$fileList.FullRowSelect = $true
$fileList.HideSelection = $false
$fileList.BorderStyle = 'Fixed3D'
[void]$fileList.Columns.Add('Path',     380)
[void]$fileList.Columns.Add('Modified', 160)
[void]$fileList.Columns.Add('Size',     90)

$fileLayout.Controls.Add($fileToolbar, 0, 0)
$fileLayout.Controls.Add($fileList, 0, 1)
$filesTab.Controls.Add($fileLayout)

# Assemble tabs
[void]$tabs.TabPages.Add($detailTab)
[void]$tabs.TabPages.Add($mailsTab)
[void]$tabs.TabPages.Add($filesTab)

$rightPanel.Controls.Add($tabs, 0, 1)
$mainSplit.Panel2.Controls.Add($rightPanel)

# === Status bar ===
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Spring = $true
$statusLabel.TextAlign = 'MiddleLeft'
[void]$statusStrip.Items.Add($statusLabel)

# === Assemble root ===
$root.Controls.Add($toolbar,     0, 0)
$root.Controls.Add($mainSplit,   0, 1)
$root.Controls.Add($statusStrip, 0, 2)
$form.Controls.Add($root)

# ============================================================
#  Init & Events
# ============================================================

Load-ShinsaData
Apply-WindowState

$uiState = Get-UiState
if ($uiState.Contains('search_text')) { $searchBox.Text = [string]$uiState['search_text'] }
if ($uiState.Contains('mail_filter')) { $mailFilterBox.Text = [string]$uiState['mail_filter'] }
if ($uiState.Contains('mail_scope')) {
    $scopeVal = [string]$uiState['mail_scope']
    if ($scopeVal -eq 'Related mails') { $scopeVal = 'Related' }
    elseif ($scopeVal -eq 'All mails') { $scopeVal = 'All' }
    if ($mailScopeCombo.Items.Contains($scopeVal)) { $mailScopeCombo.SelectedItem = $scopeVal }
}
if ($uiState.Contains('window_state')) {
    $form.WindowState = [System.Enum]::Parse([System.Windows.Forms.FormWindowState], [string]$uiState['window_state'])
}

$script:mainSplitDistance = if ($uiState.Contains('main_splitter_distance')) { [int]$uiState['main_splitter_distance'] } else { 460 }
$script:mailSplitDistance = if ($uiState.Contains('mail_splitter_distance')) { [int]$uiState['mail_splitter_distance'] } else { 200 }

# --- Toolbar events ---
$searchBox.Add_TextChanged({ Refresh-CaseGrid -SelectedCaseId (Get-SelectedCaseIdFromGrid) })

$reloadButton.Add_Click({
    Save-CurrentCase -Quiet
    Load-ShinsaData
    Refresh-CaseGrid -SelectedCaseId (Get-SelectedCaseIdFromGrid)
    Set-StatusText -Text 'Reloaded.'
})

$syncButton.Add_Click({
    try { Invoke-GuiSync }
    catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null }
})

$saveButton.Add_Click({ Save-CurrentCase })

$writebackButton.Add_Click({
    try { Invoke-GuiWriteback }
    catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null }
})

$openLedgerButton.Add_Click({
    try {
        $p = if ($script:currentCase) { ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $script:currentCase -Name 'ledger_path') } else { $script:Paths.SharePointLedgerPath }
        Start-ShinsaItem -Path $p
    }
    catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null }
})

$openCaseFolderButton.Add_Click({
    try {
        if ($null -eq $script:currentCase) { return }
        $cid = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $script:currentCase -Name 'case_id')
        Start-ShinsaItem -Path (Join-Path $script:Paths.SharePointCaseRoot $cid)
    }
    catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null }
})

# --- Grid events ---
$caseGrid.Add_SelectionChanged({
    if ($script:suppressSelection -or $caseGrid.SelectedRows.Count -eq 0) { return }
    $newId = [string]$caseGrid.SelectedRows[0].Cells['case_id'].Value
    if ($script:currentCase -and (Get-ShinsaRecordValue -Record $script:currentCase -Name 'case_id') -ne $newId) {
        Save-CurrentCase -Quiet
    }
    Select-Case -CaseId $newId -SkipGridSelection
})

# --- Mail events ---
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
        Set-StatusText -Text 'Mail linked.'
    }
    catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null }
})

$clearMailLinkButton.Add_Click({
    try {
        $mail = Get-SelectedMailRecord
        if ($null -eq $mail) { return }
        Clear-MailLink -MailId (Get-ShinsaRecordValue -Record $mail -Name 'mail_id')
        Load-ShinsaData
        if ($script:currentCase) { Select-Case -CaseId (Get-ShinsaRecordValue -Record $script:currentCase -Name 'case_id') -SkipGridSelection }
        Set-StatusText -Text 'Mail link cleared.'
    }
    catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null }
})

$openMsgButton.Add_Click({
    try {
        $mail = Get-SelectedMailRecord
        if ($null -eq $mail) { return }
        Start-ShinsaItem -Path (Get-ShinsaRecordValue -Record $mail -Name 'msg_path')
    }
    catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null }
})

$openBodyButton.Add_Click({
    try {
        $mail = Get-SelectedMailRecord
        if ($null -eq $mail) { return }
        Start-ShinsaItem -Path (Get-ShinsaRecordValue -Record $mail -Name 'body_path')
    }
    catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null }
})

$openAttachmentButton.Add_Click({
    try { Start-ShinsaItem -Path (Get-SelectedAttachmentPath) }
    catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null }
})

# --- File events ---
$openFileButton.Add_Click({
    try { Start-ShinsaItem -Path (Get-SelectedFilePath) }
    catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null }
})

$openFolderButton.Add_Click({
    try {
        if ($fileList.SelectedItems.Count -gt 0) {
            Start-ShinsaItem -Path ([System.IO.Path]::GetDirectoryName([string]$fileList.SelectedItems[0].Tag))
        }
    }
    catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null }
})

# --- Keyboard shortcuts ---
$form.Add_KeyDown({
    param($sender, $e)
    if ($e.Control -and $e.KeyCode -eq 'S') {
        Save-CurrentCase
        $e.Handled = $true
        $e.SuppressKeyPress = $true
    }
    elseif ($e.KeyCode -eq 'F5') {
        try { Invoke-GuiSync }
        catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null }
        $e.Handled = $true
        $e.SuppressKeyPress = $true
    }
    elseif ($e.Control -and $e.KeyCode -eq 'F') {
        $searchBox.Focus()
        $searchBox.SelectAll()
        $e.Handled = $true
        $e.SuppressKeyPress = $true
    }
})

# --- Form lifecycle ---
$form.Add_Shown({
    Set-SafeSplitterLayout -SplitControl $mainSplit -Panel1MinSize 340 -Panel2MinSize 500 -PreferredDistance $script:mainSplitDistance
    Set-SafeSplitterLayout -SplitControl $mailSplit -Panel1MinSize 120 -Panel2MinSize 100 -PreferredDistance $script:mailSplitDistance
    Refresh-CaseGrid
    Set-StatusText -Text ("Cases {0}  |  Mails {1}  |  Files {2}" -f $script:ledgerRecords.Count, $script:mailRecords.Count, $script:folderRecords.Count)
    # Force the form to the foreground when launched from a hidden process
    $form.TopMost = $true
    [void][Win32.NativeMethods]::SetForegroundWindow($form.Handle)
    [void][Win32.NativeMethods]::ShowWindow($form.Handle, 9)  # SW_RESTORE
    $form.TopMost = $false
})

$form.Add_FormClosing({
    Save-CurrentCase -Quiet
    Save-UiState
})

Refresh-CaseGrid
Set-StatusText -Text ("Cases {0}  |  Mails {1}  |  Files {2}" -f $script:ledgerRecords.Count, $script:mailRecords.Count, $script:folderRecords.Count)
[void]$form.ShowDialog()
