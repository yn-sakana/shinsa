Import-Module (Join-Path $PSScriptRoot '..\scripts\Common.psm1') -Force -DisableNameChecking
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Add-Type -Name NativeMethods -Namespace Win32 -MemberDefinition @'
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
'@

$ErrorActionPreference = 'Stop'
$script:AppRoot = Split-Path -Parent $PSScriptRoot
$script:Config = Get-ShinsaConfig -ScriptPath $MyInvocation.MyCommand.Path
$script:Paths = Get-ShinsaDataPaths -Config $script:Config

Ensure-ShinsaState -Paths $script:Paths

# Auto-sync on startup if JSON missing
$needsSync = $false
foreach ($p in @($script:Paths.TableJsonPath, $script:Paths.MailsJsonPath, $script:Paths.FoldersJsonPath)) {
    if (-not (Test-Path $p)) { $needsSync = $true; break }
}
if ($needsSync) {
    try { & (Join-Path $script:AppRoot 'scripts\Sync-Data.ps1') }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Sync failed: $($_.Exception.Message)", 'shinsa') | Out-Null
        exit 1
    }
}

$script:tableRecords = @()
$script:mailRecords = @()
$script:folderRecords = @()
$script:cacheState = $null
$script:currentCase = $null
$script:displayedMailRecords = @()
$script:displayedFolderRecords = @()
$script:detailControls = @{}
$script:suppressSelection = $false
$script:showAllMails = $false

# ── Helpers ──────────────────────────────────────────

function Get-FieldLabelText {
    param([string]$FieldName)
    $words = ($FieldName -replace '_', ' ') -split ' '
    ($words | ForEach-Object { if ($_.Length -gt 0) { $_.Substring(0,1).ToUpper() + $_.Substring(1) } else { $_ } }) -join ' '
}

function Get-UiState { ConvertTo-ShinsaMap -InputObject $script:cacheState.ui_state }

function Save-UiState {
    $uiState = Get-UiState
    $bounds = if ($form.WindowState -eq 'Normal') { $form.Bounds } else { $form.RestoreBounds }
    $uiState['last_case_id']            = if ($script:currentCase) { ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $script:currentCase -Name 'case_id') } else { '' }
    $uiState['window_left']             = $bounds.Left
    $uiState['window_top']              = $bounds.Top
    $uiState['window_width']            = $bounds.Width
    $uiState['window_height']           = $bounds.Height
    $uiState['window_state']            = switch ($form.WindowState) { 'Maximized' { 'Maximized' }; default { 'Normal' } }
    $uiState['main_splitter_distance']  = $mainSplit.SplitterDistance
    $uiState['content_splitter_distance'] = $contentSplit.SplitterDistance
    $uiState['search_text']             = if ($searchBox.Tag -eq 'placeholder') { '' } else { $searchBox.Text }
    $script:cacheState.ui_state = [pscustomobject]$uiState
    Save-ShinsaCache -Paths $script:Paths -Cache $script:cacheState
}

function Set-StatusText { param([string]$Text); $statusLabel.Text = $Text }

function Format-FileSize {
    param([long]$Bytes)
    if ($Bytes -ge 1GB) { return ('{0:N1} GB' -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ('{0:N1} MB' -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ('{0:N1} KB' -f ($Bytes / 1KB)) }
    '{0} B' -f $Bytes
}

function Load-ShinsaData {
    $script:tableRecords = @((Read-ShinsaJson -Path $script:Paths.TableJsonPath))
    $script:mailRecords   = @((Read-ShinsaJson -Path $script:Paths.MailsJsonPath))
    $script:folderRecords = @((Read-ShinsaJson -Path $script:Paths.FoldersJsonPath))
    $script:cacheState    = Read-ShinsaCache -Paths $script:Paths
}

function Save-CurrentCase {
    param([switch]$Quiet)
    if ($null -eq $script:currentCase) { return }
    foreach ($fn in @($script:Config.table.editable_columns)) {
        if (-not $script:detailControls.ContainsKey($fn)) { continue }
        Set-ShinsaRecordValue -Record $script:currentCase -Name $fn -Value $script:detailControls[$fn].Text.Trim()
    }
    Write-ShinsaJson -Path $script:Paths.TableJsonPath -Data $script:tableRecords
    if (-not $Quiet) { Set-StatusText -Text 'Saved.' }
}

function Get-SelectedCaseIdFromGrid {
    if ($caseGrid.SelectedRows.Count -eq 0) { return '' }
    $keyCol = [string]$script:Config.table.key_column
    [string]$caseGrid.SelectedRows[0].Cells[$keyCol].Value
}

function Get-CaseRecordById {
    param([string]$CaseId)
    $script:tableRecords | Where-Object { (Get-ShinsaRecordValue -Record $_ -Name 'case_id') -eq $CaseId } | Select-Object -First 1
}

function Set-SafeSplitterLayout {
    param([Parameter(Mandatory)][System.Windows.Forms.SplitContainer]$Ctl, [int]$Min1, [int]$Min2, [int]$Preferred)
    $Ctl.Panel1MinSize = [Math]::Max(0, $Min1)
    $avail = if ($Ctl.Orientation -eq [System.Windows.Forms.Orientation]::Vertical) { $Ctl.ClientSize.Width } else { $Ctl.ClientSize.Height }
    if ($avail -le 0) { return }
    $Ctl.Panel2MinSize = [Math]::Max(0, [Math]::Min($Min2, [Math]::Max(0, $avail - $Ctl.Panel1MinSize - 4)))
    $max = [Math]::Max($Ctl.Panel1MinSize, $avail - $Ctl.Panel2MinSize - 4)
    $Ctl.SplitterDistance = [Math]::Max($Ctl.Panel1MinSize, [Math]::Min($Preferred, $max))
}

function Build-CaseTable {
    param([string]$FilterText)
    $table = New-Object System.Data.DataTable
    foreach ($col in @($script:Config.table.display_columns)) { [void]$table.Columns.Add($col) }
    foreach ($record in @($script:tableRecords)) {
        if (-not [string]::IsNullOrWhiteSpace($FilterText)) {
            $allText = (@($record.PSObject.Properties | ForEach-Object { ConvertTo-ShinsaString -Value $_.Value }) -join ' ')
            if ($allText -notmatch [regex]::Escape($FilterText)) { continue }
        }
        $row = $table.NewRow()
        foreach ($col in @($script:Config.table.display_columns)) { $row[$col] = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $record -Name $col) }
        [void]$table.Rows.Add($row)
    }
    return ,$table
}

function Update-CaseHeader {
    if ($null -eq $script:currentCase) {
        $caseHeaderTitle.Text = ''
        return
    }
    $keyCol = [string]$script:Config.table.key_column
    $caseHeaderTitle.Text = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $script:currentCase -Name $keyCol)
}

function Refresh-CaseGrid {
    param([string]$SelectedCaseId = '')
    $keyCol = [string]$script:Config.table.key_column
    $script:suppressSelection = $true
    $filterText = if ($searchBox.Tag -eq 'placeholder') { '' } else { $searchBox.Text }
    $caseGrid.DataSource = Build-CaseTable -FilterText $filterText
    if ([string]::IsNullOrWhiteSpace($SelectedCaseId)) {
        $u = Get-UiState; if ($u.Contains('last_case_id')) { $SelectedCaseId = [string]$u['last_case_id'] }
    }
    if (-not [string]::IsNullOrWhiteSpace($SelectedCaseId)) {
        foreach ($row in $caseGrid.Rows) {
            if ([string]$row.Cells[$keyCol].Value -eq $SelectedCaseId) {
                $row.Selected = $true; $caseGrid.CurrentCell = $row.Cells[$keyCol]; break
            }
        }
    }
    if ($caseGrid.SelectedRows.Count -eq 0 -and $caseGrid.Rows.Count -gt 0) {
        $caseGrid.Rows[0].Selected = $true; $caseGrid.CurrentCell = $caseGrid.Rows[0].Cells[$keyCol]
    }
    $script:suppressSelection = $false
    if ($caseGrid.SelectedRows.Count -gt 0) {
        Select-Case -CaseId ([string]$caseGrid.SelectedRows[0].Cells[$keyCol].Value) -SkipGridSelection
    } else {
        $script:currentCase = $null; Update-CaseHeader
        foreach ($fn in $script:detailControls.Keys) { $script:detailControls[$fn].Text = '' }
        $mailList.Items.Clear(); $fileList.Items.Clear()
    }
}

function Get-MailLinkStateLabel {
    param($Rec)
    $mid = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $Rec -Name 'mail_id')
    $map = Get-ShinsaMailLinkMap -Cache $script:cacheState
    if ($map.ContainsKey($mid)) { return 'manual' }
    if ($null -ne $script:currentCase) {
        $ce = (ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $script:currentCase -Name 'contact_email')).ToLowerInvariant()
        $se = (ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $Rec -Name 'sender_email')).ToLowerInvariant()
        if (-not [string]::IsNullOrWhiteSpace($ce) -and $ce -eq $se) { return 'auto' }
    }
    ''
}

function Get-SelectedMailRecord {
    if ($mailList.SelectedItems.Count -eq 0) { return $null }
    $mid = [string]$mailList.SelectedItems[0].Tag
    $script:displayedMailRecords | Where-Object { (Get-ShinsaRecordValue -Record $_ -Name 'mail_id') -eq $mid } | Select-Object -First 1
}

function Refresh-MailList {
    $mailList.Items.Clear()
    $script:displayedMailRecords = @()
    if ($null -eq $script:currentCase) { return }
    if ($script:showAllMails) {
        $records = @($script:mailRecords)
    } else {
        $records = @(Get-ShinsaRelatedMails -CaseRecord $script:currentCase -Mails $script:mailRecords -Cache $script:cacheState)
    }
    foreach ($rec in @($records | Sort-Object received_at, mail_id -Descending)) {
        $mid = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $rec -Name 'mail_id')
        $item = New-Object System.Windows.Forms.ListViewItem((ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $rec -Name 'received_at')))
        [void]$item.SubItems.Add((ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $rec -Name 'sender_email')))
        [void]$item.SubItems.Add((ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $rec -Name 'subject')))
        [void]$item.SubItems.Add((Get-MailLinkStateLabel -Rec $rec))
        $item.Tag = $mid
        $mailList.Items.Add($item) | Out-Null
        $script:displayedMailRecords += $rec
    }
    if ($mailList.Items.Count -gt 0) { $mailList.Items[0].Selected = $true }
}

function Refresh-FileList {
    $fileTree.Nodes.Clear()
    $script:displayedFolderRecords = @()
    if ($null -eq $script:currentCase) { return }
    $cid = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $script:currentCase -Name 'case_id')
    $script:displayedFolderRecords = @($script:folderRecords | Where-Object { (Get-ShinsaRecordValue -Record $_ -Name 'case_id') -eq $cid })

    $folderPath = ''
    if ($script:displayedFolderRecords.Count -gt 0) {
        $folderPath = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $script:displayedFolderRecords[0] -Name 'folder_path')
    }
    $rootNode = New-Object System.Windows.Forms.TreeNode($cid)
    $rootNode.Tag = $folderPath

    foreach ($rec in @($script:displayedFolderRecords)) {
        $relPath = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $rec -Name 'relative_path')
        $fullPath = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $rec -Name 'file_path')
        $segments = $relPath -split '\\'
        $parent = $rootNode
        for ($i = 0; $i -lt $segments.Count; $i++) {
            $seg = $segments[$i]
            $existing = $null
            foreach ($child in $parent.Nodes) { if ($child.Text -eq $seg) { $existing = $child; break } }
            if ($i -eq $segments.Count - 1) {
                $node = New-Object System.Windows.Forms.TreeNode($seg)
                $node.Tag = $fullPath
                [void]$parent.Nodes.Add($node)
            } else {
                if ($null -eq $existing) {
                    $dirNode = New-Object System.Windows.Forms.TreeNode($seg)
                    $partialPath = Join-Path $folderPath ($segments[0..$i] -join '\')
                    $dirNode.Tag = $partialPath
                    [void]$parent.Nodes.Add($dirNode)
                    $parent = $dirNode
                } else {
                    $parent = $existing
                }
            }
        }
    }

    [void]$fileTree.Nodes.Add($rootNode)
    $fileTree.ExpandAll()
}

function Select-Case {
    param([Parameter(Mandatory)][string]$CaseId, [switch]$SkipGridSelection)
    $rec = Get-CaseRecordById -CaseId $CaseId
    if ($null -eq $rec) { return }
    $script:currentCase = $rec
    Update-CaseHeader
    foreach ($fn in $script:detailControls.Keys) {
        $script:detailControls[$fn].Text = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $rec -Name $fn)
    }
    Refresh-MailList; Refresh-FileList
    if (-not $SkipGridSelection) {
        $keyCol = [string]$script:Config.table.key_column
        foreach ($row in $caseGrid.Rows) {
            if ([string]$row.Cells[$keyCol].Value -eq $CaseId) {
                $row.Selected = $true; $caseGrid.CurrentCell = $row.Cells[$keyCol]; break
            }
        }
    }
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
    $plan = Get-ShinsaTableWritebackPlan -Config $script:Config -Paths $script:Paths
    if ($plan.case_count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('No changes to reflect.', 'shinsa') | Out-Null; return
    }
    $lines = @()
    foreach ($c in @($plan.changes)) { $lines += ('{0}: {1}' -f $c.case_id, (@($c.changes.PSObject.Properties.Name) -join ', ')) }
    $msg = "Reflect {0} case(s) / {1} field(s) to the source table?`n`n{2}" -f $plan.case_count, $plan.change_count, ($lines -join "`n")
    $r = [System.Windows.Forms.MessageBox]::Show($msg, 'shinsa', 'OKCancel', 'Question')
    if ($r -ne 'OK') { return }
    Invoke-ShinsaTableWriteback -Config $script:Config -Paths $script:Paths -Plan $plan
    Write-ShinsaJson -Path $script:Paths.TableJsonPath -Data @(Import-ShinsaTableRecords -Config $script:Config -Paths $script:Paths | Sort-Object case_id)
    Load-ShinsaData
    Refresh-CaseGrid -SelectedCaseId (Get-SelectedCaseIdFromGrid)
    Set-StatusText -Text 'Reflected to source.'
}

# ── UI Construction ──────────────────────────────────

$script:M = 8

$form = New-Object System.Windows.Forms.Form
$form.Text = $script:Config.gui.title
$form.MinimumSize = New-Object System.Drawing.Size(900, 600)
$form.Font = New-Object System.Drawing.Font($script:Config.gui.font_name, [single]$script:Config.gui.font_size)
$form.Padding = New-Object System.Windows.Forms.Padding($script:M)
$form.AutoScaleMode = 'Dpi'
$form.KeyPreview = $true

# ── Menu bar ──
$menuBar = New-Object System.Windows.Forms.MenuStrip

$menuFile = New-Object System.Windows.Forms.ToolStripMenuItem('File(&F)')
$menuReflect = New-Object System.Windows.Forms.ToolStripMenuItem('Reflect to source table...')
$menuReflect.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::W
$menuOpenTable = New-Object System.Windows.Forms.ToolStripMenuItem('Open source table')
$menuOpenFolder = New-Object System.Windows.Forms.ToolStripMenuItem('Open case folder')
$menuSync = New-Object System.Windows.Forms.ToolStripMenuItem('Reload from sources')
$menuSync.ShortcutKeys = [System.Windows.Forms.Keys]::F5
$menuQuit = New-Object System.Windows.Forms.ToolStripMenuItem('Quit')
$menuQuit.ShortcutKeys = [System.Windows.Forms.Keys]::Alt -bor [System.Windows.Forms.Keys]::F4
[void]$menuFile.DropDownItems.AddRange(@($menuSync, (New-Object System.Windows.Forms.ToolStripSeparator), $menuReflect, (New-Object System.Windows.Forms.ToolStripSeparator), $menuOpenTable, $menuOpenFolder, (New-Object System.Windows.Forms.ToolStripSeparator), $menuQuit))
[void]$menuBar.Items.Add($menuFile)
$form.MainMenuStrip = $menuBar

# ── Main split: case list | right panel ──
$mainSplit = New-Object System.Windows.Forms.SplitContainer
$mainSplit.Dock = 'Fill'
$mainSplit.Orientation = 'Vertical'

# Left panel: search + case grid
$leftPanel = New-Object System.Windows.Forms.TableLayoutPanel
$leftPanel.Dock = 'Fill'
$leftPanel.RowCount = 2
$leftPanel.ColumnCount = 1
[void]$leftPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$leftPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))

$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Dock = 'Fill'
$searchBox.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, $script:M)
$searchBox.ForeColor = [System.Drawing.SystemColors]::GrayText
$searchBox.Text = 'Search...'
$searchBox.Tag = 'placeholder'

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
$caseGrid.GridColor = [System.Drawing.Color]::FromArgb(220, 220, 220)
$caseGrid.EnableHeadersVisualStyles = $false
$caseGrid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.SystemColors]::Control
$caseGrid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font($form.Font, [System.Drawing.FontStyle]::Bold)
$caseGrid.ColumnHeadersBorderStyle = 'Raised'
$caseGrid.CellBorderStyle = 'SingleHorizontal'
$caseGrid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(248, 248, 245)

$leftPanel.Controls.Add($searchBox, 0, 0)
$leftPanel.Controls.Add($caseGrid, 0, 1)
$mainSplit.Panel1.Controls.Add($leftPanel)

# Right panel: header + tabs
$rightPanel = New-Object System.Windows.Forms.TableLayoutPanel
$rightPanel.Dock = 'Fill'
$rightPanel.RowCount = 2
$rightPanel.ColumnCount = 1
[void]$rightPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$rightPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))

# Case header
$caseHeaderPanel = New-Object System.Windows.Forms.Panel
$caseHeaderPanel.Dock = 'Fill'
$caseHeaderPanel.Padding = New-Object System.Windows.Forms.Padding($script:M, 4, $script:M, 4)
$caseHeaderPanel.Height = 30

$caseHeaderTitle = New-Object System.Windows.Forms.Label
$caseHeaderTitle.Dock = 'Fill'
$caseHeaderTitle.AutoSize = $true
$caseHeaderTitle.Font = New-Object System.Drawing.Font($form.Font, [System.Drawing.FontStyle]::Bold)

$caseHeaderPanel.Controls.Add($caseHeaderTitle)
$rightPanel.Controls.Add($caseHeaderPanel, 0, 0)

# ── Tabs: follows 正本 structure ──
$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = 'Fill'

# ── Tab: 台帳 (Table detail) ──
$tableTab = New-Object System.Windows.Forms.TabPage
$tableTab.Text = 'Table'
$tableTab.Padding = New-Object System.Windows.Forms.Padding(0)

$detailPanel = New-Object System.Windows.Forms.Panel
$detailPanel.Dock = 'Fill'
$detailPanel.AutoScroll = $true
$detailPanel.Padding = New-Object System.Windows.Forms.Padding(0, $script:M, 0, 0)

$detailCols = @($script:Config.table.detail_columns)
$editableCols = @($script:Config.table.editable_columns)
$multilineCols = @(); if ($script:Config.table.PSObject.Properties.Name -contains 'multiline_columns') { $multilineCols = @($script:Config.table.multiline_columns) }

# Measure label width with Graphics (IndexGUI style)
$gfx = $detailPanel.CreateGraphics()
$lblW = 80
foreach ($fn in $detailCols) {
    $w = [int][math]::Ceiling($gfx.MeasureString((Get-FieldLabelText -FieldName $fn), $form.Font).Width) + 8
    if ($w -gt $lblW) { $lblW = $w }
}
if ($lblW -gt 160) { $lblW = 160 }
$gfx.Dispose()

$txtX = $lblW + $script:M
$y = $script:M
$tip = New-Object System.Windows.Forms.ToolTip
$maxLblChars = 20

foreach ($fn in $detailCols) {
    $displayName = Get-FieldLabelText -FieldName $fn
    if ($displayName.Length -gt $maxLblChars) {
        $head = $displayName.Substring(0, [math]::Floor($maxLblChars / 2))
        $tail = $displayName.Substring($displayName.Length - [math]::Floor($maxLblChars / 2))
        $displayName = "$head...$tail"
    }

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $displayName
    $lbl.Location = New-Object System.Drawing.Point($script:M, ($y + 3))
    $lbl.Size = New-Object System.Drawing.Size($lblW, 22)
    $lbl.TextAlign = 'MiddleRight'
    if ($displayName -ne (Get-FieldLabelText -FieldName $fn)) { $tip.SetToolTip($lbl, (Get-FieldLabelText -FieldName $fn)) }
    $detailPanel.Controls.Add($lbl)

    $isMultiline = $fn -in $multilineCols
    $rowH = if ($isMultiline) { 48 } else { 22 }

    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Location = New-Object System.Drawing.Point($txtX, $y)
    $tb.Size = New-Object System.Drawing.Size(($detailPanel.ClientSize.Width - $txtX - $script:M), $rowH)
    $tb.Anchor = 'Top,Left,Right'
    $tb.ReadOnly = $fn -notin $editableCols
    if ($tb.ReadOnly) { $tb.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245) }
    if ($isMultiline) { $tb.Multiline = $true; $tb.ScrollBars = 'Vertical' }
    $detailPanel.Controls.Add($tb)

    $script:detailControls[$fn] = $tb
    $y += $rowH + 6
}
$tableTab.Controls.Add($detailPanel)

# ── Tab: メール (Mails) ──
$mailTab = New-Object System.Windows.Forms.TabPage
$mailTab.Text = 'Mail'

$mailSplit = New-Object System.Windows.Forms.SplitContainer
$mailSplit.Dock = 'Fill'
$mailSplit.Orientation = 'Horizontal'

$mailList = New-Object System.Windows.Forms.ListView
$mailList.Dock = 'Fill'
$mailList.View = 'Details'
$mailList.FullRowSelect = $true
$mailList.HideSelection = $false
[void]$mailList.Columns.Add('Date', 130)
[void]$mailList.Columns.Add('From', 170)
[void]$mailList.Columns.Add('Subject', 280)
[void]$mailList.Columns.Add('Link', 60)

# Mail detail: body + attachments
$mailDetailPanel = New-Object System.Windows.Forms.TableLayoutPanel
$mailDetailPanel.Dock = 'Fill'
$mailDetailPanel.RowCount = 2
$mailDetailPanel.ColumnCount = 1
[void]$mailDetailPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$mailDetailPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))

$mailPreviewBox = New-Object System.Windows.Forms.TextBox
$mailPreviewBox.Dock = 'Fill'
$mailPreviewBox.Multiline = $true
$mailPreviewBox.ScrollBars = 'Vertical'
$mailPreviewBox.ReadOnly = $true
$mailPreviewBox.BackColor = [System.Drawing.SystemColors]::Window

$mailAttachList = New-Object System.Windows.Forms.ListView
$mailAttachList.Dock = 'Fill'
$mailAttachList.View = 'List'
$mailAttachList.Height = 60
$mailAttachList.HideSelection = $false

$mailDetailPanel.Controls.Add($mailPreviewBox, 0, 0)
$mailDetailPanel.Controls.Add($mailAttachList, 0, 1)

$mailSplit.Panel1.Controls.Add($mailList)
$mailSplit.Panel2.Controls.Add($mailDetailPanel)
$mailTab.Controls.Add($mailSplit)

# ── Tab: 案件フォルダ (Case folder) ──
$folderTab = New-Object System.Windows.Forms.TabPage
$folderTab.Text = 'Folder'

$fileTree = New-Object System.Windows.Forms.TreeView
$fileTree.Dock = 'Fill'
$fileTree.HideSelection = $false
$fileTree.ShowLines = $true
$fileTree.ShowPlusMinus = $false
$fileTree.ShowRootLines = $true
$fileTree.PathSeparator = '\'
$folderTab.Controls.Add($fileTree)

[void]$tabs.TabPages.Add($tableTab)
[void]$tabs.TabPages.Add($mailTab)
[void]$tabs.TabPages.Add($folderTab)
$rightPanel.Controls.Add($tabs, 0, 1)
$mainSplit.Panel2.Controls.Add($rightPanel)

# ── Status bar ──
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Spring = $true
$statusLabel.TextAlign = 'MiddleLeft'
[void]$statusStrip.Items.Add($statusLabel)

# Assemble
$form.Controls.Add($mainSplit)
$form.Controls.Add($menuBar)
$form.Controls.Add($statusStrip)

# ── Context menus ──

# Mail context menu
$mailCtx = New-Object System.Windows.Forms.ContextMenuStrip
$mailCtxOpenMsg  = $mailCtx.Items.Add('Open MSG')
$mailCtxOpenBody = $mailCtx.Items.Add('Open body')
[void]$mailCtx.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
$mailCtxLink   = $mailCtx.Items.Add('Link to this case')
$mailCtxUnlink = $mailCtx.Items.Add('Remove link')
[void]$mailCtx.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
$mailCtxShowAll = $mailCtx.Items.Add('Show all mails')
$mailList.ContextMenuStrip = $mailCtx

# File context menu
$fileCtx = New-Object System.Windows.Forms.ContextMenuStrip
$fileCtxOpen       = $fileCtx.Items.Add('Open')
$fileCtxOpenFolder = $fileCtx.Items.Add('Open containing folder')
$fileTree.ContextMenuStrip = $fileCtx

# ── Init ──
Load-ShinsaData

$uiState = Get-UiState
if ($uiState.Contains('window_width') -and $uiState.Contains('window_height')) {
    $form.StartPosition = 'Manual'
    $form.Size = New-Object System.Drawing.Size([int]$uiState.window_width, [int]$uiState.window_height)
    if ($uiState.Contains('window_left') -and $uiState.Contains('window_top')) {
        $form.Location = New-Object System.Drawing.Point([int]$uiState.window_left, [int]$uiState.window_top)
    }
} else {
    $form.StartPosition = 'CenterScreen'
    $form.Size = New-Object System.Drawing.Size([int]$script:Config.gui.window_width, [int]$script:Config.gui.window_height)
}
if ($uiState.Contains('window_state') -and [string]$uiState['window_state'] -eq 'Maximized') { $form.WindowState = 'Maximized' }
if ($uiState.Contains('search_text') -and -not [string]::IsNullOrWhiteSpace([string]$uiState['search_text'])) {
    $searchBox.Text = [string]$uiState['search_text']; $searchBox.ForeColor = [System.Drawing.SystemColors]::WindowText; $searchBox.Tag = $null
}

$script:mainSplitDist    = if ($uiState.Contains('main_splitter_distance'))    { [int]$uiState['main_splitter_distance'] }    else { 300 }
$script:contentSplitDist = if ($uiState.Contains('content_splitter_distance')) { [int]$uiState['content_splitter_distance'] } else { 200 }
# contentSplit is mailSplit in the Mails tab
$contentSplit = $mailSplit

# ── Events ──

# Search placeholder behavior
$searchBox.Add_GotFocus({
    if ($searchBox.Tag -eq 'placeholder') {
        $searchBox.Text = ''; $searchBox.ForeColor = [System.Drawing.SystemColors]::WindowText; $searchBox.Tag = $null
    }
})
$searchBox.Add_LostFocus({
    if ([string]::IsNullOrWhiteSpace($searchBox.Text)) {
        $searchBox.Tag = 'placeholder'
        $searchBox.ForeColor = [System.Drawing.SystemColors]::GrayText
        $searchBox.Text = 'Search...'
    }
})
$searchBox.Add_TextChanged({
    if ($searchBox.Tag -ne 'placeholder') {
        Refresh-CaseGrid -SelectedCaseId (Get-SelectedCaseIdFromGrid)
    }
})

# Case grid
$caseGrid.Add_SelectionChanged({
    if ($script:suppressSelection -or $caseGrid.SelectedRows.Count -eq 0) { return }
    $keyCol = [string]$script:Config.table.key_column
    $newId = [string]$caseGrid.SelectedRows[0].Cells[$keyCol].Value
    if ($script:currentCase -and (Get-ShinsaRecordValue -Record $script:currentCase -Name 'case_id') -ne $newId) { Save-CurrentCase -Quiet }
    Select-Case -CaseId $newId -SkipGridSelection
})

# Auto-save on tab change (leaving台帳 tab)
$tabs.Add_SelectedIndexChanged({ Save-CurrentCase -Quiet })

# Mail list selection → preview + attachments
$mailList.Add_SelectedIndexChanged({
    $mailPreviewBox.Text = ''
    $mailAttachList.Items.Clear()
    $mail = Get-SelectedMailRecord
    if ($null -eq $mail) { return }
    $bp = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $mail -Name 'body_path')
    if (-not [string]::IsNullOrWhiteSpace($bp) -and (Test-Path $bp)) {
        $mailPreviewBox.Text = Get-Content -Path $bp -Raw -Encoding UTF8
    }
    $aps = @((Get-ShinsaRecordValue -Record $mail -Name 'attachment_paths'))
    foreach ($ap in $aps) {
        $apStr = [string]$ap
        if ([string]::IsNullOrWhiteSpace($apStr)) { continue }
        $item = New-Object System.Windows.Forms.ListViewItem([System.IO.Path]::GetFileName($apStr))
        $item.Tag = $apStr
        [void]$mailAttachList.Items.Add($item)
    }
})

# Attachment double-click → open
$mailAttachList.Add_DoubleClick({
    try {
        if ($mailAttachList.SelectedItems.Count -eq 0) { return }
        Start-ShinsaItem -Path ([string]$mailAttachList.SelectedItems[0].Tag)
    } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null }
})

# Mail double-click → open MSG or body
$mailList.Add_DoubleClick({
    try {
        $mail = Get-SelectedMailRecord; if ($null -eq $mail) { return }
        $msg = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $mail -Name 'msg_path')
        if (-not [string]::IsNullOrWhiteSpace($msg) -and (Test-Path $msg)) { Start-ShinsaItem -Path $msg; return }
        $bp = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $mail -Name 'body_path')
        if (-not [string]::IsNullOrWhiteSpace($bp) -and (Test-Path $bp)) { Start-ShinsaItem -Path $bp }
    } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null }
})

# Suppress all expand/collapse
$fileTree.Add_BeforeCollapse({ $_.Cancel = $true })

# File tree double-click → open
$fileTree.Add_NodeMouseDoubleClick({
    param($s, $e)
    try {
        $path = [string]$e.Node.Tag
        if (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path $path)) { Start-ShinsaItem -Path $path }
    } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null }
})

# Mail context menu actions
$mailCtx.Add_Opening({
    $mailCtxShowAll.Text = if ($script:showAllMails) { 'Show related only' } else { 'Show all mails' }
    # Build attachment sub-items
    while ($mailCtx.Items.Count -gt 7) { $mailCtx.Items.RemoveAt(7) }
    $mail = Get-SelectedMailRecord
    if ($null -ne $mail) {
        $aps = @((Get-ShinsaRecordValue -Record $mail -Name 'attachment_paths'))
        if ($aps.Count -gt 0) {
            [void]$mailCtx.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
            foreach ($ap in $aps) {
                $mi = $mailCtx.Items.Add("Open: $([System.IO.Path]::GetFileName([string]$ap))")
                $mi.Tag = [string]$ap
                $mi.Add_Click({ try { Start-ShinsaItem -Path $this.Tag } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null } })
            }
        }
    }
})
$mailCtxOpenMsg.Add_Click({
    try { $m = Get-SelectedMailRecord; if ($m) { Start-ShinsaItem -Path (Get-ShinsaRecordValue -Record $m -Name 'msg_path') } }
    catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null }
})
$mailCtxOpenBody.Add_Click({
    try { $m = Get-SelectedMailRecord; if ($m) { Start-ShinsaItem -Path (Get-ShinsaRecordValue -Record $m -Name 'body_path') } }
    catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null }
})
$mailCtxLink.Add_Click({
    try {
        if ($null -eq $script:currentCase) { return }
        $m = Get-SelectedMailRecord; if ($null -eq $m) { return }
        $ml = @(); foreach ($l in @($script:cacheState.mail_links)) { if ((Get-ShinsaRecordValue -Record $l -Name 'mail_id') -ne (Get-ShinsaRecordValue -Record $m -Name 'mail_id')) { $ml += $l } }
        $ml += [pscustomobject]@{ mail_id = (Get-ShinsaRecordValue -Record $m -Name 'mail_id'); case_id = (Get-ShinsaRecordValue -Record $script:currentCase -Name 'case_id'); mode = 'manual'; updated_at = (Get-Date).ToString('o') }
        $script:cacheState.mail_links = $ml; Save-ShinsaCache -Paths $script:Paths -Cache $script:cacheState
        Load-ShinsaData; Select-Case -CaseId (Get-ShinsaRecordValue -Record $script:currentCase -Name 'case_id') -SkipGridSelection
        Set-StatusText -Text 'Mail linked.'
    } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null }
})
$mailCtxUnlink.Add_Click({
    try {
        $m = Get-SelectedMailRecord; if ($null -eq $m) { return }
        $mid = Get-ShinsaRecordValue -Record $m -Name 'mail_id'
        $script:cacheState.mail_links = @($script:cacheState.mail_links | Where-Object { (Get-ShinsaRecordValue -Record $_ -Name 'mail_id') -ne $mid })
        Save-ShinsaCache -Paths $script:Paths -Cache $script:cacheState
        Load-ShinsaData; if ($script:currentCase) { Select-Case -CaseId (Get-ShinsaRecordValue -Record $script:currentCase -Name 'case_id') -SkipGridSelection }
        Set-StatusText -Text 'Link removed.'
    } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null }
})
$mailCtxShowAll.Add_Click({
    $script:showAllMails = -not $script:showAllMails; Refresh-MailList
})

# File context menu actions
$fileCtxOpen.Add_Click({
    try {
        $node = $fileTree.SelectedNode; if ($null -eq $node) { return }
        $path = [string]$node.Tag
        if (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path $path)) { Start-ShinsaItem -Path $path }
    } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null }
})
$fileCtxOpenFolder.Add_Click({
    try {
        $node = $fileTree.SelectedNode; if ($null -eq $node) { return }
        $path = [string]$node.Tag
        if (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path $path)) {
            if (Test-Path $path -PathType Container) { Start-ShinsaItem -Path $path }
            else { Start-ShinsaItem -Path ([System.IO.Path]::GetDirectoryName($path)) }
        }
    } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null }
})

# Menu bar actions
$menuSync.Add_Click({
    try { Invoke-GuiSync } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null }
})
$menuReflect.Add_Click({
    try { Invoke-GuiWriteback } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null }
})
$menuOpenTable.Add_Click({
    try {
        $p = if ($script:currentCase) { ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $script:currentCase -Name 'table_path') } else { $script:Paths.SharePointTablePath }
        Start-ShinsaItem -Path $p
    } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null }
})
$menuOpenFolder.Add_Click({
    try {
        if ($null -eq $script:currentCase) { return }
        $cid = ConvertTo-ShinsaString -Value (Get-ShinsaRecordValue -Record $script:currentCase -Name 'case_id')
        Start-ShinsaItem -Path (Join-Path $script:Paths.SharePointCaseRoot $cid)
    } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'shinsa') | Out-Null }
})
$menuQuit.Add_Click({ $form.Close() })

# Keyboard
$form.Add_KeyDown({
    param($s, $e)
    if ($e.Control -and $e.KeyCode -eq 'F') { $searchBox.Focus(); if ($searchBox.Tag -ne 'placeholder') { $searchBox.SelectAll() }; $e.Handled = $true; $e.SuppressKeyPress = $true }
})

# Form lifecycle
$form.Add_Shown({
    Set-SafeSplitterLayout -Ctl $mainSplit -Min1 200 -Min2 400 -Preferred $script:mainSplitDist
    Set-SafeSplitterLayout -Ctl $contentSplit -Min1 100 -Min2 60 -Preferred $script:contentSplitDist
    Refresh-CaseGrid
    Set-StatusText -Text ("{0} cases" -f $script:tableRecords.Count)
    $form.TopMost = $true
    [void][Win32.NativeMethods]::SetForegroundWindow($form.Handle)
    [void][Win32.NativeMethods]::ShowWindow($form.Handle, 9)
    $form.TopMost = $false
})
$form.Add_FormClosing({ Save-CurrentCase -Quiet; Save-UiState })

Refresh-CaseGrid
[void]$form.ShowDialog()
